using System.Buffers;

namespace FlatXlsx;

/// <summary>
/// Lets the zip container's few synchronous tail writes (deflate close, data descriptors, the
/// central directory on older targets) land in a pooled buffer instead of the destination, and
/// forwards them with the next asynchronous write. The destination stream is only ever written
/// asynchronously - which is what an ASP.NET Core response body demands.
/// </summary>
/// <remarks>The buffered bytes are bounded: entry tails and the central directory are a few
/// hundred bytes each, while the row data itself always travels the asynchronous path.
/// The caller flushes the last tail with <see cref="FlushPendingAsync"/> after the container
/// is closed; the destination stream is not owned and is left open.</remarks>
internal sealed class SyncWriteBufferingStream(Stream destination) : Stream
{
    readonly ArrayPoolBufferWriter _pending = new();

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Write(byte[] buffer, int offset, int count)
        => _pending.Write(buffer.AsSpan(offset, count));

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer)
        => _pending.Write(buffer);
#endif

    // A synchronous flush must not touch the destination; the bytes go out with the next
    // asynchronous write instead.
    public override void Flush() { }

    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
    {
        await FlushPendingAsync(cancellationToken).ConfigureAwait(false);
        await destination.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
    {
        await FlushPendingAsync(cancellationToken).ConfigureAwait(false);
        await destination.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
    }
#endif

    public override async Task FlushAsync(CancellationToken cancellationToken)
    {
        await FlushPendingAsync(cancellationToken).ConfigureAwait(false);
        await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Forwards any synchronously buffered bytes to the destination.</summary>
    public Task FlushPendingAsync(CancellationToken cancellationToken)
    {
        if (_pending.BytesWritten == 0)
            return Task.CompletedTask;
        return _pending.CopyToAsync(destination, cancellationToken);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            _pending.Dispose();
        base.Dispose(disposing);
    }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
