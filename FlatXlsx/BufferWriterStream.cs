using System.Buffers;

namespace FlatXlsx;

/// <summary>
/// Write-only <see cref="Stream"/> adapter over an <see cref="IBufferWriter{T}"/>.
/// Lets stream-based producers (e.g. ZipArchive) emit directly into a PipeWriter or other buffer writer.
/// </summary>
internal sealed class BufferWriterStream : Stream
{
    readonly IBufferWriter<byte> _writer;

    public BufferWriterStream(IBufferWriter<byte> writer)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override void Write(byte[] buffer, int offset, int count)
        => _writer.Write(buffer.AsSpan(offset, count));

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer)
        => _writer.Write(buffer);
#endif

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
}
