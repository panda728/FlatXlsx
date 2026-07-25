using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace FlatXlsx;

public class ArrayPoolBufferWriter : IBufferWriter<byte>, IDisposable
{
    byte[]? _rentedBuffer;
    int _written;
    long _committed;

    const int MinimumBufferSize = 256;

    public ArrayPoolBufferWriter(int initialCapacity = MinimumBufferSize)
    {
        if (initialCapacity <= 0)
            throw new ArgumentException(SR.InitialCapacityOutOfRange, nameof(initialCapacity));

        _rentedBuffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
        _written = 0;
        _committed = 0;
    }

    public Memory<byte> OutputAsMemory
    {
        get
        {
            CheckIfDisposed();
            return _rentedBuffer.AsMemory(0, _written);
        }
    }

    public Span<byte> OutputAsSpan
    {
        get
        {
            CheckIfDisposed();
            return _rentedBuffer.AsSpan(0, _written);
        }
    }

    public int BytesWritten
    {
        get
        {
            CheckIfDisposed();
            return _written;
        }
    }

    public long BytesCommitted
    {
        get
        {
            CheckIfDisposed();
            return _committed;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Clear()
    {
        CheckIfDisposed();
        ClearHelper();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void ClearHelper()
    {
        // No need to zero the buffer here: only [0.._written) is ever exposed,
        // and the array is cleared when returned to the pool in Dispose.
        _written = 0;
    }

    public async Task CopyToAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        CheckIfDisposed();

        if (stream == null)
        {
            ThrowArgumentNullException(nameof(stream));
            return;
        }
        if (_rentedBuffer == null)
        {
            ThrowArgumentNullException(nameof(_rentedBuffer));
            return;
        }

        await stream.WriteAsync(_rentedBuffer, 0, _written, cancellationToken).ConfigureAwait(false);

        _committed += _written;

        ClearHelper();
    }

    public void CopyTo(Stream stream)
    {
        CheckIfDisposed();

        if (stream == null)
        {
            ThrowArgumentNullException(nameof(stream));
            return;
        }

        if (_rentedBuffer == null)
        {
            ThrowArgumentNullException(nameof(_rentedBuffer));
            return;
        }

        stream.Write(_rentedBuffer, 0, _written);
        _committed += _written;
        ClearHelper();
    }

    public void Advance(int count)
    {
        CheckIfDisposed();

        if (count < 0)
        {
            ThrowArgumentNullException(nameof(count));
            return;
        }

        if (_written > (_rentedBuffer?.Length ?? 0) - count)
        {
            ThrowInvalidOperationException();
            return;
        }

        _written += count;
    }

    public void Dispose()
    {
        if (_rentedBuffer == null)
            return;

        ArrayPool<byte>.Shared.Return(_rentedBuffer, clearArray: true);
        _rentedBuffer = null;
        _written = 0;
        GC.SuppressFinalize(this);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void CheckIfDisposed()
    {
        if (_rentedBuffer == null)
        {
            ThrowObjectDisposedException();
            return;
        }
    }

    public Memory<byte> GetMemory(int sizeHint = 0)
    {
        CheckIfDisposed();

        if (sizeHint < 0)
            ThrowArgumentException(nameof(sizeHint));

        CheckAndResizeBuffer(sizeHint);
        return _rentedBuffer.AsMemory(_written);
    }

    public Span<byte> GetSpan(int sizeHint = 0)
    {
        CheckIfDisposed();

        if (sizeHint < 0)
            ThrowArgumentException(nameof(sizeHint));

        CheckAndResizeBuffer(sizeHint);
        return _rentedBuffer.AsSpan(_written);
    }

    void CheckAndResizeBuffer(int sizeHint)
    {
        if (_rentedBuffer == null)
        {
            _rentedBuffer = ArrayPool<byte>.Shared.Rent(Math.Max(MinimumBufferSize, sizeHint));
            return;
        }

        if (sizeHint == 0)
            sizeHint = MinimumBufferSize;

        int availableSpace = _rentedBuffer.Length - _written;

        if (sizeHint > availableSpace)
        {
            int growBy = sizeHint > _rentedBuffer.Length ? sizeHint : _rentedBuffer.Length;

            int newSize = checked(_rentedBuffer.Length + growBy);

            byte[] oldBuffer = _rentedBuffer;

            _rentedBuffer = ArrayPool<byte>.Shared.Rent(newSize);

            oldBuffer.AsSpan(0, _written).CopyTo(_rentedBuffer);
            ArrayPool<byte>.Shared.Return(oldBuffer, clearArray: true);
        }
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowArgumentNullException(string name)
        => throw new ArgumentNullException(name);

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowObjectDisposedException()
        => throw new ObjectDisposedException(nameof(ArrayPoolBufferWriter));

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowInvalidOperationException()
        => throw new InvalidOperationException(SR.BufferAdvanceTooFar);

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowArgumentException(string name)
        => throw new ArgumentException(SR.SizeHintNegative, name);
}

