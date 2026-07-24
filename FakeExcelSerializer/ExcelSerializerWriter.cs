using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text;

namespace FakeExcelSerializer;

public class ExcelSerializerWriter : IDisposable
{
    //const int XF_NORMAL = 0;
    const int XF_WRAP_TEXT = 1;
    const int XF_DATETIME = 2;
    const int XF_DATE = 3;
    const int XF_INT = 5;
    const int XF_NUM = 6;

    const int LEN_DATE = 10;
    const int LEN_DATETIME = 18;

#if NET6_0_OR_GREATER
    const int XF_TIME = 4;
    const int LEN_TIME = 8;
#endif

    static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
    static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
    static readonly byte[] _colStartInteger = Encoding.UTF8.GetBytes(@$"<c t=""n"" s=""{XF_INT}""><v>");
    static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@$"<c t=""n"" s=""{XF_NUM}""><v>");
    static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
    static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
    static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");

    readonly ArrayPoolBufferWriter _writer = new();
    readonly ExcelSerializerOptions _options;

    bool _countingCharLength;

    int _columnIndex = 0;
    int _currentDepth = 0;
    int _stringIndex = 0;

    public ExcelSerializerWriter(ExcelSerializerOptions options)
    {
        _options = options;
        _currentDepth = 0;
        _countingCharLength = options.AutoFitColumns;
    }

    public void Dispose()
    {
        _writer.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Maintain a dictionary of strings. Output the same value with the same ID.
    /// </summary>
    public Dictionary<string, int> SharedStrings { get; } = new();
    public ReadOnlySpan<byte> AsSpan() => _writer.OutputAsSpan;
    public ReadOnlyMemory<byte> AsMemory() => _writer.OutputAsMemory;
    public long BytesCommitted() => _writer.BytesCommitted;
    public override string ToString() => Encoding.UTF8.GetString(
#if NET5_0_OR_GREATER
        _writer.OutputAsSpan);
#else
        _writer.OutputAsSpan.ToArray());
#endif
    /// <summary>
    /// Tally the maximum number of characters per column. For automatic column width adjustment
    /// </summary>
    public Dictionary<int, int> ColumnMaxLength { get; } = new();
    public void StopCountingCharLength() => _countingCharLength = false;

    public void Clear()
    {
        _columnIndex = 0;
        _currentDepth = 0;
        _writer.Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
    public async Task CopyToAsync(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        await _writer.CopyToAsync(stream);
        Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
    public void CopyTo(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        _writer.CopyTo(stream);
        Clear();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EnterAndValidate()
    {
        _currentDepth++;
        if (_currentDepth >= _options.MaxDepth)
            ThrowReachedMaxDepth(_currentDepth);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Exit()
    {
        _currentDepth--;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteRaw(ReadOnlySpan<byte> value) => _writer.Write(value);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public int WriteEmpty()
    {
        _columnIndex++;
        _writer.Write(_emptyColumn);
        return 0;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void SetMaxLength(int length)
    {
        if (!_countingCharLength)
            return;

#if NETSTANDARD2_1_OR_GREATER
        if (!ColumnMaxLength.TryAdd(_columnIndex, length))
        {
            if (ColumnMaxLength[_columnIndex] < length)
                ColumnMaxLength[_columnIndex] = length;
        }
#else
        if (ColumnMaxLength.ContainsKey(_columnIndex))
        {
            if (ColumnMaxLength[_columnIndex] < length)
                ColumnMaxLength[_columnIndex] = length;
        }
        else
        {
            ColumnMaxLength.Add(_columnIndex, length);
        }
#endif
        _columnIndex++;
    }

    /// <summary>Write string.</summary>
    public void Write(string? value)
    {
        if (value == null || string.IsNullOrEmpty(value))
        {
            WriteEmpty();
            return;
        }

        _writer.Write(
            value.Contains(Environment.NewLine)
                ? _colStartStringWrap
                : _colStartString
        );

#if NETSTANDARD2_1_OR_GREATER
        var index = SharedStrings.TryAdd(value, _stringIndex)
            ? _stringIndex++
            : SharedStrings[value];
#else
        var index = 0;
        if (SharedStrings.ContainsKey(value))
        {
            index = SharedStrings[value];
        }
        else
        {
            SharedStrings.Add(value, _stringIndex);
            index = _stringIndex++;
        }
#endif
        WriteUtf8Bytes($"{index}");
        _writer.Write(_colEnd);
        SetMaxLength(value.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriteUtf8Bytes(string s)
    {
#if NET5_0_OR_GREATER
        Encoding.UTF8.GetBytes(s.AsSpan(), _writer);
#else
        _writer.Write(Encoding.UTF8.GetBytes(s));
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriteUtf8Bytes(ReadOnlySpan<char> s)
    {
#if NET5_0_OR_GREATER
        Encoding.UTF8.GetBytes(s, _writer);
#else
        _writer.Write(Encoding.UTF8.GetBytes(s.ToArray()));
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(char value) => Write($"{value}");

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(bool value)
    {
        _writer.Write(_colStartBoolean);
        WriteUtf8Bytes(value ? "1" : "0");
        _writer.Write(_colEnd);
        SetMaxLength(1);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterInteger(in ReadOnlySpan<char> chars)
    {
        _writer.Write(_colStartInteger);
        WriteUtf8Bytes(chars);
        _writer.Write(_colEnd);
        SetMaxLength(chars.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterNumber(in ReadOnlySpan<char> chars)
    {
        _writer.Write(_colStartNumber);
        WriteUtf8Bytes(chars);
        _writer.Write(_colEnd);

        SetMaxLength(chars.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(byte value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(sbyte value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(decimal value) => WriterNumber($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(double value) => WriterNumber($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(float value) => WriterNumber($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(int value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(uint value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(long value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(ulong value) => WriterInteger($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(short value) => WriterNumber($"{value}".AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(ushort value) => WriterNumber($"{value}".AsSpan());

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(DateTime value)
    {
        var d = value;
        if (d == DateTime.MinValue) WriteEmpty();
        if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
        {
            WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>");
            SetMaxLength(LEN_DATE);
            return;
        }

        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>");
        SetMaxLength(LEN_DATETIME);
    }

#if NET6_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(DateOnly value)
    {
        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-dd}T00:00:00</v></c>");
        SetMaxLength(LEN_DATE);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteDateTime(TimeOnly value)
    {
        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_TIME}""><v>1900-01-01T{value:HH:mm:ss}</v></c>");
        SetMaxLength(LEN_TIME);
    }
#endif

#if NETSTANDARD2_1_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowReachedMaxDepth(int depth)
    {
        throw new InvalidOperationException($"Serializer detects reached max depth:{depth}. Please check the circular reference.");
    }
}