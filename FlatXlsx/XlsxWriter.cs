using System.Buffers;
using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Text;

namespace FlatXlsx;

public class XlsxWriter(XlsxSerializerOptions options) : IDisposable
{
    const int XF_WRAP_TEXT = 1;
    const int XF_DATETIME = 2;
    const int XF_DATE = 3;
    const int XF_INT = 5;
    const int XF_NUM = 6;

    const int LEN_DATE = 10;
    const int LEN_DATETIME = 18;

    // Hard limits of the SpreadsheetML format. Exceeding any of them yields a file that Excel
    // reports as corrupt, so they are enforced up front instead of being discovered by the user.
    const int MAX_ROWS = 1_048_576;
    const int MAX_COLUMNS = 16_384;
    const int MAX_CELL_LENGTH = 32_767;

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

    static readonly byte[] _colStartInline = Encoding.UTF8.GetBytes(@"<c t=""inlineStr""><is><t>");
    static readonly byte[] _colStartInlineWrap = Encoding.UTF8.GetBytes(@$"<c t=""inlineStr"" s=""{XF_WRAP_TEXT}""><is><t>");
    static readonly byte[] _colEndInline = Encoding.UTF8.GetBytes("</t></is></c>");

#if NET8_0_OR_GREATER
    static readonly byte[] _colStartDateTime = Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>");
    static readonly byte[] _colStartDate = Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>");
    static readonly byte[] _colStartTime = Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_TIME}""><v>");
    static readonly SearchValues<char> _newlineChars = SearchValues.Create("\r\n");

    /// <summary>Bytes that must never reach an inline string cell verbatim: markup characters
    /// and the C0 controls that XML forbids entirely.</summary>
    static readonly SearchValues<byte> _inlineUnsafeBytes = SearchValues.Create(BuildInlineUnsafeBytes());

    static byte[] BuildInlineUnsafeBytes()
    {
        var bytes = new byte[0x20 + 3];
        for (var i = 0; i < 0x20; i++)
            bytes[i] = (byte)i;
        bytes[0x20] = (byte)'<';
        bytes[0x21] = (byte)'>';
        bytes[0x22] = (byte)'&';
        return bytes;
    }
#endif

    readonly ArrayPoolBufferWriter _writer = new();
    readonly XlsxSerializerOptions _options = options;

    bool _countingCharLength = options.AutoFitColumns;

    int _columnIndex = 0;
    int _currentDepth = 0;
    int _rowCount = 0;
    int _maxColumnCount = 0;

    public void Dispose()
    {
        _writer.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Interned cell strings, written to strings.xml after the sheet.
    /// </summary>
    internal SharedStringTable SharedStrings { get; } = new(options.MaxSharedStrings);

    /// <summary>Bytes currently buffered and not yet copied to the output stream.</summary>
    internal int BufferedBytes => _writer.BytesWritten;

    /// <summary>The buffered (not yet flushed) output as text - the view a custom serializer's
    /// unit test asserts against.</summary>
    public override string ToString() => Encoding.UTF8.GetString(
#if NET5_0_OR_GREATER
        _writer.OutputAsSpan);
#else
        _writer.OutputAsSpan.ToArray());
#endif
    /// <summary>
    /// Tally the maximum number of characters per column. For automatic column width adjustment
    /// </summary>
    internal Dictionary<int, int> ColumnMaxLength { get; } = [];
    internal void StopCountingCharLength() => _countingCharLength = false;

    /// <summary>Number of rows opened with <see cref="BeginRow"/> so far.</summary>
    internal int RowCount => _rowCount;

    /// <summary>Widest row written so far, in cells.</summary>
    internal int MaxColumnCount => _maxColumnCount;

    /// <summary>Starts a new row and enforces the sheet's row limit.</summary>
    internal void BeginRow()
    {
        if (_rowCount == MAX_ROWS)
            ThrowTooManyRows();
        _rowCount++;
        _columnIndex = 0;
    }

    internal void Clear()
    {
        _columnIndex = 0;
        _currentDepth = 0;
        _writer.Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
    internal async Task CopyToAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        await _writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
        Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
    internal void CopyTo(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        _writer.CopyTo(stream);
        Clear();
    }

    /// <summary>Marks the start of a nested value (a member of an object graph, an element of a
    /// collection). Guards against runaway nesting: past <see cref="XlsxSerializerOptions.MaxDepth"/>
    /// it throws, which is what stops a circular reference from writing forever.
    /// Pair every call with <see cref="ExitNested"/>.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EnterNested()
    {
        _currentDepth++;
        if (_currentDepth >= _options.MaxDepth)
            ThrowReachedMaxDepth(_currentDepth);
    }

    /// <summary>Marks the end of the nested value opened by <see cref="EnterNested"/>.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void ExitNested()
    {
        _currentDepth--;
    }

    // Internal: raw bytes bypass escaping, so exposing this would hand every custom serializer
    // a way to corrupt the workbook. The pipeline uses it for the row tags only.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal void WriteRaw(ReadOnlySpan<byte> value) => _writer.Write(value);

    /// <summary>Writes an empty cell.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteEmpty()
    {
        _writer.Write(_emptyColumn);
        SetMaxLength(0);
    }

    /// <summary>Closes the current cell: records its width for auto-fit and enforces the
    /// sheet's column limit. Every cell written must pass through here.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void SetMaxLength(int length)
    {
        if (_countingCharLength)
        {
#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
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
        }

        if (_columnIndex == MAX_COLUMNS)
            ThrowTooManyColumns();
        _columnIndex++;

        if (_columnIndex > _maxColumnCount)
            _maxColumnCount = _columnIndex;
    }

    /// <summary>Write string.</summary>
    public void Write(string? value)
    {
        if (value == null || string.IsNullOrEmpty(value))
        {
            WriteEmpty();
            return;
        }

        if (value.Length > MAX_CELL_LENGTH)
            ThrowCellTooLong(value.Length);

        var wrap = ContainsNewLine(value);

        if (!SharedStrings.TryGetOrAdd(value, out var index))
        {
            // Table full: the value goes out inline, costing file size instead of memory.
            WriteEscapedInline(value, wrap);
            return;
        }

        _writer.Write(wrap ? _colStartStringWrap : _colStartString);
#if NET8_0_OR_GREATER
        WriteUtf8Formatted(index, default);
#else
        WriteUtf8Bytes(index.ToString(CultureInfo.InvariantCulture));
#endif
        _writer.Write(_colEnd);
        SetMaxLength(value.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    static bool ContainsNewLine(string value)
#if NET8_0_OR_GREATER
        => value.AsSpan().ContainsAny(_newlineChars);
#else
        => value.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0;
#endif

    /// <summary>Writes the value as an inline string cell, escaping it as XML text.</summary>
    void WriteEscapedInline(string value, bool wrap)
    {
        _writer.Write(wrap ? _colStartInlineWrap : _colStartInline);
        XmlEscape.WriteEscaped(value, _writer);
        _writer.Write(_colEndInline);
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
    public void Write(bool value)
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

#if NET8_0_OR_GREATER
    /// <summary>Formats the value as UTF-8 directly into the buffer without intermediate string allocation.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    int WriteUtf8Formatted<T>(T value, ReadOnlySpan<char> format) where T : IUtf8SpanFormattable
    {
        int written;
        var span = _writer.GetSpan(48);
        while (!value.TryFormat(span, out written, format, CultureInfo.InvariantCulture))
            span = _writer.GetSpan(span.Length * 2);
        _writer.Advance(written);
        return written;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterFormatted<T>(T value, byte[] colStart) where T : IUtf8SpanFormattable
    {
        _writer.Write(colStart);
        var written = WriteUtf8Formatted(value, default);
        _writer.Write(_colEnd);
        SetMaxLength(written);
    }

    /// <summary>Writes the value as an inline string cell, bypassing the shared-string table.
    /// For values that are unique per row (Guid, timestamps) the table only adds allocations and
    /// dictionary growth without any dedup benefit.</summary>
    /// <remarks>The formatted text is checked for markup and control characters; anything that
    /// would need escaping falls back to the shared-string path, so a custom serializer cannot
    /// inject markup into the sheet through this method.</remarks>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WriteInlineString<T>(T value) where T : IUtf8SpanFormattable
    {
        Span<byte> tmp = stackalloc byte[128];
        if (!value.TryFormat(tmp, out var written, default, CultureInfo.InvariantCulture) || NeedsEscaping(tmp[..written]))
        {
            Write(value.ToString());
            return;
        }
        WriteInlineBytes(tmp[..written]);
    }

    /// <summary>Inline-string variant for culture-formatted values.</summary>
    public void WriteInlineString<T>(T value, IFormatProvider? provider) where T : IUtf8SpanFormattable, IFormattable
    {
        Span<byte> tmp = stackalloc byte[128];
        if (!value.TryFormat(tmp, out var written, default, provider) || NeedsEscaping(tmp[..written]))
        {
            Write(value.ToString(null, provider));
            return;
        }
        WriteInlineBytes(tmp[..written]);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    static bool NeedsEscaping(ReadOnlySpan<byte> utf8) => utf8.ContainsAny(_inlineUnsafeBytes);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriteInlineBytes(ReadOnlySpan<byte> utf8)
    {
        _writer.Write(_colStartInline);
        _writer.Write(utf8);
        _writer.Write(_colEndInline);
        SetMaxLength(utf8.Length);
    }
#endif

#if NET8_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(byte value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(sbyte value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(decimal value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(double value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(float value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(int value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(uint value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(long value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ulong value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(short value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ushort value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(Half value) => WriterFormatted(value, _colStartNumber);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(Int128 value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(UInt128 value) => WriterFormatted(value, _colStartInteger);
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(System.Numerics.BigInteger value) => WriterFormatted(value, _colStartInteger);
#else
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(byte value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(sbyte value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(decimal value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(double value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(float value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(int value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(uint value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(long value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ulong value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(short value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(ushort value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
#if NET5_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(Half value) => WriterNumber(value.ToString(CultureInfo.InvariantCulture).AsSpan());
#endif
#if NET7_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(Int128 value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(UInt128 value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
#endif
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(System.Numerics.BigInteger value) => WriterInteger(value.ToString(CultureInfo.InvariantCulture).AsSpan());
#endif

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(DateTime value)
    {
        var d = value;
        if (d == DateTime.MinValue)
        {
            WriteEmpty();
            return;
        }
        if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
        {
#if NET8_0_OR_GREATER
            _writer.Write(_colStartDate);
            WriteUtf8Formatted(d, "yyyy-MM-ddTHH:mm:ss");
            _writer.Write(_colEnd);
#else
            WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)}</v></c>");
#endif
            SetMaxLength(LEN_DATE);
            return;
        }

#if NET8_0_OR_GREATER
        _writer.Write(_colStartDateTime);
        WriteUtf8Formatted(d, "yyyy-MM-ddTHH:mm:ss");
        _writer.Write(_colEnd);
#else
        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)}</v></c>");
#endif
        SetMaxLength(LEN_DATETIME);
    }

#if NET6_0_OR_GREATER
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(DateOnly value)
    {
#if NET8_0_OR_GREATER
        _writer.Write(_colStartDate);
        WriteUtf8Formatted(value, "yyyy-MM-dd");
        _writer.Write("T00:00:00"u8);
        _writer.Write(_colEnd);
#else
        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}T00:00:00</v></c>");
#endif
        SetMaxLength(LEN_DATE);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Write(TimeOnly value)
    {
#if NET8_0_OR_GREATER
        _writer.Write(_colStartTime);
        _writer.Write("1900-01-01T"u8);
        WriteUtf8Formatted(value, "HH:mm:ss");
        _writer.Write(_colEnd);
#else
        WriteUtf8Bytes(@$"<c t=""d"" s=""{XF_TIME}""><v>1900-01-01T{value.ToString("HH:mm:ss", CultureInfo.InvariantCulture)}</v></c>");
#endif
        SetMaxLength(LEN_TIME);
    }
#endif

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowReachedMaxDepth(int depth)
    {
        throw new InvalidOperationException(SR.MaxDepthReached(depth));
    }

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowTooManyRows()
        => throw new InvalidOperationException(SR.TooManyRows(MAX_ROWS));

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowTooManyColumns()
        => throw new InvalidOperationException(SR.TooManyColumns(MAX_COLUMNS));

#if NETSTANDARD2_1_OR_GREATER || NET5_0_OR_GREATER
    [DoesNotReturn]
#endif
    static void ThrowCellTooLong(int length)
        => throw new InvalidOperationException(SR.CellTooLong(MAX_CELL_LENGTH, length));
}