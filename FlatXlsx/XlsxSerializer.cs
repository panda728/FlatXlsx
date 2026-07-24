using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace FlatXlsx;

public static class XlsxSerializer
{
    readonly static byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
    readonly static byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

    readonly static byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<bookViews><workbookView/></bookViews>
<sheets><sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>");
    readonly static byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

    readonly static string _styles = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<numFmts count=""5"">
<numFmt numFmtId=""1"" formatCode =""{0}"" />
<numFmt numFmtId=""2"" formatCode =""{1}"" />
<numFmt numFmtId=""3"" formatCode =""{2}"" />
<numFmt numFmtId=""4"" formatCode =""{3}"" />
<numFmt numFmtId=""5"" formatCode =""{4}"" />
</numFmts>
<fonts count=""1"">
<font/>
</fonts>
<fills count=""1"">
<fill/>
</fills>
<borders count=""1"">
<border/>
</borders>
<cellStyleXfs count=""1"">
<xf/>
</cellStyleXfs>
<cellXfs count=""7"">
<xf/>
<xf><alignment wrapText=""true""/></xf>
<xf numFmtId=""1""  applyNumberFormat=""1""></xf>
<xf numFmtId=""2""  applyNumberFormat=""1""></xf>
<xf numFmtId=""3""  applyNumberFormat=""1""></xf>
<xf numFmtId=""4""  applyNumberFormat=""1""></xf>
<xf numFmtId=""5""  applyNumberFormat=""1""></xf>
</cellXfs>
</styleSheet>";

    readonly static byte[] _rowStart = Encoding.UTF8.GetBytes("<row>");
    readonly static byte[] _rowEnd = Encoding.UTF8.GetBytes("</row>");
    readonly static byte[] _colStart = Encoding.UTF8.GetBytes("<cols>");
    readonly static byte[] _colEnd = Encoding.UTF8.GetBytes("</cols>");
    readonly static byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

    readonly static byte[] _sheetStart = Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
    readonly static byte[] _sheetEnd = Encoding.UTF8.GetBytes(@"</worksheet>");
    readonly static byte[] _dataStart = Encoding.UTF8.GetBytes(@"<sheetData>");
    readonly static byte[] _dataEnd = Encoding.UTF8.GetBytes(@"</sheetData>");

    readonly static byte[] _autoFilterStart = Encoding.UTF8.GetBytes(@"<autoFilter ref=""");
    readonly static byte[] _autoFilterEnd = Encoding.UTF8.GetBytes(@"""/>");

    readonly static byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
    //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
    readonly static byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
    readonly static byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
    readonly static byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

    const int COLUMN_WIDTH_MARGIN = 2;
    const int FLUSH_THRESHOLD = 32 * 1024;
    private const string CONTENT_TYPE_XML = "[Content_Types].xml";
    private const string SHEET_XML = "sheet.xml";
    private const string STRINGS_XML = "strings.xml";
    private const string RELS = "_rels";
    private const string BOOK_XML = "book.xml";
    private const string STYLES_XML = "styles.xml";
    private const string BOOK_XML_RELS = "book.xml.rels";
    private const string DOT_RELS = ".rels";

    /// <summary>Creates an .xlsx file. The output is streamed; no working folder is used.</summary>
    public static void ToFile<T>(IEnumerable<T> rows, string fileName, XlsxSerializerOptions options)
    {
        if (rows == null || !rows.Any())
            return;

        using var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);
        ToStream(rows, fs, options);
    }

    /// <summary>Writes .xlsx content to the stream. The stream does not need to be seekable
    /// (network streams are fine); it is left open after writing.</summary>
    public static void ToStream<T>(IEnumerable<T> rows, Stream stream, XlsxSerializerOptions options)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (rows == null || !rows.Any())
            return;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        WriteEntry(archive, CONTENT_TYPE_XML, _contentTypes, options.CompressionLevel);
        WriteEntry(archive, RELS + "/" + DOT_RELS, _rels, options.CompressionLevel);
        WriteEntry(archive, BOOK_XML, _book, options.CompressionLevel);
        WriteEntry(archive, RELS + "/" + BOOK_XML_RELS, _bookRels, options.CompressionLevel);
        WriteEntry(archive, STYLES_XML, Encoding.UTF8.GetBytes(string.Format(
            _styles,
            options.DateTimeFormat,
            options.DateFormat,
            options.TimeFormat,
            options.IntegerFormat,
            options.NumberFormat
        )), options.CompressionLevel);

        using var writer = new XlsxWriter(options);
        using (var sheetStream = archive.CreateEntry(SHEET_XML, options.CompressionLevel).Open())
            CreateSheet(rows, sheetStream, writer, options);
        using (var stringsStream = archive.CreateEntry(STRINGS_XML, options.CompressionLevel).Open())
            WriteSharedStrings(stringsStream, writer);
    }

    /// <summary>Writes .xlsx content to an <see cref="IBufferWriter{T}"/> such as
    /// System.IO.Pipelines.PipeWriter (e.g. ASP.NET Core's Response.BodyWriter).
    /// Flushing the underlying pipe is left to the caller.</summary>
    public static void To<T>(IEnumerable<T> rows, IBufferWriter<byte> bufferWriter, XlsxSerializerOptions options)
    {
        if (bufferWriter == null)
            throw new ArgumentNullException(nameof(bufferWriter));

        using var stream = new BufferWriterStream(bufferWriter);
        ToStream(rows, stream, options);
    }

    static void WriteEntry(ZipArchive archive, string entryName, byte[] bytes, System.IO.Compression.CompressionLevel compressionLevel)
    {
        using var s = archive.CreateEntry(entryName, compressionLevel).Open();
        s.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Creates an .xlsx file asynchronously. The output is streamed; no working folder is used.</summary>
    public static async Task ToFileAsync<T>(IEnumerable<T> rows, string fileName, XlsxSerializerOptions options, CancellationToken cancellationToken = default)
    {
        if (rows == null || !rows.Any())
            return;

        using var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
        await ToStreamAsync(rows, fs, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes .xlsx content to the stream asynchronously. The stream does not need to be
    /// seekable (network streams are fine); it is left open after writing.</summary>
    public static async Task ToStreamAsync<T>(IEnumerable<T> rows, Stream stream, XlsxSerializerOptions options, CancellationToken cancellationToken = default)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (rows == null || !rows.Any())
            return;

        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        await WriteEntryAsync(archive, CONTENT_TYPE_XML, _contentTypes, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, RELS + "/" + DOT_RELS, _rels, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, BOOK_XML, _book, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, RELS + "/" + BOOK_XML_RELS, _bookRels, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, STYLES_XML, Encoding.UTF8.GetBytes(string.Format(
            _styles,
            options.DateTimeFormat,
            options.DateFormat,
            options.TimeFormat,
            options.IntegerFormat,
            options.NumberFormat
        )), options.CompressionLevel, cancellationToken).ConfigureAwait(false);

        using var writer = new XlsxWriter(options);
        using (var sheetStream = archive.CreateEntry(SHEET_XML, options.CompressionLevel).Open())
            await CreateSheetAsync(rows, sheetStream, writer, options, cancellationToken).ConfigureAwait(false);
        using (var stringsStream = archive.CreateEntry(STRINGS_XML, options.CompressionLevel).Open())
            await WriteSharedStringsAsync(stringsStream, writer, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes .xlsx content to a <see cref="System.IO.Pipelines.PipeWriter"/>
    /// (e.g. ASP.NET Core's Response.BodyWriter). Data is flushed to the pipe as it is
    /// produced, so backpressure is honored.</summary>
    public static Task ToAsync<T>(IEnumerable<T> rows, System.IO.Pipelines.PipeWriter pipeWriter, XlsxSerializerOptions options, CancellationToken cancellationToken = default)
    {
        if (pipeWriter == null)
            throw new ArgumentNullException(nameof(pipeWriter));

        return ToStreamAsync(rows, pipeWriter.AsStream(leaveOpen: true), options, cancellationToken);
    }

    static async Task WriteEntryAsync(ZipArchive archive, string entryName, byte[] bytes, System.IO.Compression.CompressionLevel compressionLevel, CancellationToken cancellationToken)
    {
        using var s = archive.CreateEntry(entryName, compressionLevel).Open();
        await s.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
    }

    static async Task CreateSheetAsync<T>(
        IEnumerable<T> rows,
        Stream stream,
        XlsxWriter writer,
        XlsxSerializerOptions options,
        CancellationToken cancellationToken
    )
    {
        await stream.WriteAsync(_sheetStart, 0, _sheetStart.Length, cancellationToken).ConfigureAwait(false);

        if (options.HasHeaderRecord)
            await stream.WriteAsync(_frozenTitleRow, 0, _frozenTitleRow.Length, cancellationToken).ConfigureAwait(false);

        if (options.AutoFitColumns)
            await WriteCellWidthAsync(rows, stream, writer, options, cancellationToken).ConfigureAwait(false);

        await stream.WriteAsync(_dataStart, 0, _dataStart.Length, cancellationToken).ConfigureAwait(false);

        var serializer = options.GetSerializer<T>();
        if (serializer != null)
        {
            if (options.HasHeaderRecord)
            {
                writer.WriteRaw(_rowStart);
                if (options.HeaderTitles != null && options.HeaderTitles.Any())
                {
                    foreach (var t in options.HeaderTitles)
                        writer.Write(t);
                }
                else
                {
                    serializer.WriteTitle(writer, rows.First(), options);
                }
                writer.WriteRaw(_rowEnd);
                await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }

            foreach (var row in rows)
            {
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, row, options);
                writer.WriteRaw(_rowEnd);
                if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                    await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }
        }
        writer.WriteRaw(_dataEnd);

        if (options.AutoFilter)
        {
            var colName = options.HeaderTitles != null && options.HeaderTitles.Any()
                ? ToColumnName(options.HeaderTitles.Length)
                : ToColumnName(writer.ColumnMaxLength.Count);

            var range = $"A1:{colName}{rows.Count() + 1}";
            writer.WriteRaw(_autoFilterStart);
            writer.WriteRaw(Encoding.UTF8.GetBytes(range));
            writer.WriteRaw(_autoFilterEnd);
        }

        writer.WriteRaw(_sheetEnd);
        await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    static async Task WriteCellWidthAsync<T>(
        IEnumerable<T> rows,
        Stream stream,
        XlsxWriter writer,
        XlsxSerializerOptions options,
        CancellationToken cancellationToken
    )
    {
        var serializer = options.GetSerializer<T>();
        if (serializer == null) return;
        if (options.HasHeaderRecord && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                writer.Write(t);
            writer.Clear();
        }
        foreach (var row in rows.Take(options.AutoFitDepth))
        {
            serializer.Serialize(writer, row, options);
            writer.Clear();
        }
        writer.StopCountingCharLength();

        var size = 100 * writer.ColumnMaxLength.Count;
        using var buffer = new ArrayPoolBufferWriter(size);
        buffer.Write(_colStart);
        foreach (var pair in writer.ColumnMaxLength)
        {
            var id = pair.Key + 1;
            var width = Math.Min(options.AutoFitWidhtMax, pair.Value + COLUMN_WIDTH_MARGIN);

            WriteUtf8Bytes(@$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />", buffer);
        }
        buffer.Write(_colEnd);
        await buffer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    static async Task WriteSharedStringsAsync(Stream stream, XlsxWriter writer, CancellationToken cancellationToken)
    {
        await stream.WriteAsync(_sstStart, 0, _sstStart.Length, cancellationToken).ConfigureAwait(false);
        using var buffer = new ArrayPoolBufferWriter();
        foreach (var s in writer.SharedStrings.Keys)
        {
            buffer.Write(_siStart);
            WriteUtf8Bytes(SecurityElement.Escape(s), buffer);
            buffer.Write(_siEnd);
            await buffer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
        }
        await stream.WriteAsync(_sstEnd, 0, _sstEnd.Length, cancellationToken).ConfigureAwait(false);
    }

    static void CreateSheet<T>(
        IEnumerable<T> rows,
        Stream stream,
        XlsxWriter writer,
        XlsxSerializerOptions options
    )
    {
        stream.Write(_sheetStart, 0, _sheetStart.Length);

        if (options.HasHeaderRecord)
            stream.Write(_frozenTitleRow, 0, _frozenTitleRow.Length);

        if (options.AutoFitColumns)
            WriteCellWidth(rows, stream, writer, options);

        stream.Write(_dataStart, 0, _dataStart.Length);

        var serializer = options.GetSerializer<T>();
        if (serializer != null)
        {
            if (options.HasHeaderRecord)
            {
                writer.WriteRaw(_rowStart);
                if (options.HeaderTitles != null && options.HeaderTitles.Any())
                {
                    foreach (var t in options.HeaderTitles)
                        writer.Write(t);
                }
                else
                {
                    serializer.WriteTitle(writer, rows.First(), options);
                }
                writer.WriteRaw(_rowEnd);
                writer.CopyTo(stream);
            }

#if NET5_0_OR_GREATER
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), stream, writer, serializer, options);
            else if (rows is List<T> list)
                WriteRowsSpan(CollectionsMarshal.AsSpan(list), stream, writer, serializer, options);
            else
                WriteRows(rows, stream, writer, serializer, options);
#else
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), stream, writer, serializer, options);
            else
                WriteRows(rows, stream, writer, serializer, options);
#endif
        }
        writer.WriteRaw(_dataEnd);

        if (options.AutoFilter)
        {
            var colName = options.HeaderTitles != null && options.HeaderTitles.Any()
                ? ToColumnName(options.HeaderTitles.Length)
                : ToColumnName(writer.ColumnMaxLength.Count);

            var range = $"A1:{colName}{rows.Count() + 1}";
            writer.WriteRaw(_autoFilterStart);
            writer.WriteRaw(Encoding.UTF8.GetBytes(range));
            writer.WriteRaw(_autoFilterEnd);
        }

        writer.WriteRaw(_sheetEnd);
        writer.CopyTo(stream);
    }
    static string ToColumnName(int index)
    {
        if (index < 1) { return ""; }
        var list = new List<char>();
        index--;
        do
        {
            list.Add(Convert.ToChar(index % 26 + 65));
        }
        while ((index = index / 26 - 1) != -1);
        var sb = new StringBuilder();
        for (int i = list.Count - 1; i >= 0; i--)
        {
            sb.Append(list[i]);
        }
        return sb.ToString();
    }

    static void WriteRowsSpan<T>(
        Span<T> rows,
        Stream stream,
        XlsxWriter writer,
        IXlsxSerializer<T> serializer,
        XlsxSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.WriteRaw(_rowStart);
            serializer.Serialize(writer, row, options);
            writer.WriteRaw(_rowEnd);
            if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                writer.CopyTo(stream);
        }
    }

    static void WriteRows<T>(
        IEnumerable<T> rows,
        Stream stream,
        XlsxWriter writer,
        IXlsxSerializer<T> serializer,
        XlsxSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.WriteRaw(_rowStart);
            serializer.Serialize(writer, row, options);
            writer.WriteRaw(_rowEnd);
            if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                writer.CopyTo(stream);
        }
    }

    static void WriteCellWidth<T>(
        IEnumerable<T> rows,
        Stream stream,
        XlsxWriter writer,
        XlsxSerializerOptions options
    )
    {
        // Counting the number of characters in Writer's internal process
        // The result is stored in writer.ColumnMaxLength 
        var serializer = options.GetSerializer<T>();
        if (serializer == null) return;
        if (options.HasHeaderRecord && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                writer.Write(t);
            writer.Clear();
        }
        foreach (var row in rows.Take(options.AutoFitDepth))
        {
            serializer.Serialize(writer, row, options);
            writer.Clear();
        }
        writer.StopCountingCharLength();

        var size = 100 * writer.ColumnMaxLength.Count;
        using var buffer = new ArrayPoolBufferWriter(size);
        buffer.Write(_colStart);
        foreach (var pair in writer.ColumnMaxLength)
        {
            var id = pair.Key + 1;
            var width = Math.Min(options.AutoFitWidhtMax, pair.Value + COLUMN_WIDTH_MARGIN);

            WriteUtf8Bytes(@$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />", buffer);
        }
        buffer.Write(_colEnd);
        buffer.CopyTo(stream);
    }

    static void WriteSharedStrings(Stream stream, XlsxWriter writer)
    {
        stream.Write(_sstStart, 0, _sstStart.Length);
        using var buffer = new ArrayPoolBufferWriter();
        foreach (var s in writer.SharedStrings.Keys)
        {
            buffer.Write(_siStart);
            WriteUtf8Bytes(SecurityElement.Escape(s), buffer);
            buffer.Write(_siEnd);
            buffer.CopyTo(stream);
        }
        stream.Write(_sstEnd, 0, _sstEnd.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    static void WriteUtf8Bytes(string? s, ArrayPoolBufferWriter writer)
    {
        if (s == null)
            return;
#if NET5_0_OR_GREATER
        Encoding.UTF8.GetBytes(s.AsSpan(), writer);
#else
        writer.Write(Encoding.UTF8.GetBytes(s));
#endif
    }
}
