using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace FakeExcelSerializer;

public static class ExcelSerializer
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
    private const string CONTENT_TYPE_XML = "[Content_Types].xml";
    private const string SHEET_XML = "sheet.xml";
    private const string STRINGS_XML = "strings.xml";
    private const string RELS = "_rels";
    private const string BOOK_XML = "book.xml";
    private const string STYLES_XML = "styles.xml";
    private const string BOOK_XML_RELS = "book.xml.rels";
    private const string DOT_RELS = ".rels";

    public static void ToFile<T>(IEnumerable<T> rows, string fileName, ExcelSerializerOptions options)
    {
        if (rows == null || !rows.Any())
            return;

        var workPathRoot = Path.Combine(options.WorkPath, Guid.NewGuid().ToString());
        if (!Directory.Exists(workPathRoot))
            Directory.CreateDirectory(workPathRoot);
        try
        {
            using (var sheetStream = CreateStream(Path.Combine(workPathRoot, SHEET_XML)))
            using (var stringsStream = CreateStream(Path.Combine(workPathRoot, STRINGS_XML)))
            using (var writer = new ExcelSerializerWriter(options))
            {
                CreateSheet(rows, sheetStream, writer, options);
                WriteSharedStrings(stringsStream, writer);
            }

            var workRelPath = Path.Combine(workPathRoot, RELS);
            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            var _stylesBytes = Encoding.UTF8.GetBytes(string.Format(
                _styles,
                options.DateTimeFormat,
                options.DateFormat,
                options.TimeFormat,
                options.IntegerFormat,
                options.NumberFormat
            ));

            WriteStream(_contentTypes, Path.Combine(workPathRoot, CONTENT_TYPE_XML));
            WriteStream(_book, Path.Combine(workPathRoot, BOOK_XML));
            WriteStream(_stylesBytes, Path.Combine(workPathRoot, STYLES_XML));
            WriteStream(_rels, Path.Combine(workRelPath, DOT_RELS));
            WriteStream(_bookRels, Path.Combine(workRelPath, BOOK_XML_RELS));

            if (File.Exists(fileName))
                File.Delete(fileName);

            ZipFile.CreateFromDirectory(workPathRoot, fileName);
        }
        finally
        {
            try
            {
                if (Directory.Exists(workPathRoot))
                    Directory.Delete(workPathRoot, true);
            }
            catch { }
        }
    }

    static void WriteStream(byte[] bytes, string fileName)
    {
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            fs.Write(bytes, 0, bytes.Length);
    }

    static Stream CreateStream(string fileName)
        => new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);

    static void CreateSheet<T>(
        IEnumerable<T> rows,
        Stream stream,
        ExcelSerializerWriter writer,
        ExcelSerializerOptions options
    )
    {
        stream.Write(_sheetStart, 0, _sheetStart.Length);

        if (options.HasHeaderRecord)
            stream.Write(_frozenTitleRow, 0, _frozenTitleRow.Length);

        if (options.AutoFitColumns)
            WriteCellWidth(rows, ref stream, ref writer, options);

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
                    serializer.WriteTitle(ref writer, rows.First(), options);
                }
                writer.WriteRaw(_rowEnd);
                writer.CopyTo(stream);
            }

#if NET5_0_OR_GREATER
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), ref stream, ref writer, serializer, options);
            else if (rows is List<T> list)
                WriteRowsSpan(CollectionsMarshal.AsSpan(list), ref stream, ref writer, serializer, options);
            else
                WriteRows(rows, ref stream, ref writer, serializer, options);
#else
            if (rows is T[] arr)
                WriteRowsSpan(arr.AsSpan(), ref stream, ref writer, serializer, options);
            else
                WriteRows(rows, ref stream, ref writer, serializer, options);
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
        ref Stream stream,
        ref ExcelSerializerWriter writer,
        IExcelSerializer<T> serializer,
        ExcelSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.WriteRaw(_rowStart);
            serializer.Serialize(ref writer, row, options);
            writer.WriteRaw(_rowEnd);
            writer.CopyTo(stream);
        }
    }

    static void WriteRows<T>(
        IEnumerable<T> rows,
        ref Stream stream,
        ref ExcelSerializerWriter writer,
        IExcelSerializer<T> serializer,
        ExcelSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.WriteRaw(_rowStart);
            serializer.Serialize(ref writer, row, options);
            writer.WriteRaw(_rowEnd);
            writer.CopyTo(stream);
        }
    }

    static void WriteCellWidth<T>(
        IEnumerable<T> rows,
        ref Stream stream,
        ref ExcelSerializerWriter writer,
        ExcelSerializerOptions options
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
            serializer.Serialize(ref writer, row, options);
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

    static void WriteSharedStrings(Stream stream, ExcelSerializerWriter writer)
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
