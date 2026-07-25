using System.Buffers;
using System.IO;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
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

    const string _bookTemplate = @"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<bookViews><workbookView/></bookViews>
<sheets><sheet name=""{0}"" sheetId=""1"" r:id=""rId1""/></sheets>
</workbook>";
    readonly static byte[] _book = Encoding.UTF8.GetBytes(string.Format(_bookTemplate, "Sheet"));

    static readonly char[] _forbiddenSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

    static byte[] BuildBook(XlsxSerializerOptions options)
    {
        var name = options.SheetName;
        if (name == "Sheet")
            return _book;

        if (string.IsNullOrWhiteSpace(name)
            || name.Length > 31
            || name.IndexOfAny(_forbiddenSheetNameChars) >= 0
            || name[0] == '\'' || name[name.Length - 1] == '\''
            || ContainsControlChar(name))
        {
            throw new InvalidOperationException(SR.SheetNameInvalid(name));
        }

        return Encoding.UTF8.GetBytes(string.Format(_bookTemplate, EscapeAttribute(name)));
    }

    /// <summary>Escapes a value for use inside a double-quoted XML attribute. Control
    /// characters cannot be escaped in XML 1.0 at all, so every value that reaches this method
    /// has been validated not to contain them.</summary>
    static string EscapeAttribute(string value)
        => value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    static bool ContainsControlChar(string s)
    {
        foreach (var c in s)
        {
            if (c < 0x20)
                return true;
        }
        return false;
    }
    readonly static byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

    // Excel reserves numFmtId values below 164 for its built-in formats; custom codes start there.
    const int CUSTOM_NUMFMT_BASE = 164;

    /// <summary>Builds styles.xml from the options' sheet-wide formats plus the format codes
    /// the cells actually used. Written after the sheet for the same reason strings.xml is:
    /// its content is only known once the cells have been written.</summary>
    static byte[] BuildStyles(XlsxSerializerOptions options, IReadOnlyList<string> customCodes)
    {
        var sb = new StringBuilder(1024);
        sb.Append(@"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
        sb.Append(@"<numFmts count=""").Append(5 + customCodes.Count).Append(@""">");
        AppendNumFmt(sb, 1, options.DateTimeFormat);
        AppendNumFmt(sb, 2, options.DateFormat);
        AppendNumFmt(sb, 3, options.TimeFormat);
        AppendNumFmt(sb, 4, options.IntegerFormat);
        AppendNumFmt(sb, 5, options.NumberFormat);
        for (var i = 0; i < customCodes.Count; i++)
            AppendNumFmt(sb, CUSTOM_NUMFMT_BASE + i, customCodes[i]);
        sb.Append("</numFmts>");
        sb.Append("<fonts count=\"1\"><font/></fonts>");
        sb.Append("<fills count=\"1\"><fill/></fills>");
        sb.Append("<borders count=\"1\"><border/></borders>");
        sb.Append("<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>");
        sb.Append(@"<cellXfs count=""").Append(7 + customCodes.Count).Append(@""">");
        sb.Append("<xf/>");
        sb.Append(@"<xf><alignment wrapText=""true""/></xf>");
        for (var id = 1; id <= 5; id++)
            sb.Append(@"<xf numFmtId=""").Append(id).Append(@""" applyNumberFormat=""1""></xf>");
        for (var i = 0; i < customCodes.Count; i++)
            sb.Append(@"<xf numFmtId=""").Append(CUSTOM_NUMFMT_BASE + i).Append(@""" applyNumberFormat=""1""></xf>");
        sb.Append("</cellXfs>");
        sb.Append("</styleSheet>");
        return Encoding.UTF8.GetBytes(sb.ToString());

        static void AppendNumFmt(StringBuilder sb, int id, string code)
            => sb.Append(@"<numFmt numFmtId=""").Append(id).Append(@""" formatCode=""").Append(EscapeAttribute(code)).Append(@"""/>");
    }

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
    /// <remarks>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row, so downstream consumers receive a file; an
    /// empty source without titles produces no file at all.</remarks>
    public static void ToFile<T>(IEnumerable<T> rows, string fileName, XlsxSerializerOptions? options = null)
    {
        options ??= XlsxSerializerOptions.Default;
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        // The source is enumerated exactly once, so it may be a query or a forward-only reader.
        using var enumerator = rows.GetEnumerator();
        var hasRows = enumerator.MoveNext();
        if (!hasRows && options.HeaderTitles is not { Length: > 0 })
            return;

        try
        {
            using var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);
            ToStream(enumerator, hasRows, fs, options);
        }
        catch
        {
            // A failed export must not leave a plausible-looking corrupt workbook behind:
            // a file that exists is a complete file, absence plus the exception is the failure.
            TryDeletePartialFile(fileName);
            throw;
        }
    }

    /// <summary>Writes .xlsx content to the stream. The stream does not need to be seekable
    /// (network streams are fine); it is left open after writing.</summary>
    /// <remarks>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row; an empty source without titles writes nothing.</remarks>
    public static void ToStream<T>(IEnumerable<T> rows, Stream stream, XlsxSerializerOptions? options = null)
    {
        options ??= XlsxSerializerOptions.Default;
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        using var enumerator = rows.GetEnumerator();
        var hasRows = enumerator.MoveNext();
        if (!hasRows && options.HeaderTitles is not { Length: > 0 })
            return;

        ToStream(enumerator, hasRows, stream, options);
    }

    static void ToStream<T>(IEnumerator<T> rows, bool hasRows, Stream stream, XlsxSerializerOptions options)
    {
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        WriteEntry(archive, CONTENT_TYPE_XML, _contentTypes, options.CompressionLevel);
        WriteEntry(archive, RELS + "/" + DOT_RELS, _rels, options.CompressionLevel);
        WriteEntry(archive, BOOK_XML, BuildBook(options), options.CompressionLevel);
        WriteEntry(archive, RELS + "/" + BOOK_XML_RELS, _bookRels, options.CompressionLevel);

        using var writer = new XlsxCellWriter(options);
        using (var sheetStream = archive.CreateEntry(SHEET_XML, options.CompressionLevel).Open())
            CreateSheet(rows, hasRows, sheetStream, writer, options);
        using (var stringsStream = archive.CreateEntry(STRINGS_XML, options.CompressionLevel).Open())
            WriteSharedStrings(stringsStream, writer);
        WriteEntry(archive, STYLES_XML, BuildStyles(options, writer.Formats.Codes), options.CompressionLevel);
    }

    /// <summary>Writes .xlsx content to an <see cref="IBufferWriter{T}"/> such as
    /// System.IO.Pipelines.PipeWriter (e.g. ASP.NET Core's Response.BodyWriter).
    /// Flushing the underlying pipe is left to the caller.</summary>
    public static void ToBufferWriter<T>(IEnumerable<T> rows, IBufferWriter<byte> bufferWriter, XlsxSerializerOptions? options = null)
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

    static void TryDeletePartialFile(string fileName)
    {
        try
        {
            File.Delete(fileName);
        }
        catch
        {
            // The original exception is the story; a failed cleanup must not replace it.
        }
    }

    /// <summary>Creates an .xlsx file asynchronously. The output is streamed; no working folder is used.</summary>
    /// <remarks>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row; an empty source without titles produces no file.</remarks>
    public static async Task ToFileAsync<T>(IEnumerable<T> rows, string fileName, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= XlsxSerializerOptions.Default;
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        using var enumerator = rows.GetEnumerator();
        var hasRows = enumerator.MoveNext();
        if (!hasRows && options.HeaderTitles is not { Length: > 0 })
            return;

        try
        {
            using var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
            await WriteContainerAsync(fs, options,
                (s, w) => CreateSheetAsync(enumerator, hasRows, s, w, options, cancellationToken),
                cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            TryDeletePartialFile(fileName);
            throw;
        }
    }

    /// <summary>Writes .xlsx content to the stream asynchronously. The stream does not need to be
    /// seekable (network streams are fine); it is left open after writing.</summary>
    /// <remarks>
    /// <para>The stream is only ever written asynchronously - the zip container's few
    /// synchronous tail writes are buffered and forwarded with the next asynchronous write -
    /// so a stream that forbids synchronous IO, such as an ASP.NET Core response body, is a
    /// valid destination.</para>
    /// <para>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row; an empty source without titles writes nothing.</para>
    /// </remarks>
    public static async Task ToStreamAsync<T>(IEnumerable<T> rows, Stream stream, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= XlsxSerializerOptions.Default;
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        using var enumerator = rows.GetEnumerator();
        var hasRows = enumerator.MoveNext();
        if (!hasRows && options.HeaderTitles is not { Length: > 0 })
            return;

        await WriteContainerAsync(stream, options,
            (s, w) => CreateSheetAsync(enumerator, hasRows, s, w, options, cancellationToken),
            cancellationToken).ConfigureAwait(false);
    }

#if !NETSTANDARD2_0
    /// <summary>Creates an .xlsx file from an asynchronous source. Rows are awaited as they
    /// arrive - a query's <c>AsAsyncEnumerable()</c> or a streaming service response can be
    /// passed directly, without materializing it first.</summary>
    /// <remarks>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row; an empty source without titles produces no file.</remarks>
    public static async Task ToFileAsync<T>(IAsyncEnumerable<T> rows, string fileName, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= XlsxSerializerOptions.Default;
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        var enumerator = rows.GetAsyncEnumerator(cancellationToken);
        await using (enumerator.ConfigureAwait(false))
        {
            var hasRows = await enumerator.MoveNextAsync().ConfigureAwait(false);
            if (!hasRows && options.HeaderTitles is not { Length: > 0 })
                return;

            try
            {
                using var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
                await WriteContainerAsync(fs, options,
                    (s, w) => CreateSheetAsync(enumerator, hasRows, s, w, options, cancellationToken),
                    cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                TryDeletePartialFile(fileName);
                throw;
            }
        }
    }

    /// <summary>Writes .xlsx content to the stream from an asynchronous source. Rows are
    /// awaited as they arrive; the stream is only ever written asynchronously and is left
    /// open after writing.</summary>
    /// <remarks>An empty source with <see cref="XlsxSerializerOptions.HeaderTitles"/> set still
    /// produces a workbook with the header row; an empty source without titles writes nothing.</remarks>
    public static async Task ToStreamAsync<T>(IAsyncEnumerable<T> rows, Stream stream, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= XlsxSerializerOptions.Default;
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));

        var enumerator = rows.GetAsyncEnumerator(cancellationToken);
        await using (enumerator.ConfigureAwait(false))
        {
            var hasRows = await enumerator.MoveNextAsync().ConfigureAwait(false);
            if (!hasRows && options.HeaderTitles is not { Length: > 0 })
                return;

            await WriteContainerAsync(stream, options,
                (s, w) => CreateSheetAsync(enumerator, hasRows, s, w, options, cancellationToken),
                cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Writes .xlsx content to a <see cref="System.IO.Pipelines.PipeWriter"/> from an
    /// asynchronous source, flushing as data is produced.</summary>
    public static Task ToPipeWriterAsync<T>(IAsyncEnumerable<T> rows, System.IO.Pipelines.PipeWriter pipeWriter, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (pipeWriter == null)
            throw new ArgumentNullException(nameof(pipeWriter));

        return ToStreamAsync(rows, pipeWriter.AsStream(leaveOpen: true), options, cancellationToken);
    }
#endif

    static async Task WriteContainerAsync(Stream stream, XlsxSerializerOptions options, Func<Stream, XlsxCellWriter, Task> writeSheet, CancellationToken cancellationToken)
    {
        // The zip machinery performs a handful of synchronous tail writes even on its async
        // path (measured on 10.0.9: the stream ZipArchiveEntry.OpenAsync returns still closes
        // synchronously). Buffering those keeps the caller's stream strictly asynchronous,
        // which is what an ASP.NET Core response body demands.
        using var buffered = new SyncWriteBufferingStream(stream);
#if NET10_0_OR_GREATER
        var archive = await ZipArchive.CreateAsync(buffered, ZipArchiveMode.Create, leaveOpen: true, entryNameEncoding: null, cancellationToken).ConfigureAwait(false);
        try
        {
            await WriteWorkbookAsync(archive, options, writeSheet, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            await archive.DisposeAsync().ConfigureAwait(false);
        }
#else
        using (var archive = new ZipArchive(buffered, ZipArchiveMode.Create, leaveOpen: true))
        {
            await WriteWorkbookAsync(archive, options, writeSheet, cancellationToken).ConfigureAwait(false);
        }
#endif
        await buffered.FlushPendingAsync(cancellationToken).ConfigureAwait(false);
    }

    static async Task WriteWorkbookAsync(ZipArchive archive, XlsxSerializerOptions options, Func<Stream, XlsxCellWriter, Task> writeSheet, CancellationToken cancellationToken)
    {
        await WriteEntryAsync(archive, CONTENT_TYPE_XML, _contentTypes, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, RELS + "/" + DOT_RELS, _rels, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, BOOK_XML, BuildBook(options), options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        await WriteEntryAsync(archive, RELS + "/" + BOOK_XML_RELS, _bookRels, options.CompressionLevel, cancellationToken).ConfigureAwait(false);

        using var writer = new XlsxCellWriter(options);

        var sheetStream = await OpenEntryAsync(archive, SHEET_XML, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        try
        {
            await writeSheet(sheetStream, writer).ConfigureAwait(false);
        }
        finally
        {
            await DisposeEntryAsync(sheetStream).ConfigureAwait(false);
        }

        var stringsStream = await OpenEntryAsync(archive, STRINGS_XML, options.CompressionLevel, cancellationToken).ConfigureAwait(false);
        try
        {
            await WriteSharedStringsAsync(stringsStream, writer, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            await DisposeEntryAsync(stringsStream).ConfigureAwait(false);
        }

        await WriteEntryAsync(archive, STYLES_XML, BuildStyles(options, writer.Formats.Codes), options.CompressionLevel, cancellationToken).ConfigureAwait(false);
    }

    static Task<Stream> OpenEntryAsync(ZipArchive archive, string entryName, System.IO.Compression.CompressionLevel compressionLevel, CancellationToken cancellationToken)
    {
#if NET10_0_OR_GREATER
        return archive.CreateEntry(entryName, compressionLevel).OpenAsync(cancellationToken);
#else
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(archive.CreateEntry(entryName, compressionLevel).Open());
#endif
    }

    static Task DisposeEntryAsync(Stream entryStream)
    {
#if NET10_0_OR_GREATER
        return entryStream.DisposeAsync().AsTask();
#else
        entryStream.Dispose();
        return Task.CompletedTask;
#endif
    }

    /// <summary>Writes .xlsx content to a <see cref="System.IO.Pipelines.PipeWriter"/>
    /// (e.g. ASP.NET Core's Response.BodyWriter). Data is flushed to the pipe as it is
    /// produced, so backpressure is honored.</summary>
    /// <remarks>The pipe is left uncompleted: ASP.NET Core completes the response pipe itself;
    /// a hand-made <see cref="System.IO.Pipelines.Pipe"/> is completed by the caller.</remarks>
    public static Task ToPipeWriterAsync<T>(IEnumerable<T> rows, System.IO.Pipelines.PipeWriter pipeWriter, XlsxSerializerOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (pipeWriter == null)
            throw new ArgumentNullException(nameof(pipeWriter));

        return ToStreamAsync(rows, pipeWriter.AsStream(leaveOpen: true), options, cancellationToken);
    }

    static async Task WriteEntryAsync(ZipArchive archive, string entryName, byte[] bytes, System.IO.Compression.CompressionLevel compressionLevel, CancellationToken cancellationToken)
    {
        var s = await OpenEntryAsync(archive, entryName, compressionLevel, cancellationToken).ConfigureAwait(false);
        try
        {
            await s.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            await DisposeEntryAsync(s).ConfigureAwait(false);
        }
    }

    static async Task CreateSheetAsync<T>(
        IEnumerator<T> rows,
        bool hasRows,
        Stream stream,
        XlsxCellWriter writer,
        XlsxSerializerOptions options,
        CancellationToken cancellationToken
    )
    {
        await stream.WriteAsync(_sheetStart, 0, _sheetStart.Length, cancellationToken).ConfigureAwait(false);

        if (options.HasHeader)
            await stream.WriteAsync(_frozenTitleRow, 0, _frozenTitleRow.Length, cancellationToken).ConfigureAwait(false);

        var serializer = options.GetSerializer<T>();

        // When the caller found a first row it has already advanced to it. Auto-fit has to
        // measure rows before <cols> is written, so it buffers them - but never more than
        // AutoFitSampleRows, and the source is still enumerated exactly once.
        var buffered = new List<T>();
        if (hasRows)
            buffered.Add(rows.Current);
        if (options.AutoFitColumns && serializer != null)
        {
            while (hasRows && buffered.Count < options.AutoFitSampleRows && rows.MoveNext())
                buffered.Add(rows.Current);
            await WriteCellWidthAsync(buffered, stream, writer, options, cancellationToken).ConfigureAwait(false);
        }

        await stream.WriteAsync(_dataStart, 0, _dataStart.Length, cancellationToken).ConfigureAwait(false);

        if (serializer != null)
        {
            if (options.HasHeader)
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                if (options.HeaderTitles != null && options.HeaderTitles.Length > 0)
                {
                    foreach (var t in options.HeaderTitles)
                        writer.Write(t);
                }
                else if (buffered.Count > 0)
                {
                    serializer.WriteTitle(writer, buffered[0], options);
                }
                writer.WriteRaw(_rowEnd);
                await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }

            foreach (var row in buffered)
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, row, options);
                writer.WriteRaw(_rowEnd);
                if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                    await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }

            while (hasRows && rows.MoveNext())
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, rows.Current, options);
                writer.WriteRaw(_rowEnd);
                if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                    await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }
        }
        writer.WriteRaw(_dataEnd);

        if (options.AutoFilter)
        {
            // Derived from what was actually written: re-enumerating rows here would run a lazy
            // source (a query, a reader) a second time, and ColumnMaxLength is only populated
            // when AutoFitColumns is on.
            var colName = options.HeaderTitles != null && options.HeaderTitles.Any()
                ? ToColumnName(options.HeaderTitles.Length)
                : ToColumnName(writer.MaxColumnCount);

            var range = $"A1:{colName}{writer.RowCount}";
            writer.WriteRaw(_autoFilterStart);
            writer.WriteRaw(Encoding.UTF8.GetBytes(range));
            writer.WriteRaw(_autoFilterEnd);
        }

        writer.WriteRaw(_sheetEnd);
        await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }

#if !NETSTANDARD2_0
    /// <summary>The asynchronous-source twin of the enumerator overload above; the only
    /// difference is that advancing the source is awaited.</summary>
    static async Task CreateSheetAsync<T>(
        IAsyncEnumerator<T> rows,
        bool hasRows,
        Stream stream,
        XlsxCellWriter writer,
        XlsxSerializerOptions options,
        CancellationToken cancellationToken
    )
    {
        await stream.WriteAsync(_sheetStart, 0, _sheetStart.Length, cancellationToken).ConfigureAwait(false);

        if (options.HasHeader)
            await stream.WriteAsync(_frozenTitleRow, 0, _frozenTitleRow.Length, cancellationToken).ConfigureAwait(false);

        var serializer = options.GetSerializer<T>();

        var buffered = new List<T>();
        if (hasRows)
            buffered.Add(rows.Current);
        if (options.AutoFitColumns && serializer != null)
        {
            while (hasRows && buffered.Count < options.AutoFitSampleRows && await rows.MoveNextAsync().ConfigureAwait(false))
                buffered.Add(rows.Current);
            await WriteCellWidthAsync(buffered, stream, writer, options, cancellationToken).ConfigureAwait(false);
        }

        await stream.WriteAsync(_dataStart, 0, _dataStart.Length, cancellationToken).ConfigureAwait(false);

        if (serializer != null)
        {
            if (options.HasHeader)
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                if (options.HeaderTitles != null && options.HeaderTitles.Length > 0)
                {
                    foreach (var t in options.HeaderTitles)
                        writer.Write(t);
                }
                else if (buffered.Count > 0)
                {
                    serializer.WriteTitle(writer, buffered[0], options);
                }
                writer.WriteRaw(_rowEnd);
                await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }

            foreach (var row in buffered)
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, row, options);
                writer.WriteRaw(_rowEnd);
                if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                    await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }

            while (hasRows && await rows.MoveNextAsync().ConfigureAwait(false))
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, rows.Current, options);
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
                : ToColumnName(writer.MaxColumnCount);

            var range = $"A1:{colName}{writer.RowCount}";
            writer.WriteRaw(_autoFilterStart);
            writer.WriteRaw(Encoding.UTF8.GetBytes(range));
            writer.WriteRaw(_autoFilterEnd);
        }

        writer.WriteRaw(_sheetEnd);
        await writer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }
#endif

    static async Task WriteCellWidthAsync<T>(
        List<T> rows,
        Stream stream,
        XlsxCellWriter writer,
        XlsxSerializerOptions options,
        CancellationToken cancellationToken
    )
    {
        var serializer = options.GetSerializer<T>();
        if (serializer == null) return;
        if (options.HasHeader && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                writer.Write(t);
            writer.Clear();
        }
        foreach (var row in rows)
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
            var width = Math.Min(options.AutoFitMaxWidth, pair.Value + COLUMN_WIDTH_MARGIN);

            WriteUtf8Bytes(@$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />", buffer);
        }
        buffer.Write(_colEnd);
        await buffer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    static async Task WriteSharedStringsAsync(Stream stream, XlsxCellWriter writer, CancellationToken cancellationToken)
    {
        await stream.WriteAsync(_sstStart, 0, _sstStart.Length, cancellationToken).ConfigureAwait(false);
        using var buffer = new ArrayPoolBufferWriter();
        foreach (var s in writer.SharedStrings.Values)
        {
            buffer.Write(_siStart);
            XmlEscape.WriteEscaped(s, buffer);
            buffer.Write(_siEnd);
            await buffer.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
        }
        await stream.WriteAsync(_sstEnd, 0, _sstEnd.Length, cancellationToken).ConfigureAwait(false);
    }

    static void CreateSheet<T>(
        IEnumerator<T> rows,
        bool hasRows,
        Stream stream,
        XlsxCellWriter writer,
        XlsxSerializerOptions options
    )
    {
        stream.Write(_sheetStart, 0, _sheetStart.Length);

        if (options.HasHeader)
            stream.Write(_frozenTitleRow, 0, _frozenTitleRow.Length);

        var serializer = options.GetSerializer<T>();

        // When the caller found a first row it has already advanced to it. Auto-fit has to
        // measure rows before <cols> is written, so it buffers them - but never more than
        // AutoFitSampleRows, and the source is still enumerated exactly once.
        var buffered = new List<T>();
        if (hasRows)
            buffered.Add(rows.Current);
        if (options.AutoFitColumns && serializer != null)
        {
            while (hasRows && buffered.Count < options.AutoFitSampleRows && rows.MoveNext())
                buffered.Add(rows.Current);
            WriteCellWidth(buffered, stream, writer, options);
        }

        stream.Write(_dataStart, 0, _dataStart.Length);

        if (serializer != null)
        {
            if (options.HasHeader)
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                if (options.HeaderTitles != null && options.HeaderTitles.Length > 0)
                {
                    foreach (var t in options.HeaderTitles)
                        writer.Write(t);
                }
                else if (buffered.Count > 0)
                {
                    serializer.WriteTitle(writer, buffered[0], options);
                }
                writer.WriteRaw(_rowEnd);
                writer.CopyTo(stream);
            }

#if NET5_0_OR_GREATER
            WriteRowsSpan(CollectionsMarshal.AsSpan(buffered), stream, writer, serializer, options);
#else
            WriteRows(buffered, stream, writer, serializer, options);
#endif
            while (hasRows && rows.MoveNext())
            {
                writer.BeginRow();
                writer.WriteRaw(_rowStart);
                serializer.Serialize(writer, rows.Current, options);
                writer.WriteRaw(_rowEnd);
                if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                    writer.CopyTo(stream);
            }
        }
        writer.WriteRaw(_dataEnd);

        if (options.AutoFilter)
        {
            // Derived from what was actually written: re-enumerating rows here would run a lazy
            // source (a query, a reader) a second time, and ColumnMaxLength is only populated
            // when AutoFitColumns is on.
            var colName = options.HeaderTitles != null && options.HeaderTitles.Any()
                ? ToColumnName(options.HeaderTitles.Length)
                : ToColumnName(writer.MaxColumnCount);

            var range = $"A1:{colName}{writer.RowCount}";
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
        XlsxCellWriter writer,
        IXlsxSerializer<T> serializer,
        XlsxSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.BeginRow();
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
        XlsxCellWriter writer,
        IXlsxSerializer<T> serializer,
        XlsxSerializerOptions options
    )
    {
        foreach (var row in rows)
        {
            writer.BeginRow();
            writer.WriteRaw(_rowStart);
            serializer.Serialize(writer, row, options);
            writer.WriteRaw(_rowEnd);
            if (writer.BufferedBytes >= FLUSH_THRESHOLD)
                writer.CopyTo(stream);
        }
    }

    static void WriteCellWidth<T>(
        List<T> rows,
        Stream stream,
        XlsxCellWriter writer,
        XlsxSerializerOptions options
    )
    {
        // Counting the number of characters in Writer's internal process
        // The result is stored in writer.ColumnMaxLength 
        var serializer = options.GetSerializer<T>();
        if (serializer == null) return;
        if (options.HasHeader && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                writer.Write(t);
            writer.Clear();
        }
        foreach (var row in rows)
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
            var width = Math.Min(options.AutoFitMaxWidth, pair.Value + COLUMN_WIDTH_MARGIN);

            WriteUtf8Bytes(@$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />", buffer);
        }
        buffer.Write(_colEnd);
        buffer.CopyTo(stream);
    }

    static void WriteSharedStrings(Stream stream, XlsxCellWriter writer)
    {
        stream.Write(_sstStart, 0, _sstStart.Length);
        using var buffer = new ArrayPoolBufferWriter();
        foreach (var s in writer.SharedStrings.Values)
        {
            buffer.Write(_siStart);
            XmlEscape.WriteEscaped(s, buffer);
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
