using System.IO.Compression;
using System.Xml.Linq;

namespace FlatXlsx.Tests.Support;

/// <summary>
/// A deliberately small reader for the workbooks FlatXlsx produces: enough SpreadsheetML to
/// recover what a spreadsheet application would show the user.
/// </summary>
/// <remarks>
/// Tests assert against this rather than against raw markup. What FlatXlsx promises is the
/// value the reader sees, not the encoding it happens to use to get there - shared string vs
/// inline string, or which style index carries the number format, are free to change.
/// Assertions written against the markup turn every such change into a false claim.
/// </remarks>
static class Workbook
{
    public sealed record Cell(string? Text, string? NumberFormat, bool Wrapped)
    {
        public override string ToString() => Text ?? "<empty>";
    }

    public sealed record Sheet(
        IReadOnlyList<IReadOnlyList<Cell>> Rows,
        string? AutoFilterRange,
        IReadOnlyDictionary<int, double> ColumnWidths,
        int DistinctStoredStrings,
        string SheetName)
    {
        public IReadOnlyList<Cell> Row(int index) => Rows[index];

        /// <summary>Text of every cell in the row, empty cells included.</summary>
        public string?[] Texts(int rowIndex) => Rows[rowIndex].Select(c => c.Text).ToArray();
    }

    public static Sheet Read(byte[] xlsx)
    {
        using var ms = new MemoryStream(xlsx);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        var sharedStrings = ReadSharedStrings(archive);
        var styles = ReadCellStyles(archive);
        var sheet = Parse(archive, "sheet.xml");
        var ns = sheet.Root!.Name.Namespace;

        var rows = new List<IReadOnlyList<Cell>>();
        var sheetData = sheet.Root.Element(ns + "sheetData");
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements(ns + "row"))
                rows.Add(row.Elements(ns + "c").Select(c => ReadCell(c, ns, sharedStrings, styles)).ToArray());
        }

        var widths = sheet.Root.Element(ns + "cols")?
            .Elements(ns + "col")
            .ToDictionary(
                c => int.Parse(c.Attribute("min")!.Value),
                c => double.Parse(c.Attribute("width")!.Value, System.Globalization.CultureInfo.InvariantCulture))
            ?? new Dictionary<int, double>();

        var autoFilter = sheet.Root.Element(ns + "autoFilter")?.Attribute("ref")?.Value;
        return new Sheet(rows, autoFilter, widths, sharedStrings.Length, ReadSheetName(archive));
    }

    static string ReadSheetName(ZipArchive archive)
    {
        var book = Parse(archive, "book.xml");
        var ns = book.Root!.Name.Namespace;
        return book.Root.Element(ns + "sheets")!.Element(ns + "sheet")!.Attribute("name")!.Value;
    }

    /// <summary>Every part must be well-formed XML; a workbook that fails this cannot be opened.</summary>
    public static void AssertEveryPartIsWellFormed(byte[] xlsx)
    {
        using var ms = new MemoryStream(xlsx);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            using var reader = new StreamReader(entry.Open());
            XDocument.Parse(reader.ReadToEnd());
        }
    }

    sealed record CellStyle(string? NumberFormat, bool Wrapped);

    static Cell ReadCell(XElement cell, XNamespace ns, string[] sharedStrings, CellStyle[] styles)
    {
        var style = StyleOf(cell, styles);
        var type = cell.Attribute("t")?.Value;

        var text = type switch
        {
            "s" => sharedStrings[int.Parse(cell.Element(ns + "v")!.Value)],
            "inlineStr" => cell.Element(ns + "is")?.Element(ns + "t")?.Value,
            // SpreadsheetML stores booleans as 1/0; report what the user would see.
            "b" => cell.Element(ns + "v")!.Value == "1" ? "True" : "False",
            _ => cell.Element(ns + "v")?.Value,
        };
        return new Cell(text, style.NumberFormat, style.Wrapped);
    }

    static CellStyle StyleOf(XElement cell, CellStyle[] styles)
    {
        var style = cell.Attribute("s")?.Value;
        if (style == null)
            return new CellStyle(null, false);
        var index = int.Parse(style);
        return index < styles.Length ? styles[index] : new CellStyle(null, false);
    }

    static string[] ReadSharedStrings(ZipArchive archive)
    {
        var doc = Parse(archive, "strings.xml");
        var ns = doc.Root!.Name.Namespace;
        return doc.Root.Elements(ns + "si").Select(si => si.Element(ns + "t")?.Value ?? "").ToArray();
    }

    /// <summary>Resolves each cell style index to what it actually does to the cell, so tests can
    /// state "this is shown with the integer format" instead of naming a style index.</summary>
    static CellStyle[] ReadCellStyles(ZipArchive archive)
    {
        var doc = Parse(archive, "styles.xml");
        var ns = doc.Root!.Name.Namespace;

        var codes = doc.Root.Element(ns + "numFmts")?
            .Elements(ns + "numFmt")
            .ToDictionary(f => f.Attribute("numFmtId")!.Value, f => f.Attribute("formatCode")!.Value)
            ?? new Dictionary<string, string>();

        return doc.Root.Element(ns + "cellXfs")!
            .Elements(ns + "xf")
            .Select(xf =>
            {
                var id = xf.Attribute("numFmtId")?.Value;
                var code = id != null && codes.TryGetValue(id, out var c) ? c : null;
                var wrapped = xf.Element(ns + "alignment")?.Attribute("wrapText")?.Value == "true";
                return new CellStyle(code, wrapped);
            })
            .ToArray();
    }

    static XDocument Parse(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"The workbook has no {entryName} part.");
        using var reader = new StreamReader(entry.Open());
        return XDocument.Parse(reader.ReadToEnd());
    }
}
