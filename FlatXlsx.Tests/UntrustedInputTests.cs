using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises that hold no matter where the values came from.
/// </summary>
/// <remarks>
/// Claimant: the team exporting data they did not author - from a database, a form, an upload.
/// They cannot inspect every value first, so a single hostile or merely awkward one must not be
/// able to corrupt the workbook, rewrite its structure, or exhaust the process.
/// </remarks>
public class UntrustedInputTests
{
    [Fact]
    public void Markup_in_a_value_stays_a_value()
    {
        var payload = "</t></is></c><f>1+1</f>";

        var sheet = Xlsx.Read(new[] { payload, "<script>&amp;" }, XlsxSerializerOptions.Default);

        Assert.Equal(payload, sheet.Texts(0)[0]);
        Assert.Equal("<script>&amp;", sheet.Texts(1)[0]);
        Assert.Equal(2, sheet.Rows.Count);
    }

    [Fact]
    public void Markup_in_a_header_title_stays_a_title()
    {
        var options = XlsxSerializerOptions.Default with
        {
            HasHeaderRecord = true,
            HeaderTitles = new[] { "<b>Name</b>" },
        };

        var sheet = Xlsx.Read(new[] { "row1" }, options);

        Assert.Equal("<b>Name</b>", sheet.Texts(0)[0]);
    }

    [Fact]
    public void Characters_XML_cannot_carry_are_dropped_rather_than_breaking_the_file()
    {
        // A NUL in a value is illegal in XML even escaped; emitting one produces a workbook
        // that reports itself as damaged, losing the whole export rather than one character.
        var sheet = Xlsx.Read(new[] { "a\u0000b\u0001c\u000Bd" }, XlsxSerializerOptions.Default);

        Assert.Equal("abcd", sheet.Texts(0)[0]);
    }

    [Fact]
    public void Line_breaks_inside_a_value_survive_unchanged()
    {
        var sheet = Xlsx.Read(new[] { "first\r\nsecond\tthird" }, XlsxSerializerOptions.Default);

        Assert.Equal("first\r\nsecond\tthird", sheet.Texts(0)[0]);
    }

    [Fact]
    public void High_cardinality_text_stops_growing_the_shared_string_table()
    {
        // The table is held until the sheet is finished, so unbounded growth is an
        // out-of-memory risk on exports whose size the caller does not control.
        var options = XlsxSerializerOptions.Default with { MaxSharedStrings = 2 };
        var values = new[] { "a", "b", "c<x>", "d" };

        var sheet = Xlsx.Read(values, options);

        Assert.Equal(2, sheet.DistinctStoredStrings);
        Assert.Equal(values, sheet.Rows.Select(r => r[0].Text).ToArray());
    }

    [Fact]
    public void A_value_too_long_for_a_cell_is_refused_before_the_file_is_written()
    {
        var rows = new[] { new string('a', 32_768) };

        var ex = Assert.Throws<InvalidOperationException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Contains("32767", ex.Message);
    }

    [Fact]
    public void A_value_at_the_cell_limit_is_still_accepted()
    {
        var text = new string('a', 32_767);

        var sheet = Xlsx.Read(new[] { text }, XlsxSerializerOptions.Default);

        Assert.Equal(text, sheet.Texts(0)[0]);
    }

    class Node
    {
        public string Name { get; set; } = "";
        public Node? Child { get; set; }
    }

    [Fact]
    public void A_circular_reference_stops_at_the_configured_depth()
    {
        var node = new Node { Name = "a" };
        node.Child = node;

        var ex = Assert.Throws<InvalidOperationException>(
            () => Xlsx.Write(new[] { node }, XlsxSerializerOptions.Default));

        Assert.Contains("max depth", ex.Message);
    }

    [Fact]
    public void A_collection_wider_than_a_sheet_is_refused_by_name()
    {
        // Nested collections expand sideways; past the sheet's limit the file cannot be opened,
        // so the caller is told which limit they hit instead of receiving a broken workbook.
        var rows = new[] { Enumerable.Range(0, 16_385).ToArray() };

        var ex = Assert.Throws<InvalidOperationException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Contains("16384", ex.Message);
    }

    [Fact]
    public void The_row_source_is_read_exactly_once()
    {
        // Claimant: anyone exporting straight from a query or a forward-only reader, for whom a
        // second pass either costs another round trip or returns nothing at all.
        var reads = 0;
        IEnumerable<string> Source()
        {
            reads++;
            yield return "a";
            yield return "b";
        }

        var options = XlsxSerializerOptions.Default with { AutoFilter = true, HasHeaderRecord = true };
        Xlsx.Write(Source(), options);

        Assert.Equal(1, reads);
    }

    [Fact]
    public void Auto_fitting_reads_the_source_once_as_well()
    {
        var reads = 0;
        IEnumerable<string> Source()
        {
            reads++;
            for (var i = 0; i < 500; i++)
                yield return $"value-{i}";
        }

        var options = XlsxSerializerOptions.Default with { AutoFitColumns = true };
        var sheet = Workbook.Read(Xlsx.Write(Source(), options));

        Assert.Equal(1, reads);
        Assert.Equal(500, sheet.Rows.Count);
    }
}
