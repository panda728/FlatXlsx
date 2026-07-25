using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// One entry per defect that reached a build, kept so it cannot return unnoticed.
/// </summary>
/// <remarks>
/// These are not organised by feature and are not meant to be: each is a record of something
/// that was actually wrong once. A found defect is recorded here as a failing test before it is
/// fixed, which is what makes the entry evidence rather than decoration.
/// </remarks>
public class RegressionLedgerTests
{
    [Fact]
    public void The_earliest_representable_date_occupies_one_cell_not_two()
    {
        // Written as an empty cell and then again as a date, because the empty-cell branch fell
        // through. Every column to its right shifted by one for that row only.
        var sheet = Xlsx.Read(new[] { DateTime.MinValue }, XlsxSerializerOptions.Default);

        Assert.Single(sheet.Row(0));
        Assert.Null(sheet.Texts(0)[0]);
    }

    [Fact]
    public void A_carriage_return_in_a_value_is_not_turned_into_a_line_feed()
    {
        // Written literally, it was rewritten to a line feed when the file was read back, since
        // XML normalises line endings. Values arriving from Windows-authored text changed shape
        // in the export.
        var sheet = Xlsx.Read(new[] { "before\rafter" }, XlsxSerializerOptions.Default);

        Assert.Equal("before\rafter", sheet.Texts(0)[0]);
    }

    [Fact]
    public void The_filter_range_matches_the_sheet_when_columns_are_not_auto_fitted()
    {
        // The range was derived from the auto-fit tally, which is empty unless AutoFitColumns is
        // on, so the filter covered no columns at all.
        var options = XlsxSerializerOptions.Default with { AutoFilter = true, HasHeaderRecord = true };

        var sheet = Xlsx.Read(new[] { (1, "a"), (2, "b") }, options);

        Assert.Equal("A1:B3", sheet.AutoFilterRange);
    }
}
