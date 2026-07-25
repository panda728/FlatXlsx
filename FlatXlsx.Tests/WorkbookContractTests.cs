using System.Runtime.Serialization;
using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// What the finished workbook promises its reader: which columns appear, in what order, under
/// what heading, and formatted so the values read as values.
/// </summary>
public class WorkbookContractTests
{
    class Portal
    {
        [DataMember(Name = "Owner Ex", Order = 1)]
        public string Owner { get; set; } = "";
        [DataMember(Name = "Name Ex", Order = 2)]
        public string Name { get; set; } = "";
        [DataMember(Name = "Level Ex", Order = 3)]
        public int Level { get; set; }
    }

    static readonly Portal[] _portals =
    {
        new() { Owner = "panda728", Name = "Portal1", Level = 8 },
        new() { Owner = "panda728", Name = "Portal2", Level = 1 },
    };

    [Fact]
    public void Members_appear_in_the_order_their_DataMember_asks_for()
    {
        var sheet = Xlsx.Read(_portals, XlsxSerializerOptions.Default with { HasHeaderRow = true });

        Assert.Equal(new[] { "Owner Ex", "Name Ex", "Level Ex" }, sheet.Texts(0));
        Assert.Equal(new[] { "panda728", "Portal1", "8" }, sheet.Texts(1));
    }

    [Fact]
    public void Supplying_titles_is_the_whole_ask_for_a_titled_header()
    {
        // The single most common request - "put these titles on top" - must be one setting.
        // Requiring a separate header switch turns forgetting it into a silent no-header file.
        var options = XlsxSerializerOptions.Default with
        {
            HeaderTitles = new[] { "Owner", "Portal", "Lv" },
        };

        var sheet = Xlsx.Read(_portals, options);

        Assert.Equal(new[] { "Owner", "Portal", "Lv" }, sheet.Texts(0));
        Assert.Equal(new[] { "panda728", "Portal1", "8" }, sheet.Texts(1));
    }

    [Fact]
    public void Options_can_be_built_without_knowing_about_providers()
    {
        // new XlsxSerializerOptions { ... } is the first thing a newcomer writes; it must work.
        var options = new XlsxSerializerOptions { HeaderTitles = new[] { "Owner", "Portal", "Lv" } };

        var sheet = Xlsx.Read(_portals, options);

        Assert.Equal(new[] { "Owner", "Portal", "Lv" }, sheet.Texts(0));
    }

    class Employee
    {
        [XlsxColumn("部署", Order = 1)]
        public string Department { get; set; } = "";

        [XlsxColumn("氏名", Order = 0)]
        public string Name { get; set; } = "";

        public int Age { get; set; }

        [XlsxIgnore]
        public string InternalNote { get; set; } = "";
    }

    [Fact]
    public void XlsxColumn_titles_and_orders_the_columns_where_they_are_declared()
    {
        var rows = new[] { new Employee { Department = "開発", Name = "北尾", Age = 30, InternalNote = "secret" } };

        var sheet = Xlsx.Read(rows, XlsxSerializerOptions.Default with { HasHeaderRow = true });

        Assert.Equal(new[] { "氏名", "部署", "Age" }, sheet.Texts(0));
        Assert.Equal(new[] { "北尾", "開発", "30" }, sheet.Texts(1));
    }

    [Fact]
    public void An_ignored_member_never_reaches_the_file()
    {
        // Claimant: whoever marked the member - the point is that its value must not leak.
        var rows = new[] { new Employee { Department = "開発", Name = "北尾", InternalNote = "secret" } };

        var sheet = Xlsx.Read(rows, XlsxSerializerOptions.Default);

        Assert.Equal(3, sheet.Row(0).Count);
        Assert.DoesNotContain("secret", sheet.Texts(0));
    }

    class DoublyAnnotated
    {
        [XlsxColumn("Xlsx名")]
        [System.Runtime.Serialization.DataMember(Name = "DataMember名")]
        public string Value { get; set; } = "";
    }

    [Fact]
    public void XlsxColumn_wins_when_DataMember_is_also_present()
    {
        // DataMember stays honored for types already annotated for other serializers, but the
        // attribute written for this library is the more specific intention.
        var sheet = Xlsx.Read(
            new[] { new DoublyAnnotated { Value = "v" } },
            XlsxSerializerOptions.Default with { HasHeaderRow = true });

        Assert.Equal(new[] { "Xlsx名" }, sheet.Texts(0));
    }

    [Fact]
    public void Without_a_header_the_first_row_is_already_data()
    {
        var sheet = Xlsx.Read(_portals, XlsxSerializerOptions.Default);

        Assert.Equal(new[] { "panda728", "Portal1", "8" }, sheet.Texts(0));
        Assert.Equal(2, sheet.Rows.Count);
    }

    [Fact]
    public void Every_row_of_the_source_becomes_a_row_of_the_sheet()
    {
        // Enough rows to cross the writer's internal flush boundary several times: a row lost at
        // a flush would be invisible to any single-row test.
        var rows = Enumerable.Range(0, 5_000).Select(i => $"row-{i}").ToArray();

        var sheet = Xlsx.Read(rows, XlsxSerializerOptions.Default);

        Assert.Equal(rows.Length, sheet.Rows.Count);
        Assert.Equal("row-0", sheet.Texts(0)[0]);
        Assert.Equal("row-4999", sheet.Texts(rows.Length - 1)[0]);
    }

    [Fact]
    public void Whole_numbers_are_shown_with_the_integer_format()
    {
        // Claimant: the reader who expects 1,234 rather than 1234.00 in an id or count column.
        var options = XlsxSerializerOptions.Default;

        var sheet = Xlsx.Read(new[] { 1234 }, options);

        Assert.Equal(options.IntegerFormat, sheet.Row(0)[0].NumberFormat);
    }

    [Fact]
    public void Fractional_numbers_are_shown_with_the_number_format()
    {
        var options = XlsxSerializerOptions.Default;

        var sheet = Xlsx.Read(new[] { 12.34d }, options);

        Assert.Equal(options.NumberFormat, sheet.Row(0)[0].NumberFormat);
    }

    [Fact]
    public void Dates_and_times_each_get_their_own_display_format()
    {
        var options = XlsxSerializerOptions.Default;

        var dateTime = Xlsx.Read(new[] { new DateTime(2023, 1, 2, 3, 4, 5) }, options);
        var date = Xlsx.Read(new[] { new DateTime(2023, 1, 2) }, options);
        var time = Xlsx.Read(new[] { new TimeOnly(3, 4, 5) }, options);

        Assert.Equal(options.DateTimeFormat, dateTime.Row(0)[0].NumberFormat);
        Assert.Equal(options.DateFormat, date.Row(0)[0].NumberFormat);
        Assert.Equal(options.TimeFormat, time.Row(0)[0].NumberFormat);
    }

    [Fact]
    public void Custom_formats_from_the_options_reach_the_cells()
    {
        var options = XlsxSerializerOptions.Default with { IntegerFormat = "0000" };

        var sheet = Xlsx.Read(new[] { 7 }, options);

        Assert.Equal("0000", sheet.Row(0)[0].NumberFormat);
    }

    [Fact]
    public void Text_with_a_line_break_is_marked_to_wrap()
    {
        // Otherwise the second line is hidden behind the next column and the reader never
        // learns it exists.
        var sheet = Xlsx.Read(new[] { "line1\nline2", "single" }, XlsxSerializerOptions.Default);

        Assert.True(sheet.Row(0)[0].Wrapped);
        Assert.False(sheet.Row(1)[0].Wrapped);
    }

    [Fact]
    public void Auto_fitted_columns_are_wide_enough_for_their_longest_value()
    {
        var options = XlsxSerializerOptions.Default with { AutoFitColumns = true };

        var sheet = Xlsx.Read(new[] { "short", "a much longer value" }, options);

        Assert.True(
            sheet.ColumnWidths[1] >= "a much longer value".Length,
            $"column 1 was {sheet.ColumnWidths[1]} wide");
    }

    [Fact]
    public void Auto_fitted_width_stops_at_the_configured_maximum()
    {
        var options = XlsxSerializerOptions.Default with { AutoFitColumns = true, AutoFitWidthMax = 10 };

        var sheet = Xlsx.Read(new[] { new string('x', 500) }, options);

        Assert.Equal(10, sheet.ColumnWidths[1]);
    }

    [Fact]
    public void The_filter_covers_the_header_and_every_row_written()
    {
        var options = XlsxSerializerOptions.Default with { AutoFilter = true, HasHeaderRow = true };

        var sheet = Xlsx.Read(new[] { "a", "b", "c" }, options);

        Assert.Equal("A1:A4", sheet.AutoFilterRange);
    }

    [Fact]
    public void The_filter_covers_the_data_when_there_is_no_header()
    {
        var options = XlsxSerializerOptions.Default with { AutoFilter = true };

        var sheet = Xlsx.Read(new[] { "a", "b", "c" }, options);

        Assert.Equal("A1:A3", sheet.AutoFilterRange);
    }

    [Fact]
    public void An_empty_source_produces_no_workbook_at_all()
    {
        // Claimant: the caller writing to a response or a file, who must not hand over a
        // zero-row workbook that a spreadsheet application would reject as damaged.
        using var ms = new MemoryStream();

        XlsxSerializer.ToStream(Array.Empty<string>(), ms, XlsxSerializerOptions.Default);

        Assert.Equal(0, ms.Length);
    }
}
