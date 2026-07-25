using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises that must hold on the machine the export actually runs on.
/// </summary>
/// <remarks>
/// Claimant: everyone downstream of a server whose locale nobody chose deliberately. A number
/// or a date is stored in the file in a fixed, locale-independent form and only displayed
/// according to the reader's settings; if the ambient culture leaks into what is stored, the
/// same code produces a different - and often unreadable - file depending on where it runs.
/// Each case runs on its own thread so the culture it sets cannot reach any other test.
/// </remarks>
public class AmbientEnvironmentTests
{
    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]   // comma for the decimal separator
    [InlineData("ja-JP")]
    public void A_fractional_number_is_stored_with_a_dot_whatever_the_culture(string culture)
    {
        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { 12.5d }, XlsxSerializerOptions.Default));

        Assert.Equal("12.5", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]   // dot as the thousands separator
    [InlineData("ja-JP")]
    public void A_whole_number_is_stored_without_grouping_whatever_the_culture(string culture)
    {
        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { 1234567 }, XlsxSerializerOptions.Default));

        Assert.Equal("1234567", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("ar-SA")]   // a different calendar era and digit shapes
    [InlineData("th-TH")]   // Buddhist calendar: years differ by 543
    public void A_date_is_stored_in_the_western_calendar_whatever_the_culture(string culture)
    {
        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { new DateTime(2023, 1, 2, 3, 4, 5) }, XlsxSerializerOptions.Default));

        Assert.Equal("2023-01-02T03:04:05", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("ar-SA")]
    [InlineData("th-TH")]
    public void A_date_only_value_is_stored_in_the_western_calendar_whatever_the_culture(string culture)
    {
        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { new DateOnly(2023, 1, 2) }, XlsxSerializerOptions.Default));

        Assert.Equal("2023-01-02T00:00:00", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]
    [InlineData("ar-SA")]
    public void A_DateTimeOffset_is_stored_invariantly_unless_a_culture_is_chosen(string culture)
    {
        // These were the last two types whose text still followed the machine: the
        // options culture used to default to null, which means "whatever host this runs on".
        var value = new DateTimeOffset(2000, 1, 1, 10, 30, 0, TimeSpan.FromHours(9));

        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { value }, XlsxSerializerOptions.Default));

        Assert.Equal("01/01/2000 10:30:00 +09:00", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]
    public void A_Complex_is_stored_invariantly_unless_a_culture_is_chosen(string culture)
    {
        var value = new System.Numerics.Complex(1.5, 2);

        var sheet = Culture.Under(culture, null, () => Xlsx.Read(new[] { value }, XlsxSerializerOptions.Default));

        Assert.Equal(value.ToString(System.Globalization.CultureInfo.InvariantCulture), sheet.Texts(0)[0]);
    }

    [Fact]
    public void Choosing_a_culture_localizes_only_the_culture_defined_types()
    {
        // The opt-in exists, and opting in must not leak into the machine-readable cells.
        var options = new XlsxSerializerOptions { CultureInfo = new System.Globalization.CultureInfo("de-DE") };
        var offset = new DateTimeOffset(2000, 1, 1, 10, 30, 0, TimeSpan.FromHours(9));

        var sheet = Xlsx.Read(new object[] { offset, 12.5d }, options);

        Assert.Equal(offset.ToString(new System.Globalization.CultureInfo("de-DE")), sheet.Texts(0)[0]);
        Assert.Equal("12.5", sheet.Texts(1)[0]);
    }
}
