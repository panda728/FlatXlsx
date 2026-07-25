using System.Globalization;
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
    static T UnderCulture<T>(string culture, Func<T> body)
    {
        T result = default!;
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                var info = new CultureInfo(culture);
                CultureInfo.CurrentCulture = info;
                CultureInfo.CurrentUICulture = info;
                result = body();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        });
        thread.Start();
        thread.Join();
        if (failure != null)
            throw new InvalidOperationException($"failed under {culture}", failure);
        return result;
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]   // comma for the decimal separator
    [InlineData("ja-JP")]
    public void A_fractional_number_is_stored_with_a_dot_whatever_the_culture(string culture)
    {
        var sheet = UnderCulture(culture, () => Xlsx.Read(new[] { 12.5d }, XlsxSerializerOptions.Default));

        Assert.Equal("12.5", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]   // dot as the thousands separator
    [InlineData("ja-JP")]
    public void A_whole_number_is_stored_without_grouping_whatever_the_culture(string culture)
    {
        var sheet = UnderCulture(culture, () => Xlsx.Read(new[] { 1234567 }, XlsxSerializerOptions.Default));

        Assert.Equal("1234567", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("ar-SA")]   // a different calendar era and digit shapes
    [InlineData("th-TH")]   // Buddhist calendar: years differ by 543
    public void A_date_is_stored_in_the_western_calendar_whatever_the_culture(string culture)
    {
        var sheet = UnderCulture(culture, () => Xlsx.Read(new[] { new DateTime(2023, 1, 2, 3, 4, 5) }, XlsxSerializerOptions.Default));

        Assert.Equal("2023-01-02T03:04:05", sheet.Texts(0)[0]);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("ar-SA")]
    [InlineData("th-TH")]
    public void A_date_only_value_is_stored_in_the_western_calendar_whatever_the_culture(string culture)
    {
        var sheet = UnderCulture(culture, () => Xlsx.Read(new[] { new DateOnly(2023, 1, 2) }, XlsxSerializerOptions.Default));

        Assert.Equal("2023-01-02T00:00:00", sheet.Texts(0)[0]);
    }
}
