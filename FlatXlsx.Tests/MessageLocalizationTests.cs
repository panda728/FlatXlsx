using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises about the language a diagnostic message arrives in.
/// </summary>
/// <remarks>
/// Claimant: the developer reading the failure, and the support desk reading their screenshot.
/// The guarantee is narrow on purpose - the language follows the UI culture, an application that
/// ships no translation still gets a usable English message, and the numbers in the message stay
/// the same either way so they remain searchable.
/// </remarks>
public class MessageLocalizationTests
{
    static void TooLongForACell() =>
        Xlsx.Write(new[] { new string('a', 32_768) }, XlsxSerializerOptions.Default);

    [Fact]
    public void A_message_is_English_when_the_user_interface_is_English()
    {
        var message = Culture.MessageFrom("en-US", TooLongForACell);

        Assert.Contains("cannot hold more than", message);
    }

    [Fact]
    public void A_message_is_Japanese_when_the_user_interface_is_Japanese()
    {
        var message = Culture.MessageFrom("ja-JP", TooLongForACell);

        Assert.Contains("セルに格納できる文字数", message);
    }

    [Fact]
    public void An_untranslated_language_falls_back_to_English_rather_than_failing()
    {
        // Only some languages are translated, and an application is free to deploy none of the
        // satellite assemblies at all. Neither case may turn a useful message into a blank one.
        var message = Culture.MessageFrom("fr-FR", TooLongForACell);

        Assert.Contains("cannot hold more than", message);
    }

    [Fact]
    public void A_regional_variant_uses_its_parent_language()
    {
        var message = Culture.MessageFrom("ja-JP-u-ca-japanese", TooLongForACell);

        Assert.Contains("セルに格納できる文字数", message);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("ja-JP")]
    public void The_facts_in_a_message_survive_translation(string uiCulture)
    {
        // Whatever the wording, the limit and the offending length are what someone acts on -
        // and what they search the logs for.
        var message = Culture.MessageFrom(uiCulture, TooLongForACell);

        Assert.Contains("32767", message);
        Assert.Contains("32768", message);
    }

    [Fact]
    public void The_language_of_messages_does_not_change_the_workbook()
    {
        // Message language and data formatting are separate settings; a support engineer
        // switching their locale must not change the file the application produces.
        var rows = new[] { 12.5 };

        var underEnglish = Culture.Under(null, "en-US", () => Xlsx.Write(rows, XlsxSerializerOptions.Default));
        var underJapanese = Culture.Under(null, "ja-JP", () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Equal(
            Workbook.Read(underEnglish).Texts(0),
            Workbook.Read(underJapanese).Texts(0));
    }
}
