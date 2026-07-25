using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// The contract every <see cref="IXlsxSerializer{T}"/> owes, applied to every supported type.
/// </summary>
/// <remarks>
/// Claimant: anyone who opens the produced workbook. These clauses are what they are entitled
/// to regardless of which serializer handled the value, so they live with the abstraction
/// rather than with any one implementation.
/// </remarks>
public class SerializerContractTests
{
    public static TheoryData<string> SupportedTypes => SerializerCase.Names;

    [Theory]
    [MemberData(nameof(SupportedTypes))]
    public void A_value_reaches_the_reader_as_its_own_text(string type)
    {
        var subject = SerializerCase.All[type];

        var sheet = Workbook.Read(subject.Write(XlsxSerializerOptions.Default));

        Assert.Equal(subject.ExpectedCells, sheet.Texts(0));
    }

    [Theory]
    [MemberData(nameof(SupportedTypes))]
    public void A_header_names_exactly_as_many_columns_as_the_value_fills(string type)
    {
        // Without this the header row and the data rows drift apart and every column below the
        // drift is labelled with someone else's name - wrong data under a plausible heading,
        // which is worse than an obvious failure.
        var subject = SerializerCase.All[type];
        var options = XlsxSerializerOptions.Default with { HasHeaderRow = true };

        var sheet = Workbook.Read(subject.Write(options));

        Assert.Equal(subject.ExpectedCells.Length, sheet.Row(0).Count);
        Assert.Equal(subject.ExpectedCells, sheet.Texts(1));
    }

    [Theory]
    [MemberData(nameof(SupportedTypes))]
    public void A_workbook_is_well_formed_whatever_the_value(string type)
    {
        var subject = SerializerCase.All[type];

        Workbook.AssertEveryPartIsWellFormed(subject.Write(XlsxSerializerOptions.Default));
    }

    class UnregisteredShape
    {
        public string A { get; set; } = "";
        public int B { get; set; }
    }

    [Fact]
    public void An_unregistered_application_type_falls_back_to_its_public_members()
    {
        // The fallback has a boundary: an application-defined type yields one column per
        // readable instance member, while the platform's own types are refused by name
        // instead (see AnswerabilityTests) - a reflected layout for those is never what
        // the caller meant.
        var value = new UnregisteredShape { A = "ab", B = 7 };

        var sheet = Xlsx.Read(new[] { value }, XlsxSerializerOptions.Default);

        Assert.Equal(new[] { "ab", "7" }, sheet.Texts(0));
    }
}
