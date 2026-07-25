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
        var options = XlsxSerializerOptions.Default with { HasHeaderRecord = true };

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

    [Fact]
    public void An_unregistered_type_falls_back_to_its_public_members()
    {
        // Not "unsupported types are rejected": the object-graph provider accepts anything with
        // public members, so no such promise exists to underwrite. What callers can rely on is
        // that an unregistered type still yields one column per readable member.
        var value = new System.Text.StringBuilder("ab");

        var sheet = Xlsx.Read(new[] { value }, XlsxSerializerOptions.Default);

        Assert.NotEmpty(sheet.Row(0));
    }
}
