using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises about what a failure hands the person who must answer for it.
/// </summary>
/// <remarks>
/// The acceptance test is borrowed from the accountability guideline: given only what the
/// failure returned, can the operator fix the data without re-deriving anything? A data error
/// therefore carries its location and its numbers as structured properties - not only as
/// (localized) prose - and the validation scan collects the whole list in one pass, so the
/// data owner receives a work list instead of one surprise per run.
/// </remarks>
public class AnswerabilityTests
{
    [Fact]
    public void A_data_error_names_the_row_and_column_it_happened_at()
    {
        var rows = new[] { "ok", new string('a', 32_768), "also ok" };

        var ex = Assert.Throws<XlsxDataException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Equal(XlsxDataErrorKind.CellTooLong, ex.Kind);
        Assert.Equal(2, ex.Row);
        Assert.Equal(1, ex.Column);
        Assert.Equal(32_767, ex.Limit);
        Assert.Equal(32_768, ex.Actual);
        Assert.Contains("2", ex.Message);   // the location also reaches the prose
    }

    [Fact]
    public void Row_numbers_are_physical_sheet_rows_header_included()
    {
        // The operator opens the sheet and goes to the numbered row; the number must be the
        // one the spreadsheet shows, not a data-relative index.
        var options = new XlsxSerializerOptions { HeaderTitles = new[] { "Value" } };
        var rows = new[] { new string('a', 32_768) };

        var ex = Assert.Throws<XlsxDataException>(() => Xlsx.Write(rows, options));

        Assert.Equal(2, ex.Row);   // row 1 is the header
    }

    [Fact]
    public void A_row_too_wide_names_the_row_and_the_limit()
    {
        var rows = new[] { Enumerable.Range(0, 16_385).ToArray() };

        var ex = Assert.Throws<XlsxDataException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Equal(XlsxDataErrorKind.TooManyColumns, ex.Kind);
        Assert.Equal(1, ex.Row);
        Assert.Equal(16_384, ex.Limit);
    }

    class Node
    {
        public string Name { get; set; } = "";
        public Node? Child { get; set; }
    }

    [Fact]
    public void A_circular_reference_names_the_row_it_was_found_in()
    {
        var fine = new Node { Name = "fine" };
        var loop = new Node { Name = "loop" };
        loop.Child = loop;

        var ex = Assert.Throws<XlsxDataException>(
            () => Xlsx.Write(new[] { fine, loop }, XlsxSerializerOptions.Default));

        Assert.Equal(XlsxDataErrorKind.MaxDepthReached, ex.Kind);
        Assert.Equal(2, ex.Row);
    }

    [Fact]
    public void Validation_collects_every_problem_in_one_pass()
    {
        // The export meets errors one per run; the scan is where the data owner gets the whole
        // work list, each entry with its own location.
        var rows = new[] { "ok", new string('a', 40_000), "ok", new string('b', 33_000) };

        var errors = XlsxSerializer.Validate(rows);

        Assert.Equal(2, errors.Count);
        Assert.Equal(2, errors[0].Row);
        Assert.Equal(40_000, errors[0].Actual);
        Assert.Equal(4, errors[1].Row);
        Assert.Equal(33_000, errors[1].Actual);
        Assert.All(errors, e => Assert.Equal(XlsxDataErrorKind.CellTooLong, e.Kind));
    }

    [Fact]
    public void Validation_continues_past_a_row_it_cannot_walk()
    {
        // A circular row aborts that row only; the long value in the next row is still found.
        var loop = new Node { Name = "loop" };
        loop.Child = loop;
        var tooLong = new Node { Name = new string('a', 32_768) };

        var errors = XlsxSerializer.Validate(new[] { loop, tooLong });

        Assert.Equal(2, errors.Count);
        Assert.Equal(XlsxDataErrorKind.MaxDepthReached, errors[0].Kind);
        Assert.Equal(1, errors[0].Row);
        Assert.Equal(XlsxDataErrorKind.CellTooLong, errors[1].Kind);
        Assert.Equal(2, errors[1].Row);
    }

    [Fact]
    public void Clean_data_validates_to_an_empty_list()
    {
        var errors = XlsxSerializer.Validate(new[] { "a", "b", "c" });

        Assert.Empty(errors);
    }

    [Fact]
    public async Task An_async_source_validates_the_same_way()
    {
        static async IAsyncEnumerable<string> Source()
        {
            yield return "ok";
            await Task.Yield();
            yield return new string('a', 32_768);
        }

        var errors = await XlsxSerializer.ValidateAsync(Source(),
            cancellationToken: TestContext.Current.CancellationToken);

        Assert.Single(errors);
        Assert.Equal(2, errors[0].Row);
    }

    [Fact]
    public void Configuration_mistakes_still_throw_their_addressee_is_the_developer()
    {
        // The scan is for data problems; a bad sheet name is fixed in code, and deferring it
        // into the error list would misfile a developer bug as a data chore.
        var options = new XlsxSerializerOptions { SheetName = "a:b" };
        using var ms = new MemoryStream();

        var ex = Assert.Throws<InvalidOperationException>(
            () => XlsxSerializer.ToStream(new[] { "x" }, ms, options));

        Assert.IsNotType<XlsxDataException>(ex);
    }

    [Fact]
    public void A_platform_type_with_no_serializer_is_refused_by_name_instead_of_guessed_at()
    {
        // Reflecting System.Range into Start/End columns would produce a plausible-looking
        // workbook nobody asked for; the layout of .NET's own types is not the library's guess
        // to make. The refusal names the type so the fix has an address.
        System.Range[] rows = { 1..5 };

        var ex = Assert.Throws<InvalidOperationException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.IsNotType<XlsxDataException>(ex);   // addressee: the developer, not the data owner
        Assert.Contains("System.Range", ex.Message);
    }

    [Fact]
    public void A_platform_typed_member_is_refused_the_same_way()
    {
        var rows = new[] { new { Name = "a", Span = new System.Range(1, 3) } };

        var ex = Assert.Throws<InvalidOperationException>(
            () => Xlsx.Write(rows, XlsxSerializerOptions.Default));

        Assert.Contains("System.Range", ex.Message);
    }

    [Fact]
    public void Every_entry_point_warns_a_trimmed_or_AOT_build_before_it_ships()
    {
        // Claimant: the team publishing with PublishTrimmed or PublishAot. Working out a row
        // type's columns reflects over its members, which those builds can break - and the break
        // would otherwise surface in their customer's deployment, not in the build that caused
        // it. Every entry point therefore carries the annotation that moves it to compile time.
        var entryPoints = typeof(XlsxSerializer)
            .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
            .Where(m => m.Name.StartsWith("To") || m.Name.StartsWith("Validate"))
            .ToArray();

        Assert.NotEmpty(entryPoints);
        Assert.All(entryPoints, method =>
        {
            Assert.NotNull(method.GetCustomAttribute<RequiresUnreferencedCodeAttribute>());
            Assert.NotNull(method.GetCustomAttribute<RequiresDynamicCodeAttribute>());
        });
    }

    [Fact]
    public void The_remedy_the_warning_names_is_reachable_without_reflection()
    {
        // The warning tells the caller to register a serializer for every row type. That advice
        // is only honest if the registered path resolves by type test rather than by reflecting
        // over interfaces, so the value it produces is asserted here alongside the claim.
        var options = new XlsxSerializerOptions { CustomSerializers = [new UpperCaseSerializer()] };

        var sheet = Xlsx.Read(new[] { "abc" }, options);

        Assert.Equal("ABC", sheet.Texts(0)[0]);
    }

    sealed class UpperCaseSerializer : IXlsxSerializer<string>
    {
        public void WriteTitle(XlsxCellWriter writer, string value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, string value, XlsxSerializerOptions options)
            => writer.Write(value?.ToUpperInvariant());
    }

    [Fact]
    public void The_scan_treats_a_missing_serializer_as_a_developer_problem()
    {
        // Validate collects data problems; a type the library cannot place is fixed in code,
        // so it surfaces immediately rather than being filed on the data owner's work list.
        System.Range[] rows = { 1..5 };

        var ex = Assert.Throws<InvalidOperationException>(() => XlsxSerializer.Validate(rows));

        Assert.IsNotType<XlsxDataException>(ex);
    }
}
