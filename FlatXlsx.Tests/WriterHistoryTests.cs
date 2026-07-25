using CsCheck;
using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises about sequences of writes rather than about any single value.
/// </summary>
/// <remarks>
/// A cell is written once but the writer keeps state across the whole sheet - the shared string
/// table, the column tally, the buffer it flushes at intervals. Nothing about a single call can
/// show that state staying consistent, so these are exercised as histories and checked against a
/// model of what the reader should end up seeing.
/// </remarks>
public class WriterHistoryTests
{
    /// <summary>What a reader must see for a given input, expressed independently of how the
    /// writer encodes it. Characters XML cannot represent are dropped; an empty value is an
    /// empty cell rather than empty text.</summary>
    static string? Expected(string value)
        => value.Length == 0 ? null : new string(value.Where(Representable).ToArray());

    static bool Representable(char c)
        => c >= 0x20 ? c != '\uFFFE' && c != '\uFFFF' : c == '\t' || c == '\n' || c == '\r';

    /// <summary>Text built from the characters that actually cause trouble: markup, the control
    /// codes XML forbids, line breaks and non-ASCII. Bracketed by plain characters so the test
    /// says nothing about leading or trailing whitespace, which is not underwritten.</summary>
    static readonly char[] _awkwardChars =
    {
        'a', 'Z', '9',                    // ordinary text
        '<', '>', '&', '"', '\'',         // markup
        '\0', '\u0001', '\u000B', '\u001F', // control codes XML forbids outright
        '\t', '\n', '\r',                 // the control codes it allows
        '\u3042', '\u00E9',               // beyond ASCII
    };

    static readonly Gen<string> AwkwardText =
        Gen.Int[0, 16]
            .Array[0, 12]
            .Select(indexes => "x" + new string(indexes.Select(i => _awkwardChars[i]).ToArray()) + "y");

    [Fact]
    public void Any_text_reaches_the_reader_intact_apart_from_what_XML_cannot_carry()
    {
        AwkwardText.Array[1, 8].Sample(values =>
        {
            var sheet = Xlsx.Read(values, XlsxSerializerOptions.Default);

            Assert.Equal(values.Length, sheet.Rows.Count);
            for (var i = 0; i < values.Length; i++)
                Assert.Equal(Expected(values[i]), sheet.Texts(i)[0]);
        }, iter: 200);
    }

    [Fact]
    public void Repeated_text_is_stored_once_and_still_read_back_in_every_cell()
    {
        // The shared string table exists to store each distinct value once; the cells then refer
        // to it by position. Reading every cell back is what proves the references stayed lined
        // up with the table as it grew.
        Gen.Int[0, 4].Array[1, 40].Sample(picks =>
        {
            var values = picks.Select(p => $"value-{p}").ToArray();

            var sheet = Xlsx.Read(values, XlsxSerializerOptions.Default);

            for (var i = 0; i < values.Length; i++)
                Assert.Equal(values[i], sheet.Texts(i)[0]);
            Assert.Equal(values.Distinct().Count(), sheet.DistinctStoredStrings);
        }, iter: 200);
    }

    [Fact]
    public void The_number_of_rows_never_changes_what_comes_back()
    {
        // The writer flushes its buffer once it has enough bytes, so the boundary between one
        // flush and the next falls in a different place for every row count.
        Gen.Int[1, 400].Sample(count =>
        {
            var values = Enumerable.Range(0, count).Select(i => $"row-{i}-{new string('p', i % 97)}").ToArray();

            var sheet = Xlsx.Read(values, XlsxSerializerOptions.Default);

            Assert.Equal(count, sheet.Rows.Count);
            Assert.Equal(values, sheet.Rows.Select(r => r[0].Text).ToArray());
        }, iter: 40);
    }

    [Fact]
    public async Task The_asynchronous_writer_produces_the_same_workbook()
    {
        // Claimant: anyone who moves an existing export onto the async API and expects the file
        // their users receive to be unchanged.
        var values = Enumerable.Range(0, 500).Select(i => $"row-{i}").ToArray();
        var options = XlsxSerializerOptions.Default with { HasHeaderRecord = true, AutoFilter = true };

        var synchronous = Workbook.Read(Xlsx.Write(values, options));
        var asynchronous = Workbook.Read(await Xlsx.WriteAsync(values, options));

        Assert.Equal(synchronous.Rows.Count, asynchronous.Rows.Count);
        Assert.Equal(synchronous.AutoFilterRange, asynchronous.AutoFilterRange);
        for (var i = 0; i < synchronous.Rows.Count; i++)
            Assert.Equal(synchronous.Texts(i), asynchronous.Texts(i));
    }

    [Fact]
    public void Columns_stay_aligned_however_many_members_a_row_has()
    {
        // Rows are written from a fixed shape, so every row must present the same number of
        // cells; a row that quietly writes fewer shifts every later column under the wrong
        // heading.
        Gen.Int[1, 60].Sample(count =>
        {
            var rows = Enumerable.Range(0, count)
                .Select(i => (i, $"name-{i}", i % 2 == 0 ? null : $"note-{i}"))
                .ToArray();

            var sheet = Xlsx.Read(rows, XlsxSerializerOptions.Default with { HasHeaderRecord = true });

            Assert.All(sheet.Rows, row => Assert.Equal(3, row.Count));
        }, iter: 30);
    }
}
