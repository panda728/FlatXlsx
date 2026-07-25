using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises about asynchronous row sources.
/// </summary>
/// <remarks>
/// Claimant: anyone exporting straight from a query's AsAsyncEnumerable() or a streaming
/// service response. Without these overloads their only options were materializing the whole
/// source or blocking an async server thread per row - both hacks the library exists to avoid.
/// </remarks>
public class AsyncSourceTests
{
    static async IAsyncEnumerable<string> RowsAsync(int count)
    {
        for (var i = 0; i < count; i++)
        {
            await Task.Yield();  // genuinely asynchronous arrival
            yield return $"row-{i}";
        }
    }

    [Fact]
    public async Task An_async_source_produces_the_same_workbook_as_a_sync_one()
    {
        var options = XlsxSerializerOptions.Default with { HasHeaderRow = true, AutoFilter = true };
        var expected = Enumerable.Range(0, 500).Select(i => $"row-{i}").ToArray();

        using var ms = new MemoryStream();
        await XlsxSerializer.ToStreamAsync(RowsAsync(500), ms, options, TestContext.Current.CancellationToken);

        var viaAsync = Workbook.Read(ms.ToArray());
        var viaSync = Workbook.Read(Xlsx.Write(expected, options));
        Assert.Equal(viaSync.Rows.Count, viaAsync.Rows.Count);
        Assert.Equal(viaSync.AutoFilterRange, viaAsync.AutoFilterRange);
        for (var i = 0; i < viaSync.Rows.Count; i++)
            Assert.Equal(viaSync.Texts(i), viaAsync.Texts(i));
    }

    [Fact]
    public async Task Auto_fit_measures_an_async_source_without_a_second_pass()
    {
        var enumerations = 0;
        async IAsyncEnumerable<string> Source()
        {
            enumerations++;
            for (var i = 0; i < 300; i++)
            {
                await Task.Yield();
                yield return $"value-{i}";
            }
        }

        using var ms = new MemoryStream();
        await XlsxSerializer.ToStreamAsync(Source(), ms,
            XlsxSerializerOptions.Default with { AutoFitColumns = true },
            TestContext.Current.CancellationToken);

        var sheet = Workbook.Read(ms.ToArray());
        Assert.Equal(1, enumerations);
        Assert.Equal(300, sheet.Rows.Count);
    }

    [Fact]
    public async Task An_empty_async_source_with_titles_still_delivers_the_header()
    {
        var options = new XlsxSerializerOptions { HeaderTitles = new[] { "Name" } };

        using var ms = new MemoryStream();
        await XlsxSerializer.ToStreamAsync(RowsAsync(0), ms, options, TestContext.Current.CancellationToken);

        var sheet = Workbook.Read(ms.ToArray());
        Assert.Single(sheet.Rows);
        Assert.Equal(new[] { "Name" }, sheet.Texts(0));
    }

    [Fact]
    public async Task An_empty_async_source_without_titles_writes_nothing()
    {
        using var ms = new MemoryStream();

        await XlsxSerializer.ToStreamAsync(RowsAsync(0), ms, null, TestContext.Current.CancellationToken);

        Assert.Equal(0, ms.Length);
    }

    [Fact]
    public async Task Cancelling_mid_stream_stops_the_source()
    {
        using var cts = new CancellationTokenSource();
        var produced = 0;
        async IAsyncEnumerable<string> Source(
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
        {
            for (var i = 0; i < 1_000_000; i++)
            {
                ct.ThrowIfCancellationRequested();
                if (++produced == 100)
                    cts.Cancel();
                await Task.Yield();
                yield return $"row-{i}";
            }
        }

        using var ms = new MemoryStream();
#pragma warning disable xUnit1051 // the test owns this token deliberately: it cancels it mid-stream
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => XlsxSerializer.ToStreamAsync(Source(), ms, null, cts.Token));
#pragma warning restore xUnit1051

        Assert.True(produced < 1_000_000, "cancellation must stop the enumeration");
    }
}
