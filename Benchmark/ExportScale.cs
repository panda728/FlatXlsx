using BenchmarkDotNet.Attributes;
using FlatXlsx;
using System.Reflection;

namespace BenchmarkSample
{
    /// <summary>
    /// What one export costs as the row count grows, up to the neighbourhood of Excel's own
    /// 1,048,576-row limit.
    /// </summary>
    /// <remarks>
    /// ClosedXML is not measured here: at the top of this range a library that holds the whole
    /// workbook in memory needs several gigabytes, which measures the machine rather than the
    /// library.
    ///
    /// The baseline is the realistic case - a key that is unique on every row, the way an order
    /// number is - because every distinct string goes into the shared-string table and that is
    /// what a caller has to budget for. The other two say what can be done about it:
    /// <list type="bullet">
    /// <item>DistinctCapped - same data, <see cref="XlsxSerializerOptions.MaxSharedStrings"/>
    /// lowered so values that no longer fit are written into the cell instead.</item>
    /// <item>RepeatedKey - the same export where the key takes only 100 distinct values. Not a
    /// realistic export; it isolates the cost of streaming rows out from the cost of remembering
    /// distinct strings.</item>
    /// </list>
    /// </remarks>
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class ExportScale
    {
        static readonly string exePath = Assembly.GetEntryAssembly()?.Location ?? "";
        static readonly string workPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "work");
        readonly string fileName = Path.Combine(workPath, $"scale-{Guid.NewGuid()}.xlsx");

        List<Row> distinct = [];
        List<Row> repeated = [];
        XlsxSerializerOptions options = XlsxSerializerOptions.Default;
        XlsxSerializerOptions cappedOptions = XlsxSerializerOptions.Default;

        [Params(100, 1_000, 10_000, 100_000, 1_000_000)]
        public int Rows;

        [GlobalSetup]
        public void GlobalSetup()
        {
            Directory.CreateDirectory(workPath);

            // Identical rows apart from the one knob these benchmarks exist to turn.
            distinct = SampleData.Build(Rows, distinctKeys: Rows);
            repeated = SampleData.Build(Rows, distinctKeys: 100);

            options = XlsxSerializerOptions.Default with
            {
                HeaderTitles = SampleData.ColumnTitles,
                AutoFitColumns = true,
            };
            // Room for the columns that really do repeat; the unique key overflows and goes inline.
            cappedOptions = options with { MaxSharedStrings = 1_000 };
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }

        [Benchmark(Baseline = true)]
        public void Distinct() => XlsxSerializer.ToFile(distinct, fileName, options);

        [Benchmark]
        public void DistinctCapped() => XlsxSerializer.ToFile(distinct, fileName, cappedOptions);

        [Benchmark]
        public void RepeatedKey() => XlsxSerializer.ToFile(repeated, fileName, options);
    }
}
