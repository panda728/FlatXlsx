using BenchmarkDotNet.Attributes;
using FlatXlsx;
using System.Reflection;

namespace BenchmarkSample
{
    /// <summary>
    /// What repeated values were hiding: how the export behaves when every row carries a value
    /// no other row has.
    /// </summary>
    /// <remarks>
    /// <see cref="ExportScale"/> repeats a fixed block, which holds the number of *distinct*
    /// strings constant and so measures the streaming cost alone. That is a real property worth
    /// isolating, but it is not what a real export looks like: an order number or a timestamp is
    /// different on every row, and the shared-string table grows with exactly those.
    ///
    /// Three variants at the same row count:
    /// <list type="bullet">
    /// <item>Repeated - the flattering case, for reference.</item>
    /// <item>Distinct - one unique value per row, default settings. This is the honest cost.</item>
    /// <item>DistinctCapped - the same data with <see cref="XlsxSerializerOptions.MaxSharedStrings"/>
    /// lowered, so unique values are written inline instead of interned.</item>
    /// </list>
    ///
    /// Interning only pays when values repeat. For a value no other row shares, the workbook ends
    /// up storing both the string and the index that points at it, so capping the table costs
    /// nothing in file size - measured at a million rows, the capped file is in fact smaller.
    /// </remarks>
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class ExportCardinality
    {
        static readonly string exePath = Assembly.GetEntryAssembly()?.Location ?? "";
        static readonly string workPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "work");
        readonly string fileName = Path.Combine(workPath, $"cardinality-{Guid.NewGuid()}.xlsx");

        List<Row> repeated = [];
        List<Row> distinct = [];
        XlsxSerializerOptions options = XlsxSerializerOptions.Default;
        XlsxSerializerOptions cappedOptions = XlsxSerializerOptions.Default;

        [Params(100_000, 1_000_000)]
        public int Rows;

        [GlobalSetup]
        public void GlobalSetup()
        {
            Directory.CreateDirectory(workPath);

            var block = SampleData.LoadBlock();

            repeated = new List<Row>(Rows);
            distinct = new List<Row>(Rows);
            for (var i = 0; i < Rows; i++)
            {
                var source = block[i % block.Count];
                repeated.Add(source);

                // One column differs on every row - the shape of an order number or a key.
                // Everything else still repeats, as it does in real data.
                distinct.Add(new Row
                {
                    LineNum = source.LineNum,
                    HeaderID = source.HeaderID,
                    DetailID = source.DetailID,
                    Data = $"D{i:D9}",
                    Header01 = source.Header01,
                    Header02 = source.Header02,
                    Header03 = source.Header03,
                    Header04 = source.Header04,
                    Header05 = source.Header05,
                    Header06 = source.Header06,
                    Header07 = source.Header07,
                    Footer01 = source.Footer01,
                    Footer02 = source.Footer02,
                    Footer03 = source.Footer03,
                    Footer04 = source.Footer04,
                    Footer05 = source.Footer05,
                    Footer06 = source.Footer06,
                    Footer07 = source.Footer07,
                    Footer08 = source.Footer08,
                });
            }

            options = XlsxSerializerOptions.Default with
            {
                HeaderTitles = SampleData.ColumnTitles,
                AutoFitColumns = true,
            };
            // Enough for the columns that really do repeat; the unique column overflows and is
            // written inline.
            cappedOptions = options with { MaxSharedStrings = 1_000 };
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }

        [Benchmark(Baseline = true)]
        public void Repeated() => XlsxSerializer.ToFile(repeated, fileName, options);

        [Benchmark]
        public void Distinct() => XlsxSerializer.ToFile(distinct, fileName, options);

        [Benchmark]
        public void DistinctCapped() => XlsxSerializer.ToFile(distinct, fileName, cappedOptions);
    }
}
