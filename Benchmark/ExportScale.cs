using BenchmarkDotNet.Attributes;
using FlatXlsx;
using System.Reflection;

namespace BenchmarkSample
{
    /// <summary>
    /// How a single export scales with the number of rows, up to the neighbourhood of Excel's
    /// own 1,048,576-row limit.
    /// </summary>
    /// <remarks>
    /// Only FlatXlsx is measured here. The question is the shape of the curve rather than a
    /// comparison, and a library that holds the whole workbook in memory needs several gigabytes
    /// at the top of this range - which measures the machine's memory pressure, not the library.
    ///
    /// The key column takes only 100 distinct values however many rows there are, so the
    /// shared-string table stays the same size while the row count grows. That is deliberate -
    /// it isolates the streaming cost from the one part that is expected to grow with the data -
    /// and it is also the flattering case. <see cref="ExportCardinality"/> measures the other one.
    /// </remarks>
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class ExportScale
    {
        static readonly string exePath = Assembly.GetEntryAssembly()?.Location ?? "";
        static readonly string workPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "work");
        readonly string fileName = Path.Combine(workPath, $"scale-{Guid.NewGuid()}.xlsx");

        List<Row> rows = [];
        XlsxSerializerOptions options = XlsxSerializerOptions.Default;

        [Params(100, 10_000, 100_000, 1_000_000)]
        public int Rows;

        [GlobalSetup]
        public void GlobalSetup()
        {
            Directory.CreateDirectory(workPath);

            rows = SampleData.Build(Rows, distinctKeys: 100);

            options = XlsxSerializerOptions.Default with
            {
                HeaderTitles = SampleData.ColumnTitles,
                AutoFitColumns = true,
            };
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }

        [Benchmark]
        public void FlatXlsx() => XlsxSerializer.ToFile(rows, fileName, options);
    }
}
