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
    /// The rows repeat a 100-row block, so the number of *distinct* strings stays constant while
    /// the row count grows. That is deliberate: it isolates the streaming cost from the one part
    /// that is expected to grow with the data, the shared-string table.
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

            // The same block instances are referenced repeatedly; building the list must not
            // itself become the thing being measured.
            var block = SampleData.LoadBlock();
            rows = new List<Row>(Rows);
            while (rows.Count < Rows)
            {
                var take = Math.Min(block.Count, Rows - rows.Count);
                for (var i = 0; i < take; i++)
                    rows.Add(block[i]);
            }

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
