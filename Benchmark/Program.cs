using BenchmarkDotNet.Running;
using BenchmarkSample;

var workPath = "work";
if (Directory.Exists(workPath))
{
    var files = Directory.GetFiles(workPath, "*.xlsx");
    for (int i = 0; i < files.Length; i++)
        File.Delete(files[i]);
}
Directory.CreateDirectory(workPath);

#if DEBUG
var ex = new ExportExcel { N = 1 };

ex.GlobalSetup();

var sw = Stopwatch.StartNew();
ex.ExcelApplication();
sw.Stop();
Console.WriteLine($"ExcelApp : {sw.ElapsedMilliseconds:#,##0}ms");

sw.Restart();
ex.ClosedXmlNaive();
sw.Stop();
Console.WriteLine($"ClosedXmlNaive : {sw.ElapsedMilliseconds:#,##0}ms");

sw.Restart();
ex.FlatXlsx();
Console.WriteLine($"FlatXlsx : {sw.ElapsedMilliseconds:#,##0}ms");
sw.Stop();

#else
BenchmarkRunner.Run<ExportExcel>(null, args);
#endif