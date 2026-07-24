using BenchmarkDotNet.Running;
using BenchmarkSample;
using System.Diagnostics;

var workPath = "work";
if (Directory.Exists(workPath))
{
    var files = Directory.GetFiles(workPath, "*.xlsx");
    for (int i = 0; i < files.Length; i++)
        File.Delete(files[i]);
}
Directory.CreateDirectory(workPath);

#if DEBUG
var ex = new ExportExcel();

ex.N = 1;
ex.GlobalSetup();

var sw = Stopwatch.StartNew();
ex.ExcelApplication();
sw.Stop();
Console.WriteLine($"ExcelApp : {sw.ElapsedMilliseconds:#,##0}ms");

sw.Restart();
ex.ClosedXml();
sw.Stop();
Console.WriteLine($"ClosedXml : {sw.ElapsedMilliseconds:#,##0}ms");

sw.Restart();
ex.FakeExcelSerializer();
Console.WriteLine($"FakeExcelSerializer : {sw.ElapsedMilliseconds:#,##0}ms");
sw.Stop();

#else
BenchmarkRunner.Run<ExportExcel>();
#endif