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
var ex = new ExportExcel { N = 1 };

ex.GlobalSetup();

var sw = new Stopwatch();
if (OperatingSystem.IsWindows())
{
    sw.Start();
    ex.ExcelApplication();
    sw.Stop();
    Console.WriteLine($"ExcelApp : {sw.ElapsedMilliseconds:#,##0}ms");
}

sw.Restart();
ex.ClosedXmlNaive();
sw.Stop();
Console.WriteLine($"ClosedXmlNaive : {sw.ElapsedMilliseconds:#,##0}ms");

sw.Restart();
ex.FlatXlsx();
Console.WriteLine($"FlatXlsx : {sw.ElapsedMilliseconds:#,##0}ms");
sw.Stop();

#else
// Two suites: ExportExcel compares against ClosedXML at sizes ClosedXML can reach, ExportScale
// takes FlatXlsx alone up to a million rows.
// Pass e.g. --filter *ExportScale* to run just one of them.
BenchmarkSwitcher.FromTypes(new[] { typeof(ExportExcel), typeof(ExportScale) }).Run(args);
#endif