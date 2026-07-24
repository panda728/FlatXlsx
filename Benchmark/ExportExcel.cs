using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using FakeExcelSerializer;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;

namespace BenchmarkSample
{
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class ExportExcel
    {
        readonly List<Row> rows = new();
        readonly byte[] _crlf = new[] { (byte)'\r', (byte)'\n' };
        static readonly string exePath = Assembly.GetEntryAssembly()?.Location ?? "";
        static readonly string workPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "work");
        readonly string excelAppFileName = Path.Combine(workPath, $"excelapp-{Guid.NewGuid()}.xlsx");
        readonly string closedXmlFileName = Path.Combine(workPath, $"closedxml-{Guid.NewGuid()}.xlsx");
        readonly string fakeExcelFileName = Path.Combine(workPath, $"FakeExcel-{Guid.NewGuid()}.xlsx");
        const int CALCULATION_DEFAULT = 1;

        public ExportExcel()
        {
            if (!Directory.Exists(workPath))
                Directory.CreateDirectory(workPath);
        }

        [Params(1, 10, 100)]
        public int N;

        const int HEADER_LEN = 9 + 30;
        const int DETAIL_LEN = 10;
        const int FOOTER_LEN = 39;
        const int DETAIL_COUNT = 10;

        void CleanupFiles()
        {
            if (File.Exists(excelAppFileName))
                File.Delete(excelAppFileName);

            if (File.Exists(closedXmlFileName))
                File.Delete(closedXmlFileName);

            if (File.Exists(fakeExcelFileName))
                File.Delete(fakeExcelFileName);
        }

        [GlobalSetup]
        public void GlobalSetup()
        {
            CleanupFiles();

            var list = new List<Row>();
            var lineNum = 0;
            using var sr = new StreamReader("data01.dat");
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line == null)
                    break;

                if (!int.TryParse(line[..9], out var headerID))
                    throw new ApplicationException("Could not be converted to int.");
                for (int i = 0; i < DETAIL_COUNT; i++)
                {
                    list.Add(new Row
                    {
                        LineNum = lineNum++,
                        HeaderID = headerID,
                        DetailID = i + 1,
                        Data = line.Substring(HEADER_LEN + (DETAIL_LEN * i), DETAIL_LEN),
                        Header01 = line[9..13],
                        Header02 = line[13..20],
                        Header03 = line[20..25],
                        Header04 = line[25..27],
                        Header05 = line[27..30],
                        Header06 = line[30..33],
                        Header07 = line[33..39],
                        Footer01 = line[139..144],
                        Footer02 = line[144..153],
                        Footer03 = line[153..155],
                        Footer04 = line[155..158],
                        Footer05 = line[158..166],
                        Footer06 = line[166..171],
                        Footer07 = line[171..174],
                        Footer08 = line[174..178],
                    });
                }
            }

            rows.Clear();
            for (int i = 0; i < N; i++)
                rows.AddRange(list);
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            CleanupFiles();
        }

        #region Excel.Application
        //[Benchmark] too slow.
        [SupportedOSPlatform("windows")]
        public void ExcelApplication()
        {
            bool screenUpdating = false;
            bool enableEvents = false;
            int calculation = 0;

            dynamic? xlApp = null;
            dynamic? xlWbooks = null;
            dynamic? xlWbook = null;
            dynamic? xlSheets = null;
            dynamic? xlSheet = null;

            try
            {
                Type? objectClassType = Type.GetTypeFromProgID("Excel.Application") ?? throw new NullReferenceException("Excel.Application not found!");
                xlApp = Activator.CreateInstance(objectClassType) ?? throw new NullReferenceException("Excel.Application not found!");
                xlApp.ScreenUpdating = screenUpdating;
                xlApp.EnableEvents = enableEvents;

                xlWbooks = xlApp.Workbooks;
                xlWbook = xlWbooks.Add();
                xlApp.Calculation = calculation;

                xlSheets = xlWbook.Worksheets;
                xlSheet = xlSheets.Item("Sheet1");

                var cols1 = new List<string> { "行番号", "ヘッダID", "明細ID", "明細データ", "ヘッダ1", "ヘッダ2", "ヘッダ3", "ヘッダ4", "ヘッダ5", "ヘッダ6", "ヘッダ7", "フッタ1", "フッタ2", "フッタ3", "フッタ4", "フッタ5", "フッタ6", "フッタ7", "フッタ8" };
                var iCol = 1;
                foreach (var c in cols1)
                    xlSheet.Cells(1, iCol++).Value = c;

                var i = 0;
                for (int iRow = 2; iRow < rows.Count; iRow++)
                {
                    var r = rows[i++];
                    iCol = 1;
                    xlSheet.Cells(iRow, iCol++).Value = r.LineNum;
                    xlSheet.Cells(iRow, iCol++).Value = r.HeaderID;
                    xlSheet.Cells(iRow, iCol++).Value = r.DetailID;
                    xlSheet.Cells(iRow, iCol++).Value = r.Data;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header01;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header02;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header03;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header04;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header05;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header06;
                    xlSheet.Cells(iRow, iCol++).Value = r.Header07;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer01;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer02;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer03;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer04;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer05;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer06;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer07;
                    xlSheet.Cells(iRow, iCol++).Value = r.Footer08;
                }

                xlApp.Calculation = CALCULATION_DEFAULT;
                if (File.Exists(excelAppFileName))
                    File.Delete(excelAppFileName);
                xlWbook.SaveAs(excelAppFileName);
            }
            finally
            {
                if (xlWbook != null)
                {
                    xlWbook.Saved = true;
                }
                if (xlApp != null)
                {
                    xlApp.EnableEvents = true;
                    xlApp.ScreenUpdating = true;
                }

                if (xlSheet != null)
                    Marshal.ReleaseComObject(xlSheet);
                if (xlSheets != null)
                    Marshal.ReleaseComObject(xlSheets);
                if (xlWbook != null)
                    Marshal.ReleaseComObject(xlWbook);
                if (xlWbooks != null)
                    Marshal.ReleaseComObject(xlWbooks);
                if (xlApp != null)
                    xlApp.Quit();
                if (xlSheet != null)
                    Marshal.ReleaseComObject(xlApp);
            }
        }
        #endregion

        #region ClosedXml
        [Benchmark(Baseline = true)]
        public void ClosedXml()
        {
            using (var book = new XLWorkbook(XLEventTracking.Disabled))
            {
                var sheet = book.AddWorksheet("ClosedXml");
                var row = WriteTitle(sheet) + 1;
                foreach (var r in rows)
                    WriteRow(sheet, row++, r);
                sheet.ColumnsUsed().AdjustToContents();
                sheet.SheetView.FreezeRows(1);
                book.SaveAs(closedXmlFileName);
            }
        }

        int WriteTitle(IXLWorksheet sheet)
        {
            var cols1 = new List<string> { "行番号", "ヘッダID", "明細ID", "明細データ", "ヘッダ1", "ヘッダ2", "ヘッダ3", "ヘッダ4", "ヘッダ5", "ヘッダ6", "ヘッダ7", "フッタ1", "フッタ2", "フッタ3", "フッタ4", "フッタ5", "フッタ6", "フッタ7", "フッタ8" };
            var col = 0;
            foreach (var c in cols1)
                sheet.Cell(1, ++col).Value = c;

            return 1;
        }

        int WriteRow(IXLWorksheet sheet, int row, Row r)
        {
            int col = 1;
            sheet.Cell(row, col++).SetValue(r.LineNum);
            sheet.Cell(row, col++).SetValue(r.HeaderID);
            sheet.Cell(row, col++).SetValue(r.DetailID);
            sheet.Cell(row, col++).SetValue(r.Data);
            sheet.Cell(row, col++).SetValue(r.Header01).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header02).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header03).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header04).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header05).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header06).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header07).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer01).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer02).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer03).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer04).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer05).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer06).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer07).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Footer08).Style.NumberFormat.SetFormat("@");
            return col;
        }
        #endregion

        #region FakeExcelSerializer
        [Benchmark]
        public void FakeExcelSerializer()
        {
            var customOptions = ExcelSerializerOptions.Default with
            {
                HeaderTitles = new string[] { "行番号", "ヘッダID", "明細ID", "明細データ", "ヘッダ1", "ヘッダ2", "ヘッダ3", "ヘッダ4", "ヘッダ5", "ヘッダ6", "ヘッダ7", "フッタ1", "フッタ2", "フッタ3", "フッタ4", "フッタ5", "フッタ6", "フッタ7", "フッタ8" },
                HasHeaderRecord = true,
            };
            ExcelSerializer.ToFile(rows, fakeExcelFileName, customOptions);
        }
        #endregion
    }
}
