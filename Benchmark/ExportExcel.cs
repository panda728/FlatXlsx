using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using FlatXlsx;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace BenchmarkSample
{
    [MarkdownExporterAttribute.GitHub]
    [ShortRunJob]
    [MemoryDiagnoser]
    public class ExportExcel
    {
        readonly List<Row> rows = [];
        readonly byte[] _crlf = new[] { (byte)'\r', (byte)'\n' };
        static readonly string exePath = Assembly.GetEntryAssembly()?.Location ?? "";
        static readonly string workPath = Path.Combine(Path.GetDirectoryName(exePath) ?? "", "work");
        readonly string excelAppFileName = Path.Combine(workPath, $"excelapp-{Guid.NewGuid()}.xlsx");
        readonly string closedXmlFileName = Path.Combine(workPath, $"closedxml-{Guid.NewGuid()}.xlsx");
        readonly string closedXmlOptFileName = Path.Combine(workPath, $"closedxml-opt-{Guid.NewGuid()}.xlsx");
        readonly string fakeExcelFileName = Path.Combine(workPath, $"FakeExcel-{Guid.NewGuid()}.xlsx");

        static readonly string[] COLUMN_TITLES = SampleData.ColumnTitles;
        const int CALCULATION_DEFAULT = 1;

        public ExportExcel()
        {
            if (!Directory.Exists(workPath))
                Directory.CreateDirectory(workPath);
        }

        /// <summary>Rows to export. The key column is unique on every row, as it is in a real
        /// export - see <see cref="ExportCardinality"/> for what changes when it is not.</summary>
        [Params(100, 1_000, 10_000)]
        public int Rows;

        void CleanupFiles()
        {
            if (File.Exists(excelAppFileName))
                File.Delete(excelAppFileName);

            if (File.Exists(closedXmlFileName))
                File.Delete(closedXmlFileName);

            if (File.Exists(closedXmlOptFileName))
                File.Delete(closedXmlOptFileName);

            if (File.Exists(fakeExcelFileName))
                File.Delete(fakeExcelFileName);
        }

        [GlobalSetup]
        public void GlobalSetup()
        {
            CleanupFiles();

            rows.Clear();
            rows.AddRange(SampleData.Build(Rows, distinctKeys: Rows));
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
        /// <summary>Straightforward beginner-style ClosedXML: cell-by-cell writes,
        /// per-cell number formats, and AdjustToContents column sizing.</summary>
        [Benchmark(Baseline = true)]
        public void ClosedXmlNaive()
        {
            using (var book = new XLWorkbook())
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

        /// <summary>Tuned ClosedXML: bulk InsertData, column-level number formats set once,
        /// and no AdjustToContents (its per-cell measurement is ClosedXML's slowest feature).</summary>
        [Benchmark]
        public void ClosedXmlOptimized()
        {
            using (var book = new XLWorkbook())
            {
                var sheet = book.AddWorksheet("ClosedXml");
                for (var i = 0; i < COLUMN_TITLES.Length; i++)
                    sheet.Cell(1, i + 1).Value = COLUMN_TITLES[i];
                for (var c = 4; c <= 10; c++)
                    sheet.Column(c).Style.NumberFormat.SetFormat("@");
                for (var c = 12; c <= 19; c++)
                    sheet.Column(c).Style.NumberFormat.SetFormat("@");
                sheet.Cell(2, 1).InsertData(rows);
                sheet.SheetView.FreezeRows(1);
                book.SaveAs(closedXmlOptFileName);
            }
        }

        int WriteTitle(IXLWorksheet sheet)
        {
            var col = 0;
            foreach (var c in COLUMN_TITLES)
                sheet.Cell(1, ++col).Value = c;

            return 1;
        }

        int WriteRow(IXLWorksheet sheet, int row, Row r)
        {
            int col = 1;
            sheet.Cell(row, col++).SetValue(r.LineNum);
            sheet.Cell(row, col++).SetValue(r.HeaderID);
            sheet.Cell(row, col++).SetValue(r.DetailID);
            sheet.Cell(row, col++).SetValue(r.Header01).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header02).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header03).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header04).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header05).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header06).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Header07).Style.NumberFormat.SetFormat("@");
            sheet.Cell(row, col++).SetValue(r.Data);
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

        #region FlatXlsx
        /// <summary>FlatXlsx with header row and approximate column auto-fit enabled,
        /// so its feature set is comparable to the naive ClosedXML variant.</summary>
        [Benchmark]
        public void FlatXlsx()
        {
            var customOptions = XlsxSerializerOptions.Default with
            {
                HeaderTitles = COLUMN_TITLES,
                AutoFitColumns = true,
            };
            XlsxSerializer.ToFile(rows, fakeExcelFileName, customOptions);
        }
        #endregion
    }
}
