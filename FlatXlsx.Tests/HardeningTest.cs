using System.IO.Compression;
using System.Xml.Linq;

namespace FlatXlsx.Tests
{
    /// <summary>
    /// Untrusted values must never be able to break the workbook: no markup injection,
    /// no characters that make the XML unparsable, and no unbounded expansion.
    /// </summary>
    public class HardeningTest
    {
        static string SheetXml<T>(IEnumerable<T> rows, XlsxSerializerOptions options)
        {
            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(rows, ms, options);
            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
            using var reader = new StreamReader(archive.GetEntry("sheet.xml")!.Open());
            return reader.ReadToEnd();
        }

        static string StringsXml<T>(IEnumerable<T> rows, XlsxSerializerOptions options)
        {
            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(rows, ms, options);
            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
            using var reader = new StreamReader(archive.GetEntry("strings.xml")!.Open());
            return reader.ReadToEnd();
        }

        [Fact]
        public void MarkupInValues_IsEscaped_NotInjected()
        {
            var rows = new[] { "</t></is></c><f>1+1</f>", "<script>&amp;" };

            var strings = StringsXml(rows, XlsxSerializerOptions.Default);

            // Parses, and the payload survives as text rather than becoming elements.
            var doc = XDocument.Parse(strings);
            var values = doc.Root!.Elements().Select(si => si.Value).ToArray();
            Assert.Equal(rows, values);
            Assert.DoesNotContain("<f>", strings);
        }

        [Fact]
        public void ControlCharacters_AreDropped_SoTheFileStaysParsable()
        {
            // NUL and friends are illegal in XML 1.0 even when escaped, and turn up routinely
            // in data pulled out of databases.
            var rows = new[] { "a\u0000b\u0001c\u000Bd", "keep\ttab\nand\r\nnewline" };

            var strings = StringsXml(rows, XlsxSerializerOptions.Default);

            var doc = XDocument.Parse(strings);
            var values = doc.Root!.Elements().Select(si => si.Value).ToArray();
            Assert.Equal("abcd", values[0]);
            Assert.Contains("\t", values[1]);
        }

        [Fact]
        public void MarkupInHeaderTitles_IsEscaped()
        {
            var options = XlsxSerializerOptions.Default with
            {
                HasHeaderRecord = true,
                HeaderTitles = new[] { "<b>Name</b>", "A & B" },
            };

            var strings = StringsXml(new[] { "row1" }, options);

            var doc = XDocument.Parse(strings);
            Assert.Contains("<b>Name</b>", doc.Root!.Elements().Select(si => si.Value));
        }

        [Fact]
        public void SharedStringCap_FallsBackToInlineStrings()
        {
            var options = XlsxSerializerOptions.Default with { MaxSharedStrings = 2 };
            var rows = new[] { "a", "b", "c<x>", "d" };

            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(rows, ms, options);
            ms.Position = 0;
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

            using var sheetReader = new StreamReader(archive.GetEntry("sheet.xml")!.Open());
            var sheet = sheetReader.ReadToEnd();
            using var stringsReader = new StreamReader(archive.GetEntry("strings.xml")!.Open());
            var strings = stringsReader.ReadToEnd();

            // Only the first two distinct values reached the table; the rest went inline,
            // still escaped.
            Assert.Equal(2, XDocument.Parse(strings).Root!.Elements().Count());
            Assert.Contains("inlineStr", sheet);
            Assert.Contains("c&lt;x&gt;", sheet);
            XDocument.Parse(sheet);
        }

        [Fact]
        public void CellLongerThanExcelAllows_Throws()
        {
            var rows = new[] { new string('a', 32_768) };

            var ex = Assert.Throws<InvalidOperationException>(
                () => SheetXml(rows, XlsxSerializerOptions.Default));
            Assert.Contains("32767", ex.Message);
        }

        [Fact]
        public void CellAtExcelLimit_IsAccepted()
        {
            var rows = new[] { new string('a', 32_767) };

            var sheet = SheetXml(rows, XlsxSerializerOptions.Default);

            XDocument.Parse(sheet);
        }

        class Node
        {
            public string Name { get; set; } = "";
            public Node? Child { get; set; }
        }

        [Fact]
        public void CircularReference_StopsAtMaxDepth()
        {
            var a = new Node { Name = "a" };
            a.Child = a;

            var ex = Assert.Throws<InvalidOperationException>(
                () => SheetXml(new[] { a }, XlsxSerializerOptions.Default));
            Assert.Contains("max depth", ex.Message);
        }

        [Fact]
        public void WideCollection_StopsAtColumnLimit()
        {
            // A nested collection expands across columns; beyond Excel's limit the file would
            // be unopenable, so this has to fail loudly instead.
            var rows = new[] { Enumerable.Range(0, 16_385).ToArray() };

            var ex = Assert.Throws<InvalidOperationException>(
                () => SheetXml(rows, XlsxSerializerOptions.Default));
            Assert.Contains("16384", ex.Message);
        }

        [Fact]
        public void LazySource_IsEnumeratedOnce_WhenAutoFitIsOff()
        {
            var enumerations = 0;
            IEnumerable<string> Source()
            {
                enumerations++;
                yield return "a";
                yield return "b";
            }

            var options = XlsxSerializerOptions.Default with { AutoFilter = true, HasHeaderRecord = true };
            SheetXml(Source(), options);

            // Any() + First() + the write loop + Count() used to add up to four passes over a
            // sequence that may be backed by a query or a reader.
            Assert.Equal(1, enumerations);
        }

        [Fact]
        public void AutoFilterRange_CoversWhatWasActuallyWritten()
        {
            var options = XlsxSerializerOptions.Default with { AutoFilter = true, HasHeaderRecord = true };

            var sheet = SheetXml(new[] { "a", "b", "c" }, options);

            // 1 header + 3 data rows, one column.
            Assert.Contains("<autoFilter ref=\"A1:A4\"/>", sheet);
        }
    }
}
