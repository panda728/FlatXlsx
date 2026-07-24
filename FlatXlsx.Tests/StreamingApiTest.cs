using System.Buffers;
using System.IO.Compression;
using System.Xml.Linq;

namespace FlatXlsx.Tests
{
    public class StreamingApiTest
    {
        class Portal
        {
            public string Name { get; set; } = "";
            public string Owner { get; set; } = "";
            public int Level { get; set; }
        }

        static readonly Portal[] _rows = new[]
        {
            new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
            new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
        };

        static readonly string[] _expectedEntries = new[]
        {
            "[Content_Types].xml", "_rels/.rels", "book.xml", "_rels/book.xml.rels",
            "styles.xml", "sheet.xml", "strings.xml",
        };

        static Dictionary<string, string> ReadEntries(Stream zipStream)
        {
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);
            var result = new Dictionary<string, string>();
            foreach (var entry in archive.Entries)
            {
                using var reader = new StreamReader(entry.Open());
                result[entry.FullName] = reader.ReadToEnd();
            }
            return result;
        }

        [Fact]
        public void ToStream_WritesValidWorkbook()
        {
            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(_rows, ms, XlsxSerializerOptions.Default);
            ms.Position = 0;

            var entries = ReadEntries(ms);
            Assert.Equivalent(_expectedEntries, entries.Keys);

            foreach (var xml in entries.Values)
                XDocument.Parse(xml); // throws if malformed

            var sheet = XDocument.Parse(entries["sheet.xml"]);
            var ns = sheet.Root!.Name.Namespace;
            var sheetRows = sheet.Root.Element(ns + "sheetData")!.Elements(ns + "row").ToArray();
            Assert.Equal(2, sheetRows.Length);

            var strings = XDocument.Parse(entries["strings.xml"]);
            var shared = strings.Root!.Elements().Select(si => si.Value).ToArray();
            Assert.Contains("Portal1", shared);
            Assert.Contains("Portal2", shared);
            Assert.Contains("panda728", shared);
        }

        [Fact]
        public void ToStream_NonSeekableStream_Works()
        {
            using var inner = new MemoryStream();
            using (var forwardOnly = new WriteOnlyStream(inner))
                XlsxSerializer.ToStream(_rows, forwardOnly, XlsxSerializerOptions.Default);

            using var readBack = new MemoryStream(inner.ToArray());
            var entries = ReadEntries(readBack);
            Assert.Equivalent(_expectedEntries, entries.Keys);
        }

        [Fact]
        public void To_BufferWriter_ProducesSameEntriesAsToStream()
        {
            var buffer = new ArrayBufferWriter<byte>();
            XlsxSerializer.To(_rows, buffer, XlsxSerializerOptions.Default);

            using var fromBuffer = new MemoryStream(buffer.WrittenSpan.ToArray());
            var bufferEntries = ReadEntries(fromBuffer);

            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(_rows, ms, XlsxSerializerOptions.Default);
            ms.Position = 0;
            var streamEntries = ReadEntries(ms);

            Assert.Equal(streamEntries, bufferEntries);
        }

        [Fact]
        public void ToFile_WritesValidWorkbook_WithoutWorkFolder()
        {
            var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_test_{Guid.NewGuid():N}.xlsx");
            try
            {
                XlsxSerializer.ToFile(_rows, path, XlsxSerializerOptions.Default);
                using var fs = File.OpenRead(path);
                var entries = ReadEntries(fs);
                Assert.Equivalent(_expectedEntries, entries.Keys);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ToStream_EmptyRows_WritesNothing()
        {
            using var ms = new MemoryStream();
            XlsxSerializer.ToStream(Array.Empty<Portal>(), ms, XlsxSerializerOptions.Default);
            Assert.Equal(0, ms.Length);
        }

        /// <summary>Simulates a network stream: write-only, non-seekable.</summary>
        sealed class WriteOnlyStream : Stream
        {
            readonly Stream _inner;
            public WriteOnlyStream(Stream inner) => _inner = inner;
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }
            public override void Flush() => _inner.Flush();
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }
    }
}
