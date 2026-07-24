using System.Buffers;
using System.IO.Compression;
using System.Xml.Linq;
using FluentAssertions;

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
            ExcelSerializer.ToStream(_rows, ms, ExcelSerializerOptions.Default);
            ms.Position = 0;

            var entries = ReadEntries(ms);
            entries.Keys.Should().BeEquivalentTo(_expectedEntries);

            foreach (var xml in entries.Values)
                XDocument.Parse(xml); // throws if malformed

            var sheet = XDocument.Parse(entries["sheet.xml"]);
            var ns = sheet.Root!.Name.Namespace;
            var sheetRows = sheet.Root.Element(ns + "sheetData")!.Elements(ns + "row").ToArray();
            sheetRows.Should().HaveCount(2);

            var strings = XDocument.Parse(entries["strings.xml"]);
            var shared = strings.Root!.Elements().Select(si => si.Value).ToArray();
            shared.Should().Contain(new[] { "Portal1", "Portal2", "panda728" });
        }

        [Fact]
        public void ToStream_NonSeekableStream_Works()
        {
            using var inner = new MemoryStream();
            using (var forwardOnly = new WriteOnlyStream(inner))
                ExcelSerializer.ToStream(_rows, forwardOnly, ExcelSerializerOptions.Default);

            using var readBack = new MemoryStream(inner.ToArray());
            var entries = ReadEntries(readBack);
            entries.Keys.Should().BeEquivalentTo(_expectedEntries);
        }

        [Fact]
        public void To_BufferWriter_ProducesSameEntriesAsToStream()
        {
            var buffer = new ArrayBufferWriter<byte>();
            ExcelSerializer.To(_rows, buffer, ExcelSerializerOptions.Default);

            using var fromBuffer = new MemoryStream(buffer.WrittenSpan.ToArray());
            var bufferEntries = ReadEntries(fromBuffer);

            using var ms = new MemoryStream();
            ExcelSerializer.ToStream(_rows, ms, ExcelSerializerOptions.Default);
            ms.Position = 0;
            var streamEntries = ReadEntries(ms);

            bufferEntries.Should().Equal(streamEntries);
        }

        [Fact]
        public void ToFile_WritesValidWorkbook_WithoutWorkFolder()
        {
            var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_test_{Guid.NewGuid():N}.xlsx");
            try
            {
                ExcelSerializer.ToFile(_rows, path, ExcelSerializerOptions.Default);
                using var fs = File.OpenRead(path);
                var entries = ReadEntries(fs);
                entries.Keys.Should().BeEquivalentTo(_expectedEntries);
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
            ExcelSerializer.ToStream(Array.Empty<Portal>(), ms, ExcelSerializerOptions.Default);
            ms.Length.Should().Be(0);
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
