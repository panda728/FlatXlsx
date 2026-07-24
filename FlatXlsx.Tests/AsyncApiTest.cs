using System.IO.Compression;
using System.IO.Pipelines;

namespace FlatXlsx.Tests
{
    public class AsyncApiTest
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
        public async Task ToStreamAsync_ProducesSameEntriesAsSync()
        {
            using var syncMs = new MemoryStream();
            XlsxSerializer.ToStream(_rows, syncMs, XlsxSerializerOptions.Default);
            syncMs.Position = 0;
            var syncEntries = ReadEntries(syncMs);

            using var asyncMs = new MemoryStream();
            await XlsxSerializer.ToStreamAsync(_rows, asyncMs, XlsxSerializerOptions.Default);
            asyncMs.Position = 0;
            var asyncEntries = ReadEntries(asyncMs);

            Assert.Equal(syncEntries, asyncEntries);
        }

        [Fact]
        public async Task ToFileAsync_WritesValidWorkbook()
        {
            var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_async_{Guid.NewGuid():N}.xlsx");
            try
            {
                await XlsxSerializer.ToFileAsync(_rows, path, XlsxSerializerOptions.Default);
                using var fs = File.OpenRead(path);
                var entries = ReadEntries(fs);
                Assert.Equal(7, entries.Count);
                Assert.Contains("sheet.xml", entries.Keys);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task ToAsync_PipeWriter_StreamsAndFlushes()
        {
            var pipe = new Pipe();

            var writeTask = Task.Run(async () =>
            {
                await XlsxSerializer.ToAsync(_rows, pipe.Writer, XlsxSerializerOptions.Default);
                await pipe.Writer.CompleteAsync();
            });

            using var received = new MemoryStream();
            await pipe.Reader.CopyToAsync(received);
            await pipe.Reader.CompleteAsync();
            await writeTask;

            received.Position = 0;
            var entries = ReadEntries(received);
            Assert.Equal(7, entries.Count);
            Assert.Contains("strings.xml", entries.Keys);
        }

        [Fact]
        public async Task ToStreamAsync_Cancelled_Throws()
        {
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            using var ms = new MemoryStream();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => XlsxSerializer.ToStreamAsync(_rows, ms, XlsxSerializerOptions.Default, cts.Token));
        }

        [Fact]
        public async Task ToStreamAsync_EmptyRows_WritesNothing()
        {
            using var ms = new MemoryStream();
            await XlsxSerializer.ToStreamAsync(Array.Empty<Portal>(), ms, XlsxSerializerOptions.Default);
            Assert.Equal(0, ms.Length);
        }
    }
}
