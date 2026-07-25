using System.Buffers;
using System.IO.Pipelines;
using FlatXlsx.Tests.Support;

namespace FlatXlsx.Tests;

/// <summary>
/// Promises about where the workbook can be sent, as opposed to what it contains.
/// </summary>
/// <remarks>
/// Claimant: the calling application. It picks a destination - a file, a response body, a
/// network stream - and needs the writer to behave itself there: not to seek, not to close what
/// it does not own, and not to leave a file behind when there was nothing to write.
/// </remarks>
public class OutputTargetTests
{
    static readonly string[] _rows = { "Portal1", "Portal2" };

    [Fact]
    public void A_stream_that_cannot_seek_is_still_a_valid_destination()
    {
        // Response bodies and sockets cannot seek. A writer that rewinds to patch a header
        // would work in tests and fail in production.
        using var inner = new MemoryStream();
        using (var forwardOnly = new ForwardOnlyStream(inner))
            XlsxSerializer.ToStream(_rows, forwardOnly, XlsxSerializerOptions.Default);

        var sheet = Workbook.Read(inner.ToArray());

        Assert.Equal(_rows, sheet.Rows.Select(r => r[0].Text).ToArray());
    }

    [Fact]
    public void The_callers_stream_is_left_open()
    {
        using var ms = new MemoryStream();

        XlsxSerializer.ToStream(_rows, ms, XlsxSerializerOptions.Default);

        Assert.True(ms.CanWrite, "the caller still owns the stream and may keep writing to it");
    }

    [Fact]
    public void A_buffer_writer_receives_the_same_workbook_as_a_stream()
    {
        var buffer = new ArrayBufferWriter<byte>();

        XlsxSerializer.To(_rows, buffer, XlsxSerializerOptions.Default);

        var viaBuffer = Workbook.Read(buffer.WrittenSpan.ToArray());
        var viaStream = Workbook.Read(Xlsx.Write(_rows, XlsxSerializerOptions.Default));
        Assert.Equal(viaStream.Rows.Count, viaBuffer.Rows.Count);
        Assert.Equal(viaStream.Texts(0), viaBuffer.Texts(0));
    }

    [Fact]
    public async Task A_pipe_receives_the_workbook_while_it_is_still_being_written()
    {
        // The point of the pipe overload is that the reader can start consuming before the
        // export finishes; if the writer only flushed at the end, a large export would sit in
        // memory instead of streaming out.
        var pipe = new Pipe();

        var writing = Task.Run(async () =>
        {
            await XlsxSerializer.ToAsync(_rows, pipe.Writer, XlsxSerializerOptions.Default);
            await pipe.Writer.CompleteAsync();
        });

        using var received = new MemoryStream();
        await pipe.Reader.CopyToAsync(received);
        await pipe.Reader.CompleteAsync();
        await writing;

        var sheet = Workbook.Read(received.ToArray());
        Assert.Equal(_rows, sheet.Rows.Select(r => r[0].Text).ToArray());
    }

    [Fact]
    public void A_file_is_written_where_it_was_asked_for()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_{Guid.NewGuid():N}.xlsx");
        try
        {
            XlsxSerializer.ToFile(_rows, path, XlsxSerializerOptions.Default);

            var sheet = Workbook.Read(File.ReadAllBytes(path));
            Assert.Equal(_rows, sheet.Rows.Select(r => r[0].Text).ToArray());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task An_awaited_export_writes_the_same_file()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_{Guid.NewGuid():N}.xlsx");
        try
        {
            await XlsxSerializer.ToFileAsync(_rows, path, XlsxSerializerOptions.Default);

            var sheet = Workbook.Read(File.ReadAllBytes(path));
            Assert.Equal(_rows, sheet.Rows.Select(r => r[0].Text).ToArray());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Nothing_to_export_leaves_no_file_behind()
    {
        var path = Path.Combine(Path.GetTempPath(), $"flatxlsx_{Guid.NewGuid():N}.xlsx");
        try
        {
            XlsxSerializer.ToFile(Array.Empty<string>(), path, XlsxSerializerOptions.Default);

            Assert.False(File.Exists(path), "an empty export must not leave an unopenable file");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task A_cancelled_export_stops_instead_of_finishing()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        using var ms = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => XlsxSerializer.ToStreamAsync(_rows, ms, XlsxSerializerOptions.Default, cts.Token));
    }

    /// <summary>Stands in for a response body or socket: write-only and unable to seek.</summary>
    sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() => inner.Flush();
        public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
