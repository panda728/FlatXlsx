namespace FlatXlsx.Tests.Support;

static class Xlsx
{
    public static byte[] Write<T>(IEnumerable<T> rows, XlsxSerializerOptions options)
    {
        using var ms = new MemoryStream();
        XlsxSerializer.ToStream(rows, ms, options);
        return ms.ToArray();
    }

    public static async Task<byte[]> WriteAsync<T>(IEnumerable<T> rows, XlsxSerializerOptions options)
    {
        using var ms = new MemoryStream();
        await XlsxSerializer.ToStreamAsync(rows, ms, options);
        return ms.ToArray();
    }

    public static Workbook.Sheet Read<T>(IEnumerable<T> rows, XlsxSerializerOptions options)
        => Workbook.Read(Write(rows, options));
}
