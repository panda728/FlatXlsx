using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;

namespace FlatXlsx;

public record XlsxSerializerOptions(IXlsxSerializerProvider Provider)
{
    public static XlsxSerializerOptions Default { get; } = new XlsxSerializerOptions(XlsxSerializerProvider.Default);

    public CultureInfo? CultureInfo { get; init; }

    /// <summary>Guards against circular references and deeply nested object graphs.</summary>
    public int MaxDepth { get; init; } = 64;

    /// <summary>Upper bound on the number of distinct strings kept in the shared-string table.
    /// The table lives in memory until the sheet is finished, so this caps memory use when the
    /// data has high cardinality. Values beyond the cap are written as inline strings instead,
    /// which costs file size but never fails.</summary>
    public int MaxSharedStrings { get; init; } = 1_000_000;

    public bool AutoFilter { get; init; } = false;
    public bool AutoFitColumns { get; init; } = false;
    public int AutoFitDepth { get; init; } = 200;
    public int AutoFitWidhtMax { get; init; } = 100;
    [Obsolete("Output is now streamed directly; no working folder is used and this value is ignored.")]
    public string WorkPath { get; init; } = "";

    /// <summary>Compression level for the xlsx (zip) container.
    /// Use <see cref="System.IO.Compression.CompressionLevel.Fastest"/> to trade a slightly
    /// larger file for significantly faster serialization of large data sets.</summary>
    public System.IO.Compression.CompressionLevel CompressionLevel { get; init; } = System.IO.Compression.CompressionLevel.Optimal;

    public string DateTimeFormat { get; init; } = "yyyy/mm/dd hh:mm;@";
    public string DateFormat { get; init; } = "yyyy/mm/dd;@";
    public string TimeFormat { get; init; } = "hh:mm;@";
    public string IntegerFormat { get; init; } = "#,##0;[Red]\\-#,##0";
    public string NumberFormat { get; init; } = "#,##0.00;[Red]\\-#,##0.00";

    public bool HasHeaderRecord { get; init; } = false;
    public string[]? HeaderTitles { get; init; }

    public IXlsxSerializer<T>? GetSerializer<T>()
        => Provider.GetSerializer<T>();

    public IXlsxSerializer<T> GetRequiredSerializer<T>()
    {
        var serializer = Provider.GetSerializer<T>();
        if (serializer == null) Throw(typeof(T));
        return serializer!;
    }

#if !NETSTANDARD2_0
    [DoesNotReturn]
#endif
    void Throw(Type type)
    {
        throw new InvalidOperationException($"Type is not found in provider. Type:{type}");
    }
}
