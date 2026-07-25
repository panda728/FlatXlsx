using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;

namespace FlatXlsx;

public record XlsxSerializerOptions
{
    public static XlsxSerializerOptions Default { get; } = new();

    /// <summary>Resolves the serializer for each type. The default provider covers the
    /// supported built-in types plus object graphs; replace it only to plug in custom
    /// serializers.</summary>
    public IXlsxSerializerProvider Provider { get; init; } = XlsxSerializerProvider.Default;

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
    public int AutoFitWidthMax { get; init; } = 100;

    /// <summary>Compression level for the xlsx (zip) container.
    /// Use <see cref="System.IO.Compression.CompressionLevel.Fastest"/> to trade a slightly
    /// larger file for significantly faster serialization of large data sets.</summary>
    public System.IO.Compression.CompressionLevel CompressionLevel { get; init; } = System.IO.Compression.CompressionLevel.Optimal;

    public string DateTimeFormat { get; init; } = "yyyy/mm/dd hh:mm;@";
    public string DateFormat { get; init; } = "yyyy/mm/dd;@";
    public string TimeFormat { get; init; } = "hh:mm;@";
    public string IntegerFormat { get; init; } = "#,##0;[Red]\\-#,##0";
    public string NumberFormat { get; init; } = "#,##0.00;[Red]\\-#,##0.00";

    /// <summary>Adds a frozen header row titled from the members' names
    /// (or their <see cref="XlsxColumnAttribute"/> / DataMember names).</summary>
    /// <remarks>Not needed when <see cref="HeaderTitles"/> is set: supplying titles is already
    /// asking for a header row.</remarks>
    public bool HasHeaderRow { get; init; } = false;

    /// <summary>Adds a frozen header row with exactly these titles. Setting this is sufficient
    /// on its own; no other option needs to change.</summary>
    public string[]? HeaderTitles { get; init; }

    /// <summary>A header row is written when either <see cref="HasHeaderRow"/> is set or
    /// <see cref="HeaderTitles"/> supplies the titles - asking for titles and separately asking
    /// for the row to exist would be two settings for one intention.</summary>
    internal bool HasHeader => HasHeaderRow || HeaderTitles is { Length: > 0 };

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
        throw new InvalidOperationException(SR.SerializerNotFound(type));
    }
}
