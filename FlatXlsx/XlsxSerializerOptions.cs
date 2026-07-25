using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace FlatXlsx;

public record XlsxSerializerOptions
{
    public static XlsxSerializerOptions Default { get; } = new();

    /// <summary>Resolves the serializer for each type. The default provider covers the
    /// supported built-in types plus object graphs. To add or replace serializers for a few
    /// types, <see cref="CustomSerializers"/> is the simpler road; replace the provider only
    /// for full control over resolution.</summary>
    public IXlsxSerializerProvider Provider { get; init; } = XlsxSerializerProvider.Default;

    /// <summary>Serializers consulted before the built-in ones. Setting this is sufficient on
    /// its own - no provider wiring is needed.</summary>
    /// <remarks>Matching is by the serializer's exact value type: a serializer registered for a
    /// base type is not applied to derived types.
    /// <code>
    /// new XlsxSerializerOptions { CustomSerializers = new IXlsxSerializer[] { new MoneySerializer() } }
    /// </code></remarks>
    public IXlsxSerializer[]? CustomSerializers { get; init; }

    /// <summary>Formats the few types whose text is culture-defined: DateTimeOffset and
    /// Complex. Everything else - numbers, dates, times - is always stored culture-invariantly,
    /// so the same code produces the same file on every machine. Set this only to deliberately
    /// localize those two types' text.</summary>
    public CultureInfo CultureInfo { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>Name of the (single) worksheet. Excel's naming rules apply: 1-31 characters,
    /// none of <c>: \ / ? * [ ]</c>, and no leading or trailing apostrophe - violations fail
    /// at write time rather than producing a workbook Excel rejects.</summary>
    public string SheetName { get; init; } = "Sheet";

    /// <summary>Guards against circular references and deeply nested object graphs.</summary>
    public int MaxDepth { get; init; } = 64;

    /// <summary>Upper bound on the number of distinct strings kept in the shared-string table.
    /// The table lives in memory until the sheet is finished, so this caps memory use when the
    /// data has high cardinality. Values beyond the cap are written as inline strings instead,
    /// which costs file size but never fails.</summary>
    public int MaxSharedStrings { get; init; } = 1_000_000;

    /// <summary>Adds an Excel auto-filter covering the header and every row written.</summary>
    public bool AutoFilter { get; init; } = false;

    /// <summary>Sizes each column to its content, approximated by character count.</summary>
    public bool AutoFitColumns { get; init; } = false;

    /// <summary>How many leading rows are measured for <see cref="AutoFitColumns"/>. Widths are
    /// written before the data, so measuring is bounded to this sample rather than buffering
    /// the whole sequence; a longer value in a later row does not widen the column.</summary>
    public int AutoFitSampleRows { get; init; } = 200;

    /// <summary>Upper bound for an auto-fitted column's width, in characters.</summary>
    public int AutoFitMaxWidth { get; init; } = 100;

    /// <summary>Compression level for the xlsx (zip) container.
    /// Use <see cref="System.IO.Compression.CompressionLevel.Fastest"/> to trade a slightly
    /// larger file for significantly faster serialization of large data sets.</summary>
    public System.IO.Compression.CompressionLevel CompressionLevel { get; init; } = System.IO.Compression.CompressionLevel.Optimal;

    /// <summary>Additional Excel number-format codes, referenced from a custom serializer by
    /// index: <c>writer.Write(value, customFormat: 0)</c> uses <c>CustomFormats[0]</c>.
    /// The built-in formats below stay untouched, so one percent column no longer means
    /// hijacking <see cref="NumberFormat"/> for every number in the sheet.</summary>
    /// <remarks><code>
    /// options = new XlsxSerializerOptions { CustomFormats = new[] { "0.0%" } };
    /// // in a serializer:  writer.Write(ratio, customFormat: 0);
    /// </code></remarks>
    public string[]? CustomFormats { get; init; }

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

    // The composed provider is rebuilt whenever the inputs are observed to have changed rather
    // than cached unconditionally: `with` copies fields, so an unconditional cache taken from
    // the original record would leak into copies that carry different serializers.
    IXlsxSerializerProvider? _effective;
    IXlsxSerializer[]? _effectiveSerializers;
    IXlsxSerializerProvider? _effectiveBase;

    IXlsxSerializerProvider EffectiveProvider
    {
        get
        {
            if (CustomSerializers is not { Length: > 0 })
                return Provider;

            var effective = _effective;
            if (effective == null
                || !ReferenceEquals(_effectiveSerializers, CustomSerializers)
                || !ReferenceEquals(_effectiveBase, Provider))
            {
                effective = XlsxSerializerProvider.Create(CustomSerializers, new[] { Provider });
                _effectiveSerializers = CustomSerializers;
                _effectiveBase = Provider;
                _effective = effective;
            }
            return effective;
        }
    }

    public IXlsxSerializer<T>? GetSerializer<T>()
        => EffectiveProvider.GetSerializer<T>();

    public IXlsxSerializer<T> GetRequiredSerializer<T>()
    {
        var serializer = EffectiveProvider.GetSerializer<T>();
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
