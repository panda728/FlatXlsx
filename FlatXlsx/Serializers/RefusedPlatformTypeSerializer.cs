namespace FlatXlsx.Serializers;

/// <summary>What a platform type with no registered serializer resolves to: both operations
/// refuse, naming the type and the remedy. A reflected member layout for the platform's own
/// types is never what the caller meant, so the refusal replaces a silently wrong workbook.
/// The message is built at throw time so it follows the UI culture of the failure, like every
/// other diagnostic.</summary>
internal sealed class RefusedPlatformTypeSerializer<T> : IXlsxSerializer<T>
{
    public void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value")
        => throw Refusal();

    public void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options)
        => throw Refusal();

    static InvalidOperationException Refusal()
    {
#if NETSTANDARD
        // On the netstandard builds the likeliest arrival is a type the net8.0+ targets carry
        // built-in serializers for; those callers are told the faster fix first.
        switch (typeof(T).FullName)
        {
            case "System.DateOnly":
            case "System.TimeOnly":
            case "System.Half":
            case "System.Int128":
            case "System.UInt128":
            case "System.Text.Rune":
                return new InvalidOperationException(SR.PlatformTypeNeedsNewerTarget(typeof(T)));
        }
#endif
        return new InvalidOperationException(SR.PlatformTypeNotSupported(typeof(T)));
    }
}
