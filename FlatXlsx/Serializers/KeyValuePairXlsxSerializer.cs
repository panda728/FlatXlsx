namespace FlatXlsx.Serializers;

/// <summary>Writes a <see cref="KeyValuePair{TKey, TValue}"/> as two consecutive cells titled
/// "Key" and "Value". A dictionary enumerates as pairs, so dictionary exports flow through here
/// one entry at a time.</summary>
public sealed class KeyValuePairXlsxSerializer<TKey, TValue> : IXlsxSerializer<KeyValuePair<TKey, TValue>>
{
    public void WriteTitle(XlsxCellWriter writer, KeyValuePair<TKey, TValue> value, XlsxSerializerOptions options, string name = "value")
    {
        writer.EnterNested();
        options.GetRequiredSerializer<TKey>().WriteTitle(writer, value.Key, options, "Key");
        options.GetRequiredSerializer<TValue>().WriteTitle(writer, value.Value, options, "Value");
        writer.ExitNested();
    }

    public void Serialize(XlsxCellWriter writer, KeyValuePair<TKey, TValue> value, XlsxSerializerOptions options)
    {
        writer.EnterNested();
        options.GetRequiredSerializer<TKey>().Serialize(writer, value.Key, options);
        options.GetRequiredSerializer<TValue>().Serialize(writer, value.Value, options);
        writer.ExitNested();
    }
}
