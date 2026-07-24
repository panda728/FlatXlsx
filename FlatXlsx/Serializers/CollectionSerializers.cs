namespace FlatXlsx.Serializers;

public sealed class EnumerableXlsxSerializer<TCollection, TElement> : IXlsxSerializer<TCollection>
    where TCollection : IEnumerable<TElement>
{
    public void WriteTitle(XlsxWriter writer, TCollection value, XlsxSerializerOptions options, string name = "value")
    {
        writer.EnterAndValidate();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.WriteTitle(writer, item, options, name);
        }
        writer.Exit();
    }

    public void Serialize(XlsxWriter writer, TCollection value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterAndValidate();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.Serialize(writer, item, options);
        }
        writer.Exit();
    }
}

public sealed class DictionaryXlsxSerializer<TDictionary, TKey, TValue> : IXlsxSerializer<TDictionary>
    where TDictionary : IDictionary<TKey, TValue>
{
    public void WriteTitle(XlsxWriter writer, TDictionary value, XlsxSerializerOptions options, string name = "value")
    {

        writer.EnterAndValidate();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(writer, item.Key, options, "key");
            valueSerializer.WriteTitle(writer, item.Value, options, name);
        }
        writer.Exit();
    }

    public void Serialize(XlsxWriter writer, TDictionary value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterAndValidate();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();

        foreach (var item in value)
        {
            if (item.Value == null)
            {
                writer.WriteEmpty();
                continue;
            }

            keySerializer.Serialize(writer, item.Key, options);
            valueSerializer.Serialize(writer, item.Value, options);
        }
        writer.Exit();
    }
}

public sealed class EnumerableKeyValuePairXlsxSerializer<TCollection, TKey, TValue> : IXlsxSerializer<TCollection>
    where TCollection : IEnumerable<KeyValuePair<TKey, TValue>>
{
    public void WriteTitle(XlsxWriter writer, TCollection value, XlsxSerializerOptions options, string name = "value")
    {
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        writer.EnterAndValidate();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(writer, item.Key, options, "key");
            valueSerializer.WriteTitle(writer, item.Value, options, name);
        }
        writer.Exit();
    }

    public void Serialize(XlsxWriter writer, TCollection value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        writer.EnterAndValidate();

        foreach (var item in value)
        {
            if (item.Value == null)
            {
                writer.WriteEmpty();
                continue;
            }
            keySerializer.Serialize(writer, item.Key, options);
            valueSerializer.Serialize(writer, item.Value, options);
        }
        writer.Exit();
    }
}