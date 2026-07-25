namespace FlatXlsx.Serializers;

public sealed class EnumerableXlsxSerializer<TCollection, TElement> : IXlsxSerializer<TCollection>
    where TCollection : IEnumerable<TElement>
{
    public void WriteTitle(XlsxWriter writer, TCollection value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterNested();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.WriteTitle(writer, item, options, name);
        }
        writer.ExitNested();
    }

    public void Serialize(XlsxWriter writer, TCollection value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterNested();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.Serialize(writer, item, options);
        }
        writer.ExitNested();
    }
}

public sealed class DictionaryXlsxSerializer<TDictionary, TKey, TValue> : IXlsxSerializer<TDictionary>
    where TDictionary : IDictionary<TKey, TValue>
{
    public void WriteTitle(XlsxWriter writer, TDictionary value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterNested();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(writer, item.Key, options, "key");
            valueSerializer.WriteTitle(writer, item.Value, options, name);
        }
        writer.ExitNested();
    }

    public void Serialize(XlsxWriter writer, TDictionary value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        writer.EnterNested();
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
        writer.ExitNested();
    }
}

public sealed class EnumerableKeyValuePairXlsxSerializer<TCollection, TKey, TValue> : IXlsxSerializer<TCollection>
    where TCollection : IEnumerable<KeyValuePair<TKey, TValue>>
{
    public void WriteTitle(XlsxWriter writer, TCollection value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        writer.EnterNested();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(writer, item.Key, options, "key");
            valueSerializer.WriteTitle(writer, item.Value, options, name);
        }
        writer.ExitNested();
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
        writer.EnterNested();

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
        writer.ExitNested();
    }
}