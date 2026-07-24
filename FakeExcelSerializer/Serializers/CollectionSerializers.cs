namespace FakeExcelSerializer.Serializers;

public sealed class EnumerableExcelSerializer<TCollection, TElement> : IExcelSerializer<TCollection>
    where TCollection : IEnumerable<TElement>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, TCollection value, ExcelSerializerOptions options, string name = "value")
    {
        writer.EnterAndValidate();
        var serializer = options.GetRequiredSerializer<TElement>();
        foreach (var item in value)
        {
            serializer.WriteTitle(ref writer, item, options, name);
        }
        writer.Exit();
    }

    public void Serialize(ref ExcelSerializerWriter writer, TCollection value, ExcelSerializerOptions options)
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
            serializer.Serialize(ref writer, item, options);
        }
        writer.Exit();
    }
}

public sealed class DictionaryExcelSerializer<TDictionary, TKey, TValue> : IExcelSerializer<TDictionary>
    where TDictionary : IDictionary<TKey, TValue>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, TDictionary value, ExcelSerializerOptions options, string name = "value")
    {

        writer.EnterAndValidate();
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(ref writer, item.Key, options, "key");
            valueSerializer.WriteTitle(ref writer, item.Value, options, name);
        }
        writer.Exit();
    }

    public void Serialize(ref ExcelSerializerWriter writer, TDictionary value, ExcelSerializerOptions options)
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

            keySerializer.Serialize(ref writer, item.Key, options);
            valueSerializer.Serialize(ref writer, item.Value, options);
        }
        writer.Exit();
    }
}

public sealed class EnumerableKeyValuePairExcelSerializer<TCollection, TKey, TValue> : IExcelSerializer<TCollection>
    where TCollection : IEnumerable<KeyValuePair<TKey, TValue>>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, TCollection value, ExcelSerializerOptions options, string name = "value")
    {
        var keySerializer = options.GetRequiredSerializer<TKey>();
        var valueSerializer = options.GetRequiredSerializer<TValue>();
        writer.EnterAndValidate();
        foreach (var item in value)
        {
            keySerializer.WriteTitle(ref writer, item.Key, options, "key");
            valueSerializer.WriteTitle(ref writer, item.Value, options, name);
        }
        writer.Exit();
    }

    public void Serialize(ref ExcelSerializerWriter writer, TCollection value, ExcelSerializerOptions options)
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
            keySerializer.Serialize(ref writer, item.Key, options);
            valueSerializer.Serialize(ref writer, item.Value, options);
        }
        writer.Exit();
    }
}