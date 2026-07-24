using FakeExcelSerializer.Serializers;

namespace FakeExcelSerializer.Providers;

public sealed class CollectionExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new CollectionExcelSerializerProvider();

    CollectionExcelSerializerProvider()
    {

    }

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        return Cache<T>.Serializer;
    }

    static IExcelSerializer? CreateSerializer(Type type)
    {
        try
        {
            // Wellknown specialized types
            if (type == typeof(Dictionary<string, string>))
            {
                return new DictionaryExcelSerializer<Dictionary<string, string>, string, string>();
            }
            else if (type == typeof(Dictionary<string, object>))
            {
                return new DictionaryExcelSerializer<Dictionary<string, object>, string, object>();
            }
            else if (type == typeof(KeyValuePair<string, string>[]))
            {
                return new EnumerableKeyValuePairExcelSerializer<KeyValuePair<string, string>[], string, string>();
            }
            else if (type == typeof(KeyValuePair<string, object>[]))
            {
                return new EnumerableKeyValuePairExcelSerializer<KeyValuePair<string, object>[], string, object>();
            }

            if (type.IsGenericType || type.IsArray)
            {
                // Generic Dictionary
                var dictionaryDef = type.GetImplementedGenericType(typeof(IDictionary<,>));
                if (dictionaryDef != null)
                {
                    var keyType = dictionaryDef.GenericTypeArguments[0];
                    var valueType = dictionaryDef.GenericTypeArguments[1];
                    return CreateInstance(typeof(DictionaryExcelSerializer<,,>), new[] { type, keyType, valueType });
                }

                // Generic Collections
                var enumerableDef = type.GetImplementedGenericType(typeof(IEnumerable<>));
                if (enumerableDef != null)
                {
                    var elementType = enumerableDef.GenericTypeArguments[0];
                    if (elementType.IsGenericType && elementType.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
                    {
                        var keyType = elementType.GenericTypeArguments[0];
                        var valueType = elementType.GenericTypeArguments[1];
                        return CreateInstance(typeof(EnumerableKeyValuePairExcelSerializer<,,>), new[] { type, keyType, valueType });
                    }
                    else
                    {
                        return CreateInstance(typeof(EnumerableExcelSerializer<,>), new[] { type, elementType });
                    }
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    static IExcelSerializer? CreateInstance(Type genericType, Type[] genericTypeArguments, params object[] arguments)
    {
        return (IExcelSerializer?)Activator.CreateInstance(genericType.MakeGenericType(genericTypeArguments), arguments);
    }

    static class Cache<T>
    {
        public static readonly IExcelSerializer<T>? Serializer = (IExcelSerializer<T>?)CreateSerializer(typeof(T));
    }
}
