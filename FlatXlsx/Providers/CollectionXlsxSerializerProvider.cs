// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class CollectionXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new CollectionXlsxSerializerProvider();

    CollectionXlsxSerializerProvider()
    {

    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        return Cache<T>.Serializer;
    }

    static IXlsxSerializer? CreateSerializer(Type type)
    {
        try
        {
            // Wellknown specialized types
            if (type == typeof(Dictionary<string, string>))
            {
                return new DictionaryXlsxSerializer<Dictionary<string, string>, string, string>();
            }
            else if (type == typeof(Dictionary<string, object>))
            {
                return new DictionaryXlsxSerializer<Dictionary<string, object>, string, object>();
            }
            else if (type == typeof(KeyValuePair<string, string>[]))
            {
                return new EnumerableKeyValuePairXlsxSerializer<KeyValuePair<string, string>[], string, string>();
            }
            else if (type == typeof(KeyValuePair<string, object>[]))
            {
                return new EnumerableKeyValuePairXlsxSerializer<KeyValuePair<string, object>[], string, object>();
            }

            if (type.IsGenericType || type.IsArray)
            {
                // Generic Dictionary
                var dictionaryDef = type.GetImplementedGenericType(typeof(IDictionary<,>));
                if (dictionaryDef != null)
                {
                    var keyType = dictionaryDef.GenericTypeArguments[0];
                    var valueType = dictionaryDef.GenericTypeArguments[1];
                    return CreateInstance(typeof(DictionaryXlsxSerializer<,,>), new[] { type, keyType, valueType });
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
                        return CreateInstance(typeof(EnumerableKeyValuePairXlsxSerializer<,,>), new[] { type, keyType, valueType });
                    }
                    else
                    {
                        return CreateInstance(typeof(EnumerableXlsxSerializer<,>), new[] { type, elementType });
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

    static IXlsxSerializer? CreateInstance(Type genericType, Type[] genericTypeArguments, params object[] arguments)
    {
        return (IXlsxSerializer?)Activator.CreateInstance(genericType.MakeGenericType(genericTypeArguments), arguments);
    }

    static class Cache<T>
    {
        public static readonly IXlsxSerializer<T>? Serializer = (IXlsxSerializer<T>?)CreateSerializer(typeof(T));
    }
}
