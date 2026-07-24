using System.Collections.Concurrent;

namespace FlatXlsx.Providers;

public sealed class AdhocXlsxSerializerProvider : IXlsxSerializerProvider
{
    readonly IXlsxSerializer[] serializers;
    readonly ConcurrentDictionary<Type, IXlsxSerializer?> cache = new();
    readonly Func<Type, IXlsxSerializer?> factory;

    public AdhocXlsxSerializerProvider(IXlsxSerializer[] serializers)
    {
        this.serializers = serializers;
        this.factory = CreateSerializer;
    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        return (IXlsxSerializer<T>?)cache.GetOrAdd(typeof(T), factory);
    }

    IXlsxSerializer? CreateSerializer(Type type)
    {
        foreach (var serializer in serializers)
        {
            var serializerType = serializer.GetType().GetImplementedGenericType(typeof(IXlsxSerializer<>));
            if (serializerType != null && serializerType.GenericTypeArguments[0] == type)
            {
                return serializer;
            }
        }

        return null;
    }
}

