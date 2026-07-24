using FakeExcelSerializer.Providers;
using FakeExcelSerializer.Serializers;
using System.Collections.Concurrent;

namespace FakeExcelSerializer;

public interface IExcelSerializerProvider
{
    IExcelSerializer<T>? GetSerializer<T>();
}

public static class ExcelSerializerProvider
{
    public static IExcelSerializerProvider Default { get; } = new DefaultExcelSerializerProvider();

    public static IExcelSerializerProvider Create(params IExcelSerializerProvider[] providers)
    {
        return new CompositeSerializerProvider(providers);
    }

    public static IExcelSerializerProvider Create(IExcelSerializer[] serializers, IExcelSerializerProvider[] providers)
    {
        var adhocProvider = new AdhocExcelSerializerProvider(serializers);
        return new CompositeSerializerProvider(providers.Prepend(adhocProvider).ToArray());
    }
}

public class DefaultExcelSerializerProvider : IExcelSerializerProvider
{
    static readonly IExcelSerializerProvider[] providers = new[]
    {
            PrimitiveExcelSerializerProvider.Instance,
            BuiltinExcelSerializerProvider.Instance,
            AttributeExcelSerializerProvider.Instance,
            GenericsExcelSerializerProvider.Instance,
            CollectionExcelSerializerProvider.Instance,
            ObjectFallbackExcelSerializerProvider.Instance,
            ObjectGraphExcelSerializerProvider.Instance
        };

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        return Cache<T>.Serializer;
    }

    static class Cache<T>
    {
        public static readonly IExcelSerializer<T>? Serializer;

        static Cache()
        {
            try
            {
                foreach (var provider in providers)
                {
                    var serializer = provider.GetSerializer<T>();
                    if (serializer != null)
                    {
                        Serializer = serializer;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Serializer = new ErrorSerializer<T>(ex);
            }
        }
    }
}

internal class CompositeSerializerProvider : IExcelSerializerProvider
{
    readonly IExcelSerializerProvider[] providers;
    readonly ConcurrentDictionary<Type, IExcelSerializer?> cache;

    public CompositeSerializerProvider(IExcelSerializerProvider[] providers)
    {
        this.providers = providers;
        this.cache = new ConcurrentDictionary<Type, IExcelSerializer?>();
    }

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        if (!cache.TryGetValue(typeof(T), out var serializer))
        {
            serializer = CreateSerializer<T>();
            if (!cache.TryAdd(typeof(T), serializer))
            {
                serializer = cache[typeof(T)];
            }
        }

        return (IExcelSerializer<T>?)serializer;
    }

    IExcelSerializer? CreateSerializer<T>()
    {
        foreach (var provider in providers)
        {
            var serializer = provider.GetSerializer<T>();
            if (serializer != null)
            {
                return serializer;
            }
        }

        return null;
    }
}
