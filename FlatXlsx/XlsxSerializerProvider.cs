using FlatXlsx.Providers;
using FlatXlsx.Serializers;
using System.Collections.Concurrent;

namespace FlatXlsx;

public interface IXlsxSerializerProvider
{
    IXlsxSerializer<T>? GetSerializer<T>();
}

public static class XlsxSerializerProvider
{
    public static IXlsxSerializerProvider Default { get; } = new DefaultXlsxSerializerProvider();

    /// <summary>Composes providers into one; earlier providers win.</summary>
    public static IXlsxSerializerProvider Create(params IXlsxSerializerProvider[] providers)
    {
        return new CompositeSerializerProvider(providers);
    }

    /// <summary>These serializers first, then the default provider - no need to pass
    /// <see cref="Default"/> yourself. Equivalent to setting
    /// <see cref="XlsxSerializerOptions.CustomSerializers"/>.</summary>
    public static IXlsxSerializerProvider Create(params IXlsxSerializer[] serializers)
    {
        return Create(serializers, new[] { Default });
    }

    /// <summary>These serializers first, then the given providers in order.</summary>
    public static IXlsxSerializerProvider Create(IXlsxSerializer[] serializers, IXlsxSerializerProvider[] providers)
    {
        var adhocProvider = new AdhocXlsxSerializerProvider(serializers);
        return new CompositeSerializerProvider(providers.Prepend(adhocProvider).ToArray());
    }
}

/// <summary>The resolution chain behind <see cref="XlsxSerializerProvider.Default"/>. Internal:
/// the instance is what callers compose with; the class itself offers nothing to construct or
/// override.</summary>
internal sealed class DefaultXlsxSerializerProvider : IXlsxSerializerProvider
{
    static readonly IXlsxSerializerProvider[] providers = new[]
    {
            PrimitiveXlsxSerializerProvider.Instance,
            BuiltinXlsxSerializerProvider.Instance,
            AttributeXlsxSerializerProvider.Instance,
            GenericsXlsxSerializerProvider.Instance,
            CollectionXlsxSerializerProvider.Instance,
            ObjectFallbackXlsxSerializerProvider.Instance,
            ObjectGraphXlsxSerializerProvider.Instance
        };

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        return Cache<T>.Serializer;
    }

    static class Cache<T>
    {
        public static readonly IXlsxSerializer<T>? Serializer;

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

internal sealed class CompositeSerializerProvider(IXlsxSerializerProvider[] providers) : IXlsxSerializerProvider
{
    readonly ConcurrentDictionary<Type, IXlsxSerializer?> cache = new();

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (!cache.TryGetValue(typeof(T), out var serializer))
        {
            serializer = CreateSerializer<T>();
            if (!cache.TryAdd(typeof(T), serializer))
            {
                serializer = cache[typeof(T)];
            }
        }

        return (IXlsxSerializer<T>?)serializer;
    }

    IXlsxSerializer? CreateSerializer<T>()
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
