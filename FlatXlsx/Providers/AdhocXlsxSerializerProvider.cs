using System.Collections.Concurrent;

namespace FlatXlsx.Providers;

/// <summary>Serves a fixed list of serializer instances, matched by value type. Construction
/// plumbing behind <see cref="XlsxSerializerProvider.Create(IXlsxSerializer[], IXlsxSerializerProvider[])"/>
/// and <see cref="XlsxSerializerOptions.CustomSerializers"/>.</summary>
/// <remarks>
/// Matching is a type test rather than a reflected interface list, because this is the one
/// resolution path that a trimmed or ahead-of-time-compiled application can rely on: registering
/// a serializer for every row type is what the trimming warning on the entry points tells callers
/// to do, so this path must not itself reflect. <see cref="IXlsxSerializer{T}"/> is invariant, so
/// the test matches exactly the type the serializer was written for - the documented behaviour.
/// </remarks>
internal sealed class AdhocXlsxSerializerProvider(IXlsxSerializer[] serializers) : IXlsxSerializerProvider
{
    readonly ConcurrentDictionary<Type, IXlsxSerializer?> cache = new();

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (cache.TryGetValue(typeof(T), out var cached))
            return (IXlsxSerializer<T>?)cached;

        IXlsxSerializer? match = null;
        foreach (var serializer in serializers)
        {
            if (serializer is IXlsxSerializer<T> typed)
            {
                match = typed;
                break;
            }
        }

        // A racing caller computes the same answer, so last write wins harmlessly.
        cache[typeof(T)] = match;
        return (IXlsxSerializer<T>?)match;
    }
}
