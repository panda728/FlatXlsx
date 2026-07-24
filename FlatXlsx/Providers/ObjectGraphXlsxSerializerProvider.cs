using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class ObjectGraphXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new ObjectGraphXlsxSerializerProvider();

    ObjectGraphXlsxSerializerProvider()
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
            return (IXlsxSerializer?)Activator.CreateInstance(typeof(CompiledObjectGraphXlsxSerializer<>).MakeGenericType(type));
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    static class Cache<T>
    {
        public static readonly IXlsxSerializer<T>? Serializer = (IXlsxSerializer<T>?)CreateSerializer(typeof(T));
    }
}