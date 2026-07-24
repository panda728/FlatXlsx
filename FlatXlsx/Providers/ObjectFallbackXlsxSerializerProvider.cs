using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public class ObjectFallbackXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new ObjectFallbackXlsxSerializerProvider();

    ObjectFallbackXlsxSerializerProvider()
    {

    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (typeof(T) == typeof(object))
        {
            return (IXlsxSerializer<T>)new ObjectFallbackXlsxSerializer();
        }

        return null;
    }
}
