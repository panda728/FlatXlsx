using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public class ObjectFallbackExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new ObjectFallbackExcelSerializerProvider();

    ObjectFallbackExcelSerializerProvider()
    {

    }

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        if (typeof(T) == typeof(object))
        {
            return (IExcelSerializer<T>)new ObjectFallbackExcelSerializer();
        }

        return null;
    }
}
