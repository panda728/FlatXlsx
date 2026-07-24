using FakeExcelSerializer.Serializers;

namespace FakeExcelSerializer.Providers;

public sealed class ObjectGraphExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new ObjectGraphExcelSerializerProvider();

    ObjectGraphExcelSerializerProvider()
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
            return (IExcelSerializer?)Activator.CreateInstance(typeof(CompiledObjectGraphExcelSerializer<>).MakeGenericType(type));
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    static class Cache<T>
    {
        public static readonly IExcelSerializer<T>? Serializer = (IExcelSerializer<T>?)CreateSerializer(typeof(T));
    }
}