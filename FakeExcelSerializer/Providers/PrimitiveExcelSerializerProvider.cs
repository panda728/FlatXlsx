using FakeExcelSerializer.Serializers;

namespace FakeExcelSerializer.Providers;

public sealed partial class PrimitiveExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new PrimitiveExcelSerializerProvider();
    readonly Dictionary<Type, IExcelSerializer> serializers = new Dictionary<Type, IExcelSerializer>();

    internal partial void InitPrimitives(); // implement from PrimitiveSerializers.cs

    PrimitiveExcelSerializerProvider()
    {
        InitPrimitives();
    }

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        if (serializers.TryGetValue(typeof(T), out var value))
        {
            return (IExcelSerializer<T>)value;
        }
        return null;
    }
}