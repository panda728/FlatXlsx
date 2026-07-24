using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed partial class PrimitiveXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new PrimitiveXlsxSerializerProvider();
    readonly Dictionary<Type, IXlsxSerializer> serializers = new Dictionary<Type, IXlsxSerializer>();

    internal partial void InitPrimitives(); // implement from PrimitiveSerializers.cs

    PrimitiveXlsxSerializerProvider()
    {
        InitPrimitives();
    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (serializers.TryGetValue(typeof(T), out var value))
        {
            return (IXlsxSerializer<T>)value;
        }
        return null;
    }
}