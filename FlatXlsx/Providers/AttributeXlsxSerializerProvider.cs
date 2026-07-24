using FlatXlsx.Serializers;
using System.Reflection;

namespace FlatXlsx.Providers;

public sealed class AttributeXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new AttributeXlsxSerializerProvider();

    AttributeXlsxSerializerProvider()
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
            var attr = type.GetCustomAttribute<XlsxSerializerAttribute>();
            if (attr != null)
            {
                attr.Validate(type);
                return (IXlsxSerializer?)Activator.CreateInstance(attr.Type);
            }

            return null;
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