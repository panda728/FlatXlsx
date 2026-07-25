using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class GenericsXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new GenericsXlsxSerializerProvider();

    GenericsXlsxSerializerProvider()
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
            if (type.IsGenericType)
            {
                // Nullable<T>
                var nullableUnderlying = Nullable.GetUnderlyingType(type);
                if (nullableUnderlying != null)
                {
                    return CreateInstance(typeof(NullableXlsxSerializer<>), new[] { nullableUnderlying });
                }

                // KeyValuePair<TKey, TValue> - the shape a dictionary enumerates as
                if (type.GetGenericTypeDefinition() == typeof(KeyValuePair<,>))
                {
                    return CreateInstance(typeof(KeyValuePairXlsxSerializer<,>), type.GetGenericArguments());
                }

                // Tuple/ValueTuple
                var fullName = type.FullName;
                if (fullName != null && (fullName.StartsWith("System.Tuple") || fullName.StartsWith("System.ValueTuple")))
                {
                    var serializerType = (type.IsValueType)
                        ? TupleXlsxSerializer.GetValueTupleXlsxSerializerType(type.GenericTypeArguments.Length)
                        : TupleXlsxSerializer.GetTupleXlsxSerializerType(type.GenericTypeArguments.Length);

                    return CreateInstance(serializerType, type.GetGenericArguments());
                }
            }
            else if (type.IsEnum)
            {
                return CreateInstance(typeof(EnumStringXlsxSerializer<>), new[] { type });
            }

            return null;
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    static IXlsxSerializer? CreateInstance(Type genericType, Type[] genericTypeArguments, params object[] arguments)
    {
        return (IXlsxSerializer?)Activator.CreateInstance(genericType.MakeGenericType(genericTypeArguments), arguments);
    }

    static class Cache<T>
    {
        public static readonly IXlsxSerializer<T>? Serializer = (IXlsxSerializer<T>?)CreateSerializer(typeof(T));
    }
}