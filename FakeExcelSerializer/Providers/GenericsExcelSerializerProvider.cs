using FakeExcelSerializer.Serializers;

namespace FakeExcelSerializer.Providers;

public sealed class GenericsExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new GenericsExcelSerializerProvider();

    GenericsExcelSerializerProvider()
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
            if (type.IsGenericType)
            {
                // Nullable<T>
                var nullableUnderlying = Nullable.GetUnderlyingType(type);
                if (nullableUnderlying != null)
                {
                    return CreateInstance(typeof(NullableExcelSerializer<>), new[] { nullableUnderlying });
                }

                // Tuple/ValueTuple
                var fullName = type.FullName;
                if (fullName != null && (fullName.StartsWith("System.Tuple") || fullName.StartsWith("System.ValueTuple")))
                {
                    var serializerType = (type.IsValueType)
                        ? TupleExcelSerializer.GetValueTupleExcelSerializerType(type.GenericTypeArguments.Length)
                        : TupleExcelSerializer.GetTupleExcelSerializerType(type.GenericTypeArguments.Length);

                    return CreateInstance(serializerType, type.GetGenericArguments());
                }
            }
            else if (type.IsEnum)
            {
                return CreateInstance(typeof(EnumStringExcelSerializer<>), new[] { type });
            }

            return null;
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    static IExcelSerializer? CreateInstance(Type genericType, Type[] genericTypeArguments, params object[] arguments)
    {
        return (IExcelSerializer?)Activator.CreateInstance(genericType.MakeGenericType(genericTypeArguments), arguments);
    }

    static class Cache<T>
    {
        public static readonly IExcelSerializer<T>? Serializer = (IExcelSerializer<T>?)CreateSerializer(typeof(T));
    }
}