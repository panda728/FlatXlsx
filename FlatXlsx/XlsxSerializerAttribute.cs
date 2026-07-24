namespace FlatXlsx;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Enum, AllowMultiple = false)]
public sealed class XlsxSerializerAttribute(Type type) : Attribute
{
    public Type Type { get; } = type;

    internal void Validate(Type targetType)
    {
        var serializerType = Type.GetImplementedGenericType(typeof(IXlsxSerializer<>));
        if (serializerType == null)
        {
            throw new InvalidOperationException($"Type is not implemented IXlsxSerializer<T>, Type:{Type.FullName}");
        }

        var attrType = serializerType.GenericTypeArguments[0];
        if (attrType != targetType)
        {
            throw new InvalidOperationException($"Attribute XlsxSerializer type is not same as target type. AttrType:{attrType.FullName} TargetType:{targetType.FullName}");
        }
    }
}

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class IgnoreXlsxSerializeAttribute : Attribute
{
}