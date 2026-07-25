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
            throw new InvalidOperationException(SR.SerializerTypeNotImplemented(Type));
        }

        var attrType = serializerType.GenericTypeArguments[0];
        if (attrType != targetType)
        {
            throw new InvalidOperationException(SR.SerializerTypeMismatch(attrType, targetType));
        }
    }
}

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class IgnoreXlsxSerializeAttribute : Attribute
{
}