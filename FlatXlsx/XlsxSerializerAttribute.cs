// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
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

/// <summary>Excludes the member from the output.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class XlsxIgnoreAttribute : Attribute
{
}

/// <summary>
/// Sets the column title and position of a member, right where the member is declared.
/// </summary>
/// <remarks>
/// <code>
/// [XlsxColumn("部署", Order = 1)]
/// public string Department { get; set; }
/// </code>
/// <c>[DataMember(Name, Order)]</c> is also honored for types already annotated for other
/// serializers; when both are present, this attribute wins.
/// </remarks>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public sealed class XlsxColumnAttribute(string? name = null) : Attribute
{
    /// <summary>The column title. null keeps the member's own name.</summary>
    public string? Name { get; } = name;

    /// <summary>Position of the column; members are sorted by this, then by declaration order.
    /// Unset means declaration order.</summary>
    public int Order { get; set; } = -1;
}