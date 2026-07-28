// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using System.Runtime.CompilerServices;

namespace FlatXlsx;

internal static class ReflectionExtensions
{
    public static bool IsNullable(this Type type)
    {
        return Nullable.GetUnderlyingType(type) != null;
    }

    public static Type? GetImplementedGenericType(this Type type, Type genericTypeDefinition)
    {
        return type.GetInterfaces().FirstOrDefault(x => x.IsConstructedGenericType && x.GetGenericTypeDefinition() == genericTypeDefinition);
    }
}
