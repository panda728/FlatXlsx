using System.Reflection;
using System.Runtime.CompilerServices;

namespace FakeExcelSerializer;

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