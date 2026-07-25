using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class ObjectGraphXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new ObjectGraphXlsxSerializerProvider();

    ObjectGraphXlsxSerializerProvider()
    {
    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        return Cache<T>.Serializer;
    }

    static IXlsxSerializer? CreateSerializer(Type type)
    {
        if (IsPlatformType(type))
        {
            return (IXlsxSerializer?)Activator.CreateInstance(typeof(RefusedPlatformTypeSerializer<>).MakeGenericType(type));
        }

        try
        {
            return (IXlsxSerializer?)Activator.CreateInstance(typeof(CompiledObjectGraphXlsxSerializer<>).MakeGenericType(type));
        }
        catch (System.Reflection.TargetInvocationException ex) when (ex.InnerException != null)
        {
            // Activator wraps what the graph compiler threw; the inner exception is the one
            // that says something.
            return ErrorSerializer.Create(type, ex.InnerException);
        }
        catch (Exception ex)
        {
            return ErrorSerializer.Create(type, ex);
        }
    }

    /// <summary>Inferring a column layout from members is for application-defined shapes. For
    /// the platform's own types a reflected layout is never what the caller meant - DateOnly,
    /// for example, would fan out into numeric member columns - so a platform type with no
    /// registered serializer is refused by name instead of silently expanded.</summary>
    static bool IsPlatformType(Type type)
    {
        if (type.Assembly == typeof(object).Assembly)
            return true;

        var assembly = type.Assembly.GetName().Name;
        return assembly != null
            && (assembly == "System" || assembly.StartsWith("System.", StringComparison.Ordinal));
    }

    static class Cache<T>
    {
        public static readonly IXlsxSerializer<T>? Serializer = (IXlsxSerializer<T>?)CreateSerializer(typeof(T));
    }
}