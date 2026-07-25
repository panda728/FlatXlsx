using System.Runtime.ExceptionServices;

namespace FlatXlsx.Serializers;

/// <summary>Defers a provider-resolution failure to the moment the serializer is used, so the
/// original exception surfaces at the call that needed it. Internal plumbing, not a name for
/// users to meet in IntelliSense.</summary>
internal sealed class ErrorSerializer<T>(Exception exception) : IXlsxSerializer<T>
{
    readonly ExceptionDispatchInfo exception = ExceptionDispatchInfo.Capture(exception);

    public void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value")
    {
        exception.Throw();
    }

    public void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options)
    {
        exception.Throw();
    }
}

internal static class ErrorSerializer
{
    public static IXlsxSerializer Create(Type type, Exception exception)
    {
        return (IXlsxSerializer)Activator.CreateInstance(typeof(ErrorSerializer<>).MakeGenericType(type), exception)!;
    }
}
