using System.Buffers;
using System.Runtime.ExceptionServices;

namespace FlatXlsx.Serializers;

public sealed class ErrorSerializer<T>(Exception exception) : IXlsxSerializer<T>
{
    readonly ExceptionDispatchInfo exception = ExceptionDispatchInfo.Capture(exception);

    public void WriteTitle(XlsxWriter writer, T value, XlsxSerializerOptions options, string name = "value")
    {
        exception.Throw();
    }

    public void Serialize(XlsxWriter writer, T value, XlsxSerializerOptions options)
    {
        exception.Throw();
    }
}

public static class ErrorSerializer
{
    public static IXlsxSerializer Create(Type type, Exception exception)
    {
        return (IXlsxSerializer)Activator.CreateInstance(typeof(ErrorSerializer<>).MakeGenericType(type), exception)!;
    }
}
