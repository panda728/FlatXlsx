using System.Buffers;
using System.Runtime.ExceptionServices;

namespace FakeExcelSerializer.Serializers;

public sealed class ErrorSerializer<T> : IExcelSerializer<T>
{
    readonly ExceptionDispatchInfo exception;

    public ErrorSerializer(Exception exception)
    {
        this.exception = ExceptionDispatchInfo.Capture(exception);
    }
    public void WriteTitle(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options, string name = "value")
    {
        exception.Throw();
    }

    public void Serialize(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options)
    {
        exception.Throw();
    }
}

public static class ErrorSerializer
{
    public static IExcelSerializer Create(Type type, Exception exception)
    {
        return (IExcelSerializer)Activator.CreateInstance(typeof(ErrorSerializer<>).MakeGenericType(type), exception)!;
    }
}
