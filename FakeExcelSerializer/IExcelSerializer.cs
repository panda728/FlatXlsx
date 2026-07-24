using System.Buffers;

namespace FakeExcelSerializer;

public interface IExcelSerializer { }

public interface IExcelSerializer<T> : IExcelSerializer
{
    void WriteTitle(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options, string name = "value");
    void Serialize(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options);
}
