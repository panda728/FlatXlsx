using System.Buffers;

namespace FlatXlsx;

public interface IXlsxSerializer { }

public interface IXlsxSerializer<T> : IXlsxSerializer
{
    void WriteTitle(XlsxWriter writer, T value, XlsxSerializerOptions options, string name = "value");
    void Serialize(XlsxWriter writer, T value, XlsxSerializerOptions options);
}
