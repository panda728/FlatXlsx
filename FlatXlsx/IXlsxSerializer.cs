
namespace FlatXlsx;

public interface IXlsxSerializer { }

public interface IXlsxSerializer<T> : IXlsxSerializer
{
    void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value");
    void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options);
}
