using System.Buffers;

namespace FlatXlsx.Serializers;

public sealed class NullableXlsxSerializer<T> : IXlsxSerializer<T?>
    where T : struct
{
    public void WriteTitle(XlsxWriter writer, T? value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.Write(name);
            return;
        }
        options.GetRequiredSerializer<T>().WriteTitle(writer, value.Value, options, name);
    }

    public void Serialize(XlsxWriter writer, T? value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }
        options.GetRequiredSerializer<T>().Serialize(writer, value.Value, options);
    }
}