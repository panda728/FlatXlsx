// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.

namespace FlatXlsx.Serializers;

public sealed class NullableXlsxSerializer<T> : IXlsxSerializer<T?>
    where T : struct
{
    public void WriteTitle(XlsxCellWriter writer, T? value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.Write(name);
            return;
        }
        options.GetRequiredSerializer<T>().WriteTitle(writer, value.Value, options, name);
    }

    public void Serialize(XlsxCellWriter writer, T? value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }
        options.GetRequiredSerializer<T>().Serialize(writer, value.Value, options);
    }
}
