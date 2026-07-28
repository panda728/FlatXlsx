// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.

namespace FlatXlsx;

public interface IXlsxSerializer { }

public interface IXlsxSerializer<T> : IXlsxSerializer
{
    void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value");
    void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options);
}
