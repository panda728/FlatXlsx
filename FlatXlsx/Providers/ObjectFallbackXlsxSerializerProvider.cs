// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public class ObjectFallbackXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new ObjectFallbackXlsxSerializerProvider();

    ObjectFallbackXlsxSerializerProvider()
    {

    }

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (typeof(T) == typeof(object))
        {
            return (IXlsxSerializer<T>)new ObjectFallbackXlsxSerializer();
        }

        return null;
    }
}
