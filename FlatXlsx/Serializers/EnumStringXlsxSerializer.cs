// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.Serialization;

namespace FlatXlsx.Serializers;

public sealed class EnumStringXlsxSerializer<T> : IXlsxSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);

    public void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options)
    {
        var str = stringCache.GetOrAdd(value, toStringFactory);
        writer.Write(str);
    }

    static string EnumToString(T value)
    {
        var str = value.ToString();
        var field = value.GetType().GetField(str);
        if (field != null)
        {
            var enumMember = field.GetCustomAttribute<EnumMemberAttribute>();
            if (enumMember != null && enumMember.Value != null)
            {
                str = enumMember.Value;
            }
        }
        return str;
    }
}

public sealed class EnumNumberXlsxSerializer<T> : IXlsxSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);

    public void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options)
    {
        var str = stringCache.GetOrAdd(value, toStringFactory);
        writer.Write(str);
    }

    static string EnumToString(T value)
    {
        return Convert.ChangeType(value, Enum.GetUnderlyingType(typeof(T))).ToString()!;
    }
}
