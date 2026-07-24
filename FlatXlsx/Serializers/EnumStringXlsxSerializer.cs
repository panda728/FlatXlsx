using System.Buffers;
using System.Collections.Concurrent;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text.Encodings.Web;

namespace FlatXlsx.Serializers;

public sealed class EnumStringXlsxSerializer<T> : IXlsxSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(XlsxWriter writer, T value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);

    public void Serialize(XlsxWriter writer, T value, XlsxSerializerOptions options)
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

public sealed class EnumValueXlsxSerializer<T> : IXlsxSerializer<T>
    where T : Enum
{
    static readonly ConcurrentDictionary<T, string> stringCache = new();
    static readonly Func<T, string> toStringFactory = EnumToString;

    public void WriteTitle(XlsxWriter writer, T value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);

    public void Serialize(XlsxWriter writer, T value, XlsxSerializerOptions options)
    {
        var str = stringCache.GetOrAdd(value, toStringFactory);
        writer.Write(str);
    }

    static string EnumToString(T value)
    {
        return Convert.ChangeType(value, Enum.GetUnderlyingType(typeof(T))).ToString()!;
    }
}