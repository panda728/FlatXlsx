using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace FlatXlsx.Serializers;

internal class ObjectFallbackXlsxSerializer : IXlsxSerializer<object>
{
    delegate void WriteTitleDelegate(XlsxWriter writer, object value, XlsxSerializerOptions options, string name);
    static readonly ConcurrentDictionary<Type, WriteTitleDelegate> nongenericWriteTitles = new();
    static readonly Func<Type, WriteTitleDelegate> factoryWriteTitle = CompileWriteTitleDelegate;

    delegate void SerializeDelegate(XlsxWriter writer, object value, XlsxSerializerOptions options);
    static readonly ConcurrentDictionary<Type, SerializeDelegate> nongenericSerializers = new();
    static readonly Func<Type, SerializeDelegate> factory = CompileSerializeDelegate;

    public void WriteTitle(XlsxWriter writer, object value, XlsxSerializerOptions options, string name = "value")
    {
        if (value == null)
        {
            writer.Write(name);
            return;
        }

        var type = value.GetType();
        if (type == typeof(object))
        {
            writer.Write(name);
            return;
        }

        var writeTitle = nongenericWriteTitles.GetOrAdd(type, factoryWriteTitle);
        writeTitle.Invoke(writer, value, options, name);
    }

    public void Serialize(XlsxWriter writer, object value, XlsxSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        var type = value.GetType();
        if (type == typeof(object))
        {
            writer.WriteEmpty();
            return;
        }

        var serializer = nongenericSerializers.GetOrAdd(type, factory);
        serializer.Invoke(writer, value, options);
    }

    static WriteTitleDelegate CompileWriteTitleDelegate(Type type)
    {
        var writer = Expression.Parameter(typeof(XlsxWriter));
        var value = Expression.Parameter(typeof(object));
        var options = Expression.Parameter(typeof(XlsxSerializerOptions));
        var name = Expression.Parameter(typeof(string));

        var getRequiredSerializer = typeof(XlsxSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        var writeTitle = typeof(IXlsxSerializer<>).MakeGenericType(type).GetMethod("WriteTitle")!;
        var body = Expression.Call(
            Expression.Call(options, getRequiredSerializer),
            writeTitle,
            writer,
            Expression.Convert(value, type),
            options,
            name);

        var lambda = Expression.Lambda<WriteTitleDelegate>(body, writer, value, options, name);
        return lambda.Compile();
    }

    static SerializeDelegate CompileSerializeDelegate(Type type)
    {
        // Serialize(XlsxWriter writer, object value, XlsxSerializerOptions options)
        //   options.GetRequiredSerializer<T>().Serialize(writer, (T)value, options)

        var writer = Expression.Parameter(typeof(XlsxWriter));
        var value = Expression.Parameter(typeof(object));
        var options = Expression.Parameter(typeof(XlsxSerializerOptions));

        var getRequiredSerializer = typeof(XlsxSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        var serialize = typeof(IXlsxSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;

        var body = Expression.Call(
            Expression.Call(options, getRequiredSerializer),
            serialize,
            writer,
            Expression.Convert(value, type),
            options);

        var lambda = Expression.Lambda<SerializeDelegate>(body, writer, value, options);
        return lambda.Compile();
    }
}
