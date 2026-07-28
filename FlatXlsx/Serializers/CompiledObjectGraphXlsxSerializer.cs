// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;

namespace FlatXlsx.Serializers;

internal sealed class CompiledObjectGraphXlsxSerializer<T> : IXlsxSerializer<T>
{
    delegate void WriteTitleMethod(XlsxCellWriter writer, IXlsxSerializer?[]? alternateSerializers, T value, XlsxSerializerOptions options);
    delegate void SerializeMethod(XlsxCellWriter writer, IXlsxSerializer?[]? alternateSerializers, T value, XlsxSerializerOptions options);

    static readonly IXlsxSerializer?[]? alternateSerializers;
    static readonly WriteTitleMethod writeTitle;
    static readonly SerializeMethod serialize;
    static readonly bool isReferenceType;

    static CompiledObjectGraphXlsxSerializer()
    {
        isReferenceType = !typeof(T).IsValueType;

        // Instance members only: the parameterless overloads also return statics, which are not
        // row data and cannot be read through the row instance anyway.
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.GetIndexParameters().Length == 0);
        var fields = typeof(T).GetFields(BindingFlags.Public | BindingFlags.Instance);
        var members = props.Cast<MemberInfo>().Concat(fields)
            .Where(x => x.GetCustomAttribute<XlsxIgnoreAttribute>() == null)
            .Select((x, i) => new SerializableMemberInfo(x, i))
            .OrderBy(x => x.Order)
            .ThenBy(x => x.Name)
            .ToArray();

        if (members.Any(x => x.XlsxSerializer != null))
        {
            alternateSerializers = members.Select(x => x.XlsxSerializer).ToArray();
        }
        writeTitle = CompileTitleWriter(typeof(T), members);
        serialize = CompileSerializer(typeof(T), members);
    }

    public void WriteTitle(XlsxCellWriter writer, T value, XlsxSerializerOptions options, string name = "value")
    {
        writer.EnterNested();
        writeTitle(writer, alternateSerializers, value, options);
        writer.ExitNested();
    }

    public void Serialize(XlsxCellWriter writer, T value, XlsxSerializerOptions options)
    {
        if (isReferenceType)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
        }

        writer.EnterNested();
        serialize(writer, alternateSerializers, value, options);
        writer.ExitNested();
    }

    static WriteTitleMethod CompileTitleWriter(Type valueType, SerializableMemberInfo[] memberInfos)
    {
        // foreach(members)
        //     options.GetRequiredSerializer<T>() || ((IXlsxSerializer<T>)alternateSerializers[0] .WriteTitle(writer, value.Foo, options, propertyName)
        var argWriter = Expression.Parameter(typeof(XlsxCellWriter));
        var argAlternateSerializers = Expression.Parameter(typeof(IXlsxSerializer[]));
        var argValue = Expression.Parameter(valueType);
        var argOptions = Expression.Parameter(typeof(XlsxSerializerOptions));
        var foreachBodies = new List<Expression>();

        var i = 0;
        foreach (var memberInfo in memberInfos)
        {
            Expression serializer = memberInfo.XlsxSerializer == null
                ? Expression.Call(argOptions, ReflectionInfos.XlsxSerializerOptions_GetRequiredSerializer(memberInfo.MemberType))
                : Expression.Convert(
                    Expression.ArrayIndex(argAlternateSerializers, Expression.Constant(i, typeof(int))),
                    typeof(IXlsxSerializer<>).MakeGenericType(memberInfo.MemberType)
                );

            var callWriteMember = Expression.Call(
                serializer,
                ReflectionInfos.IXlsxSerializer_WriteTitle(memberInfo.MemberType),
                argWriter,
                memberInfo.GetMemberExpression(argValue),
                argOptions,
                Expression.Constant(memberInfo.Name)
            );

            foreachBodies.Add(callWriteMember);
            i++;
        }

        var body = Expression.Block(foreachBodies);
        var lambda = Expression.Lambda<WriteTitleMethod>(body, argWriter, argAlternateSerializers, argValue, argOptions);
        return lambda.Compile();
    }

    static SerializeMethod CompileSerializer(Type valueType, SerializableMemberInfo[] memberInfos)
    {
        // foreach(members)
        //   if (value.Foo != null) // reference type || nullable type
        //     options.GetRequiredSerializer<T>() || ((IXlsxSerializer<T>)alternateSerializers[0] .Serialize(writer, value.Foo, options)
        //   else
        //     writer.WriteEmpty();
        var argWriter = Expression.Parameter(typeof(XlsxCellWriter));
        var argAlternateSerializers = Expression.Parameter(typeof(IXlsxSerializer[]));
        var argValue = Expression.Parameter(valueType);
        var argOptions = Expression.Parameter(typeof(XlsxSerializerOptions));
        var foreachBodies = new List<Expression>();

        var i = 0;
        foreach (var memberInfo in memberInfos)
        {
            Expression serializer = memberInfo.XlsxSerializer == null
                ? Expression.Call(argOptions, ReflectionInfos.XlsxSerializerOptions_GetRequiredSerializer(memberInfo.MemberType))
                : Expression.Convert(
                    Expression.ArrayIndex(argAlternateSerializers, Expression.Constant(i, typeof(int))),
                    typeof(IXlsxSerializer<>).MakeGenericType(memberInfo.MemberType)
                );

            var callSerializer = Expression.Call(
                serializer,
                ReflectionInfos.IXlsxSerializer_Serialize(memberInfo.MemberType),
                argWriter,
                memberInfo.GetMemberExpression(argValue),
                argOptions
            );

            if (!memberInfo.MemberType.IsValueType || memberInfo.MemberType.IsNullable())
            {
                var nullExpr = Expression.Constant(null, memberInfo.MemberType);
                var ifBody = Expression.IfThenElse(
                    Expression.NotEqual(memberInfo.GetMemberExpression(argValue), nullExpr),
                    callSerializer,
                    Expression.Call(argWriter, ReflectionInfos.Writer_Empty)
                );
                foreachBodies.Add(ifBody);
            }
            else
            {
                foreachBodies.Add(callSerializer);
            }
            i++;
        }

        var body = Expression.Block(foreachBodies);
        var lambda = Expression.Lambda<SerializeMethod>(body, argWriter, argAlternateSerializers, argValue, argOptions);
        return lambda.Compile();

    }

    internal static class ReflectionInfos
    {
        internal static MethodInfo Writer_Empty { get; } = typeof(XlsxCellWriter).GetMethod("WriteEmpty")!;
        internal static MethodInfo XlsxSerializerOptions_GetRequiredSerializer(Type type) => typeof(XlsxSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        internal static MethodInfo IXlsxSerializer_Serialize(Type type) => typeof(IXlsxSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;
        internal static MethodInfo IXlsxSerializer_WriteTitle(Type type) => typeof(IXlsxSerializer<>).MakeGenericType(type).GetMethod("WriteTitle")!;
    }
}

internal sealed class SerializableMemberInfo
{
    public string Name { get; }
    public int Order { get; }
    public IXlsxSerializer? XlsxSerializer { get; }
    public Type MemberType { get; }
    public MemberInfo MemberInfo { get; }

    public SerializableMemberInfo(MemberInfo member, int i)
    {
        var column = member.GetCustomAttribute<XlsxColumnAttribute>();
        var dataMember = member.GetCustomAttribute<DataMemberAttribute>();

        MemberInfo = member;
        Name = column?.Name ?? dataMember?.Name ?? member.Name;
        Order = column != null
            ? (column.Order >= 0 ? column.Order : i)
            : dataMember?.Order ?? i;

        MemberType = member switch
        {
            PropertyInfo pi => pi.PropertyType,
            FieldInfo fi => fi.FieldType,
            _ => throw new InvalidOperationException(SR.MemberKindNotSupported)
        };

        var serializerAttr = member.GetCustomAttribute<XlsxSerializerAttribute>();
        if (serializerAttr != null)
        {
            serializerAttr.Validate(MemberType);
            XlsxSerializer = (IXlsxSerializer?)Activator.CreateInstance(serializerAttr.Type);
        }
    }

    public MemberExpression GetMemberExpression(Expression expression)
    {
        if (MemberInfo is FieldInfo fi)
        {
            return Expression.Field(expression, fi);
        }
        else if (MemberInfo is PropertyInfo pi)
        {
            return Expression.Property(expression, pi);
        }
        throw new InvalidOperationException(SR.MemberKindNotSupported);
    }
}