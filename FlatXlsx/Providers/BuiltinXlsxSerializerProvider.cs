// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class BuiltinXlsxSerializerProvider : IXlsxSerializerProvider
{
    public static IXlsxSerializerProvider Instance { get; } = new BuiltinXlsxSerializerProvider();
    readonly Dictionary<Type, IXlsxSerializer> serializers = new()
        {
            { typeof(string), new BuiltinSerializers.StringXlsxSerializer() },
            { typeof(char), new BuiltinSerializers.CharXlsxSerializer() },
            { typeof(Guid), new  BuiltinSerializers.GuidXlsxSerializer() },
            { typeof(Enum), new  BuiltinSerializers.EnumXlsxSerializer() },
            { typeof(DateTime), new  BuiltinSerializers.DateTimeXlsxSerializer() },
            { typeof(DateTimeOffset), new  BuiltinSerializers.DateTimeOffsetXlsxSerializer() },
            { typeof(TimeSpan), new  BuiltinSerializers.TimeSpanXlsxSerializer() },
            { typeof(Uri), new  BuiltinSerializers.UriXlsxSerializer() },
            { typeof(Version), new  BuiltinSerializers.VersionXlsxSerializer() },
            { typeof(System.Numerics.BigInteger), new  BuiltinSerializers.BigIntegerXlsxSerializer() },
            { typeof(System.Numerics.Complex), new  BuiltinSerializers.ComplexXlsxSerializer() },
            { typeof(IntPtr), new  BuiltinSerializers.IntPtrXlsxSerializer() },
            { typeof(UIntPtr), new  BuiltinSerializers.UIntPtrXlsxSerializer() },
#if NET5_0_OR_GREATER
            { typeof(System.Text.Rune), new  BuiltinSerializers.RuneXlsxSerializer() },
#endif
#if NET5_0_OR_GREATER
            { typeof(Half), new  BuiltinSerializers.HalfXlsxSerializer() },
#endif
#if NET7_0_OR_GREATER
            { typeof(Int128), new  BuiltinSerializers.Int128XlsxSerializer() },
            { typeof(UInt128), new  BuiltinSerializers.UInt128XlsxSerializer() },
#endif
#if NET6_0_OR_GREATER
            { typeof(DateOnly), new  BuiltinSerializers.DateOnlyXlsxSerializer() },
            { typeof(TimeOnly), new  BuiltinSerializers.TimeOnlyXlsxSerializer() },
#endif
    };

    public IXlsxSerializer<T>? GetSerializer<T>()
    {
        if (serializers.TryGetValue(typeof(T), out var value))
        {
            return (IXlsxSerializer<T>)value;
        }
        return null;
    }
}