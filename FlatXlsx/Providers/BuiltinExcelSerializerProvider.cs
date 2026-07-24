using FlatXlsx.Serializers;

namespace FlatXlsx.Providers;

public sealed class BuiltinExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new BuiltinExcelSerializerProvider();
    readonly Dictionary<Type, IExcelSerializer> serializers = new()
        {
            { typeof(string), new BuiltinSerializers.StringExcelSerializer() },
            { typeof(char), new BuiltinSerializers.CharExcelSerializer() },
            { typeof(Guid), new  BuiltinSerializers.GuidExcelSerializer() },
            { typeof(Enum), new  BuiltinSerializers.EnumExcelSerializer() },
            { typeof(DateTime), new  BuiltinSerializers.DateTimeExcelSerializer() },
            { typeof(DateTimeOffset), new  BuiltinSerializers.DateTimeOffsetExcelSerializer() },
            { typeof(TimeSpan), new  BuiltinSerializers.TimeSpanExcelSerializer() },
            { typeof(Uri), new  BuiltinSerializers.UriExcelSerializer() },
            { typeof(Version), new  BuiltinSerializers.VersionExcelSerializer() },
            { typeof(System.Numerics.BigInteger), new  BuiltinSerializers.BigIntegerExcelSerializer() },
            { typeof(System.Numerics.Complex), new  BuiltinSerializers.ComplexExcelSerializer() },
            { typeof(IntPtr), new  BuiltinSerializers.IntPtrExcelSerializer() },
            { typeof(UIntPtr), new  BuiltinSerializers.UIntPtrExcelSerializer() },
#if NET5_0_OR_GREATER
            { typeof(System.Text.Rune), new  BuiltinSerializers.RuneExcelSerializer() },
#endif
#if NET5_0_OR_GREATER
            { typeof(Half), new  BuiltinSerializers.HalfExcelSerializer() },
#endif
#if NET7_0_OR_GREATER
            { typeof(Int128), new  BuiltinSerializers.Int128ExcelSerializer() },
            { typeof(UInt128), new  BuiltinSerializers.UInt128ExcelSerializer() },
#endif
#if NET6_0_OR_GREATER
            { typeof(DateOnly), new  BuiltinSerializers.DateOnlyExcelSerializer() },
            { typeof(TimeOnly), new  BuiltinSerializers.TimeOnlyExcelSerializer() },
#endif
    };

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        if (serializers.TryGetValue(typeof(T), out var value))
        {
            return (IExcelSerializer<T>)value;
        }
        return null;
    }
}