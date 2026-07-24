using System.Numerics;

namespace FlatXlsx.Serializers;

internal class BuiltinSerializers
{
    public sealed class StringExcelSerializer : IExcelSerializer<string?>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, string? value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, string? value, ExcelSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class CharExcelSerializer : IExcelSerializer<char>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, char value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, char value, ExcelSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class GuidExcelSerializer : IExcelSerializer<Guid>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Guid value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Guid value, ExcelSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class EnumExcelSerializer : IExcelSerializer<Enum>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Enum value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Enum value, ExcelSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class DateTimeExcelSerializer : IExcelSerializer<DateTime>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class DateTimeOffsetExcelSerializer : IExcelSerializer<DateTimeOffset>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, DateTimeOffset value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, DateTimeOffset value, ExcelSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class TimeSpanExcelSerializer : IExcelSerializer<TimeSpan>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, TimeSpan value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, TimeSpan value, ExcelSerializerOptions options)
            => writer.Write(value.ToString());
    }

    public sealed class UriExcelSerializer : IExcelSerializer<Uri>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Uri value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Uri value, ExcelSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
            writer.Write($"{value}");
        }
    }

    public sealed class VersionExcelSerializer : IExcelSerializer<Version>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Version value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Version value, ExcelSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
            writer.Write(value.ToString());
        }
    }

    public sealed class BigIntegerExcelSerializer : IExcelSerializer<BigInteger>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, BigInteger value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, BigInteger value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value);
    }

    public sealed class ComplexExcelSerializer : IExcelSerializer<Complex>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Complex value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Complex value, ExcelSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class IntPtrExcelSerializer : IExcelSerializer<IntPtr>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, IntPtr value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, IntPtr value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value.ToInt64());
    }

    public sealed class UIntPtrExcelSerializer : IExcelSerializer<UIntPtr>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, UIntPtr value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, UIntPtr value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value.ToUInt64());
    }

#if NET5_0_OR_GREATER
    public sealed class RuneExcelSerializer : IExcelSerializer<System.Text.Rune>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, System.Text.Rune value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, System.Text.Rune value, ExcelSerializerOptions options)
            => writer.Write(value.ToString());
    }
#endif

#if NET5_0_OR_GREATER
    public sealed class HalfExcelSerializer : IExcelSerializer<Half>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Half value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Half value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value);
    }
#endif

#if NET7_0_OR_GREATER
    public sealed class Int128ExcelSerializer : IExcelSerializer<Int128>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, Int128 value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, Int128 value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value);
    }

    public sealed class UInt128ExcelSerializer : IExcelSerializer<UInt128>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, UInt128 value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, UInt128 value, ExcelSerializerOptions options)
            => writer.WritePrimitive(value);
    }
#endif

#if NET6_0_OR_GREATER
    public sealed class DateOnlyExcelSerializer : IExcelSerializer<DateOnly>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, DateOnly value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, DateOnly value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class TimeOnlyExcelSerializer : IExcelSerializer<TimeOnly>
    {
        public void WriteTitle(ref ExcelSerializerWriter writer, TimeOnly value, ExcelSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(ref ExcelSerializerWriter writer, TimeOnly value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }
#endif
}
