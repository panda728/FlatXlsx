using System.Numerics;

namespace FlatXlsx.Serializers;

internal class BuiltinSerializers
{
    public sealed class StringXlsxSerializer : IXlsxSerializer<string?>
    {
        public void WriteTitle(XlsxWriter writer, string? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, string? value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class CharXlsxSerializer : IXlsxSerializer<char>
    {
        public void WriteTitle(XlsxWriter writer, char value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, char value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class GuidXlsxSerializer : IXlsxSerializer<Guid>
    {
        public void WriteTitle(XlsxWriter writer, Guid value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Guid value, XlsxSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class EnumXlsxSerializer : IXlsxSerializer<Enum>
    {
        public void WriteTitle(XlsxWriter writer, Enum value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Enum value, XlsxSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class DateTimeXlsxSerializer : IXlsxSerializer<DateTime>
    {
        public void WriteTitle(XlsxWriter writer, DateTime value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, DateTime value, XlsxSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class DateTimeOffsetXlsxSerializer : IXlsxSerializer<DateTimeOffset>
    {
        public void WriteTitle(XlsxWriter writer, DateTimeOffset value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, DateTimeOffset value, XlsxSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class TimeSpanXlsxSerializer : IXlsxSerializer<TimeSpan>
    {
        public void WriteTitle(XlsxWriter writer, TimeSpan value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, TimeSpan value, XlsxSerializerOptions options)
            => writer.Write(value.ToString());
    }

    public sealed class UriXlsxSerializer : IXlsxSerializer<Uri?>
    {
        public void WriteTitle(XlsxWriter writer, Uri? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Uri? value, XlsxSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
            writer.Write($"{value}");
        }
    }

    public sealed class VersionXlsxSerializer : IXlsxSerializer<Version?>
    {
        public void WriteTitle(XlsxWriter writer, Version? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Version? value, XlsxSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
            writer.Write(value.ToString());
        }
    }

    public sealed class BigIntegerXlsxSerializer : IXlsxSerializer<BigInteger>
    {
        public void WriteTitle(XlsxWriter writer, BigInteger value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, BigInteger value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value);
    }

    public sealed class ComplexXlsxSerializer : IXlsxSerializer<Complex>
    {
        public void WriteTitle(XlsxWriter writer, Complex value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Complex value, XlsxSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class IntPtrXlsxSerializer : IXlsxSerializer<IntPtr>
    {
        public void WriteTitle(XlsxWriter writer, IntPtr value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, IntPtr value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value.ToInt64());
    }

    public sealed class UIntPtrXlsxSerializer : IXlsxSerializer<UIntPtr>
    {
        public void WriteTitle(XlsxWriter writer, UIntPtr value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, UIntPtr value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value.ToUInt64());
    }

#if NET5_0_OR_GREATER
    public sealed class RuneXlsxSerializer : IXlsxSerializer<System.Text.Rune>
    {
        public void WriteTitle(XlsxWriter writer, System.Text.Rune value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, System.Text.Rune value, XlsxSerializerOptions options)
            => writer.Write(value.ToString());
    }
#endif

#if NET5_0_OR_GREATER
    public sealed class HalfXlsxSerializer : IXlsxSerializer<Half>
    {
        public void WriteTitle(XlsxWriter writer, Half value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Half value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value);
    }
#endif

#if NET7_0_OR_GREATER
    public sealed class Int128XlsxSerializer : IXlsxSerializer<Int128>
    {
        public void WriteTitle(XlsxWriter writer, Int128 value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, Int128 value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value);
    }

    public sealed class UInt128XlsxSerializer : IXlsxSerializer<UInt128>
    {
        public void WriteTitle(XlsxWriter writer, UInt128 value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, UInt128 value, XlsxSerializerOptions options)
            => writer.WritePrimitive(value);
    }
#endif

#if NET6_0_OR_GREATER
    public sealed class DateOnlyXlsxSerializer : IXlsxSerializer<DateOnly>
    {
        public void WriteTitle(XlsxWriter writer, DateOnly value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, DateOnly value, XlsxSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class TimeOnlyXlsxSerializer : IXlsxSerializer<TimeOnly>
    {
        public void WriteTitle(XlsxWriter writer, TimeOnly value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxWriter writer, TimeOnly value, XlsxSerializerOptions options)
            => writer.WriteDateTime(value);
    }
#endif
}
