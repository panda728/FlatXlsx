// Derived from Cysharp/WebSerializer (MIT License, Copyright (c) 2022 Cysharp, Inc.).
// See THIRD-PARTY-NOTICES.txt in the repository root.
using System.Numerics;

namespace FlatXlsx.Serializers;

internal class BuiltinSerializers
{
    public sealed class StringXlsxSerializer : IXlsxSerializer<string?>
    {
        public void WriteTitle(XlsxCellWriter writer, string? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, string? value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class CharXlsxSerializer : IXlsxSerializer<char>
    {
        public void WriteTitle(XlsxCellWriter writer, char value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, char value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class GuidXlsxSerializer : IXlsxSerializer<Guid>
    {
        public void WriteTitle(XlsxCellWriter writer, Guid value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Guid value, XlsxSerializerOptions options)
#if NET8_0_OR_GREATER
            => writer.WriteInlineString(value);
#else
            => writer.Write($"{value}");
#endif
    }

    public sealed class EnumXlsxSerializer : IXlsxSerializer<Enum>
    {
        public void WriteTitle(XlsxCellWriter writer, Enum value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Enum value, XlsxSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class DateTimeXlsxSerializer : IXlsxSerializer<DateTime>
    {
        public void WriteTitle(XlsxCellWriter writer, DateTime value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, DateTime value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class DateTimeOffsetXlsxSerializer : IXlsxSerializer<DateTimeOffset>
    {
        public void WriteTitle(XlsxCellWriter writer, DateTimeOffset value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, DateTimeOffset value, XlsxSerializerOptions options)
#if NET8_0_OR_GREATER
            => writer.WriteInlineString(value, options.CultureInfo);
#else
            => writer.Write(value.ToString(options.CultureInfo));
#endif
    }

    public sealed class TimeSpanXlsxSerializer : IXlsxSerializer<TimeSpan>
    {
        public void WriteTitle(XlsxCellWriter writer, TimeSpan value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, TimeSpan value, XlsxSerializerOptions options)
#if NET8_0_OR_GREATER
            => writer.WriteInlineString(value);
#else
            => writer.Write(value.ToString());
#endif
    }

    public sealed class UriXlsxSerializer : IXlsxSerializer<Uri?>
    {
        public void WriteTitle(XlsxCellWriter writer, Uri? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Uri? value, XlsxSerializerOptions options)
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
        public void WriteTitle(XlsxCellWriter writer, Version? value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Version? value, XlsxSerializerOptions options)
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
        public void WriteTitle(XlsxCellWriter writer, BigInteger value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, BigInteger value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class ComplexXlsxSerializer : IXlsxSerializer<Complex>
    {
        public void WriteTitle(XlsxCellWriter writer, Complex value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Complex value, XlsxSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class IntPtrXlsxSerializer : IXlsxSerializer<IntPtr>
    {
        public void WriteTitle(XlsxCellWriter writer, IntPtr value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, IntPtr value, XlsxSerializerOptions options)
            => writer.Write(value.ToInt64());
    }

    public sealed class UIntPtrXlsxSerializer : IXlsxSerializer<UIntPtr>
    {
        public void WriteTitle(XlsxCellWriter writer, UIntPtr value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, UIntPtr value, XlsxSerializerOptions options)
            => writer.Write(value.ToUInt64());
    }

#if NET5_0_OR_GREATER
    public sealed class RuneXlsxSerializer : IXlsxSerializer<System.Text.Rune>
    {
        public void WriteTitle(XlsxCellWriter writer, System.Text.Rune value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, System.Text.Rune value, XlsxSerializerOptions options)
            => writer.Write(value.ToString());
    }
#endif

#if NET5_0_OR_GREATER
    public sealed class HalfXlsxSerializer : IXlsxSerializer<Half>
    {
        public void WriteTitle(XlsxCellWriter writer, Half value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Half value, XlsxSerializerOptions options)
            => writer.Write(value);
    }
#endif

#if NET7_0_OR_GREATER
    public sealed class Int128XlsxSerializer : IXlsxSerializer<Int128>
    {
        public void WriteTitle(XlsxCellWriter writer, Int128 value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, Int128 value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class UInt128XlsxSerializer : IXlsxSerializer<UInt128>
    {
        public void WriteTitle(XlsxCellWriter writer, UInt128 value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, UInt128 value, XlsxSerializerOptions options)
            => writer.Write(value);
    }
#endif

#if NET6_0_OR_GREATER
    public sealed class DateOnlyXlsxSerializer : IXlsxSerializer<DateOnly>
    {
        public void WriteTitle(XlsxCellWriter writer, DateOnly value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, DateOnly value, XlsxSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class TimeOnlyXlsxSerializer : IXlsxSerializer<TimeOnly>
    {
        public void WriteTitle(XlsxCellWriter writer, TimeOnly value, XlsxSerializerOptions options, string name = "value")
            => writer.Write(name);
        public void Serialize(XlsxCellWriter writer, TimeOnly value, XlsxSerializerOptions options)
            => writer.Write(value);
    }
#endif
}
