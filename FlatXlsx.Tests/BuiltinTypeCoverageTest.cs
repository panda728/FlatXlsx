using System.Numerics;
using FluentAssertions;

namespace FlatXlsx.Tests
{
    public class BuiltinTypeCoverageTest
    {
        static (string Xml, string[] SharedStrings) Serialize<T>(T value, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer!.Serialize(ref writer, value, option);
                return (writer.ToString(), writer.SharedStrings.Keys.ToArray());
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void Serializer_Version()
        {
            var (xml, strings) = Serialize(new Version(1, 2, 3, 4), ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"s\"><v>0</v></c>");
            strings.Should().Equal("1.2.3.4");
        }

        [Fact]
        public void Serializer_BigInteger()
        {
            var (xml, _) = Serialize(
                BigInteger.Parse("123456789012345678901234567890"),
                ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"n\" s=\"5\"><v>123456789012345678901234567890</v></c>");
        }

        [Fact]
        public void Serializer_BigInteger_Negative()
        {
            var (xml, _) = Serialize(new BigInteger(-42), ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"n\" s=\"5\"><v>-42</v></c>");
        }

        [Fact]
        public void Serializer_Complex()
        {
            var (xml, strings) = Serialize(new Complex(1, 2), ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"s\"><v>0</v></c>");
            strings.Should().Equal(new Complex(1, 2).ToString(ExcelSerializerOptions.Default.CultureInfo));
        }

        [Fact]
        public void Serializer_IntPtr()
        {
            var (xml, _) = Serialize((IntPtr)12345, ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"n\" s=\"5\"><v>12345</v></c>");
        }

        [Fact]
        public void Serializer_UIntPtr()
        {
            var (xml, _) = Serialize((UIntPtr)12345, ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"n\" s=\"5\"><v>12345</v></c>");
        }

        [Fact]
        public void Serializer_Nint()
        {
            nint value = -777;
            var (xml, _) = Serialize(value, ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"n\" s=\"5\"><v>-777</v></c>");
        }

        [Fact]
        public void Serializer_Rune()
        {
            var (xml, strings) = Serialize(new System.Text.Rune('A'), ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"s\"><v>0</v></c>");
            strings.Should().Equal("A");
        }

        [Fact]
        public void Serializer_Rune_Emoji()
        {
            var (xml, strings) = Serialize(new System.Text.Rune(0x1F600), ExcelSerializerOptions.Default);
            xml.Should().Be("<c t=\"s\"><v>0</v></c>");
            strings.Should().Equal("\U0001F600");
        }

        [Fact]
        public void Serializer_Nullable_BigInteger_Null()
        {
            var (xml, _) = Serialize((BigInteger?)null, ExcelSerializerOptions.Default);
            xml.Should().Be("<c></c>");
        }
    }
}
