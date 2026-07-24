using System.Numerics;

namespace FlatXlsx.Tests
{
    public class BuiltinTypeCoverageTest
    {
        static (string Xml, string[] SharedStrings) Serialize<T>(T value, XlsxSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            var writer = new XlsxWriter(option);
            try
            {
                serializer!.Serialize(writer, value, option);
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
            var (xml, strings) = Serialize(new Version(1, 2, 3, 4), XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"s\"><v>0</v></c>", xml);
            Assert.Equal(new[] { "1.2.3.4" }, strings);
        }

        [Fact]
        public void Serializer_BigInteger()
        {
            var (xml, _) = Serialize(
                BigInteger.Parse("123456789012345678901234567890"),
                XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"n\" s=\"5\"><v>123456789012345678901234567890</v></c>", xml);
        }

        [Fact]
        public void Serializer_BigInteger_Negative()
        {
            var (xml, _) = Serialize(new BigInteger(-42), XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"n\" s=\"5\"><v>-42</v></c>", xml);
        }

        [Fact]
        public void Serializer_Complex()
        {
            var (xml, strings) = Serialize(new Complex(1, 2), XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"s\"><v>0</v></c>", xml);
            Assert.Equal(new[] { new Complex(1, 2).ToString(XlsxSerializerOptions.Default.CultureInfo) }, strings);
        }

        [Fact]
        public void Serializer_IntPtr()
        {
            var (xml, _) = Serialize((IntPtr)12345, XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"n\" s=\"5\"><v>12345</v></c>", xml);
        }

        [Fact]
        public void Serializer_UIntPtr()
        {
            var (xml, _) = Serialize((UIntPtr)12345, XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"n\" s=\"5\"><v>12345</v></c>", xml);
        }

        [Fact]
        public void Serializer_Nint()
        {
            nint value = -777;
            var (xml, _) = Serialize(value, XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"n\" s=\"5\"><v>-777</v></c>", xml);
        }

        [Fact]
        public void Serializer_Rune()
        {
            var (xml, strings) = Serialize(new System.Text.Rune('A'), XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"s\"><v>0</v></c>", xml);
            Assert.Equal(new[] { "A" }, strings);
        }

        [Fact]
        public void Serializer_Rune_Emoji()
        {
            var (xml, strings) = Serialize(new System.Text.Rune(0x1F600), XlsxSerializerOptions.Default);
            Assert.Equal("<c t=\"s\"><v>0</v></c>", xml);
            Assert.Equal(new[] { "\U0001F600" }, strings);
        }

        [Fact]
        public void Serializer_Nullable_BigInteger_Null()
        {
            var (xml, _) = Serialize((BigInteger?)null, XlsxSerializerOptions.Default);
            Assert.Equal("<c></c>", xml);
        }
    }
}
