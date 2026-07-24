
namespace FlatXlsx.Tests
{
    public class WriterFormattingTest
    {
        static string Serialize<T>(T value, XlsxSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            var writer = new XlsxWriter(option);
            try
            {
                serializer!.Serialize(writer, value, option);
                return writer.ToString();
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DateTime_MinValue_WritesSingleEmptyCell()
        {
            Assert.Equal("<c></c>", Serialize(DateTime.MinValue, XlsxSerializerOptions.Default));
        }

        [Fact]
        public void DateTime_DateOnlyValue_UsesDateStyle()
        {
            Assert.Equal("<c t=\"d\" s=\"3\"><v>2023-01-02T00:00:00</v></c>", Serialize(new DateTime(2023, 1, 2), XlsxSerializerOptions.Default));
        }

        [Fact]
        public void DateTime_WithTime_UsesDateTimeStyle()
        {
            Assert.Equal("<c t=\"d\" s=\"2\"><v>2023-01-02T03:04:05</v></c>", Serialize(new DateTime(2023, 1, 2, 3, 4, 5), XlsxSerializerOptions.Default));
        }

        [Fact]
        public void DateOnly_WritesDateCell()
        {
            Assert.Equal("<c t=\"d\" s=\"3\"><v>2023-01-02T00:00:00</v></c>", Serialize(new DateOnly(2023, 1, 2), XlsxSerializerOptions.Default));
        }

        [Fact]
        public void TimeOnly_WritesTimeCell()
        {
            Assert.Equal("<c t=\"d\" s=\"4\"><v>1900-01-01T03:04:05</v></c>", Serialize(new TimeOnly(3, 4, 5), XlsxSerializerOptions.Default));
        }

        [Fact]
        public void String_WithNewline_UsesWrapTextStyle()
        {
            var xml = Serialize("line1\nline2", XlsxSerializerOptions.Default);
            Assert.StartsWith("<c t=\"s\" s=\"1\">", xml);
        }

        [Fact]
        public void String_WithoutNewline_UsesPlainStringCell()
        {
            var xml = Serialize("plain", XlsxSerializerOptions.Default);
            Assert.StartsWith("<c t=\"s\">", xml);
        }

        [Fact]
        public void Decimal_MinValue_RoundTripsAllDigits()
        {
            Assert.Equal("<c t=\"n\" s=\"6\"><v>-79228162514264337593543950335</v></c>", Serialize(decimal.MinValue, XlsxSerializerOptions.Default));
        }
    }
}
