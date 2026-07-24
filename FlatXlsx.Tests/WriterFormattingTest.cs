using FluentAssertions;

namespace FlatXlsx.Tests
{
    public class WriterFormattingTest
    {
        static string Serialize<T>(T value, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer!.Serialize(ref writer, value, option);
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
            Serialize(DateTime.MinValue, ExcelSerializerOptions.Default)
                .Should().Be("<c></c>");
        }

        [Fact]
        public void DateTime_DateOnlyValue_UsesDateStyle()
        {
            Serialize(new DateTime(2023, 1, 2), ExcelSerializerOptions.Default)
                .Should().Be("<c t=\"d\" s=\"3\"><v>2023-01-02T00:00:00</v></c>");
        }

        [Fact]
        public void DateTime_WithTime_UsesDateTimeStyle()
        {
            Serialize(new DateTime(2023, 1, 2, 3, 4, 5), ExcelSerializerOptions.Default)
                .Should().Be("<c t=\"d\" s=\"2\"><v>2023-01-02T03:04:05</v></c>");
        }

        [Fact]
        public void DateOnly_WritesDateCell()
        {
            Serialize(new DateOnly(2023, 1, 2), ExcelSerializerOptions.Default)
                .Should().Be("<c t=\"d\" s=\"3\"><v>2023-01-02T00:00:00</v></c>");
        }

        [Fact]
        public void TimeOnly_WritesTimeCell()
        {
            Serialize(new TimeOnly(3, 4, 5), ExcelSerializerOptions.Default)
                .Should().Be("<c t=\"d\" s=\"4\"><v>1900-01-01T03:04:05</v></c>");
        }

        [Fact]
        public void String_WithNewline_UsesWrapTextStyle()
        {
            var xml = Serialize("line1\nline2", ExcelSerializerOptions.Default);
            xml.Should().StartWith("<c t=\"s\" s=\"1\">");
        }

        [Fact]
        public void String_WithoutNewline_UsesPlainStringCell()
        {
            var xml = Serialize("plain", ExcelSerializerOptions.Default);
            xml.Should().StartWith("<c t=\"s\">");
        }

        [Fact]
        public void Decimal_MinValue_RoundTripsAllDigits()
        {
            Serialize(decimal.MinValue, ExcelSerializerOptions.Default)
                .Should().Be("<c t=\"n\" s=\"6\"><v>-79228162514264337593543950335</v></c>");
        }
    }
}
