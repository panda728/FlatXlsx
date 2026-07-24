using FluentAssertions;

namespace FakeExcelSerializer.Tests
{
    public class BuiltinSerializersTest
    {
        void RunStringColumnTest<T>(
            T value1, T value2,
            string value1ShouldBe, string value2ShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                serializer.Serialize(ref writer, value2, option);
                serializer.Serialize(ref writer, value1, option);

                Assert.Equal(2, writer.SharedStrings.Count);

                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;
                var sharedString2 = writer.SharedStrings.Skip(1).First().Key;

                columnXml.Should().Be("<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c>");
                sharedString1.Should().Be(value1ShouldBe);
                sharedString2.Should().Be(value2ShouldBe);
            }
            catch
            {
                throw;
            }
            finally
            {
                writer.Dispose();
            }
        }

        void RunColumnTest<T>(
            T value1, string value1ShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);

                Assert.Empty(writer.SharedStrings);

                writer.ToString().Should().Be(value1ShouldBe);
            }
            catch
            {
                throw;
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void Serializer_string()
        {
            RunStringColumnTest(
                "column1", "column2",
                "column1", "column2",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_char()
        {
            RunStringColumnTest(
                'A', 'Z',
                "A", "Z",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Guid()
        {
            var guid1 = Guid.NewGuid();
            var guid2 = Guid.NewGuid();

            RunStringColumnTest(
                guid1, guid2,
                guid1.ToString(), guid2.ToString(),
                ExcelSerializerOptions.Default);
        }

        enum DayOfWeek
        {
            Mon, Tue, Wed, Thu, Fri, Sat, Sun
        }

        [Fact]
        public void Serializer_Enum()
        {
            RunStringColumnTest(
                DayOfWeek.Mon, DayOfWeek.Tue,
                DayOfWeek.Mon.ToString(), DayOfWeek.Tue.ToString(),
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_DateTime()
        {
            var value = new DateTime(2000, 1, 1);
            RunColumnTest(
                value,
                "<c t=\"d\" s=\"3\"><v>2000-01-01T00:00:00</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_DateTimeOffset()
        {
            var option = ExcelSerializerOptions.Default;
            var value1 = DateTimeOffset.Now;
            var value2 = DateTimeOffset.UtcNow;
            RunStringColumnTest(
                value1, value2,
                value1.ToString(option.CultureInfo), value2.ToString(option.CultureInfo),
                option);
        }

        [Fact]
        public void Serializer_TimeSpan()
        {
            var value1 = DateTime.Today.AddHours(10) - DateTime.Today;
            var value2 = DateTime.Today.AddHours(-10) - DateTime.Today;
            RunStringColumnTest(
                value1, value2,
                "10:00:00", "-10:00:00",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_Uri()
        {
            var value1 = new Uri("http://hoge.com/fuga");
            var value2 = new Uri("http://hoge.com/fugafuga");
            RunStringColumnTest(
                value1, value2,
                "http://hoge.com/fuga", "http://hoge.com/fugafuga",
                ExcelSerializerOptions.Default);
        }
#if NET6_0_OR_GREATER
        [Fact]
        public void Serializer_DateOnly()
        {
            var option = ExcelSerializerOptions.Default;
            var value1 = DateOnly.FromDateTime(new DateTime(2000, 1, 1));
            var value2 = DateOnly.FromDateTime(new DateTime(9999, 12, 31));
            RunColumnTest(value1, "<c t=\"d\" s=\"3\"><v>2000-01-01T00:00:00</v></c>", option);
            RunColumnTest(value2, "<c t=\"d\" s=\"3\"><v>9999-12-31T00:00:00</v></c>", option);
        }
        [Fact]
        public void Serializer_TimeOnly()
        {
            var option = ExcelSerializerOptions.Default;
            var value1 = TimeOnly.FromDateTime(new DateTime(2000, 1, 1, 0, 0, 0));
            var value2 = TimeOnly.FromDateTime(new DateTime(9999, 12, 31, 23, 59, 59));
            RunColumnTest(value1, "<c t=\"d\" s=\"4\"><v>1900-01-01T00:00:00</v></c>", option);
            RunColumnTest(value2, "<c t=\"d\" s=\"4\"><v>1900-01-01T23:59:59</v></c>", option);
        }
#endif
    }
}
