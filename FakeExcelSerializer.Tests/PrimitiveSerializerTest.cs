using FluentAssertions;

namespace FakeExcelSerializer.Tests
{
    public partial class PrimitiveSerializerTest
    {
        internal void RunIntegerTest<T>(T value1, T value2, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                serializer.Serialize(ref writer, value2, option);
                Assert.Empty(writer.SharedStrings);
                writer.ToString().Should().Be($"<c t=\"n\" s=\"5\"><v>{value1}</v></c><c t=\"n\" s=\"5\"><v>{value2}</v></c>");
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
        internal void RunNumberTest<T>(T value1, T value2, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                serializer.Serialize(ref writer, value2, option);
                Assert.Empty(writer.SharedStrings);
                writer.ToString().Should().Be($"<c t=\"n\" s=\"6\"><v>{value1}</v></c><c t=\"n\" s=\"6\"><v>{value2}</v></c>");
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
        public void Serializer_Boolean()
        {
            var option = ExcelSerializerOptions.Default;
            var serializer = option.GetSerializer<Boolean>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var value1 = true;
            var value2 = false;

            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                serializer.Serialize(ref writer, value2, option);
                Assert.Empty(writer.SharedStrings);
                writer.ToString().Should().Be($"<c t=\"b\"><v>1</v></c><c t=\"b\"><v>0</v></c>");
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
    }
}
