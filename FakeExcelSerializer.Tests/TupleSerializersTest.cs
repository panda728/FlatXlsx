using FluentAssertions;
using Xunit.Sdk;

namespace FakeExcelSerializer.Tests
{
    public partial class TupleSerializersTest
    {
        void RunTest<T>(
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
    }
}
