using Xunit.Sdk;

namespace FlatXlsx.Tests
{
    public partial class TupleSerializersTest
    {
        void RunTest<T>(
            T value1, string value1ShouldBe,
            XlsxSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new XlsxWriter(option);
            try
            {
                serializer.Serialize(writer, value1, option);
                Assert.Empty(writer.SharedStrings);
                Assert.Equal(value1ShouldBe, writer.ToString());
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
