
namespace FlatXlsx.Tests
{
    public class NewNetTypesTest
    {
        void RunColumnTest<T>(
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
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void Serializer_Half()
        {
            RunColumnTest((Half)1.5, "<c t=\"n\" s=\"6\"><v>1.5</v></c>",
                XlsxSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Int128()
        {
            RunColumnTest(Int128.MaxValue, "<c t=\"n\" s=\"5\"><v>170141183460469231731687303715884105727</v></c>",
                XlsxSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Int128_Negative()
        {
            RunColumnTest((Int128)(-12345), "<c t=\"n\" s=\"5\"><v>-12345</v></c>",
                XlsxSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_UInt128()
        {
            RunColumnTest(UInt128.MaxValue, "<c t=\"n\" s=\"5\"><v>340282366920938463463374607431768211455</v></c>",
                XlsxSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Nullable_Int128()
        {
            RunColumnTest((Int128?)42, "<c t=\"n\" s=\"5\"><v>42</v></c>",
                XlsxSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Nullable_Half_Null()
        {
            RunColumnTest((Half?)null, "<c></c>",
                XlsxSerializerOptions.Default);
        }
    }
}
