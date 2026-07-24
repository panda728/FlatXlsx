using FluentAssertions;

namespace FlatXlsx.Tests
{
    public class NewNetTypesTest
    {
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
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void Serializer_Half()
        {
            RunColumnTest((Half)1.5, "<c t=\"n\" s=\"6\"><v>1.5</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Int128()
        {
            RunColumnTest(Int128.MaxValue, "<c t=\"n\" s=\"5\"><v>170141183460469231731687303715884105727</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Int128_Negative()
        {
            RunColumnTest((Int128)(-12345), "<c t=\"n\" s=\"5\"><v>-12345</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_UInt128()
        {
            RunColumnTest(UInt128.MaxValue, "<c t=\"n\" s=\"5\"><v>340282366920938463463374607431768211455</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Nullable_Int128()
        {
            RunColumnTest((Int128?)42, "<c t=\"n\" s=\"5\"><v>42</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_Nullable_Half_Null()
        {
            RunColumnTest((Half?)null, "<c></c>",
                ExcelSerializerOptions.Default);
        }
    }
}
