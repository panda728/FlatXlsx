using FluentAssertions;
using System.Collections.ObjectModel;

namespace FakeExcelSerializer.Tests
{
    public class SerializersTest
    {
        void RunTest<T>(T value, string value1ShouldBe, string columnXmlShouldBe, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value, option);
                Assert.Single(writer.SharedStrings);
                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;

                columnXml.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe);
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
        void RunTest<T>(T value, string value1ShouldBe1, string value1ShouldBe2, string columnXmlShouldBe, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value, option);
                Assert.Equal(2, writer.SharedStrings.Count);
                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;
                var sharedString2 = writer.SharedStrings.Skip(1).First().Key;

                columnXml.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe1);
                sharedString2.Should().Be(value1ShouldBe2);
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
        public void Serializer_TCollection()
        {
            var dinosaurs = new Collection<string>
            {
                "Psitticosaurus",
                "Caudipteryx"
            };
            RunTest(dinosaurs, "Psitticosaurus", "Caudipteryx",
                "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_IDictionary()
        {
            var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
            RunTest(dic, "key1", "key2",
                "<c t=\"s\"><v>0</v></c><c t=\"n\" s=\"5\"><v>1</v></c><c t=\"s\"><v>1</v></c><c t=\"n\" s=\"5\"><v>2</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_KeyValuePair()
        {
            var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
            RunTest(dic.First(), "key1",
                "<c t=\"s\"><v>0</v></c><c t=\"n\" s=\"5\"><v>1</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_ObjectFallback()
        {
            var value = (object)"key1";
            RunTest(value, "key1",
                "<c t=\"s\"><v>0</v></c>",
                ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_CompiledObject()
        {
            var potals1 = new Portal { Name = "Portal1", Owner = null, Level = 8 };
            CompiledObjectTest(potals1, "Portal1", "<c t=\"s\"><v>0</v></c><c></c><c t=\"n\" s=\"5\"><v>8</v></c>", ExcelSerializerOptions.Default);
        }

        void CompiledObjectTest<T>(
            T value,
            string value1ShouldBe,
            string columnXmlShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value, option);
                Assert.Single(writer.SharedStrings);
                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;

                columnXml.Should().Be(columnXmlShouldBe);
                sharedString1.Should().Be(value1ShouldBe);
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

        public class Portal
        {
            public string Name { get; set; } = "";
            public string? Owner { get; set; }
            public int Level { get; set; }
        }

    }
}
