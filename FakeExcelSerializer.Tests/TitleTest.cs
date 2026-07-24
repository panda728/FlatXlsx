using FluentAssertions;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;
using System.Text;

namespace FakeExcelSerializer.Tests
{
    public class TitleTest
    {
        void RunStringColumnTest<T>(
            T value1,
            string value1ShouldBe,
            string value2ShouldBe,
            string value3ShouldBe,
            string titleShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.WriteTitle(ref writer, value1, option);
                Assert.Equal(3, writer.SharedStrings.Count);

                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;
                var sharedString2 = writer.SharedStrings.Skip(1).First().Key;
                var sharedString3 = writer.SharedStrings.Skip(2).First().Key;

                columnXml.Should().Be(titleShouldBe);
                sharedString1.Should().Be(value1ShouldBe);
                sharedString2.Should().Be(value2ShouldBe);
                sharedString3.Should().Be(value3ShouldBe);
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

        void RunTest<T>(T value, string value1ShouldBe, string columnXmlShouldBe, ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.WriteTitle(ref writer, value, option);
                Assert.NotEmpty(writer.SharedStrings);
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

        [Fact]
        public void Serializer_WriteTitle()
        {
            var list = new List<TestData>()
            {
                new TestData(){Title  = "Title1", Name = "Name1", Address="Address1"},
                new TestData(){Title  = "Title2", Name = "Name2", Address="Address2"},
                new TestData(){Title  = "Title3", Name = "Name3", Address="Address3"},
            };

            var option = ExcelSerializerOptions.Default with
            {
                HasHeaderRecord = true,
            };

            RunStringColumnTest(
                list,
                "Address Ex",
                "Title Ex",
                "Name Ex",
                "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>2</v></c>",
                option);
        }

        [Fact]
        public void Serializer_ObjectFallback()
        {
            var value = (object)"key1";
            RunTest(value, "value",
                "<c t=\"s\"><v>0</v></c>",
                ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_tuple2()
        {
            var t = Tuple.Create(1, 2);
            RunTest(t, "value", "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>", ExcelSerializerOptions.Default);
        }
        [Fact]
        public void Serializer_IDictionary()
        {
            var dic = new Dictionary<string, int> { { "key1", 1 }, { "key2", 2 } };
            RunTest(dic, "key",
                "<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c>",
                ExcelSerializerOptions.Default);
        }
    }

    public class TestData
    {
        [DataMember(Name = "Title Ex", Order = 2)]
        public string Title { get; set; }
        [DataMember(Name = "Name Ex", Order = 3)]
        public string Name { get; set; }
        [DataMember(Name = "Address Ex", Order = 1)]
        public string Address { get; set; }
    }
}
