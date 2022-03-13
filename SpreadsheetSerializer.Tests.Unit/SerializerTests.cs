using FluentAssertions;
using JsonDirectorySerializer;
using Newtonsoft.Json;
using NUnit.Framework;
using SpreadsheetSerializer.AsposeCells;
using SpreadsheetSerializer.Examples;
using System.IO;

namespace SpreadsheetSerializer.Tests.Unit
{
    public class SerializerTests
    {
        private const string DirectoryPath = @"..\..\..\DataSource1";
        private ExampleWorkbook workbook;

        [SetUp]
        public void Setup()
        {
            string jsonDirectoryPath = Path.Combine(DirectoryPath, "ExampleWorkbook");
            workbook = JsonDirectoryToClassDeserializer<ExampleWorkbook>.Deserialize(jsonDirectoryPath);
            workbook = JsonConvert.DeserializeObject<ExampleWorkbook>(JsonDirectoryToStringDeserializer.Deserialize(jsonDirectoryPath));
        }

        [Test]
        public void Test1()
        {
            Assert.Pass();
        }

        [Test]
        public void DeserializeFromExcelMatchesDeserializeFromTextFiles()
        {
            var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>();
            var myWorkbookExample = wbDeserializer.Deserialize(Path.Combine(DirectoryPath, "ExampleWorkbook.xlsx"));

            myWorkbookExample.Should().BeEquivalentTo(workbook);
        }
    }
}