using FluentAssertions;
using JsonDirectorySerializer;
using NUnit.Framework;
using SpreadsheetSerializer.AsposeCells;
using SpreadsheetSerializer.Examples;
using System.IO;

namespace SpreadsheetSerializer.Tests.Unit
{
    public class SerializerTests
    {
        // JsonFileFromListSerializer requires the \ on the end of the directory path or it will create a file named DataSource1.json
        // none of the other tools require the \ at the end
        private const string DirectoryPath = @"..\..\..\DataSource1\";
        private ExampleWorkbook workbook;
        private readonly SingleJsonFileListTests singleJsonFileListTests = new SingleJsonFileListTests();

        /// <summary>
        /// Setup tests by setting the Aspose license and reading the DataSource1 JSON directory into the workbook object.
        /// Uncomment the Excel Editing section of this method to update test cases from saved changes in Excel.
        /// </summary>
        [SetUp]
        public void Setup()
        {
            // The Aspose license file should be dropped into the build folder (Ex: bin/Debug/net5.0/). It should never be checked into git.
            SpreadsheetSerializer.AsposeCells.License.SetLicense("Aspose.Total.lic");

            // Uncomment this section when editing test cases in Excel.
            // Otherwise comment out this section because the single source of truth is JSON, which can be checked into git and diffed
            // Excel Editing:
            //var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>();
            //var workbookExample = wbDeserializer.Deserialize(Path.Combine(DirectoryPath, nameof(ExampleWorkbook)));
            //JsonDirectorySerializer.JsonDirectorySerializer.Serialize(workbookExample, DirectoryPath);


            // Method 1: Json directory to class object instance
            string dataSource1ClassPath = Path.Combine(DirectoryPath, nameof(ExampleWorkbook));
            workbook = JsonDirectoryToClassDeserializer<ExampleWorkbook>.Deserialize(dataSource1ClassPath);

            // Method 2: Json directory to json string
            //string json = JsonDirectoryToStringDeserializer.Deserialize(dataSource1ClassPath);
            //workbook = JsonConvert.DeserializeObject<ExampleWorkbook>(json);
        }


        /// <summary>
        /// Uncomment this Test to create an Excel workbook from the JSON directory so that the test cases can be edited in Excel.
        /// Comment it out again after Excel workbook has been made.
        /// The Excel workbook does not need to be checked into git
        /// </summary>
        [Test]
        public void CreateExcelWorkbook()
        {
            //string dataSource1ClassPath = Path.Combine(DirectoryPath, nameof(ExampleWorkbook));
            //var workbookExample = JsonDirectoryToClassDeserializer<ExampleWorkbook>.Deserialize(dataSource1ClassPath);

            //var myWorkbookSerializer = new WorkbookSerializer<ExampleWorkbook>();
            //myWorkbookSerializer.Serialize(workbookExample, dataSource1ClassPath);
        }

        /// <summary>
        /// Uncomment this Test to create a JSON directory from an Excel workbook so that the test cases can be edited in JSON.
        /// Comment it out again after the JSON directory has been made because the SetUp method will read from JSON by default.
        /// The JSON directory test cases should be checked into git, but the Excel workbook does not need to be checked in.
        /// Other developers can re-create the Excel workbook by uncommenting the CreateExcelWorkbook() Test
        /// </summary>
        [Test]
        public void CreateJsonDirectory()
        {
            //var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>();
            //var workbookExample = wbDeserializer.Deserialize(Path.Combine(DirectoryPath, nameof(ExampleWorkbook)));
            //JsonDirectorySerializer.JsonDirectorySerializer.Serialize(workbookExample, DirectoryPath);
        }

        /// <summary>
        /// This test can be used to verify that the Excel workbook matches the JSON directory
        /// (after the Excel workbook has been created by executing the CreateExcelWorkbook() Test)
        /// </summary>
        [Test]
        public void DeserializeFromExcelMatchesDeserializeFromTextFiles()
        {
            // Uncomment this section or run the CreateExcelWorkbook() Test to create the Excel workbook from JSON directory first
            //var myWorkbookSerializer = new WorkbookSerializer<ExampleWorkbook>();
            //myWorkbookSerializer.Serialize(workbook, DirectoryPath);

            var wbDeserializer = new WorkbookDeserializer<ExampleWorkbook>();
            var myWorkbookExample = wbDeserializer.Deserialize(Path.Combine(DirectoryPath, nameof(ExampleWorkbook)));

            myWorkbookExample.Should().BeEquivalentTo(workbook);
        }
    }
}