using Aspose.Cells;
using System.Data;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class AsposeWorkbook : Aspose.Cells.Workbook
    {
        public string WorkbookName { get; set; }

        public AsposeWorkbook()
        {
        }

        public AsposeWorkbook(FileFormatType fileFormatType) : base(fileFormatType)
        {
        }

        public AsposeWorkbook(string file) : base(file)
        {
        }

        public AsposeWorkbook(Stream stream) : base(stream)
        {
        }

        public AsposeWorkbook(string file, LoadOptions loadOptions) : base(file, loadOptions)
        {
        }

        public AsposeWorkbook(Stream stream, LoadOptions loadOptions) : base(stream, loadOptions)
        {
        }

        public AsposeWorkbook WithWorkbookName(string workbookName)
        {
            WorkbookName = workbookName;
            return this;
        }

        public void WriteDataTableToExcelTab(DataTable dataTable, string worksheetName)
        {
            var sheet = Worksheets.Add(worksheetName);
            var dataOptions = new Aspose.Cells.ImportTableOptions();

            sheet.Cells.ImportData(dataTable, 0, 0, dataOptions);
            sheet.AutoFitColumns();
        }

        public void Delete()
        {
            File.Delete(WorkbookName + ".xlsx");
        }

        public void RemoveDefaultTab()
        {
            var sheet = Worksheets["Sheet1"];
            Worksheets.RemoveAt(sheet.Index);
        }

        public void Save()
        {
            Save(WorkbookName + ".xlsx");
        }
    }
}
