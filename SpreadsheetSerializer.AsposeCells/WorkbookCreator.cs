using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookCreator : Aspose.Cells.Workbook
    {
        private readonly string workbookName;

        public WorkbookCreator(string workbookName)
        {
            this.workbookName = workbookName;

        }

        //public WorkbookCreator CreateWorkbook(string workbookName)
        //{
        //    this.workbookName = workbookName;
        //    var workbook = new WorkbookCreator();
        //    return workbook;
        //}

        public void Delete()
        {
            File.Delete(workbookName + ".xlsx");
        }

        public void RemoveDefaultTab()
        {
            var sheet = Worksheets["Sheet1"];
            Worksheets.RemoveAt(sheet.Index);
        }

        public void Save()
        {
            Save(workbookName + ".xlsx");
        }
    }
}
