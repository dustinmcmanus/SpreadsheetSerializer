using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class ExcelWorkbookSerializer
    {
        private WorkbookCreator workbook;

        public string WorkbookName { get; set; } = "";
        public string ResourceName { get; set; } = "spreadsheet.xlsx";
        public List<IWorksheetSerializer> WorkSheetSerializers;

        public ExcelWorkbookSerializer()
        {
        }

        public ExcelWorkbookSerializer(string workbookName)
        {
            SetWorkbookName(workbookName);
        }

        public ExcelWorkbookSerializer WithWorkbookName(string workbookName)
        {
            SetWorkbookName(workbookName);
            return this;
        }

        public void Serialize()
        {
            foreach (var worksheetSerializer in WorkSheetSerializers)
            {
                worksheetSerializer.Serialize();
            }
        }

        public void Deserialize()
        {
            foreach (var worksheetSerializer in WorkSheetSerializers)
            {
                worksheetSerializer.Deserialize();
            }
        }

        private void SetWorkbookName(string workbookName)
        {
            // TODO: handle full file path
            string workbookNameWithoutExtension = workbookName;
            if (workbookName.EndsWith(".xlsx"))
            {
                workbookNameWithoutExtension = Path.GetFileNameWithoutExtension(workbookName);
            }

            this.WorkbookName = workbookNameWithoutExtension;
        }
    }
}
