using Aspose.Cells;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class ExcelWorksheetSerializer<T> : ISpreadsheetSerializer<T>, IWorksheetSerializer
    {
        private WorkbookCreator workbook;
        private List<T> recordsToSerialize;
        private List<T> deserializedRecords;

        public string WorkbookName { get; set; } = "";
        public string ResourceName { get; set; } = "spreadsheet.xlsx";


        public ExcelWorksheetSerializer()
        {
        }

        public ExcelWorksheetSerializer(List<T> recordsToSerialize)
        {
            this.recordsToSerialize = recordsToSerialize;
        }

        public ExcelWorksheetSerializer(ref List<T> listReceivingDeserializedRecords)
        {
            this.deserializedRecords = listReceivingDeserializedRecords;
        }

        public ExcelWorksheetSerializer<T> WithWorkbookName(string workbookName)
        {
            // TODO: handle full file path
            string workbookNameWithoutExtension = workbookName;
            if (workbookName.EndsWith(".xlsx"))
            {
                workbookNameWithoutExtension = Path.GetFileNameWithoutExtension(workbookName);
            }

            this.WorkbookName = workbookNameWithoutExtension;
            return this;
        }

        public void Serialize(List<T> records, string spreadsheetName)
        {
            // Create excel workbook
            using (workbook = new WorkbookCreator(spreadsheetName))
            {
                CreateExcelTabForRecords(records, typeof(T).Name);
                workbook.RemoveDefaultTab();

                workbook.Delete();
                workbook.Save();
            }
            workbook = null;
        }

        public List<T> Deserialize(string tabName = "")
        {
            List<T> listOfSchedules;
            // Create excel workbook
            var assembly = typeof(WorkbookCreator).Assembly; //Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(ResourceName))
            {
                using (var workbook = new Aspose.Cells.Workbook(stream))
                {
                    if (string.IsNullOrEmpty(tabName))
                    {
                        tabName = typeof(T).Name;
                    }
                    var sheet = workbook.Worksheets[tabName];
                    int lastRow = sheet.Cells.Rows.Count;
                    int columnCount = typeof(T).GetProperties().Length;
                    var options = new ExportTableOptions();
                    options.ExportColumnName = true;
                    options.ExportAsString = true;

                    string json;
                    using (DataTable dt = sheet.Cells.ExportDataTable(0, 0, lastRow, columnCount, options))
                    {
                        json = JsonConvert.SerializeObject(dt, Formatting.Indented);
                    }
                    listOfSchedules = JsonConvert.DeserializeObject<List<T>>(json, new FormatConverter());
                }

            }
            workbook = null;
            return listOfSchedules;
        }

        public List<T> DeserializeWithNullStrings()
        {
            List<T> listOfSchedules;
            // Create excel workbook
            var assembly = typeof(WorkbookCreator).Assembly; //Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(ResourceName))
            {
                using (var workbook = new Aspose.Cells.Workbook(stream))
                {
                    string tabName = typeof(T).Name;
                    var sheet = workbook.Worksheets[tabName];
                    int lastRow = sheet.Cells.Rows.Count;
                    int columnCount = typeof(T).GetProperties().Length;
                    var options = new ExportTableOptions();
                    options.ExportColumnName = true;
                    options.ExportAsString = true;

                    string json;
                    using (DataTable dt = sheet.Cells.ExportDataTable(0, 0, lastRow, columnCount, options))
                    {
                        json = JsonConvert.SerializeObject(dt, Formatting.Indented);
                    }
                    listOfSchedules = JsonConvert.DeserializeObject<List<T>>(json, new FormatConverterWithNullStrings());
                }

            }
            workbook = null;
            return listOfSchedules;
        }

        private void CreateExcelTabForRecords(List<T> records, string tabName)
        {
            var dataTableCreator = new DataTableConverter<T>();
            var dataTable = dataTableCreator.CreateDataTableFor(records);
            WriteDataTableToExcelTab(dataTable, tabName);
        }

        private void WriteDataTableToExcelTab(DataTable dataTable, string tabName)
        {
            var sheet = workbook.Worksheets.Add(tabName);
            var dataOptions = new Aspose.Cells.ImportTableOptions();

            sheet.Cells.ImportData(dataTable, 0, 0, dataOptions);
            sheet.AutoFitColumns();
        }

        public void Serialize()
        {
            string tabName = typeof(T).Name;
            Serialize(recordsToSerialize, tabName);
        }

        public void Deserialize()
        {
            string tabName = typeof(T).Name;
            deserializedRecords = Deserialize(tabName);
        }
    }
}
