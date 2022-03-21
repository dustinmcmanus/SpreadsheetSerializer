using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookSerializerSingleSheet<T>
    {
        private string WorkbookName { get; set; } = "";
        private string FilePath { get; set; } = "";
        private string WorksheetName { get; set; } = "";

        private WorksheetSerializer worksheetSerializer;

        /// <summary> 
        /// Writes out List object to an Excel Workbook file with a single Worksheet having rows for each object in the list.
        /// Only Lists should be serialized.
        /// Defaults to current working directory and [nameof(T)].xlsx filename and the worksheet named as nameof(T) also.
        /// </summary>
        /// <param name="list"></param>
        /// <param name="filePath"></param>
        /// <param name="worksheetName"></param>
        public void Serialize(IEnumerable<T> list, string filePath = "", string worksheetName = "")
        {
            var records = list.ToList();
            SetFileProperties(filePath);
            WorksheetName = worksheetName;

            if (string.IsNullOrEmpty(WorksheetName))
            {
                WorksheetName = typeof(T).Name;
            }

            // Populate WorksheetSerializers from Workbook properties on list
            worksheetSerializer = GetWorksheetSerializers(records);
            // Create excel workbook
            WriteWorkbook();
            DisposeDataTables();
        }

        private WorksheetSerializer GetWorksheetSerializers(List<T> list)
        {
            var dataTableCreator = new DataTableConverter();
            Type genericListType = typeof(T);
            var dataTable = dataTableCreator.CreateDataTableFor(list, genericListType);

            var worksheetCreator = new WorksheetSerializer(dataTable, WorksheetName);

            return worksheetCreator;
        }

        private void WriteWorkbook()
        {
            // Force creation through WorkbookCreator and WorkbookRetriever factory methods
            using (var workbook = WorkbookCreator.CreateWorkbookWithFilePath(FilePath))
            {
                CreateSheets(workbook);
                workbook.RemoveDefaultTab();
                workbook.Delete();
                workbook.Save();
            }
        }

        private void CreateSheets(AsposeWorkbook workbook)
        {
            worksheetSerializer.CreateSheet(workbook);
        }

        private void DisposeDataTables()
        {
            worksheetSerializer.DisposeDataTable();
        }

        private void SetFileProperties(string filePath)
        {
            FilePath = filePath;
            WorkbookName = Path.GetFileNameWithoutExtension(filePath);
            if (string.IsNullOrEmpty(WorkbookName))
            {
                WorkbookName = typeof(T).Name;
                FilePath = Path.Combine(FilePath, WorkbookName);
            }

            // if the file name does not have an extension, then add a default one for Excel
            if (!Path.HasExtension(FilePath))
            {
                FilePath += ".xlsx";
            }
        }
    }
}