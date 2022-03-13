using System;
using System.Collections;
using System.Collections.Generic;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookSerializer<T>
    {
        public string WorkbookName { get; set; } = "";
        public string FilePath { get; set; } = "";

        private List<WorksheetSerializer> worksheetSerializers = new List<WorksheetSerializer>();

        public void Serialize(T workbookClass, string filePath = @"")
        {
            SetWorkbookNameIfUnspecified(filePath);
            // Populate WorksheetSerializers from Workbook properties on workbookClass
            worksheetSerializers = GetWorksheetSerializers(workbookClass);
            // Create excel workbook
            WriteWorkbook();
            DisposeDataTables();
        }

        private List<WorksheetSerializer> GetWorksheetSerializers(T workbookClass)
        {
            List<WorksheetSerializer> serializers = new List<WorksheetSerializer>();
            var propertyInfos = typeof(T).GetProperties();
            foreach (var propertyInfo in propertyInfos)
            {
                string worksheetName = propertyInfo.Name;
                Type propertyType = propertyInfo.PropertyType;

                var genericListType = propertyType.GetGenericArguments()[0];
                var propertyListValue = (IList)propertyInfo.GetValue(workbookClass);

                var dataTableCreator = new DataTableConverter();
                var dataTable = dataTableCreator.CreateDataTableFor(propertyListValue, genericListType);

                var worksheetCreator = new WorksheetSerializer(dataTable, worksheetName);
                serializers.Add(worksheetCreator);
            }

            return serializers;
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
            foreach (var worksheetSerializer in worksheetSerializers)
            {
                worksheetSerializer.CreateSheet(workbook);
            }
        }

        private void DisposeDataTables()
        {
            foreach (var worksheetSerializer in worksheetSerializers)
            {
                worksheetSerializer.DisposeDataTable();
            }
        }

        private void SetWorkbookNameIfUnspecified(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                var workbookClassName = typeof(T).Name;
                WorkbookName = workbookClassName;
                FilePath = workbookClassName + ".xlsx";
            }
        }
    }
}
