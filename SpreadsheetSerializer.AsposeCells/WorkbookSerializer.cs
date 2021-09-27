using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookSerializer<T>
    {
        public string WorkbookName { get; set; } = "";
        private List<WorksheetSerializer> worksheetSerializers = new List<WorksheetSerializer>();

        public WorkbookSerializer()
        {
        }

        public WorkbookSerializer(string workbookName)
        {
            SetWorkbookName(workbookName);
        }

        public WorkbookSerializer<T> WithWorkbookName(string workbookName)
        {
            SetWorkbookName(workbookName);
            return this;
        }

        private void SetWorkbookName(string workbookName)
        {
            string workbookNameWithoutExtension = workbookName;
            if (workbookName.EndsWith(".xlsx"))
            {
                workbookNameWithoutExtension = Path.GetFileNameWithoutExtension(workbookName);
            }

            this.WorkbookName = workbookNameWithoutExtension;
        }

        public void Serialize(T workbookClass)
        {
            // TODO: handle saving to full file path
            // Create excel workbook
            SetWorkbookNameIfUnspecified();
            //PopulateWorksheetSerializersFromWorkbookProperties(workbookClass);
            worksheetSerializers = GetWorksheetSerializers(workbookClass);
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
            // TODO: prevent instantiating AsposeWorkbook() directly without a name.
            // Force creation through WorkbookCreator and WorkbookRetriever factory methods
            using (var workbook = WorkbookCreator.CreateWorkbookWithName(WorkbookName))
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

        private void SetWorkbookNameIfUnspecified()
        {
            if (string.IsNullOrEmpty(WorkbookName))
            {
                var workbookClassName = typeof(T).Name;
                WorkbookName = workbookClassName;
            }
        }
    }
}
