using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookDeserializerSingleSheet<T>
    {
        private WorksheetDeserializer worksheetDeserializer;
        private JsonConverter jsonConverter;

        public WorkbookDeserializerSingleSheet()
        {
        }

        public WorkbookDeserializerSingleSheet<T> WithJsonConverterForNullStrings()
        {
            this.jsonConverter = new JsonConverterWithNullStrings();
            return this;
        }

        // allow user to override with any custom NewtonSoft JsonConverter
        public WorkbookDeserializerSingleSheet<T> WithJsonConverter(JsonConverter converter)
        {
            this.jsonConverter = converter;
            return this;
        }

        public List<T> Deserialize(string workbookFilePath = "", string worksheetName = "")
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                worksheetName = typeof(T).Name; // Path.GetFileNameWithoutExtension(workbookFilePath);
            }

            string path = workbookFilePath;
            string fileName = Path.GetFileName(workbookFilePath);
            if (string.IsNullOrEmpty(fileName))
            {
                string directory = Path.GetDirectoryName(path);
                string fileNameWithoutExtension = typeof(T).Name;
                if (!string.IsNullOrEmpty(directory))
                {
                    path = Path.Combine(directory, fileNameWithoutExtension);
                }
                else
                {
                    path = fileNameWithoutExtension;
                }
            }

            // if the file name does not have an extension, then add a default one for Excel
            if (!Path.HasExtension(path))
            {
                path += ".xlsx";
            }


            List<T> workbookClass = new List<T>(); //(T)Activator.CreateInstance(typeof(T));
            worksheetDeserializer = GetWorksheetDeserializers(workbookClass, worksheetName);
            DeserializeWorkboookFromFilePath(path);
            return workbookClass;
        }

        private WorksheetDeserializer GetWorksheetDeserializers(List<T> list, string worksheetName)
        {
            Type genericListType = typeof(T);

            var worksheetCreator = new WorksheetDeserializer(worksheetName, list, genericListType);

            if (jsonConverter != null)
            {
                worksheetCreator.SetJsonConverter(jsonConverter);
            }

            return worksheetCreator;
        }

        private void DeserializeWorkboookFromFilePath(string workbookFilePath)
        {
            using (var workbook = WorkbookRetriever.GetWorkbookFromFilePath(workbookFilePath))
            {
                PopulateWorkbookClassListsFromWorkbook(workbook);
            }
        }

        private void PopulateWorkbookClassListsFromWorkbook(AsposeWorkbook workbook)
        {
            worksheetDeserializer.Deserialize(workbook);
        }
    }
}
