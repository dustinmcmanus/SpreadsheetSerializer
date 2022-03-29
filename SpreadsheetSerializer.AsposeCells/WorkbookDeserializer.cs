using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookDeserializer<T>
    {
        private List<WorksheetDeserializer> worksheetDeserializers = new List<WorksheetDeserializer>();
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();

        public WorkbookDeserializer<T> WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }

        public WorkbookDeserializer()
        {
        }


        public T Deserialize(string workbookFilePath)
        {
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

            T workbookClass = (T)Activator.CreateInstance(typeof(T));
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            DeserializeWorkboookFromFilePath(path);
            return workbookClass;
        }

        public T Deserialize(Stream asposeWorkbookStream)
        {
            T workbookClass = (T)Activator.CreateInstance(typeof(T));
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            DeserializeWorkboookFromStream(asposeWorkbookStream);
            return workbookClass;
        }

        private List<WorksheetDeserializer> GetWorksheetDeserializers(T workbookClass)
        {
            List<WorksheetDeserializer> deserializers = new List<WorksheetDeserializer>();
            var propertyInfos = typeof(T).GetProperties();
            foreach (var propertyInfo in propertyInfos)
            {
                string worksheetName = propertyInfo.Name;
                Type propertyType = propertyInfo.PropertyType;

                var genericListType = propertyType.GetGenericArguments()[0];
                var propertyListValue = (IList)propertyInfo.GetValue(workbookClass);

                if (propertyListValue == null)
                {
                    propertyListValue = CreateList(genericListType);
                    propertyInfo.SetValue(workbookClass, propertyListValue);
                }

                var worksheetCreator = new WorksheetDeserializer(worksheetName, propertyListValue, genericListType).WithJsonSerializerSettings(serializerSettings);
                deserializers.Add(worksheetCreator);
            }

            return deserializers;
        }

        private void DeserializeWorkboookFromFilePath(string workbookFilePath)
        {
            using (var workbook = WorkbookRetriever.GetWorkbookFromFilePath(workbookFilePath))
            {
                PopulateWorkbookClassListsFromWorkbook(workbook);
            }
        }

        private void DeserializeWorkboookFromStream(Stream workbookStream)
        {
            using (var workbook = WorkbookRetriever.GetWorkbookFromStream(workbookStream))
            {
                PopulateWorkbookClassListsFromWorkbook(workbook);
            }
        }

        private void PopulateWorkbookClassListsFromWorkbook(AsposeWorkbook workbook)
        {
            foreach (var worksheetDeserializer in worksheetDeserializers)
            {
                worksheetDeserializer.Deserialize(workbook);
            }
        }

        // https://stackoverflow.com/questions/2493215/create-list-of-variable-type
        private IList CreateList(Type genericListType)
        {
            Type listType = typeof(List<>).MakeGenericType(genericListType);
            return (IList)Activator.CreateInstance(listType);
        }
    }
}
