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
        private JsonConverter jsonConverter;

        public WorkbookDeserializer()
        {
        }

        public WorkbookDeserializer<T> WithNullStringJsonConverter()
        {
            this.jsonConverter = new JsonConverterWithNullStrings();
            return this;
        }

        // allow user to override with any custom NewtonSoft JsonConverter
        public WorkbookDeserializer<T> WithJsonConverter(JsonConverter jsonConverter)
        {
            this.jsonConverter = jsonConverter;
            return this;
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

                var worksheetCreator = new WorksheetDeserializer(worksheetName, propertyListValue, genericListType);

                if (jsonConverter != null)
                {
                    worksheetCreator.SetJsonConverter(jsonConverter);
                }
                deserializers.Add(worksheetCreator);
            }

            return deserializers;
        }

        public T Deserialize(string workbookFilePath)
        {
            T workbookClass = (T)Activator.CreateInstance(typeof(T));
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            DeserializeWorkboookFromFilePath(workbookFilePath);
            return workbookClass;
        }

        public T Deserialize(Stream workbookStream)
        {
            T workbookClass = (T)Activator.CreateInstance(typeof(T));
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            DeserializeWorkboookFromStream(workbookStream);
            return workbookClass;
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
