using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookDeserializer<T>
    {
        public string WorkbookName { get; set; } = "";
        private List<WorksheetDeserializer> worksheetDeserializers = new List<WorksheetDeserializer>();

        public WorkbookDeserializer()
        {
        }

        public WorkbookDeserializer(string workbookName)
        {
            SetWorkbookName(workbookName);
        }

        public WorkbookDeserializer<T> WithWorkbookName(string workbookName)
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
                deserializers.Add(worksheetCreator);
            }

            return deserializers;
        }

        //public T Deserialize(string workbookName = "")
        //{
        //    throw new System.NotImplementedException();
        //}

        public void Deserialize(ref T workbookClass, string workbookFilePath)
        {
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            using (var workbook = WorkbookRetriever.GetWorkbookFromFilePath(workbookFilePath))
            {
                foreach (var worksheetDeserializer in worksheetDeserializers)
                {
                    worksheetDeserializer.Deserialize(workbook);
                }
            }
        }

        public void DeserializeWithNullStrings(ref T workbookClass, string workbookFilePath)
        {
            worksheetDeserializers = GetWorksheetDeserializers(workbookClass);
            using (var workbook = WorkbookRetriever.GetWorkbookFromFilePath(workbookFilePath))
            {
                foreach (var worksheetDeserializer in worksheetDeserializers)
                {
                    worksheetDeserializer.DeserializeWithNullStrings(workbook);
                }
            }
        }

        // https://stackoverflow.com/questions/2493215/create-list-of-variable-type
        public IList CreateList(Type genericListType)
        {
            Type listType = typeof(List<>).MakeGenericType(genericListType);
            return (IList)Activator.CreateInstance(listType);
        }
    }
}
