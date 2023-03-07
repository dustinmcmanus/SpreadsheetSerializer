using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorkbookDeserializer<T>
    {
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();
        private Dictionary<string, PropertyInfo> propertyInfoByName = new Dictionary<string, PropertyInfo>();
        private T workbookClass;

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

            workbookClass = (T)Activator.CreateInstance(typeof(T));
            GetWorksheetDeserializers();
            DeserializeWorkboookFromFilePath(path);
            return workbookClass;
        }

        public T Deserialize(Stream asposeWorkbookStream)
        {
            workbookClass = (T)Activator.CreateInstance(typeof(T));
            GetWorksheetDeserializers();
            DeserializeWorkboookFromStream(asposeWorkbookStream);
            return workbookClass;
        }

        private void GetWorksheetDeserializers()
        {
            var propertyInfos = typeof(T).GetProperties();
            foreach (var propertyInfo in propertyInfos)
            {
                Type propertyType = propertyInfo.PropertyType;

                if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
                {
                    propertyInfoByName.Add(propertyInfo.Name, propertyInfo);
                }

            }
        }

        private IList AddWorkbookRowsToList(Worksheet worksheet, PropertyInfo propertyInfo)
        {
            IList list = null;

            string worksheetName = propertyInfo.Name;
            Type propertyType = propertyInfo.PropertyType;

            var genericListType = propertyType.GetGenericArguments()[0];
            list = (IList)propertyInfo.GetValue(workbookClass);

            using (DataTable dt = GetDataTable(worksheet, genericListType))
            {
                if (dt == null)
                {
                    list = null;
                    return list;
                }

                if (list == null)
                {
                    list = CreateList(genericListType);
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var obj = GetObjectFromDataTableRow(dt, i, genericListType);
                    list.Add(obj);
                }
            }

            return list;
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
            foreach (var worksheet in workbook.Worksheets)
            {
                if (propertyInfoByName.TryGetValue(worksheet.Name, out var propertyInfo))
                {
                    var list = AddWorkbookRowsToList(worksheet, propertyInfo);
                    propertyInfo.SetValue(workbookClass, list);
                }
            }
        }

        // https://stackoverflow.com/questions/2493215/create-list-of-variable-type
        private IList CreateList(Type genericListType)
        {
            Type listType = typeof(List<>).MakeGenericType(genericListType);
            return (IList)Activator.CreateInstance(listType);
        }

        private DataTable GetDataTable(Worksheet sheet, Type genericListType)
        {

            int lastRow = sheet.Cells.Rows.Count;
            int columnCount = genericListType.GetProperties().Length;
            var options = new ExportTableOptions();
            options.ExportColumnName = true;
            options.ExportAsString = true;
            return sheet.Cells.ExportDataTable(0, 0, lastRow, columnCount, options);
        }

        private object GetObjectFromDataTableRow(DataTable dt, int rowIndex, Type genericListType)
        {
            // adapted from https://stackoverflow.com/questions/43315700/json-net-how-to-serialize-just-one-row-from-a-datatable-object-without-it-being
            string json = new JObject(
                dt.Columns.Cast<DataColumn>()
                    .Select(c => new JProperty(c.ColumnName, JToken.FromObject(dt.Rows[rowIndex][c])))
            ).ToString(Formatting.None);

            var obj = JsonConvert.DeserializeObject(json, genericListType, serializerSettings);
            return obj;
        }

    }
}
