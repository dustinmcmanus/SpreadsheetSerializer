using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Data;
using System.Linq;

namespace SpreadsheetSerializer.AsposeCells
{
    internal class WorksheetDeserializer
    {
        private readonly string worksheetName;
        private readonly Type genericListType;
        private readonly IList list;

        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings()
        {
            Converters = { new JsonConverterDefault() }
        };

        public WorksheetDeserializer WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }


        public WorksheetDeserializer(string worksheetName, IList list, Type genericListType)
        {
            this.worksheetName = worksheetName ?? throw new ArgumentNullException(nameof(worksheetName));
            this.list = list ?? throw new ArgumentNullException(nameof(list));
            this.genericListType = genericListType ?? throw new ArgumentNullException(nameof(genericListType));
        }

        public void Deserialize(AsposeWorkbook workbook)
        {
            list.Clear();
            AddWorkbookRowsToList(workbook);
        }

        private void AddWorkbookRowsToList(AsposeWorkbook workbook)
        {
            using (DataTable dt = GetDataTable(workbook))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var obj = GetObjectFromDataTableRow(dt, i);
                    list.Add(obj);
                }
            }
        }

        private DataTable GetDataTable(AsposeWorkbook workbook)
        {
            var sheet = workbook.Worksheets[worksheetName];
            int lastRow = sheet.Cells.Rows.Count;
            int columnCount = genericListType.GetProperties().Length;
            var options = new ExportTableOptions();
            options.ExportColumnName = true;
            options.ExportAsString = true;
            return sheet.Cells.ExportDataTable(0, 0, lastRow, columnCount, options);
        }

        private object GetObjectFromDataTableRow(DataTable dt, int rowIndex)
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
