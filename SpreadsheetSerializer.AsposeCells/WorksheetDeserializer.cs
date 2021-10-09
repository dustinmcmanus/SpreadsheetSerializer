using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Data;
using System.Linq;

namespace SpreadsheetSerializer.AsposeCells
{
    public class WorksheetDeserializer
    {
        private readonly string worksheetName;
        private readonly Type genericListType;
        private readonly IList list;
        private JsonConverter jsonConverter;

        public WorksheetDeserializer(string worksheetName, IList list, Type genericListType)
        {
            this.worksheetName = worksheetName ?? throw new ArgumentNullException(nameof(worksheetName));
            this.list = list ?? throw new ArgumentNullException(nameof(list));
            this.genericListType = genericListType ?? throw new ArgumentNullException(nameof(genericListType));
            jsonConverter = new DefaultJsonConverter();
        }

        public void SetJsonConverter(JsonConverter jsonConverter)
        {
            this.jsonConverter = jsonConverter;
        }

        public void Deserialize(AsposeWorkbook workbook)
        {
            list.Clear();
            AddWorkbookRowsToList(workbook);
        }

        public void DeserializeWithNullStrings(AsposeWorkbook workbook)
        {
            jsonConverter = new JsonConverterWithNullStrings();
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
            ).ToString(Formatting.Indented);

            var obj = JsonConvert.DeserializeObject(json, genericListType, jsonConverter);
            return obj;
        }
    }
}
