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

        public WorksheetDeserializer(string worksheetName, IList list, Type genericListType)
        {
            this.worksheetName = worksheetName ?? throw new ArgumentNullException(nameof(worksheetName));
            this.list = list ?? throw new ArgumentNullException(nameof(list));
            this.genericListType = genericListType ?? throw new ArgumentNullException(nameof(genericListType));
        }

        public void Deserialize(AsposeWorkbook workbook)
        {
            var formatConverter = new FormatConverter();
            list.Clear();
            AddWorkbookRowsToList(genericListType, workbook, formatConverter);
        }

        public void DeserializeWithNullStrings(AsposeWorkbook workbook)
        {
            var formatConverter = new FormatConverterWithNullStrings();
            list.Clear();
            AddWorkbookRowsToList(genericListType, workbook, formatConverter);
        }

        private void AddWorkbookRowsToList(Type genericListType, AsposeWorkbook workbook, JsonConverter formatConverter)
        {
            using (DataTable dt = GetDataTable(genericListType, workbook))
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    var obj = GetObjectFromDataTableRow(genericListType, dt, i, formatConverter);
                    list.Add(obj);
                }
            }
        }

        private DataTable GetDataTable(Type genericListType, AsposeWorkbook workbook)
        {
            var sheet = workbook.Worksheets[worksheetName];
            int lastRow = sheet.Cells.Rows.Count;
            int columnCount = genericListType.GetProperties().Length;
            var options = new ExportTableOptions();
            options.ExportColumnName = true;
            options.ExportAsString = true;
            return sheet.Cells.ExportDataTable(0, 0, lastRow, columnCount, options);
        }

        private static object GetObjectFromDataTableRow(Type genericListType, DataTable dt, int rowIndex, JsonConverter formatConverter)
        {
            // adapted from https://stackoverflow.com/questions/43315700/json-net-how-to-serialize-just-one-row-from-a-datatable-object-without-it-being
            string json = new JObject(
                dt.Columns.Cast<DataColumn>()
                    .Select(c => new JProperty(c.ColumnName, JToken.FromObject(dt.Rows[rowIndex][c])))
            ).ToString(Formatting.Indented);

            var obj = JsonConvert.DeserializeObject(json, genericListType, formatConverter);
            return obj;
        }
    }
}
