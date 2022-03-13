using System.Collections;
using System.Data;

namespace SpreadsheetSerializer
{
    public class DataTableConverter<T> : DataTableConverter
    {
        public DataTable CreateDataTableFor(IList records)
        {
            var typeOfObject = typeof(T);
            return base.CreateDataTableFor(records, typeOfObject);
        }

    }
}
