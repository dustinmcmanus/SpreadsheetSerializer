using System;
using System.Collections;
using System.Data;
using System.Linq;
using System.Reflection;

namespace SpreadsheetSerializer
{
    public class DataTableConverter
    {
        public DataTable CreateDataTableFor(IList records, Type typeOfObject)
        {
            var objectPropertiesUnordered = typeOfObject.GetProperties();
            var properties = from property in objectPropertiesUnordered
                             where Attribute.IsDefined(property, typeof(OrderAttribute))
                             orderby ((OrderAttribute)property
                                 .GetCustomAttributes(typeof(OrderAttribute), false)
                                 .Single()).Order
                             select property;

            var values = new object[objectPropertiesUnordered.Length];
            var propertiesList = properties.ToList();

            DataTable dataTable = new DataTable(typeOfObject.Name);
            foreach (PropertyInfo prop in propertiesList)
            {
                //Setting column names as Property names    
                dataTable.Columns.Add(prop.Name);
            }

            foreach (var record in records)
            {
                for (int i = 0; i < propertiesList.Count; i++)
                {
                    //inserting property values to datatable rows    
                    values[i] = propertiesList[i].GetValue(record, null);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

    }
}
