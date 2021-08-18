using System.Collections.Generic;

namespace SpreadsheetSerializer
{
    public interface ISpreadsheetSerializer<T>
    {
        void Serialize(List<T> records, string spreadsheetName);

        List<T> Deserialize(string tabName = "");
    }
}
