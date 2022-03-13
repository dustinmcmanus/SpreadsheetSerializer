using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace JsonDirectorySerializer
{
    public static class JsonFileFromListSerializer<T>
    {
        public static void Serialize(IList<T> list, string filePath = @"")
        {
            string className = typeof(T).Name;
            string jsonFilePath = filePath;

            if (string.IsNullOrEmpty(jsonFilePath))
            {
                jsonFilePath = Path.Combine(filePath, className);
            }

            if (!Path.HasExtension(jsonFilePath))
            {
                jsonFilePath += ".json";
            }

            T lastItem = list.Last();

            var sb = new StringBuilder();
            sb.Append("[");

            foreach (var item in list)
            {
                if (item != null)
                {
                    string record = JsonConvert.SerializeObject(item, Formatting.None);
                    sb.Append(record);

                    if (!ReferenceEquals(item, lastItem))
                    {
                        sb.AppendLine(",");
                    }

                }
            }
            sb.Append("]");

            System.IO.File.WriteAllText(jsonFilePath, sb.ToString());
        }
    }
}