using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace JsonDirectorySerializer
{
    public static class JsonFileFromListSerializer<T>
    {
        public static void Serialize(IEnumerable<T> list, string filePath = "")
        {
            var items = list.ToList();
            string path = filePath;
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = typeof(T).Name;
                path = Path.Combine(path, fileName);
            }

            // if the file name does not have an extension, then add a default one for Excel
            if (!Path.HasExtension(path))
            {
                path += ".json";
            }


            var sb = new StringBuilder();
            sb.Append("[");

            if (items.Any())
            {
                T lastItem = items.Last();
                foreach (var item in items)
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
            }

            sb.Append("]");

            string directoryPath = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            System.IO.File.WriteAllText(path, sb.ToString());
        }
    }
}