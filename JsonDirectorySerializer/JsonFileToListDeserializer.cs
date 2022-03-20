using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace JsonDirectorySerializer
{
    public static class JsonFileToListDeserializer<T>
    {
        public static List<T> Deserialize(string jsonFilePath = "")
        {
            string path = jsonFilePath;
            string fileName = Path.GetFileNameWithoutExtension(jsonFilePath);
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

            string json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<List<T>>(json);
        }
    }
}