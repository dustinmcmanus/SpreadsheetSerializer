using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace JsonDirectorySerializer
{
    public class JsonFileToListDeserializer<T>
    {
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();

        public JsonFileToListDeserializer<T> WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }

        public List<T> Deserialize(string jsonFilePath = "")
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
            return JsonConvert.DeserializeObject<List<T>>(json, serializerSettings);
        }
    }
}