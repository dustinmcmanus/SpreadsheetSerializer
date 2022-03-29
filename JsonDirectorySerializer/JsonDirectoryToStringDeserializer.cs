using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;

namespace JsonDirectorySerializer
{
    public class JsonDirectoryToStringDeserializer
    {
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();

        public JsonDirectoryToStringDeserializer WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }

        public string Deserialize(string jsonDirectoryPath)
        {
            string[] fileNames = Directory.GetFiles(jsonDirectoryPath);
            JObject jObject = new JObject();

            foreach (var fileName in fileNames)
            {
                string fileText = File.ReadAllText(fileName);
                JArray records = JArray.Parse(fileText);
                string keyName = Path.GetFileNameWithoutExtension(fileName);
                jObject[keyName] = records;
            }

            var json = JsonConvert.SerializeObject(jObject, serializerSettings);
            return json;
        }
    }
}