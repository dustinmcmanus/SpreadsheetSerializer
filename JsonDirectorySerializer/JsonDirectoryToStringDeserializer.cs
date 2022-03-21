using Newtonsoft.Json.Linq;
using System.IO;

namespace JsonDirectorySerializer
{
    public class JsonDirectoryToStringDeserializer
    {
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

            var json = jObject.ToString();
            return json;
        }
    }
}