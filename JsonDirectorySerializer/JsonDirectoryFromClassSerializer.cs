using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;

namespace JsonDirectorySerializer
{
    public class JsonDirectoryFromClassSerializer
    {
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();

        public JsonDirectoryFromClassSerializer WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }

        public void Serialize(object classInstanceToSerialize, string directoryPath = "")
        {
            string className = classInstanceToSerialize.GetType().Name;
            string newDirectoryPath = Path.Combine(directoryPath, className);
            Directory.CreateDirectory(newDirectoryPath);
            string json = JsonConvert.SerializeObject(classInstanceToSerialize, serializerSettings);

            Serialize(json, newDirectoryPath);
        }

        private void Serialize(string json, string directoryPath)
        {
            var jObject = JObject.Parse(json);

            foreach (var child in jObject)
            {
                var key = child.Key;
                var value = child.Value;

                if (value != null)
                {
                    var sb = new StringBuilder();
                    sb.Append("[");

                    foreach (var token in value)
                    {
                        string record = token.ToString(Formatting.None);
                        sb.Append(record);

                        if (!ReferenceEquals(token, value.Last))
                        {
                            sb.AppendLine(",");
                        }
                    }

                    sb.Append("]");

                    string filePath = Path.Combine(directoryPath, key + ".json");
                    System.IO.File.WriteAllText(filePath, sb.ToString());
                }
            }
        }
    }
}