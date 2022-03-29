using Newtonsoft.Json;

namespace JsonDirectorySerializer
{
    public class JsonDirectoryToClassDeserializer<T>
    {
        private JsonSerializerSettings serializerSettings = new JsonSerializerSettings();

        public JsonDirectoryToClassDeserializer<T> WithJsonSerializerSettings(JsonSerializerSettings settings)
        {
            this.serializerSettings = settings;
            return this;
        }

        public T Deserialize(string jsonDirectoryPath)
        {
            return JsonConvert.DeserializeObject<T>(new JsonDirectoryToStringDeserializer().Deserialize(jsonDirectoryPath), serializerSettings);
        }
    }
}