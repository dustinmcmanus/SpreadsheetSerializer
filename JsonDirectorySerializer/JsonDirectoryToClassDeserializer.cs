using Newtonsoft.Json;

namespace JsonDirectorySerializer
{
    public class JsonDirectoryToClassDeserializer<T>
    {
        public T Deserialize(string jsonDirectoryPath)
        {
            return JsonConvert.DeserializeObject<T>(new JsonDirectoryToStringDeserializer().Deserialize(jsonDirectoryPath));
        }
    }
}