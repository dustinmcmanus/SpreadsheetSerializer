using Newtonsoft.Json;

namespace JsonDirectorySerializer
{
    public static class JsonDirectoryToClassDeserializer<T>
    {
        public static T Deserialize(string jsonDirectoryPath)
        {
            return JsonConvert.DeserializeObject<T>(JsonDirectoryToStringDeserializer.Deserialize(jsonDirectoryPath));
        }
    }
}