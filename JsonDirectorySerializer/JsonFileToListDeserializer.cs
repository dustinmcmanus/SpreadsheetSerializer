using Newtonsoft.Json;

namespace JsonDirectorySerializer
{
    public static class JsonFileToClassDeserializer<T>
    {
        public static T Deserialize(string jsonFilePath)
        {
            return JsonConvert.DeserializeObject<T>(jsonFilePath);
        }
    }
}