using Newtonsoft.Json;
using System;

namespace SpreadsheetSerializer
{

    // from https://stackoverflow.com/questions/17199500/jsonconvert-string-to-integer-about-digit-grouping-symbol/17200536
    public class JsonConverterWithNullStrings : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (objectType == typeof(int))
            {
                if (reader.Value is null)
                {
                    return GetDefaultValue(objectType);
                }
                return Convert.ToInt32(string.IsNullOrWhiteSpace(reader.Value.ToString()) ? "0" : reader.Value.ToString());
            }
            else if (objectType == typeof(long))
            {
                if (reader.Value is null)
                {
                    return GetDefaultValue(objectType);
                }
                return Convert.ToInt64(string.IsNullOrWhiteSpace(reader.Value.ToString()) ? "0" : reader.Value.ToString());
            }

            return reader.Value;
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(int) || objectType == typeof(long);
        }

        private object GetDefaultValue(Type t)
        {
            if (t.IsValueType)
                return Activator.CreateInstance(t);

            return null;
        }
    }
}
