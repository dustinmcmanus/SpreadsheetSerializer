using Newtonsoft.Json;
using System;

namespace SpreadsheetSerializer
{

    // from https://stackoverflow.com/questions/17199500/jsonconvert-string-to-integer-about-digit-grouping-symbol/17200536
    public class FormatConverterWithNullStrings : JsonConverter
    {
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (objectType == typeof(int)) // || objectType == typeof(long)
            {
                if (reader.Value is null)
                {
                    return GetDefaultValue(objectType);
                }
                return Convert.ToInt32(reader.Value.ToString().Replace(".", string.Empty));
            }
            //else if (objectType == typeof(string))
            //{
            //    if (reader.Value is null)
            //    {
            //        return string.Empty;
            //    }
            //}

            return reader.Value;
        }

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(int); // || objectType == typeof(string)
        }

        private object GetDefaultValue(Type t)
        {
            if (t.IsValueType)
                return Activator.CreateInstance(t);

            return null;
        }
    }
}
