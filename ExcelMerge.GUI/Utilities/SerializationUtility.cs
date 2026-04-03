using YamlDotNet.Serialization;

namespace ExcelMerge.GUI.Utilities
{
    public static class SerializationUtility
    {
        public static T DeepClone<T>(T obj)
        {
            if (obj == null) return default;

            var serializer = new SerializerBuilder().Build();
            var yaml = serializer.Serialize(obj);
            var deserializer = new DeserializerBuilder().Build();
            return deserializer.Deserialize<T>(yaml);
        }
    }
}
