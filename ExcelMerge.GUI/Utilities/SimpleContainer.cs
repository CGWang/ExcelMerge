using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge.GUI.Utilities
{
    public interface IContainer
    {
        IContainer RegisterInstance<T>(string key, T instance);
        T Resolve<T>(string key);
        IEnumerable<T> ResolveAll<T>();
    }

    public class SimpleContainer : IContainer
    {
        private readonly Dictionary<(Type, string), object> _instances = new();

        public IContainer RegisterInstance<T>(string key, T instance)
        {
            _instances[(typeof(T), key)] = instance;
            return this;
        }

        public T Resolve<T>(string key)
        {
            if (_instances.TryGetValue((typeof(T), key), out var instance))
                return (T)instance;

            throw new InvalidOperationException($"No instance registered for type {typeof(T).Name} with key '{key}'.");
        }

        public IEnumerable<T> ResolveAll<T>()
        {
            return _instances
                .Where(kv => kv.Key.Item1 == typeof(T))
                .Select(kv => (T)kv.Value);
        }
    }
}
