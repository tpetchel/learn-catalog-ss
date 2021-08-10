using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using YamlDotNet.Serialization;

public abstract class MetadataBase : Dictionary<string, object>
{
    protected static T Convert<T>(Dictionary<string, object> data) where T : MetadataBase, new()
    {
        var t = new T();
        foreach (var kvp in data)
        {
            t[kvp.Key] = kvp.Value;
        }
        return t;
    }

    protected T Lookup<T>(string key) where T : class
    {
        if (ContainsKey(key))
        {
            return this[key] as T;
        }
        return default(T);
    }
}

public class ModuleMetadata : MetadataBase
{
    public string Uid { get => Lookup<string>("uid"); }

    public string Title { get => Lookup<string>("title"); }

    public string Author { get => Lookup<Dictionary<object, object>>("metadata")["ms.author"].ToString(); }

    public string Date { get
        {
            try
            {
                return Lookup<Dictionary<object, object>>("metadata")["ms.date"].ToString();
            }
            catch (KeyNotFoundException)
            {
                return string.Empty;
            }
        }
    }

    public List<string> Products { get => Lookup<List<object>>("products").Select(product => product.ToString()).ToList(); }

    public List<string> Units { get => Lookup<List<object>>("units").Select(unit => unit.ToString()).ToList(); }

    public void Dump()
    {
        Dump(this, 0);
    }

    private static void Dump<TKey>(Dictionary<TKey, object> dictionary, int level)
    {
        var indent = new string(' ', 2 * level);
        foreach (var kvp in dictionary)
        {
            if (kvp.Value is Dictionary<object, object>)
            {
                Console.WriteLine($"{indent}{kvp.Key}:");
                Dump(kvp.Value as Dictionary<object, object>, level + 1);
            }
            else
            {
                Console.WriteLine($"{indent}{kvp.Key}: {kvp.Value.ToString()}");
            }
        }
    }

    public static ModuleMetadata Load(string path)
    {
        using (var reader = new StreamReader(path))
        {
            var deserializer = new Deserializer();
            return deserializer.Deserialize<ModuleMetadata>(reader.ReadToEnd());
        }
    }
}