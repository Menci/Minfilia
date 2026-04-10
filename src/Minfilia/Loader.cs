using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;

namespace Minfilia;

public static class Loader
{
    private static readonly List<Assembly> _loadedAssemblies = new List<Assembly>();

    private static Assembly? OnAssemblyResolve(object sender, ResolveEventArgs e)
    {
        var requestedName = new AssemblyName(e.Name).Name;
        return _loadedAssemblies.FirstOrDefault(a => a.GetName().Name == requestedName);
    }

    public static void Main(string[] args)
    {
        var currentAssembly = Assembly.GetExecutingAssembly();
        var resourceNames = currentAssembly.GetManifestResourceNames();

        byte[]? packData = null;
        if (resourceNames.Contains("ResourcePack.gzip"))
        {
            using var stream = currentAssembly.GetManifestResourceStream("ResourcePack.gzip");
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            packData = DecompressGZip(ms.ToArray());
        }
        else if (resourceNames.Contains("ResourcePack.raw"))
        {
            using var stream = currentAssembly.GetManifestResourceStream("ResourcePack.raw");
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            packData = ms.ToArray();
        }

        if (packData == null)
            throw new Exception("ResourcePack not found in assembly. Was BundleTask run?");

        var (entryName, assemblies) = DeserializeResourcePack(packData);

        AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

        // Load all assemblies
        Assembly? entryAssembly = null;
        foreach (var kv in assemblies)
        {
            var loaded = Assembly.Load(kv.Value);
            _loadedAssemblies.Add(loaded);
            if (kv.Key == entryName)
                entryAssembly = loaded;
        }

        if (entryAssembly == null)
            throw new Exception($"Entry assembly '{entryName}' not found in resource pack.");

        var programType = entryAssembly.GetType("Minfilia.Program")
            ?? throw new Exception("Minfilia.Program not found in entry assembly.");

        var mainMethod = programType.GetMethod("Main", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)
            ?? throw new Exception("Main method not found in Minfilia.Program.");

        var result = mainMethod.Invoke(null, new object[] { args });
        if (result is System.Threading.Tasks.Task task)
            task.GetAwaiter().GetResult();
    }

    private static (string entryName, Dictionary<string, byte[]> assemblies) DeserializeResourcePack(byte[] data)
    {
        var assemblies = new Dictionary<string, byte[]>();
        using var ms = new MemoryStream(data);
        using var br = new BinaryReader(ms);
        var entryName = br.ReadString(); // explicit entry assembly name
        var count = br.ReadInt32();
        for (var i = 0; i < count; i++)
        {
            var name = br.ReadString();
            var bytes = br.ReadBytes(br.ReadInt32());
            assemblies[name] = bytes;
        }
        return (entryName, assemblies);
    }

    private static byte[] DecompressGZip(byte[] data)
    {
        using var input = new MemoryStream(data);
        using var gzip = new GZipStream(input, CompressionMode.Decompress);
        using var output = new MemoryStream();
        gzip.CopyTo(output);
        return output.ToArray();
    }
}
