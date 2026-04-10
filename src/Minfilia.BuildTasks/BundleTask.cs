using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.Build.Framework;
using Mono.Cecil;

namespace Minfilia.BuildTasks;

public class BundleTask : Microsoft.Build.Utilities.Task
{
    [Required]
    public string EntryPoint { get; set; } = string.Empty;

    [Required]
    public string Target { get; set; } = string.Empty;

    public override bool Execute()
    {
        try
        {
            ExecuteImpl();
            return true;
        }
        catch (Exception exception)
        {
            Log.LogErrorFromException(exception, true);
            return false;
        }
    }

    private void ExecuteImpl()
    {
        var outputDirectory = Path.GetDirectoryName(EntryPoint);
        var assemblyFiles = Directory.EnumerateFiles(outputDirectory)
            .Where(path => path.EndsWith(".dll") || path.EndsWith(".exe"))
            .ToDictionary(
                path => Path.GetFileName(path),
                path => File.ReadAllBytes(path));

        var assemblyModules = new Dictionary<string, ModuleDefinition>();
        foreach (var kv in assemblyFiles)
        {
            try
            {
                assemblyModules[kv.Key] = ModuleDefinition.ReadModule(new MemoryStream(kv.Value));
            }
            catch
            {
                // Skip non-.NET files
            }
        }

        var entryFileName = Path.GetFileName(EntryPoint);
        if (!assemblyModules.TryGetValue(entryFileName, out var entryModule))
            throw new Exception($"Entry {EntryPoint} not found in {outputDirectory}");

        // Include ALL non-system DLLs from output directory (not just BFS-reachable).
        // Many shim/facade assemblies are needed at runtime but not directly referenced.
        var systemPrefixes = new[] { "System.", "mscorlib", "netstandard", "Microsoft.CSharp" };
        var resourceFiles = new Dictionary<string, byte[]>();

        // Always include the entry point exe
        resourceFiles[entryFileName] = assemblyFiles[entryFileName];

        // Include all DLLs that are not framework assemblies
        foreach (var kv in assemblyFiles)
        {
            if (kv.Key == entryFileName) continue;
            if (!kv.Key.EndsWith(".dll")) continue;

            var isSystem = false;
            foreach (var prefix in systemPrefixes)
            {
                if (kv.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    // But include System.* packages from NuGet (they have corresponding modules)
                    if (assemblyModules.ContainsKey(kv.Key))
                    {
                        resourceFiles[kv.Key] = kv.Value;
                    }
                    isSystem = true;
                    break;
                }
            }
            if (!isSystem)
            {
                resourceFiles[kv.Key] = kv.Value;
            }
        }

        // Serialize resource pack
        var serialized = SerializeResourcePack(resourceFiles, entryFileName);

        // Compress with GZip
        var compressed = CompressGZip(serialized);
        Log.LogMessage(MessageImportance.High,
            $"BundleTask: {resourceFiles.Count} assemblies, {serialized.Length / 1024}KB -> {compressed.Length / 1024}KB (GZip)");

        // Embed into target exe
        using var targetStream = new FileStream(Target, FileMode.Open, FileAccess.ReadWrite);
        var targetModule = ModuleDefinition.ReadModule(targetStream);
        targetModule.Resources.Add(
            new EmbeddedResource("ResourcePack.gzip", ManifestResourceAttributes.Private, compressed));
        targetModule.Write(targetStream);
    }

    private static byte[] SerializeResourcePack(Dictionary<string, byte[]> files, string entryAssemblyName)
    {
        using var ms = new MemoryStream();
        using var bw = new BinaryWriter(ms);
        bw.Write(entryAssemblyName); // explicit entry assembly name
        bw.Write(files.Count);
        foreach (var kv in files)
        {
            var data = kv.Value;
            bw.Write(kv.Key);
            bw.Write(data.Length);
            bw.Write(data);
        }
        return ms.ToArray();
    }

    private static byte[] CompressGZip(byte[] data)
    {
        using var ms = new MemoryStream();
        using (var gzip = new GZipStream(ms, CompressionLevel.Optimal, leaveOpen: true))
        {
            gzip.Write(data, 0, data.Length);
        }
        return ms.ToArray();
    }
}
