using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Minfilia.Http;
using Minfilia.Outlook;
using Minfilia.Tools;

namespace Minfilia;

internal static class Program
{
    private const int DefaultPort = 3027;

    static async Task Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        int port;
        try
        {
            port = ParsePort(args);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return;
        }

        Console.WriteLine("Initializing Outlook COM session...");
        using var session = new OutlookSession();

        try
        {
            await session.WaitForReadyAsync();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to connect to Outlook: {ex.Message}");
            Console.Error.WriteLine("Make sure Outlook is running.");
            return;
        }

        Console.WriteLine("Outlook connected.");

        // Explicit tool registration with injected dependencies
        var toolInstances = new object[]
        {
            new StoreTools(session),
            new SearchTools(session),
            new EmailTools(session),
            new CalendarTools(session),
            new ContactTools(session),
        };

        var tools = new McpServerPrimitiveCollection<McpServerTool>();
        foreach (var instance in toolInstances)
        {
            var type = instance.GetType();
            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            {
                if (method.GetCustomAttribute<McpServerToolAttribute>() != null)
                    tools.Add(McpServerTool.Create(method, instance));
            }
        }

        Console.WriteLine($"Registered {tools.Count} tools.");

        var serverOptions = new McpServerOptions
        {
            ServerInfo = new Implementation { Name = "Minfilia", Version = GetServerVersion() },
            Capabilities = new ServerCapabilities { Tools = new ToolsCapability() },
            ToolCollection = tools,
        };

        var prefix = $"http://localhost:{port}/";
        using var httpServer = new McpHttpServer(serverOptions);
        using var cts = new CancellationTokenSource();

        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

        try
        {
            await httpServer.RunAsync(prefix, cts.Token);
        }
        catch (OperationCanceledException) { }

        Console.WriteLine("Server stopped.");
    }

    private static int ParsePort(string[] args)
    {
        if (args.Length == 0)
            return DefaultPort;

        if (args.Length > 1)
            throw new ArgumentException("Invalid arguments: expected at most one port argument.");

        if (!int.TryParse(args[0], out var port))
            throw new ArgumentException($"Invalid port: '{args[0]}'. Expected an integer between 1 and 65535.");

        if (port < 1 || port > 65535)
            throw new ArgumentException($"Invalid port: {port}. Must be between 1 and 65535.");

        return port;
    }

    private static string GetServerVersion()
    {
        var version = typeof(Program).Assembly.GetName().Version;
        if (version == null)
            return "1.0.0";

        return version.Revision == 0 ? version.ToString(3) : version.ToString(4);
    }
}
