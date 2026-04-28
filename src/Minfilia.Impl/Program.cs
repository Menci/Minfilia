using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Minfilia.Http;
using Minfilia.Outlook;

namespace Minfilia;

internal static class Program
{
    private const int DefaultPort = 3027;

    static async Task Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        var executor = new OutlookOperationExecutor();
        var commands = ToolCatalog.CreateCommands(executor);

        try
        {
            Environment.ExitCode = await RunAsync(args, commands);
        }
        catch (CliUsageException ex)
        {
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine("Run `Minfilia.exe --help` for usage.");
            Environment.ExitCode = 2;
        }
        catch (McpException ex)
        {
            Console.Error.WriteLine(ex.Message);
            Environment.ExitCode = 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Environment.ExitCode = 1;
        }
    }

    private static async Task<int> RunAsync(string[] args, IReadOnlyList<ToolCommand> commands)
    {
        if (args.Length == 0)
        {
            CliRunner.WriteGlobalHelp(commands, DefaultPort);
            return 0;
        }

        if (CliRunner.IsHelpToken(args[0]))
        {
            CliRunner.WriteGlobalHelp(commands, DefaultPort);
            return 0;
        }

        if (string.Equals(args[0], "help", StringComparison.OrdinalIgnoreCase))
            return HandleHelpCommand(args, commands);

        if (string.Equals(args[0], "mcp", StringComparison.OrdinalIgnoreCase))
            return await RunMcpCommandAsync(GetRemainingArgs(args, 1), commands).ConfigureAwait(false);

        var command = FindCommand(commands, args[0])
            ?? throw new CliUsageException($"Unknown command: '{args[0]}'.");

        if (args.Length > 1 && CliRunner.IsHelpToken(args[1]))
        {
            if (args.Length > 2)
                throw new CliUsageException("Command help does not accept additional arguments.");

            CliRunner.WriteCommandHelp(command);
            return 0;
        }

        var invocationArguments = CliRunner.ParseArguments(command, GetRemainingArgs(args, 1));
        var result = await CliRunner.InvokeAsync(command, invocationArguments).ConfigureAwait(false);
        CliRunner.WriteJson(result);
        return 0;
    }

    private static int HandleHelpCommand(string[] args, IReadOnlyList<ToolCommand> commands)
    {
        if (args.Length == 1)
        {
            CliRunner.WriteGlobalHelp(commands, DefaultPort);
            return 0;
        }

        if (args.Length > 2)
            throw new CliUsageException("Usage: Minfilia.exe help [command]");

        if (string.Equals(args[1], "mcp", StringComparison.OrdinalIgnoreCase))
        {
            CliRunner.WriteMcpHelp(DefaultPort);
            return 0;
        }

        var command = FindCommand(commands, args[1])
            ?? throw new CliUsageException($"Unknown command: '{args[1]}'.");

        CliRunner.WriteCommandHelp(command);
        return 0;
    }

    private static async Task<int> RunMcpCommandAsync(string[] args, IReadOnlyList<ToolCommand> commands)
    {
        if (args.Length > 0 && CliRunner.IsHelpToken(args[0]))
        {
            if (args.Length > 1)
                throw new CliUsageException("Usage: Minfilia.exe mcp [--port <int>]");

            CliRunner.WriteMcpHelp(DefaultPort);
            return 0;
        }

        var port = ParseMcpPort(args);
        var tools = ToolCatalog.CreateMcpTools(commands);
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
            await httpServer.RunAsync(prefix, cts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
        }

        Console.WriteLine("Server stopped.");
        return 0;
    }

    private static int ParseMcpPort(string[] args)
    {
        if (args.Length == 0)
            return DefaultPort;

        if (args.Length == 1)
        {
            const string portPrefix = "--port=";
            if (args[0].StartsWith(portPrefix, StringComparison.OrdinalIgnoreCase))
                return ParsePort(args[0].Substring(portPrefix.Length));

            throw new CliUsageException("Usage: Minfilia.exe mcp [--port <int>]");
        }

        if (args.Length != 2 || !string.Equals(args[0], "--port", StringComparison.OrdinalIgnoreCase))
            throw new CliUsageException("Usage: Minfilia.exe mcp [--port <int>]");

        return ParsePort(args[1]);
    }

    private static ToolCommand? FindCommand(IReadOnlyList<ToolCommand> commands, string name)
    {
        foreach (var command in commands)
        {
            if (string.Equals(command.Name, name, StringComparison.OrdinalIgnoreCase))
                return command;
        }

        return null;
    }

    private static int ParsePort(string value)
    {
        if (!int.TryParse(value, out var port))
            throw new CliUsageException($"Invalid port: '{value}'. Expected an integer between 1 and 65535.");

        if (port < 1 || port > 65535)
            throw new CliUsageException($"Invalid port: {port}. Must be between 1 and 65535.");

        return port;
    }

    private static string[] GetRemainingArgs(string[] args, int startIndex)
    {
        if (startIndex >= args.Length)
            return Array.Empty<string>();

        var result = new string[args.Length - startIndex];
        Array.Copy(args, startIndex, result, 0, result.Length);
        return result;
    }

    private static string GetServerVersion()
    {
        var version = typeof(Program).Assembly.GetName().Version;
        if (version == null)
            return "1.0.0";

        return version.Revision == 0 ? version.ToString(3) : version.ToString(4);
    }
}
