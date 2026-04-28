using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace Minfilia;

internal static class CliRunner
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static bool IsHelpToken(string value)
    {
        return string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "-h", StringComparison.OrdinalIgnoreCase);
    }

    public static void WriteGlobalHelp(IReadOnlyList<ToolCommand> commands, int defaultPort)
    {
        Console.WriteLine("Usage:");
        Console.WriteLine($"  Minfilia.exe mcp [--port <int>]     Start MCP server on port {defaultPort} by default.");
        Console.WriteLine("  Minfilia.exe <command> [--name value ...]  Run one CLI command and print JSON.");
        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.WriteLine("  mcp                  Start the stateless HTTP MCP server.");

        foreach (var command in commands)
            Console.WriteLine($"  {command.Name.PadRight(20)} {command.Description}");

        Console.WriteLine();
        Console.WriteLine("Use `Minfilia.exe help <command>` or `Minfilia.exe <command> --help` for details.");
        Console.WriteLine("Boolean options accept `true` or `false`; a bare `--flag` means `true`.");
    }

    public static void WriteMcpHelp(int defaultPort)
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  Minfilia.exe mcp [--port <int>]");
        Console.WriteLine();
        Console.WriteLine("Starts the stateless HTTP MCP server.");
        Console.WriteLine($"Default port: {defaultPort}");
    }

    public static void WriteCommandHelp(ToolCommand command)
    {
        Console.Write("Usage:\n  Minfilia.exe ");
        Console.Write(command.Name);

        foreach (var parameter in command.Parameters)
        {
            Console.Write(' ');
            Console.Write(FormatUsageParameter(parameter));
        }

        Console.WriteLine();

        if (!string.IsNullOrWhiteSpace(command.Description))
        {
            Console.WriteLine();
            Console.WriteLine(command.Description);
        }

        if (command.Parameters.Length == 0)
            return;

        Console.WriteLine();
        Console.WriteLine("Options:");
        foreach (var parameter in command.Parameters)
        {
            var name = GetParameterName(parameter);
            var description = parameter.GetCustomAttribute<DescriptionAttribute>()?.Description ?? string.Empty;
            var requirement = parameter.HasDefaultValue
                ? $"Optional. Default: {FormatDefaultValue(parameter.DefaultValue)}."
                : "Required.";
            var boolHint = IsBooleanParameter(parameter)
                ? " Bare `--flag` form means `true`."
                : string.Empty;

            Console.WriteLine($"  --{name} <{GetTypeName(parameter.ParameterType)}>  {description} {requirement}{boolHint}".TrimEnd());
        }
    }

    public static object?[] ParseArguments(ToolCommand command, IReadOnlyList<string> args)
    {
        var parameters = command.Parameters;
        var parameterByName = new Dictionary<string, ParameterInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var parameter in parameters)
            parameterByName.Add(GetParameterName(parameter), parameter);

        var rawValues = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Count; i++)
        {
            var token = args[i];
            if (!token.StartsWith("--", StringComparison.Ordinal))
                throw new CliUsageException($"Unexpected argument: '{token}'. Options must use --<name>.");

            var separatorIndex = token.IndexOf('=');
            var optionName = separatorIndex >= 0 ? token.Substring(2, separatorIndex - 2) : token.Substring(2);
            if (optionName.Length == 0)
                throw new CliUsageException($"Invalid option syntax: '{token}'.");

            if (!parameterByName.TryGetValue(optionName, out var parameter))
                throw new CliUsageException($"Unknown option '--{optionName}' for command '{command.Name}'.");

            var parameterName = GetParameterName(parameter);
            if (rawValues.ContainsKey(parameterName))
                throw new CliUsageException($"Option '--{parameterName}' was provided more than once.");

            string? rawValue;
            if (separatorIndex >= 0)
            {
                rawValue = token.Substring(separatorIndex + 1);
            }
            else if (IsBooleanParameter(parameter) && (i + 1 >= args.Count || args[i + 1].StartsWith("--", StringComparison.Ordinal)))
            {
                rawValue = "true";
            }
            else
            {
                if (i + 1 >= args.Count)
                    throw new CliUsageException($"Option '--{parameterName}' requires a value.");

                rawValue = args[++i];
            }

            rawValues.Add(parameterName, rawValue);
        }

        var invocationArguments = new object?[parameters.Length];
        for (var i = 0; i < parameters.Length; i++)
        {
            var parameter = parameters[i];
            var parameterName = GetParameterName(parameter);

            if (rawValues.TryGetValue(parameterName, out var rawValue))
            {
                invocationArguments[i] = ConvertValue(parameter, rawValue);
                continue;
            }

            if (parameter.HasDefaultValue)
            {
                invocationArguments[i] = NormalizeDefaultValue(parameter.DefaultValue);
                continue;
            }

            throw new CliUsageException($"Missing required option '--{parameterName}'.");
        }

        return invocationArguments;
    }

    public static async Task<object?> InvokeAsync(ToolCommand command, object?[] arguments)
    {
        Task task;
        try
        {
            task = (Task?)command.Method.Invoke(command.Instance, arguments)
                ?? throw new InvalidOperationException($"Command '{command.Name}' did not return a Task.");
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }

        await task.ConfigureAwait(false);
        return GetTaskResult(command.Method.ReturnType, task);
    }

    public static void WriteJson(object? value)
    {
        Console.WriteLine(JsonSerializer.Serialize(value, value?.GetType() ?? typeof(object), JsonOptions));
    }

    private static string FormatUsageParameter(ParameterInfo parameter)
    {
        var name = GetParameterName(parameter);
        var option = $"--{name} <{GetTypeName(parameter.ParameterType)}>";
        return parameter.HasDefaultValue ? $"[{option}]" : option;
    }

    private static object? ConvertValue(ParameterInfo parameter, string? rawValue)
    {
        var targetType = Nullable.GetUnderlyingType(parameter.ParameterType) ?? parameter.ParameterType;
        var parameterName = GetParameterName(parameter);

        if (targetType == typeof(string))
            return rawValue;

        if (rawValue == null)
            return null;

        if (targetType == typeof(int))
        {
            if (int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var intValue))
                return intValue;

            throw new CliUsageException($"Invalid value for '--{parameterName}': '{rawValue}'. Expected an integer.");
        }

        if (targetType == typeof(bool))
        {
            if (bool.TryParse(rawValue, out var boolValue))
                return boolValue;

            throw new CliUsageException($"Invalid value for '--{parameterName}': '{rawValue}'. Expected 'true' or 'false'.");
        }

        throw new InvalidOperationException(
            $"Unsupported parameter type '{parameter.ParameterType.FullName}' for command '{parameter.Member.Name}'.");
    }

    private static object? GetTaskResult(Type returnType, Task task)
    {
        if (!returnType.IsGenericType || returnType.GetGenericTypeDefinition() != typeof(Task<>))
            return null;

        return returnType.GetProperty("Result")?.GetValue(task);
    }

    private static string GetTypeName(Type parameterType)
    {
        var targetType = Nullable.GetUnderlyingType(parameterType) ?? parameterType;
        if (targetType == typeof(string))
            return "string";
        if (targetType == typeof(int))
            return "int";
        if (targetType == typeof(bool))
            return "bool";

        return targetType.Name;
    }

    private static string GetParameterName(ParameterInfo parameter)
    {
        return parameter.Name ?? throw new InvalidOperationException(
            $"Command parameter for '{parameter.Member.Name}' does not have a name.");
    }

    private static bool IsBooleanParameter(ParameterInfo parameter)
    {
        var targetType = Nullable.GetUnderlyingType(parameter.ParameterType) ?? parameter.ParameterType;
        return targetType == typeof(bool);
    }

    private static object? NormalizeDefaultValue(object? value)
    {
        return value == DBNull.Value ? null : value;
    }

    private static string FormatDefaultValue(object? value)
    {
        if (value == null || value == DBNull.Value)
            return "null";

        return value switch
        {
            bool boolValue => boolValue ? "true" : "false",
            string stringValue => $"'{stringValue}'",
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? value.ToString() ?? string.Empty,
        };
    }
}

internal sealed class CliUsageException(string message) : Exception(message);
