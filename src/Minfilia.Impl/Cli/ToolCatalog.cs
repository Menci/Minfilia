using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using ModelContextProtocol.Server;
using Minfilia.Outlook;
using Minfilia.Tools;

namespace Minfilia;

internal static class ToolCatalog
{
    public static List<ToolCommand> CreateCommands(OutlookOperationExecutor executor)
    {
        var commands = new List<ToolCommand>();

        foreach (var instance in CreateToolInstances(executor))
        {
            var type = instance.GetType();
            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            {
                var toolAttribute = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttribute == null)
                    continue;

                var name = toolAttribute.Name;
                if (string.IsNullOrWhiteSpace(name))
                    throw new InvalidOperationException($"Tool method '{type.FullName}.{method.Name}' is missing a public tool name.");

                var toolName = name!;

                commands.Add(new ToolCommand
                {
                    Name = toolName,
                    Description = method.GetCustomAttribute<DescriptionAttribute>()?.Description ?? string.Empty,
                    Instance = instance,
                    Method = method,
                    Parameters = method.GetParameters(),
                });
            }
        }

        commands.Sort(static (left, right) => string.Compare(left.Name, right.Name, StringComparison.OrdinalIgnoreCase));
        return commands;
    }

    public static McpServerPrimitiveCollection<McpServerTool> CreateMcpTools(IEnumerable<ToolCommand> commands)
    {
        var tools = new McpServerPrimitiveCollection<McpServerTool>();
        foreach (var command in commands)
            tools.Add(McpServerTool.Create(command.Method, command.Instance));
        return tools;
    }

    private static object[] CreateToolInstances(OutlookOperationExecutor executor)
    {
        return
        [
            new StoreTools(executor),
            new SearchTools(executor),
            new EmailTools(executor),
            new CalendarTools(executor),
            new ContactTools(executor),
        ];
    }
}

internal sealed class ToolCommand
{
    public required string Name { get; init; }
    public required string Description { get; init; }
    public required object Instance { get; init; }
    public required MethodInfo Method { get; init; }
    public required ParameterInfo[] Parameters { get; init; }
}
