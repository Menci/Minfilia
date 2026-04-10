using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using ModelContextProtocol;
using Minfilia.Models;

namespace Minfilia.Outlook;

/// <summary>
/// Folder path resolution and tree enumeration.
/// All methods run inside OutlookSession.ExecuteAsync on the STA thread.
/// </summary>
internal static class FolderResolver
{
    private const int OlMailItem = 0;
    private const int OlAppointmentItem = 1;
    private static readonly ConcurrentDictionary<string, byte> LoggedCountFailures = new();

    public static dynamic ResolveFolder(dynamic ns, string path)
    {
        var parts = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
            throw new McpException("Empty folder path.");

        dynamic? storeRoot = null;
        var storeCount = (int)ns.Stores.Count;
        for (var i = 1; i <= storeCount; i++)
        {
            var store = ns.Stores.Item(i);
            if (string.Equals((string)store.DisplayName, parts[0], StringComparison.OrdinalIgnoreCase))
            {
                storeRoot = store.GetRootFolder();
                break;
            }
        }

        if (storeRoot == null)
            throw new McpException($"Store not found: '{parts[0]}'");

        dynamic current = storeRoot;
        for (var i = 1; i < parts.Length; i++)
        {
            dynamic? found = null;
            var folderCount = (int)current.Folders.Count;
            for (var j = 1; j <= folderCount; j++)
            {
                var folder = current.Folders.Item(j);
                if (string.Equals((string)folder.Name, parts[i], StringComparison.OrdinalIgnoreCase))
                {
                    found = folder;
                    break;
                }
            }

            if (found == null)
                throw new McpException($"Folder not found: '{parts[i]}' in path '{path}'");

            current = found;
        }

        return current;
    }

    public static dynamic ResolveMailFolder(dynamic ns, string path)
    {
        var folder = ResolveFolder(ns, path);
        var itemType = (int)folder.DefaultItemType;
        if (itemType != OlMailItem)
            throw new McpException($"'{path}' is not a mail folder (type={ItemTypeName(itemType)}).");
        return folder;
    }

    public static dynamic ResolveCalendarFolder(dynamic ns, string path)
    {
        var folder = ResolveFolder(ns, path);
        var itemType = (int)folder.DefaultItemType;
        if (itemType != OlAppointmentItem)
            throw new McpException($"'{path}' is not a calendar folder (type={ItemTypeName(itemType)}).");
        return folder;
    }

    public static dynamic FindDefaultInbox(dynamic ns)
    {
        try { return ns.GetDefaultFolder(6); }
        catch (Exception ex) { throw new McpException($"Default Inbox not available: {ex.Message}"); }
    }

    public static dynamic FindDefaultCalendar(dynamic ns)
    {
        try { return ns.GetDefaultFolder(9); }
        catch (Exception ex) { throw new McpException($"Default calendar not available: {ex.Message}"); }
    }

    public static dynamic ResolveMailFolderOrDefault(dynamic ns, string? folderPath)
    {
        return folderPath != null ? ResolveMailFolder(ns, folderPath) : FindDefaultInbox(ns);
    }

    public static void CollectFolders(dynamic folder, string parentPath, List<FolderInfo> result, int depth, int maxDepth)
    {
        var name = (string)folder.Name;
        var path = $"{parentPath}/{name}";
        var itemCount = TryReadCount(() => (int)folder.Items.Count, "item count", path, logFailure: true);
        var unreadCount = TryReadCount(() => (int)folder.UnReadItemCount, "unread count", path, logFailure: true);

        result.Add(new FolderInfo { Path = path, Name = name, ItemCount = itemCount, UnreadCount = unreadCount });

        if (depth < maxDepth)
        {
            var subCount = (int)folder.Folders.Count;
            for (var i = 1; i <= subCount; i++)
                CollectFolders(folder.Folders.Item(i), path, result, depth + 1, maxDepth);
        }
    }

    public static void CollectCalendars(dynamic folder, string parentPath, List<CalendarInfo> result, string storeName, string defaultCalendarPath)
    {
        var name = (string)folder.Name;
        var currentPath = $"{parentPath}/{name}";

        if ((int)folder.DefaultItemType == OlAppointmentItem)
        {
            var itemCount = TryReadCount(() => (int)folder.Items.Count, "calendar item count", currentPath, logFailure: true);
            result.Add(new CalendarInfo
            {
                Name = name,
                StoreName = storeName,
                Path = currentPath,
                ItemCount = itemCount,
                IsDefault = string.Equals(currentPath, defaultCalendarPath, StringComparison.OrdinalIgnoreCase),
            });
        }

        var subCount = (int)folder.Folders.Count;
        for (var i = 1; i <= subCount; i++)
            CollectCalendars(folder.Folders.Item(i), currentPath, result, storeName, defaultCalendarPath);
    }

    public static string EscapeDasl(string value) => value.Replace("'", "''");

    public static string GetContractPath(dynamic folder, string context)
    {
        try
        {
            var folderPath = (string)folder.FolderPath;
            var parts = folderPath.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
                throw new McpException($"{context} path is empty.");

            return string.Join("/", parts);
        }
        catch (McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new McpException($"{context} path could not be read: {ex.Message}");
        }
    }

    private static int? TryReadCount(Func<int> read, string kind, string path, bool logFailure)
    {
        try
        {
            return read();
        }
        catch (Exception ex)
        {
            if (logFailure)
            {
                var key = $"{kind}|{ex.GetType().FullName}";
                if (LoggedCountFailures.TryAdd(key, 0))
                    Console.Error.WriteLine($"Warning: cannot read {kind} for '{path}', returning null: {ex.Message}");
            }
            return null;
        }
    }

    private static string ItemTypeName(int type) => type switch
    {
        0 => "Mail", 1 => "Calendar", 2 => "Contact", _ => $"Unknown({type})",
    };
}
