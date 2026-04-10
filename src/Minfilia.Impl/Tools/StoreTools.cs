using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Minfilia.Models;
using Minfilia.Outlook;

namespace Minfilia.Tools;

[McpServerToolType]
internal sealed class StoreTools(OutlookSession _session)
{
    [McpServerTool(Name = "list_stores", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List Outlook stores. Email may be missing when Outlook does not expose SMTP or account-to-store binding.")]
    public async Task<List<StoreInfo>> ListStores()
    {
        return await _session.ExecuteAsync<List<StoreInfo>>(ns =>
        {
            var stores = new List<StoreInfo>();
            var storeCount = (int)ns.Stores.Count;
            for (var i = 1; i <= storeCount; i++)
            {
                var store = ns.Stores.Item(i);
                var displayName = (string)store.DisplayName;

                stores.Add(new StoreInfo { StoreName = displayName });
            }

            var accountCount = (int)ns.Accounts.Count;
            for (var i = 1; i <= accountCount; i++)
            {
                var acct = ns.Accounts.Item(i);
                var accountContext = $"account #{i}";
                var accountDisplayName = ReadOptionalString(() => (string)acct.DisplayName, "DisplayName", accountContext) ?? accountContext;
                var email = ReadOptionalString(() => (string)acct.SmtpAddress, "SmtpAddress", accountDisplayName);
                var deliveryStoreName = ReadOptionalString(() => (string)acct.DeliveryStore.DisplayName, "DeliveryStore.DisplayName", accountDisplayName);

                if (email == null)
                {
                    Console.Error.WriteLine($"Warning: cannot read SMTP address for '{accountDisplayName}', leaving email null.");
                    continue;
                }

                if (deliveryStoreName == null)
                {
                    Console.Error.WriteLine($"Warning: cannot read delivery store for '{accountDisplayName}', leaving email null.");
                    continue;
                }

                var matched = false;

                for (var j = 0; j < stores.Count; j++)
                {
                    if (stores[j].StoreName == deliveryStoreName)
                    {
                        stores[j] = new StoreInfo { StoreName = stores[j].StoreName, Email = email };
                        matched = true;
                        break;
                    }
                }

                if (!matched)
                    Console.Error.WriteLine($"Warning: delivery store '{deliveryStoreName}' for '{accountDisplayName}' was not found in Outlook stores.");
            }

            return stores;
        });
    }

    private static string? ReadOptionalString(Func<string> read, string propertyName, string context)
    {
        try
        {
            return read();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: cannot read {propertyName} for '{context}', returning null: {ex.Message}");
            return null;
        }
    }

    [McpServerTool(Name = "list_folders", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List folders in an Outlook store. Item counts may be missing when Outlook does not expose them.")]
    public async Task<List<FolderInfo>> ListFolders(
        [Description("Store display name to filter. If omitted, lists all stores.")] string? storeName = null,
        [Description("Parent folder path (e.g. 'Mailbox Name/Inbox'). If omitted, lists from store root.")] string? parentPath = null,
        [Description("Maximum folder depth to recurse. Default 3, max 10.")] int maxDepth = 3)
    {
        storeName = InputValidator.NullIfWhiteSpace(storeName);
        parentPath = InputValidator.NullIfWhiteSpace(parentPath);
        maxDepth = InputValidator.ValidateMaxResults(maxDepth, 0, 10, "maxDepth");

        return await _session.ExecuteAsync(ns =>
        {
            var result = new List<FolderInfo>();

            if (parentPath != null)
            {
                var parent = FolderResolver.ResolveFolder(ns, parentPath);
                var subCount = (int)parent.Folders.Count;
                for (var k = 1; k <= subCount; k++)
                    FolderResolver.CollectFolders(parent.Folders.Item(k), parentPath, result, 0, maxDepth);
                return result;
            }

            var storeCount = (int)ns.Stores.Count;
            for (var i = 1; i <= storeCount; i++)
            {
                var store = ns.Stores.Item(i);
                var currentStoreName = (string)store.DisplayName;
                if (storeName != null && !string.Equals(currentStoreName, storeName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var root = store.GetRootFolder();
                var subCount = (int)root.Folders.Count;
                for (var k = 1; k <= subCount; k++)
                    FolderResolver.CollectFolders(root.Folders.Item(k), currentStoreName, result, 0, maxDepth);
            }

            return result;
        });
    }
}
