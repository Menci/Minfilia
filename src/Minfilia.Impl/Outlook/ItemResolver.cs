using System;
using ModelContextProtocol;

namespace Minfilia.Outlook;

/// <summary>
/// Item lookup helpers for Outlook OOM quirks.
/// All methods run inside OutlookSession.ExecuteAsync on the STA thread.
/// </summary>
internal static class ItemResolver
{
    public static dynamic ResolveItemById(dynamic ns, string id)
    {
        try
        {
            return ns.GetItemFromID(id);
        }
        catch (Exception firstEx)
        {
            var storeCount = (int)ns.Stores.Count;
            for (var i = 1; i <= storeCount; i++)
            {
                var store = ns.Stores.Item(i);
                string? outlookStoreId;

                try { outlookStoreId = (string)store.StoreID; }
                catch { continue; }

                try
                {
                    return ns.GetItemFromID(id, outlookStoreId);
                }
                catch { }
            }

            throw new McpException($"Item not found for EntryID '{id}': {firstEx.Message}");
        }
    }
}
