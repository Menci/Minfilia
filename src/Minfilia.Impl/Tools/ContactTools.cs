using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Minfilia.Models;
using Minfilia.Outlook;

namespace Minfilia.Tools;

[McpServerToolType]
internal sealed class ContactTools(OutlookSession _session)
{
    [McpServerTool(Name = "search_contacts", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Search Outlook contacts by name, email, company, or job title. Returns truncated=true when maxResults cuts off additional matches, skippedFolders for contact folders that could not be enumerated, and mayHaveMissedMatches=true when some items could not be fully inspected during matching.")]
    public async Task<ContactSearchResult> SearchContacts(
        [Description("Search query to match against contact name, email, company, or job title.")] string query,
        [Description("Maximum number of results. Default 50, max 200.")] int maxResults = 50)
    {
        query = InputValidator.RequireNonBlank(query, "query");
        maxResults = InputValidator.ValidateMaxResults(maxResults);

        return await _session.ExecuteAsync(ns =>
        {
            var state = new ContactSearchState();
            var queryLower = query.ToLowerInvariant();
            var storeCount = (int)ns.Stores.Count;

            for (var i = 1; i <= storeCount; i++)
            {
                if (state.Results.Count >= maxResults)
                {
                    state.Truncated = true;
                    break;
                }

                var store = ns.Stores.Item(i);
                var root = store.GetRootFolder();
                var storeName = (string)store.DisplayName;
                var subCount = (int)root.Folders.Count;
                for (var j = 1; j <= subCount; j++)
                {
                    if (state.Results.Count >= maxResults)
                    {
                        state.Truncated = true;
                        break;
                    }

                    SearchContactsInFolder(root.Folders.Item(j), storeName, queryLower, state, maxResults);
                }
            }

            return new ContactSearchResult
            {
                Results = state.Results,
                Truncated = state.Truncated,
                SkippedFolders = state.SkippedFolders,
                MayHaveMissedMatches = state.MayHaveMissedMatches,
            };
        });
    }

    private static void SearchContactsInFolder(dynamic folder, string parentPath, string queryLower,
        ContactSearchState state, int maxResults)
    {
        var name = (string)folder.Name;
        var currentPath = $"{parentPath}/{name}";

        // DefaultItemType 2 = olContactItem
        if ((int)folder.DefaultItemType == 2)
        {
            try
            {
                var items = folder.Items;
                foreach (var item in items)
                {
                    if (state.Results.Count >= maxResults)
                    {
                        state.Truncated = true;
                        break;
                    }

                    bool isContactItem;
                    if (!ItemMapper.TryIsContactItem(item, out isContactItem))
                    {
                        state.MayHaveMissedMatches = true;
                        continue;
                    }

                    if (!isContactItem)
                        continue;

                    string? fullName;
                    if (!PropertyReader.TryOptionalString(item, "FullName", out fullName))
                        state.MayHaveMissedMatches = true;

                    string? email;
                    if (!PropertyReader.TryOptionalString(item, "Email1Address", out email))
                        state.MayHaveMissedMatches = true;

                    string? company;
                    if (!PropertyReader.TryOptionalString(item, "CompanyName", out company))
                        state.MayHaveMissedMatches = true;

                    string? jobTitle;
                    if (!PropertyReader.TryOptionalString(item, "JobTitle", out jobTitle))
                        state.MayHaveMissedMatches = true;

                    var text = $"{fullName ?? ""} {email ?? ""} {company ?? ""} {jobTitle ?? ""}".ToLowerInvariant();
                    if (text.Contains(queryLower))
                        state.Results.Add(ItemMapper.ToContactInfo(item));
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: cannot enumerate contacts folder '{currentPath}', skipping folder: {ex.Message}");
                state.SkippedFolders.Add(currentPath);
            }
        }

        var subCount = (int)folder.Folders.Count;
        for (var i = 1; i <= subCount; i++)
        {
            if (state.Results.Count >= maxResults)
            {
                state.Truncated = true;
                break;
            }

            SearchContactsInFolder(folder.Folders.Item(i), currentPath, queryLower, state, maxResults);
        }
    }

    private sealed class ContactSearchState
    {
        public List<ContactInfo> Results { get; } = [];
        public List<string> SkippedFolders { get; } = [];
        public bool MayHaveMissedMatches { get; set; }
        public bool Truncated { get; set; }
    }
}
