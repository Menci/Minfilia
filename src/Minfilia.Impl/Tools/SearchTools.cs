using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Minfilia.Models;
using Minfilia.Outlook;

namespace Minfilia.Tools;

[McpServerToolType]
internal sealed class SearchTools(OutlookSession _session)
{
    private const int ScanLimit = 5000;

    [McpServerTool(Name = "search_emails", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Search emails in Outlook folders by keyword, sender, subject, and date. Returns truncated=true when the client-side scan cap is hit, and mayHaveMissedMatches=true when some items could not be fully inspected for keyword matching.")]
    public async Task<SearchResult> SearchEmails(
        [Description("Keyword query to match in subject and body.")] string? query = null,
        [Description("Folder path to search in (e.g. 'Mailbox Name/Inbox'). Omit to search default inbox.")] string? folderPath = null,
        [Description("Filter by sender name or email.")] string? from = null,
        [Description("Filter by subject keyword (server-side, fast).")] string? subject = null,
        [Description("Start date in yyyy-MM-dd format (e.g. '2026-04-07').")] string? dateFrom = null,
        [Description("End date in yyyy-MM-dd format, inclusive.")] string? dateTo = null,
        [Description("Only return emails with attachments.")] bool? hasAttachment = null,
        [Description("Only return unread emails.")] bool? unreadOnly = null,
        [Description("Maximum number of results. Default 25, max 100.")] int maxResults = 25,
        [Description("Pagination cursor from previous search result.")] string? cursor = null)
    {
        query = InputValidator.NullIfWhiteSpace(query);
        folderPath = InputValidator.NullIfWhiteSpace(folderPath);
        from = InputValidator.NullIfWhiteSpace(from);
        subject = InputValidator.NullIfWhiteSpace(subject);
        maxResults = InputValidator.ValidateMaxResults(maxResults, 1, 100);
        var offset = InputValidator.ParseCursor(cursor);
        var dateFromValue = dateFrom != null ? InputValidator.ParseDate(dateFrom, "dateFrom") : (DateTime?)null;
        var dateToValue = dateTo != null ? InputValidator.ParseDate(dateTo, "dateTo") : (DateTime?)null;
        var dateToExclusiveValue = dateToValue?.AddDays(1);
        if (dateFromValue != null && dateToValue != null)
            InputValidator.ValidateDateRange(dateFromValue.Value, dateToValue.Value, dateFrom, dateTo);

        return await _session.ExecuteAsync(ns =>
        {
            var targetFolder = FolderResolver.ResolveMailFolderOrDefault(ns, folderPath);

            var filters = new List<string>();
            if (subject != null)
                filters.Add($"\"urn:schemas:httpmail:subject\" LIKE '%{FolderResolver.EscapeDasl(subject)}%'");
            if (from != null)
                filters.Add($"(\"urn:schemas:httpmail:senderemail\" LIKE '%{FolderResolver.EscapeDasl(from)}%'" +
                             $" OR \"urn:schemas:httpmail:sendername\" LIKE '%{FolderResolver.EscapeDasl(from)}%'" +
                             $" OR \"urn:schemas:httpmail:fromemail\" LIKE '%{FolderResolver.EscapeDasl(from)}%'" +
                             $" OR \"urn:schemas:httpmail:displayfrom\" LIKE '%{FolderResolver.EscapeDasl(from)}%')");
            if (dateFromValue != null)
                filters.Add($"\"urn:schemas:httpmail:datereceived\" >= '{InputValidator.FormatDate(dateFromValue.Value)}'");
            if (dateToExclusiveValue != null)
                filters.Add($"\"urn:schemas:httpmail:datereceived\" < '{InputValidator.FormatDate(dateToExclusiveValue.Value)}'");
            if (hasAttachment == true)
                filters.Add("\"urn:schemas:httpmail:hasattachment\" = 1");
            if (unreadOnly == true)
                filters.Add("\"urn:schemas:httpmail:read\" = 0");

            dynamic items = targetFolder.Items;
            if (filters.Count > 0)
                items = items.Restrict("@SQL=" + string.Join(" AND ", filters));
            items.Sort("[ReceivedTime]", true);

            var queryLower = query?.ToLowerInvariant();
            var results = new List<EmailSummary>();
            var scanned = 0;
            var matched = 0;
            var skipped = 0;
            var truncated = false;
            var mayHaveMissedMatches = false;

            foreach (var item in items)
            {
                if (matched >= maxResults) break;
                if (scanned >= ScanLimit) { truncated = true; break; }
                scanned++;

                bool isMailItem;
                if (!ItemMapper.TryIsMailItem(item, out isMailItem))
                {
                    mayHaveMissedMatches = true;
                    continue;
                }

                if (!isMailItem)
                    continue;

                if (queryLower != null)
                {
                    string? subjectText;
                    if (!PropertyReader.TryOptionalString(item, "Subject", out subjectText))
                        mayHaveMissedMatches = true;

                    string? bodyText;
                    if (!PropertyReader.TryOptionalString(item, "Body", out bodyText))
                        mayHaveMissedMatches = true;

                    var matchText = $"{subjectText ?? ""} {bodyText ?? ""}".ToLowerInvariant();
                    if (!matchText.Contains(queryLower))
                        continue;
                }

                if (skipped < offset) { skipped++; continue; }

                results.Add(ItemMapper.ToEmailSummary(item));
                matched++;
            }

            // Always provide cursor if there might be more results
            string? nextCursor = null;
            if (matched > 0 && (matched == maxResults || truncated))
                nextCursor = Convert.ToBase64String(Encoding.UTF8.GetBytes((offset + matched).ToString()));

            return new SearchResult
            {
                Results = results,
                NextCursor = nextCursor,
                Truncated = truncated,
                MayHaveMissedMatches = mayHaveMissedMatches,
            };
        });
    }
}
