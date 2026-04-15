using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using Minfilia.Models;
using Minfilia.Outlook;

namespace Minfilia.Tools;

[McpServerToolType]
internal sealed class CalendarTools(OutlookOperationExecutor _executor)
{
    [McpServerTool(Name = "list_calendars", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List all calendar folders across Outlook stores. The default calendar is flagged with isDefault=true and listed first. itemCount may be missing when Outlook does not expose it.")]
    public async Task<List<CalendarInfo>> ListCalendars(
        [Description("Store display name to filter. Omit to list all.")] string? storeName = null)
    {
        storeName = InputValidator.NullIfWhiteSpace(storeName);

        return await _executor.ExecuteAsync(ns =>
        {
            var result = new List<CalendarInfo>();
            var storeCount = (int)ns.Stores.Count;
            var defaultCalendarPath = FolderResolver.GetContractPath(FolderResolver.FindDefaultCalendar(ns), "Default calendar");

            for (var i = 1; i <= storeCount; i++)
            {
                var store = ns.Stores.Item(i);
                var currentStoreName = (string)store.DisplayName;
                if (storeName != null && !string.Equals(currentStoreName, storeName, StringComparison.OrdinalIgnoreCase))
                    continue;

                var root = store.GetRootFolder();
                var subCount = (int)root.Folders.Count;
                for (var j = 1; j <= subCount; j++)
                    FolderResolver.CollectCalendars(root.Folders.Item(j), currentStoreName, result, currentStoreName, defaultCalendarPath);
            }

            result.Sort(static (left, right) =>
            {
                var defaultComparison = right.IsDefault.CompareTo(left.IsDefault);
                if (defaultComparison != 0)
                    return defaultComparison;

                var storeComparison = string.Compare(left.StoreName, right.StoreName, StringComparison.OrdinalIgnoreCase);
                if (storeComparison != 0)
                    return storeComparison;

                return string.Compare(left.Path, right.Path, StringComparison.OrdinalIgnoreCase);
            });

            return result;
        });
    }

    [McpServerTool(Name = "get_calendar_events", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Get calendar events within a date range using exact local overlap semantics. Omit calendarPath to use the default calendar. Body is omitted unless includeBody=true. Returns truncated=true when maxResults cuts off additional events.")]
    public async Task<CalendarEventsResult> GetCalendarEvents(
        [Description("Start date in yyyy-MM-dd format (e.g. '2026-04-07').")] string dateFrom,
        [Description("End date in yyyy-MM-dd format (e.g. '2026-04-08'). The range is inclusive.")] string dateTo,
        [Description("Calendar folder path (e.g. 'Mailbox Name/Calendar'). Omit for default calendar.")] string? calendarPath = null,
        [Description("Maximum number of results. Default 50, max 200.")] int maxResults = 50,
        [Description("Include truncated calendar body text in results. Default false.")] bool includeBody = false)
    {
        calendarPath = InputValidator.NullIfWhiteSpace(calendarPath);
        maxResults = InputValidator.ValidateMaxResults(maxResults);
        var rangeStartValue = InputValidator.ParseDate(dateFrom, "dateFrom");
        var rangeEndValue = InputValidator.ParseDate(dateTo, "dateTo");
        InputValidator.ValidateDateRange(rangeStartValue, rangeEndValue, dateFrom, dateTo);

        var exactRangeStart = rangeStartValue;
        var exactRangeEndExclusive = rangeEndValue.AddDays(1);
        var prefilterStart = InputValidator.FormatDate(exactRangeStart.AddDays(-1));
        var prefilterEnd = InputValidator.FormatDate(exactRangeEndExclusive.AddDays(1));

        return await _executor.ExecuteAsync(ns =>
        {
            var calendar = calendarPath != null
                ? FolderResolver.ResolveCalendarFolder(ns, calendarPath)
                : FolderResolver.FindDefaultCalendar(ns);

            dynamic items = calendar.Items;
            items.IncludeRecurrences = true;
            items.Sort("[Start]");

            // Outlook Restrict can shift recurrence boundaries across dates. Use a widened server-side
            // prefilter, then enforce the exact local overlap rule on the materialized occurrences.
            var filter = $"[Start] < '{prefilterEnd}' AND [End] > '{prefilterStart}'";
            dynamic filtered = items.Restrict(filter);

            var result = new List<CalendarEvent>();
            var truncated = false;
            foreach (var item in filtered)
            {
                if (!OverlapsExactLocalRange(item, exactRangeStart, exactRangeEndExclusive))
                    continue;

                if (result.Count >= maxResults)
                {
                    truncated = true;
                    break;
                }

                result.Add(ItemMapper.ToCalendarEvent(item, includeBody));
            }

            return new CalendarEventsResult { Results = result, Truncated = truncated };
        });
    }

    private static bool OverlapsExactLocalRange(dynamic item, DateTime rangeStart, DateTime rangeEndExclusive)
    {
        var itemStart = (DateTime)item.Start;
        var itemEnd = (DateTime)item.End;
        return itemStart < rangeEndExclusive && itemEnd > rangeStart;
    }
}
