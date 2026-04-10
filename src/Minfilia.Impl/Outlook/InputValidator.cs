using System;
using System.Globalization;
using ModelContextProtocol;

namespace Minfilia.Outlook;

/// <summary>
/// Shared input validation. Call before entering COM thread.
/// </summary>
internal static class InputValidator
{
    private const string DateFormat = "yyyy-MM-dd";

    public static string RequireNonBlank(string value, string paramName)
    {
        var normalized = NullIfWhiteSpace(value);
        if (normalized == null)
            throw new McpException($"Invalid {paramName}: value must not be blank.");
        return normalized;
    }

    public static string? NullIfWhiteSpace(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
    }

    public static DateTime ParseDate(string value, string paramName)
    {
        if (!DateTime.TryParseExact(value, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
            throw new McpException($"Invalid {paramName}: '{value}'. Expected {DateFormat}.");
        return result.Date;
    }

    public static void ValidateDateRange(string dateFrom, string dateTo)
    {
        ValidateDateRange(ParseDate(dateFrom, "dateFrom"), ParseDate(dateTo, "dateTo"), dateFrom, dateTo);
    }

    public static void ValidateDateRange(DateTime from, DateTime to, string? rawFrom = null, string? rawTo = null)
    {
        if (from > to)
            throw new McpException($"Invalid date range: dateFrom '{rawFrom ?? FormatDate(from)}' is after dateTo '{rawTo ?? FormatDate(to)}'.");
    }

    public static string FormatLocalDateTime(DateTime value)
    {
        var localValue = value.Kind switch
        {
            DateTimeKind.Local => value,
            DateTimeKind.Utc => value.ToLocalTime(),
            _ => DateTime.SpecifyKind(value, DateTimeKind.Local),
        };

        return new DateTimeOffset(localValue).ToString("yyyy-MM-ddTHH:mm:sszzz", CultureInfo.InvariantCulture);
    }

    public static int ValidateMaxResults(int maxResults, int min = 1, int max = 200, string paramName = "maxResults")
    {
        if (maxResults < min)
            throw new McpException($"Invalid {paramName}: {maxResults}. Must be at least {min}.");
        if (maxResults > max)
            throw new McpException($"Invalid {paramName}: {maxResults}. Must be at most {max}.");
        return maxResults;
    }

    public static int ParseCursor(string? cursor)
    {
        if (cursor == null) return 0;
        try
        {
            var decoded = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(cursor));
            if (int.TryParse(decoded, out var offset) && offset >= 0)
                return offset;
        }
        catch { }

        throw new McpException($"Invalid cursor: '{cursor}'. Must be a pagination cursor from a previous search.");
    }

    public static string FormatDate(DateTime value) => value.ToString(DateFormat, CultureInfo.InvariantCulture);
}
