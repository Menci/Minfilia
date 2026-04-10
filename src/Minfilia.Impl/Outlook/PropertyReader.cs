using System;
using System.Collections.Concurrent;

namespace Minfilia.Outlook;

/// <summary>
/// COM property access with explicit required/optional semantics.
/// All methods run inside OutlookSession.ExecuteAsync on the STA thread.
/// </summary>
internal static class PropertyReader
{
    private static readonly ConcurrentDictionary<string, byte> LoggedOptionalFailures = new();

    /// <summary>
    /// Reads a required COM property. Throws McpException on failure.
    /// Use for identity/semantic fields: EntryID, ReceivedTime, Class, Importance, UnRead, Start, End, IsRecurring.
    /// </summary>
    public static T Required<T>(dynamic obj, string property)
    {
        try
        {
            var value = ((object)obj).GetType().InvokeMember(
                property, System.Reflection.BindingFlags.GetProperty, null, obj, null);
            if (value == null)
                throw new ModelContextProtocol.McpException($"Required property '{property}' is null.");

            return (T)value;
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Required property '{property}' could not be read: {ex.Message}");
        }
    }

    /// <summary>
    /// Reads an optional string COM property. Returns null on read failure.
    /// Use for display-only fields that may be absent or unreadable: CC, Location, SenderName, etc.
    /// </summary>
    public static string? OptionalString(dynamic obj, string property)
    {
        string? value;
        return TryOptionalString(obj, property, out value) ? value : null;
    }

    /// <summary>
    /// Reads an optional string COM property. Returns true for successful reads, including blank values.
    /// </summary>
    public static bool TryOptionalString(dynamic obj, string property, out string? value)
    {
        try
        {
            value = (string?)(((object)obj).GetType().InvokeMember(
                property, System.Reflection.BindingFlags.GetProperty, null, obj, null) ?? "");
            return true;
        }
        catch (Exception ex)
        {
            value = null;
            var objectType = ((object)obj).GetType().FullName ?? "<unknown>";
            var key = $"{objectType}|{property}|{ex.GetType().FullName}";
            if (LoggedOptionalFailures.TryAdd(key, 0))
                Console.Error.WriteLine($"Warning: optional property '{property}' on '{objectType}' could not be read, returning null: {ex.Message}");
            return false;
        }
    }
}
