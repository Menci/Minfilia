namespace Minfilia.Models;

internal sealed class CalendarEvent
{
    public required string EntryId { get; init; }
    public string? Subject { get; init; }
    public required string StartTime { get; init; }
    public required string EndTime { get; init; }
    public string? Location { get; init; }
    public string? Organizer { get; init; }
    public required bool IsRecurring { get; init; }
    public required bool IsCancelled { get; init; }
    public string? Body { get; init; }
}
