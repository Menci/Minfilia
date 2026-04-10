namespace Minfilia.Models;

internal sealed class CalendarInfo
{
    public required string Name { get; init; }
    public required string StoreName { get; init; }
    public required string Path { get; init; }
    public int? ItemCount { get; init; }
    public bool IsDefault { get; init; }
}
