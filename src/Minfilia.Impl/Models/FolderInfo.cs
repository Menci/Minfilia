namespace Minfilia.Models;

internal sealed class FolderInfo
{
    public required string Path { get; init; }
    public required string Name { get; init; }
    public int? ItemCount { get; init; }
    public int? UnreadCount { get; init; }
}
