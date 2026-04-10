namespace Minfilia.Models;

internal sealed class EmailSummary
{
    public required string Id { get; init; }
    public string? Subject { get; init; }
    public string? From { get; init; }
    public required string Date { get; init; }
    public string? Preview { get; init; }
}
