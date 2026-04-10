using System.Collections.Generic;

namespace Minfilia.Models;

internal sealed class SearchResult
{
    public List<EmailSummary> Results { get; init; } = [];
    public string? NextCursor { get; init; }
    public bool Truncated { get; init; }
    public bool MayHaveMissedMatches { get; init; }
}
