using System.Collections.Generic;

namespace Minfilia.Models;

internal sealed class ContactSearchResult
{
    public List<ContactInfo> Results { get; init; } = [];
    public bool Truncated { get; init; }
    public List<string> SkippedFolders { get; init; } = [];
    public bool MayHaveMissedMatches { get; init; }
}
