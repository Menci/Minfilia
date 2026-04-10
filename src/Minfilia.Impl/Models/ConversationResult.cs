using System.Collections.Generic;

namespace Minfilia.Models;

internal sealed class ConversationResult
{
    public List<EmailSummary> Results { get; init; } = [];
    public bool Truncated { get; init; }
}
