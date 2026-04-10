using System.Collections.Generic;

namespace Minfilia.Models;

internal sealed class CalendarEventsResult
{
    public List<CalendarEvent> Results { get; init; } = [];
    public bool Truncated { get; init; }
}
