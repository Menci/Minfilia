using System.Collections.Generic;

namespace Minfilia.Models;

internal sealed class EmailDetail
{
    public required string Id { get; init; }
    public required string Subject { get; init; }
    public string? From { get; init; }
    public string? FromEmail { get; init; }
    public string? To { get; init; }
    public string? Cc { get; init; }
    public required string Date { get; init; }
    public string? Body { get; init; }
    public string? HtmlBody { get; init; }
    public required int Importance { get; init; }
    public required bool IsRead { get; init; }
    public string? ConversationTopic { get; init; }
    public string? ConversationId { get; init; }
    public List<AttachmentInfo> Attachments { get; init; } = [];
    public bool AttachmentsReadFailed { get; init; }
}

internal sealed class AttachmentInfo
{
    public string? Name { get; init; }
    public long Size { get; init; }
}
