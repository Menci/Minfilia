using System;
using System.Collections.Generic;
using ModelContextProtocol;
using Minfilia.Models;

namespace Minfilia.Outlook;

/// <summary>
/// Maps Outlook COM items to typed DTOs.
/// All methods run inside OutlookSession.ExecuteAsync on the STA thread.
/// </summary>
internal static class ItemMapper
{
    private const int OlAppointmentItemClass = 26;
    private const int OlContactItemClass = 40;
    private const int OlMailItemClass = 43;

    /// <summary>
    /// Validates that a COM item is a mail item (Class == 43), then maps to EmailDetail.
    /// </summary>
    public static EmailDetail ToEmailDetail(dynamic item, string id)
    {
        ValidateItemClass(item, OlMailItemClass, $"EntryID '{id}'", "mail item");

        var attachments = MapAttachments((object)item, id);

        return new EmailDetail
        {
            Id = id,
            Subject = PropertyReader.Required<string>(item, "Subject"),
            From = PropertyReader.OptionalString(item, "SenderName"),
            FromEmail = PropertyReader.OptionalString(item, "SenderEmailAddress"),
            To = PropertyReader.OptionalString(item, "To"),
            Cc = PropertyReader.OptionalString(item, "CC"),
            Date = InputValidator.FormatLocalDateTime(PropertyReader.Required<DateTime>(item, "ReceivedTime")),
            Body = PropertyReader.OptionalString(item, "Body"),
            HtmlBody = PropertyReader.OptionalString(item, "HTMLBody"),
            Importance = PropertyReader.Required<int>(item, "Importance"),
            IsRead = !PropertyReader.Required<bool>(item, "UnRead"),
            ConversationTopic = PropertyReader.OptionalString(item, "ConversationTopic"),
            ConversationId = PropertyReader.OptionalString(item, "ConversationID"),
            Attachments = attachments.Attachments,
            AttachmentsReadFailed = attachments.ReadFailed,
        };
    }

    /// <summary>
    /// Maps a COM mail item to an EmailSummary DTO for search results.
    /// </summary>
    public static EmailSummary ToEmailSummary(dynamic item)
    {
        ValidateItemClass(item, OlMailItemClass, "Search result item", "mail item");

        var body = PropertyReader.OptionalString(item, "Body");

        return new EmailSummary
        {
            Id = PropertyReader.Required<string>(item, "EntryID"),
            Subject = PropertyReader.OptionalString(item, "Subject"),
            From = PropertyReader.OptionalString(item, "SenderName"),
            Date = InputValidator.FormatLocalDateTime(PropertyReader.Required<DateTime>(item, "ReceivedTime")),
            Preview = TruncateBody(body, 200),
        };
    }

    /// <summary>
    /// Maps a COM calendar item to a CalendarEvent DTO.
    /// </summary>
    public static CalendarEvent ToCalendarEvent(dynamic item, bool includeBody)
    {
        ValidateItemClass(item, OlAppointmentItemClass, "Calendar result item", "calendar item");

        return new CalendarEvent
        {
            EntryId = PropertyReader.Required<string>(item, "EntryID"),
            Subject = PropertyReader.OptionalString(item, "Subject"),
            StartTime = InputValidator.FormatLocalDateTime(PropertyReader.Required<DateTime>(item, "Start")),
            EndTime = InputValidator.FormatLocalDateTime(PropertyReader.Required<DateTime>(item, "End")),
            Location = PropertyReader.OptionalString(item, "Location"),
            Organizer = PropertyReader.OptionalString(item, "Organizer"),
            IsRecurring = PropertyReader.Required<bool>(item, "IsRecurring"),
            IsCancelled = IsCanceledMeeting(PropertyReader.Required<int>(item, "MeetingStatus")),
            Body = includeBody ? TruncateBody(PropertyReader.OptionalString(item, "Body"), 500) : null,
        };
    }

    public static bool TryIsMailItem(dynamic item, out bool isMailItem)
    {
        int itemClass;
        if (!TryReadItemClass(item, out itemClass))
        {
            isMailItem = false;
            return false;
        }

        isMailItem = itemClass == OlMailItemClass;
        return true;
    }

    public static bool TryIsContactItem(dynamic item, out bool isContactItem)
    {
        int itemClass;
        if (!TryReadItemClass(item, out itemClass))
        {
            isContactItem = false;
            return false;
        }

        isContactItem = itemClass == OlContactItemClass;
        return true;
    }

    public static ContactInfo ToContactInfo(dynamic item)
    {
        ValidateItemClass(item, OlContactItemClass, "Search result item", "contact item");

        return new ContactInfo
        {
            Name = PropertyReader.OptionalString(item, "FullName"),
            Email = PropertyReader.OptionalString(item, "Email1Address"),
            Company = PropertyReader.OptionalString(item, "CompanyName"),
            Phone = ReadPhone(item),
            JobTitle = PropertyReader.OptionalString(item, "JobTitle"),
        };
    }

    private static (List<AttachmentInfo> Attachments, bool ReadFailed) MapAttachments(object itemObject, string id)
    {
        dynamic item = itemObject;
        var result = new List<AttachmentInfo>();
        try
        {
            var count = (int)item.Attachments.Count;
            for (var i = 1; i <= count; i++)
            {
                var att = item.Attachments.Item(i);
                result.Add(new AttachmentInfo
                {
                    Name = PropertyReader.OptionalString(att, "FileName"),
                    Size = (long)att.Size,
                });
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: cannot read attachments for '{id}': {ex.Message}");
            return (result, true);
        }
        return (result, false);
    }

    private static void ValidateItemClass(dynamic item, int expectedClass, string context, string expectedType)
    {
        var itemClass = (int)item.Class;
        if (itemClass != expectedClass)
            throw new McpException($"{context} is not a {expectedType} (class={itemClass}). Use the appropriate tool for this item type.");
    }

    private static bool TryReadItemClass(dynamic item, out int itemClass)
    {
        try
        {
            itemClass = (int)item.Class;
            return true;
        }
        catch
        {
            itemClass = 0;
            return false;
        }
    }

    private static string? TruncateBody(string? body, int maxLen)
    {
        if (body == null)
            return null;

        var normalized = body.Replace("\r\n", " ").Replace("\n", " ").Trim();
        return normalized.Length <= maxLen ? normalized : normalized.Substring(0, maxLen).Trim();
    }

    private static bool IsCanceledMeeting(int meetingStatus) => meetingStatus == 5 || meetingStatus == 7;

    private static string? ReadPhone(dynamic item)
    {
        string? businessPhone;
        if (PropertyReader.TryOptionalString(item, "BusinessTelephoneNumber", out businessPhone))
        {
            if (!string.IsNullOrEmpty(businessPhone))
                return businessPhone;

            string? mobilePhone;
            return PropertyReader.TryOptionalString(item, "MobileTelephoneNumber", out mobilePhone)
                ? mobilePhone ?? businessPhone
                : businessPhone;
        }

        string? fallbackPhone;
        if (PropertyReader.TryOptionalString(item, "MobileTelephoneNumber", out fallbackPhone))
            return fallbackPhone;

        return null;
    }
}
