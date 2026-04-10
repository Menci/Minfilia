using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Minfilia.Models;
using Minfilia.Outlook;

namespace Minfilia.Tools;

[McpServerToolType]
internal sealed class EmailTools(OutlookSession _session)
{
    [McpServerTool(Name = "get_email", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Get the full content of an email by its EntryID.")]
    public async Task<EmailDetail> GetEmail(
        [Description("The EntryID of the email to retrieve.")] string id)
    {
        id = InputValidator.RequireNonBlank(id, "id");

        return await _session.ExecuteAsync(ns =>
        {
            var item = ItemResolver.ResolveItemById(ns, id);
            return ItemMapper.ToEmailDetail(item, id);
        });
    }

    [McpServerTool(Name = "get_conversation", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Get emails in a conversation thread by conversation topic. Returns truncated=true when maxResults cuts off additional messages.")]
    public async Task<ConversationResult> GetConversation(
        [Description("The conversation topic string to search for.")] string conversationTopic,
        [Description("Folder path to search in (e.g. 'Mailbox Name/Inbox'). Omit to search default inbox.")] string? folderPath = null,
        [Description("Maximum number of results. Default 50, max 200.")] int maxResults = 50)
    {
        conversationTopic = InputValidator.RequireNonBlank(conversationTopic, "conversationTopic");
        folderPath = InputValidator.NullIfWhiteSpace(folderPath);
        maxResults = InputValidator.ValidateMaxResults(maxResults);

        return await _session.ExecuteAsync(ns =>
        {
            var targetFolder = FolderResolver.ResolveMailFolderOrDefault(ns, folderPath);

            var escaped = FolderResolver.EscapeDasl(conversationTopic);
            var dasl = $"@SQL=\"urn:schemas:httpmail:thread-topic\" = '{escaped}'";

            dynamic items = targetFolder.Items.Restrict(dasl);
            items.Sort("[ReceivedTime]", true);

            var result = new List<EmailSummary>();
            var truncated = false;
            foreach (var item in items)
            {
                bool isMailItem;
                if (!ItemMapper.TryIsMailItem(item, out isMailItem))
                    throw new ModelContextProtocol.McpException("Conversation result item could not be inspected: required property 'Class' could not be read.");

                if (result.Count >= maxResults)
                {
                    truncated = true;
                    break;
                }

                if (!isMailItem)
                    continue;

                result.Add(ItemMapper.ToEmailSummary(item));
            }

            return new ConversationResult { Results = result, Truncated = truncated };
        });
    }
}
