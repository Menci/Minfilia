<p align="center">
  <img src="assets/icon.svg" width="200" height="200" alt="Minfilia">
</p>

<h1 align="center">Minfilia</h1>

<p align="center">
  Single-file Outlook MCP server for Windows via COM interface<br>
  The name is from the FFXIV character [Minfilia](https://ffxiv.consolegameswiki.com/wiki/Minfilia_Warde).
</p>

---

## Tools

| Tool | Description |
|---|---|
| `list_stores` | List Outlook stores; `email` may be missing if Outlook does not expose SMTP or account-to-store binding |
| `list_folders` | List folder hierarchy with item counts, which may be missing when Outlook refuses them |
| `search_emails` | Keyword substring search (subject + body) with date/sender filters, stateless pagination, and `mayHaveMissedMatches` |
| `get_email` | Read full email content by EntryID |
| `get_conversation` | Get a conversation thread by topic, with explicit `truncated` |
| `list_calendars` | List calendar folders; the default calendar is flagged with `isDefault` and listed first; `itemCount` may be missing when Outlook refuses it |
| `get_calendar_events` | Query events by exact local-date overlap, with structured `isCancelled`, optional body via `includeBody`, local-time timestamps with UTC offset, and explicit `truncated` |
| `search_contacts` | Search contacts by name, email, company, or job title, with explicit `truncated`, `skippedFolders`, and `mayHaveMissedMatches` |

## Usage

Download `Minfilia.exe` from [Releases](../../releases). Windows and desktop Outlook are required, and Outlook must already be running.

```
Minfilia.exe [port]
```

Default port is 3027 and MCP endpoint is `http://localhost:3027/` (Streamable HTTP, stateless).

You can also build the project with .NET SDK 8.0+ (target .NET Framework 4.8):

```
dotnet build src/Minfilia/Minfilia.csproj -c Release
```

For more detailed instructions, just ask your AI coding agent (Codex / Claude Code).

## License

MIT
