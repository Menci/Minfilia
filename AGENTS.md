# AGENTS.md

Read-only MCP server for Microsoft Outlook (COM interface). .NET Framework 4.8, C#.

## Maintenance Policy

- This repo does not preserve backward compatibility. Ever.
- If a public tool name, parameter name, field name, or payload shape is misleading, replace it immediately. Do not add aliases, fallbacks, deprecation layers, or compatibility branches.
- Do not carry technical debt for old clients. The current contract must stay minimal, explicit, and semantically honest.
- Breaking changes are preferred over keeping wrong names alive.

## Architecture

```
Minfilia.BuildTasks    → MSBuild task: bundles all DLLs into Minfilia.exe via Mono.Cecil + GZip
Minfilia.Impl          → Main project: HTTP server + MCP tools + Outlook COM wrapper
Minfilia               → Single-file exe entry point (Loader.cs unpacks embedded assemblies)
```

## Key files

- `src/Minfilia.Impl/Program.cs` — Entry point, tool registration, server startup
- `src/Minfilia.Impl/Http/McpHttpServer.cs` — HttpListener ↔ StreamableHttpServerTransport glue (~120 lines)
- `src/Minfilia.Impl/Outlook/OutlookSession.cs` — COM singleton on dedicated STA thread, all COM calls marshaled via BlockingCollection
- `src/Minfilia.Impl/Tools/*.cs` — MCP tool implementations (StoreTools, SearchTools, EmailTools, CalendarTools, ContactTools)
- `src/Minfilia/Loader.cs` — Runtime for the `Minfilia` single-file entry point: decompress GZip resource pack, load assemblies via AssemblyResolve hook
- `src/Minfilia.BuildTasks/BundleTask.cs` — Build-time: scan output DLLs, serialize + GZip, embed into exe via Mono.Cecil

## Patterns

- **Performance model**: Local, single-user, low-concurrency server. Prioritize Outlook COM correctness, contract honesty, and round-trip usability over throughput.
- **COM threading**: All Outlook COM calls go through `OutlookSession.ExecuteAsync<T>()` which marshals to a dedicated STA thread. Never call COM directly from MCP handler threads.
- **Late-bound COM**: Uses `dynamic` / `Type.GetTypeFromProgID("Outlook.Application")` — no Office PIA dependency.
- **AI-agent-facing contract**: Keep schemas plain. Optional display fields are just nullable strings and may be omitted from JSON when null. Do not add field-level "failed" / "unavailable" metadata unless it changes the meaning of the result set.
- **Mail search**: `search_emails` uses DASL `Restrict` for server-side filtering (subject, date, sender), then client-side `Contains` on subject/body for keyword matching. Mail date filters use whole-day semantics (`dateTo` inclusive). Pagination is stateless; `truncated=true` explicitly exposes the scan cap, and `mayHaveMissedMatches=true` means some items could not be fully inspected for client-side keyword matching.
- **Contact search**: `search_contacts` intentionally keeps the matching path simple for this local low-volume server: enumerate contact folders and match client-side on a few display fields from actual contact items. `skippedFolders` means a contact folder could not be enumerated, and `mayHaveMissedMatches=true` means some contact items could not be fully inspected during matching.
- **Calendar listing**: `list_calendars` marks the default calendar with `isDefault=true` and returns it first so agents do not have to infer the primary calendar from store order.
- **Calendar range semantics**: `get_calendar_events` uses a widened Outlook `Restrict` prefilter, then enforces exact local overlap client-side to avoid recurrence/time-zone spillover across adjacent days. Calendar bodies are omitted unless explicitly requested.
- **Time semantics**: Email and calendar timestamps are emitted in local Outlook/Windows time with an explicit UTC offset. Do not emit floating local timestamps without an offset.
- **Stateless MCP**: Each HTTP POST creates a fresh `StreamableHttpServerTransport` + `McpServer`, processes one JSON-RPC request, then disposes. No session state.
- **Identity contract**: Store names and folder paths use Outlook display names for readable MCP round-trip behavior. Do not swap them to opaque Outlook IDs unless the entire input/output path contract changes together.
- **Naming contract**: Public names should be explicit. Prefer `storeName`, `folderPath`, and `calendarPath` over vague legacy names.
- **Error semantics**: Required identity/semantic fields fail closed. Blank required strings and out-of-range limits fail closed instead of being silently normalized. Optional display fields are best-effort instead of being wrapped in extra diagnostic metadata: blank strings stay blank, unreadable fields become missing values in MCP JSON output, and the first read failure for each optional-property shape is logged. Folder/calendar counts may be missing when Outlook does not expose them. `list_stores.email` may be missing when Outlook does not expose SMTP or account-to-store binding information. Bounded result tools expose `truncated` instead of silently cutting results; `search_emails` and `search_contacts` expose `mayHaveMissedMatches` when some items could not be fully inspected, and contact search also exposes `skippedFolders` when a folder had to be skipped.
- **Single-file packaging**: Marina-style BundleTask embeds all managed DLLs as a GZip-compressed resource. Loader.cs extracts and loads them at startup via `AppDomain.AssemblyResolve`.

## Code style

Follows [Marina](https://github.com/Menci/Marina) conventions:
- Primary constructors with `_camelCase` parameters as private fields
- PascalCase public members, `_camelCase` private fields
- `<Nullable>enable</Nullable>` globally, with `CS0436` suppressed for the known `dotnetCampus.LatestCSharpFeatures` generated `Index`/`Range` conflict on net48
- File-scoped namespaces
- 4-space indent, LF line endings

## Build

```bash
dotnet build src/Minfilia/Minfilia.csproj -c Release
```

This builds all 3 projects in dependency order (BuildTasks → Minfilia.Impl → Minfilia + BundleTask).

## Testing

Start the server, then verify with curl:

```bash
Minfilia.exe 3027

# In another terminal:
curl -X POST http://localhost:3027/ \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"0.1"}}}'
```
