"""
Smoke tests for Minfilia, an Outlook MCP server.

Covers all behavioral fixes from review rounds 1-3:
  - Path resolution: invalid paths fail closed, no silent fallback
  - Path round-trip: list_* paths can be fed back to get_*/search_*
  - Calendar overlap semantics: single-day query finds all-day events
  - Store binding: email bound to correct store
  - Folder type validation: calendar tools reject mail folders and vice versa
  - Item type validation: get_email rejects non-mail items
  - Input validation: blank required strings, bad cursor, bad dates, inverted ranges, out-of-range limits
  - Truncation: scan cap is explicit in response
  - search_emails dateTo is inclusive for whole-day filtering
  - Optional display fields are best-effort instead of carrying per-field diagnostic metadata
  - Description honesty: list_folders says "folders" not "mail folders"
  - Contract clarity: explicit tool/field/parameter names replace vague legacy names
  - Models are typed and expose explicit failure states

Usage:
    1. Start the server:  Minfilia.exe [port]
    2. Run tests:         python test_smoke.py [port]
"""

import json
import sys
import urllib.request
from datetime import datetime, timedelta

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 3027
URL = f"http://localhost:{PORT}/"
TIMEOUT = 120
ID = 0
PASS = FAIL = 0
FAILURES = []


def call(method, params=None):
    global ID
    ID += 1
    req = urllib.request.Request(
        URL,
        json.dumps({"jsonrpc": "2.0", "id": ID, "method": method, "params": params or {}}).encode(),
        headers={"Content-Type": "application/json", "Accept": "application/json, text/event-stream"},
    )
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        body = r.read().decode()
    for line in reversed(body.strip().split("\n")):
        if line.startswith("data:"):
            return json.loads(line[5:])
    return json.loads(body)


def tool(name, args=None):
    r = call("tools/call", {"name": name, "arguments": args or {}})
    if "error" in r:
        message = (r.get("error") or {}).get("message", "")
        return {"_error": True, "_msg": [message] if message else []}

    result = r.get("result", r)
    if result.get("isError"):
        return {"_error": True, "_msg": [c.get("text", "") for c in result.get("content", [])]}
    c = result.get("content", [])
    if c and c[0].get("type") == "text":
        return json.loads(c[0]["text"])
    return c


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        FAILURES.append(name)
        print(f"  FAIL  {name}  {detail}")


def is_str_or_null(value):
    return isinstance(value, (str, type(None)))


def overlaps_local_date_range(event, date_from, date_to):
    start = datetime.fromisoformat(event["startTime"])
    end = datetime.fromisoformat(event["endTime"])
    range_start = datetime.fromisoformat(date_from)
    range_end_exclusive = datetime.fromisoformat(date_to) + timedelta(days=1)
    return start < range_end_exclusive and end > range_start


def has_utc_offset(value):
    return datetime.fromisoformat(value).tzinfo is not None


def has_keys(obj, keys):
    return isinstance(obj, dict) and all(key in obj for key in keys)


def is_success_result(result):
    return isinstance(result, dict) and not result.get("_error")


def get_input_properties(tool_def):
    return ((tool_def or {}).get("inputSchema") or {}).get("properties") or {}


def find_mail_folder(folders):
    for folder in folders or []:
        path = folder.get("path")
        if not path:
            continue

        result = tool("search_emails", {"folderPath": path, "maxResults": 1})
        if is_success_result(result):
            return path

    return None


def find_calendar_event(calendars):
    candidates = sorted(calendars or [], key=lambda c: (c.get("itemCount") or 0), reverse=True)
    for calendar in candidates:
        path = calendar.get("path")
        if not path:
            continue

        result = tool("get_calendar_events", {
            "dateFrom": "2025-01-01",
            "dateTo": "2026-12-31",
            "calendarPath": path,
            "maxResults": 1,
        })
        if is_success_result(result) and result.get("results"):
            return path, result["results"][0]

    return None, None


def main():
    print(f"Minfilia smoke tests — server at {URL}\n")

    # --- Initialize ---
    init = call("initialize", {
        "protocolVersion": "2024-11-05",
        "capabilities": {},
        "clientInfo": {"name": "smoke-test", "version": "1.0"},
    })
    server_info = init["result"]["serverInfo"]
    print(f"  Server: {server_info['name']} v{server_info['version']}\n")

    # =========================================================================
    # Basic functionality
    # =========================================================================
    print("--- Basic functionality ---")

    stores = tool("list_stores")
    check("list_stores returns list", isinstance(stores, list) and len(stores) > 0)

    folders = tool("list_folders", {"maxDepth": 1})
    check("list_folders returns list", isinstance(folders, list) and len(folders) > 0)

    calendars = tool("list_calendars")
    check("list_calendars returns list", isinstance(calendars, list) and len(calendars) > 0)
    if isinstance(calendars, list) and calendars:
        check("CalendarInfo.isDefault is bool", all(isinstance(c.get("isDefault"), bool) for c in calendars))
        default_calendars = [c for c in calendars if c.get("isDefault")]
        check("at most one default calendar", len(default_calendars) <= 1, str(default_calendars[:2]))
        if default_calendars:
            check("default calendar is listed first", calendars[0]["path"] == default_calendars[0]["path"])

    emails = tool("search_emails", {"maxResults": 1})
    check("search_emails returns results", isinstance(emails, dict) and "results" in emails)

    contacts = tool("search_contacts", {"query": "test"})
    check("search_contacts returns results envelope",
          isinstance(contacts, dict)
          and isinstance(contacts.get("results"), list)
          and isinstance(contacts.get("truncated"), bool)
          and isinstance(contacts.get("mayHaveMissedMatches"), bool))

    mail_folder = find_mail_folder(folders)
    _calendar_path_with_event, calendar_event = find_calendar_event(calendars)

    # =========================================================================
    # Path resolution — invalid paths must error (Review 1 #1)
    # =========================================================================
    print("\n--- Invalid path → error (not fallback) ---")

    r = tool("search_emails", {"folderPath": "NoSuchStore/NoSuchFolder"})
    check("search_emails invalid folder → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_conversation", {"conversationTopic": "test", "folderPath": "NoSuchStore/X"})
    check("get_conversation invalid folder → error", isinstance(r, dict) and r.get("_error"))

    r = tool("list_folders", {"parentPath": "NoSuchStore/Calendar/NoSuchFolder"})
    check("list_folders invalid parentPath → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "2026-04-08", "calendarPath": "NoSuch/Cal"})
    check("get_calendar_events invalid path → error", isinstance(r, dict) and r.get("_error"))

    # =========================================================================
    # Path round-trip — list_* paths feed back to get_* (Review 1 #2)
    # =========================================================================
    print("\n--- Path round-trip ---")

    if calendars:
        nested = [c for c in calendars if c["path"].count("/") >= 2]
        if nested:
            cal = nested[0]
            events = tool("get_calendar_events", {
                "dateFrom": "2025-01-01", "dateTo": "2026-12-31",
                "calendarPath": cal["path"], "maxResults": 1,
            })
            check(f"calendar path round-trips: {cal['path']}",
                  isinstance(events, dict) and isinstance(events.get("results"), list))

    if folders:
        f = folders[0]
        sub = tool("list_folders", {"parentPath": f["path"], "maxDepth": 0})
        check(f"folder path round-trips: {f['path']}", isinstance(sub, list))

    # =========================================================================
    # Calendar overlap semantics — single-day all-day event (Review 1 #3)
    # =========================================================================
    print("\n--- Calendar overlap semantics ---")

    holiday_cals = [c for c in (calendars or []) if "holiday" in c["name"].lower() and (c.get("itemCount") or 0) > 0]
    if holiday_cals:
        events_result = tool("get_calendar_events", {
            "dateFrom": "2026-01-19", "dateTo": "2026-01-19",
            "calendarPath": holiday_cals[0]["path"],
        })
        events = events_result.get("results", []) if isinstance(events_result, dict) else []
        found = [e for e in events if "king" in (e.get("subject") or "").lower()]
        check("MLK day found with single-day query", len(found) > 0,
              f"events: {[(e.get('subject') or '') for e in events[:3]]}")

    # =========================================================================
    # Store binding — email on correct store (Review 1 #4)
    # =========================================================================
    print("\n--- Store binding ---")

    if stores:
        with_email = [s for s in stores if s.get("email")]
        primary = [s for s in with_email if "archive" not in s["storeName"].lower()]
        check("primary store has email (not archive)",
              len(primary) > 0 or len(with_email) == 0,
              f"stores: {[(s['storeName'][:30], s.get('email')) for s in with_email]}")

    # =========================================================================
    # Folder type validation (Review 2 #2)
    # =========================================================================
    print("\n--- Folder type validation ---")

    if mail_folder:
        r = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "2026-04-08", "calendarPath": mail_folder})
        check("calendar tool rejects mail folder", isinstance(r, dict) and r.get("_error"))

    if calendars:
        r = tool("search_emails", {"folderPath": calendars[0]["path"], "maxResults": 1})
        check("search tool rejects calendar folder", isinstance(r, dict) and r.get("_error"))

    # =========================================================================
    # Item type validation — get_email rejects non-mail items (Review 2 #1)
    # =========================================================================
    print("\n--- Item type validation ---")

    if calendar_event:
        r = tool("get_email", {"id": calendar_event.get("entryId", "")})
        check("get_email rejects calendar item", isinstance(r, dict) and r.get("_error"))

    r = tool("get_email", {"id": "INVALID_ENTRY_ID"})
    check("get_email rejects bad EntryID", isinstance(r, dict) and r.get("_error"))

    # =========================================================================
    # Input validation (Review 2 #3, Review 3 #5)
    # =========================================================================
    print("\n--- Input validation ---")

    r = tool("search_emails", {"cursor": "not-base64!!!"})
    check("bad cursor → error", isinstance(r, dict) and r.get("_error"))

    r = tool("search_contacts", {"query": "   "})
    check("blank contact query → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_conversation", {"conversationTopic": "   "})
    check("blank conversation topic → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_email", {"id": "   "})
    check("blank EntryID → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "bad", "dateTo": "2026-04-08"})
    check("bad dateFrom → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "04/07/2026", "dateTo": "2026-04-08"})
    check("non-ISO dateFrom → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "bad"})
    check("bad dateTo → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "2026-04-30", "dateTo": "2026-04-01"})
    check("inverted date range → error", isinstance(r, dict) and r.get("_error"))

    r = tool("search_emails", {"dateFrom": "2026-04-30", "dateTo": "2026-04-01"})
    check("search inverted range → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "2026-04-08", "maxResults": -1})
    check("negative maxResults → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_conversation", {"conversationTopic": "test", "maxResults": -1})
    check("get_conversation neg maxResults → error", isinstance(r, dict) and r.get("_error"))

    r = tool("search_emails", {"maxResults": 101})
    check("search maxResults above cap → error", isinstance(r, dict) and r.get("_error"))

    r = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "2026-04-08", "maxResults": 201})
    check("calendar maxResults above cap → error", isinstance(r, dict) and r.get("_error"))

    r = tool("list_folders", {"maxDepth": 11})
    check("list_folders maxDepth above cap → error", isinstance(r, dict) and r.get("_error"))

    r = tool("search_emails", {"query": "   ", "folderPath": "   ", "maxResults": 1})
    check("blank optional filters normalize to omitted", isinstance(r, dict) and "results" in r)

    if emails and emails.get("results"):
        sample_email = emails["results"][0]
        sample_date = sample_email.get("date", "")[:10]
        if sample_date:
            r = tool("search_emails", {"dateFrom": sample_date, "dateTo": sample_date, "maxResults": 100})
            check("search_emails dateTo is inclusive",
                  is_success_result(r) and any(item.get("id") == sample_email.get("id") for item in r.get("results", [])),
                  str(r)[:120])

            if sample_email.get("from"):
                r = tool("search_emails", {
                    "dateFrom": sample_date,
                    "dateTo": sample_date,
                    "from": sample_email["from"],
                    "maxResults": 100,
                })
                check("search_emails from filter round-trips from summary",
                      is_success_result(r) and any(item.get("id") == sample_email.get("id") for item in r.get("results", [])),
                      str(r)[:120])

    # =========================================================================
    # Truncation is explicit (Review 3 #3)
    # =========================================================================
    print("\n--- Truncation ---")

    r = tool("search_emails", {"maxResults": 1, "cursor": "NDk5OQ=="})  # offset 4999
    if isinstance(r, dict) and not r.get("_error"):
        check("truncated field present in SearchResult", "truncated" in r)
    else:
        check("high-offset search handled", True)

    # High offset beyond scan limit: truncated=true, no nextCursor, empty results
    r = tool("search_emails", {"maxResults": 1, "cursor": "NjAwMA=="})  # offset 6000
    if isinstance(r, dict) and not r.get("_error"):
        check("offset>5000: truncated=true", r.get("truncated") == True)
        check("offset>5000: results empty", len(r.get("results", [])) == 0)
        check("offset>5000: no nextCursor (can't continue)", r.get("nextCursor") is None,
              f"nextCursor={r.get('nextCursor')}")
    else:
        check("offset>5000 handled gracefully", not r.get("_error"), str(r)[:100])

    # =========================================================================
    # Description honesty (Review 3 #6)
    # =========================================================================
    print("\n--- Contract descriptions ---")

    tools_list = call("tools/list").get("result", {}).get("tools", [])
    tool_names = {t.get("name") for t in tools_list}
    check("tool list exposes list_stores", "list_stores" in tool_names)
    check("tool list removed list_accounts", "list_accounts" not in tool_names)

    legacy_accounts = tool("list_accounts")
    check("legacy list_accounts call errors", isinstance(legacy_accounts, dict) and legacy_accounts.get("_error"))

    lf = next((t for t in tools_list if t["name"] == "list_folders"), None)
    if lf:
        check("list_folders desc: no 'mail'", "mail" not in lf["description"].lower(), lf["description"][:60])
        check("list_folders desc mentions missing counts", "may be missing" in lf["description"].lower())
        lf_props = get_input_properties(lf)
        check("list_folders schema uses storeName", "storeName" in lf_props and "storeId" not in lf_props)

    ls = next((t for t in tools_list if t["name"] == "list_stores"), None)
    if ls:
        check("list_stores desc mentions missing email", "may be missing" in ls["description"].lower())

    lc = next((t for t in tools_list if t["name"] == "list_calendars"), None)
    if lc:
        check("list_calendars desc mentions missing itemCount", "may be missing" in lc["description"].lower())
        lc_props = get_input_properties(lc)
        check("list_calendars schema uses storeName", "storeName" in lc_props and "storeId" not in lc_props)

    se = next((t for t in tools_list if t["name"] == "search_emails"), None)
    if se:
        check("search_emails desc: no 'full-text'", "full-text" not in se["description"].lower())
        check("search_emails desc mentions mayHaveMissedMatches", "mayhavemissedmatches" in se["description"].lower())
        se_props = get_input_properties(se)
        check("search_emails schema uses folderPath", "folderPath" in se_props and "folder" not in se_props)

    gc = next((t for t in tools_list if t["name"] == "get_conversation"), None)
    if gc:
        gc_props = get_input_properties(gc)
        check("get_conversation schema uses folderPath", "folderPath" in gc_props and "folder" not in gc_props)

    ge = next((t for t in tools_list if t["name"] == "get_calendar_events"), None)
    if ge:
        ge_props = get_input_properties(ge)
        check("get_calendar_events schema uses calendarPath", "calendarPath" in ge_props and "calendarName" not in ge_props)

    sc = next((t for t in tools_list if t["name"] == "search_contacts"), None)
    if sc:
        check("search_contacts desc mentions mayHaveMissedMatches", "mayhavemissedmatches" in sc["description"].lower())
        check("search_contacts desc mentions company/job title", "company" in sc["description"].lower() and "job title" in sc["description"].lower())

    # =========================================================================
    # Typed models (Review 1 #7, Review 3 #11)
    # =========================================================================
    print("\n--- Typed models ---")

    if stores:
        s = stores[0]
        check("StoreInfo has core fields", has_keys(s, ["storeName"]))
        check("StoreInfo.email is str-or-null", isinstance(s.get("email"), (str, type(None))))
        check("StoreInfo removed legacy field names", "name" not in s and "storeId" not in s)

    if folders:
        f = folders[0]
        check("FolderInfo has core fields", has_keys(f, ["path", "name"]))
        check("FolderInfo counts are int-or-null",
              isinstance(f.get("itemCount"), (int, type(None))) and isinstance(f.get("unreadCount"), (int, type(None))))

    if calendars:
        c = calendars[0]
        check("CalendarInfo has core fields", has_keys(c, ["name", "storeName", "path"]))
        check("CalendarInfo.itemCount is int-or-null", isinstance(c.get("itemCount"), (int, type(None))))
        check("CalendarInfo removed legacy store field", "store" not in c)

    if isinstance(contacts, dict) and not contacts.get("_error"):
        check("ContactSearchResult has expected fields",
              has_keys(contacts, ["results", "truncated", "skippedFolders", "mayHaveMissedMatches"]))
        check("ContactSearchResult.skippedFolders is list", isinstance(contacts.get("skippedFolders"), list))
        check("ContactSearchResult.mayHaveMissedMatches is bool", isinstance(contacts.get("mayHaveMissedMatches"), bool))
        check("ContactSearchResult has no unavailableMatchFields", "unavailableMatchFields" not in contacts)
        if contacts.get("results"):
            contact = contacts["results"][0]
            check("ContactInfo has no unavailableFields", "unavailableFields" not in contact)
            check("ContactInfo display fields are str-or-null",
                  all(is_str_or_null(contact.get(k)) for k in ["name", "email", "company", "phone", "jobTitle"]))

    if emails and emails.get("results"):
        e = emails["results"][0]
        check("SearchResult has expected fields", has_keys(emails, ["results", "truncated", "mayHaveMissedMatches"]))
        check("EmailSummary has core fields", has_keys(e, ["id", "date"]))
        check("EmailSummary display fields are str-or-null",
              all(is_str_or_null(e.get(k)) for k in ["subject", "from", "preview"]))
        check("EmailSummary has no unavailableFields", "unavailableFields" not in e)
        check("EmailSummary.date has UTC offset", has_utc_offset(e["date"]), e["date"])
        check("SearchResult.mayHaveMissedMatches is bool", isinstance(emails.get("mayHaveMissedMatches"), bool))
        check("SearchResult has no unavailableMatchFields", "unavailableMatchFields" not in emails)

    email_id = emails["results"][0]["id"] if emails and emails.get("results") else None
    if email_id:
        detail = tool("get_email", {"id": email_id})
        check("get_email on search result succeeds", is_success_result(detail), str(detail)[:120])
        if is_success_result(detail):
            check("EmailDetail has core fields", has_keys(detail, [
                "id", "subject", "date", "importance", "isRead", "attachments", "attachmentsReadFailed"
            ]))
            check("EmailDetail.date is non-empty", detail.get("date", "") != "")
            check("EmailDetail.date has UTC offset", has_utc_offset(detail["date"]), detail["date"])
            check("EmailDetail.importance is int", isinstance(detail.get("importance"), int))
            check("EmailDetail.attachmentsReadFailed is bool", isinstance(detail.get("attachmentsReadFailed"), bool))
            check("EmailDetail has no unavailableFields", "unavailableFields" not in detail)
            check("EmailDetail optional fields are str-or-null",
                  all(is_str_or_null(detail.get(k)) for k in [
                      "from", "fromEmail", "to", "cc", "body", "htmlBody", "conversationTopic", "conversationId"
                  ]))
            if detail.get("attachments"):
                attachment = detail["attachments"][0]
                check("AttachmentInfo has core fields", has_keys(attachment, ["size"]))
                check("AttachmentInfo.name is str-or-null", is_str_or_null(attachment.get("name")))
                check("AttachmentInfo has no unavailableFields", "unavailableFields" not in attachment)

            topic = detail.get("conversationTopic")
            if topic:
                conversation = tool("get_conversation", {"conversationTopic": topic, "maxResults": 1})
                check("get_conversation on detail topic succeeds", is_success_result(conversation), str(conversation)[:120])
                check("ConversationResult exposes results/truncated",
                      isinstance(conversation, dict) and isinstance(conversation.get("results"), list) and isinstance(conversation.get("truncated"), bool))

    calendar_result = tool("get_calendar_events", {"dateFrom": "2026-04-07", "dateTo": "2026-04-08", "maxResults": 1})
    if isinstance(calendar_result, dict) and not calendar_result.get("_error"):
        check("CalendarEventsResult has expected fields", has_keys(calendar_result, ["results", "truncated"]))
        check("CalendarEventsResult exposes results/truncated",
              isinstance(calendar_result.get("results"), list) and isinstance(calendar_result.get("truncated"), bool))
        if calendar_result.get("results"):
            event = calendar_result["results"][0]
            check("CalendarEvent has core fields",
                  has_keys(event, ["entryId", "startTime", "endTime", "isRecurring", "isCancelled"]))
            check("CalendarEvent has no unavailableFields", "unavailableFields" not in event)
            check("CalendarEvent optional fields are str-or-null",
                  all(is_str_or_null(event.get(k)) for k in ["subject", "location", "organizer", "body"]))
            check("CalendarEvent.startTime has UTC offset", has_utc_offset(event["startTime"]), event["startTime"])
            check("CalendarEvent.endTime has UTC offset", has_utc_offset(event["endTime"]), event["endTime"])
            check("CalendarEvent.isCancelled is bool", isinstance(event.get("isCancelled"), bool))
            check("calendar event overlaps requested local date range",
                  overlaps_local_date_range(event, "2026-04-07", "2026-04-08"),
                  str(event))
            check("calendar body omitted by default", "body" not in event or event.get("body") is None)

    # =========================================================================
    # StoreName round-trip
    # =========================================================================
    print("\n--- StoreName round-trip ---")

    if stores:
        store_name = stores[0]["storeName"]
        r = tool("list_folders", {"storeName": store_name, "maxDepth": 0})
        check(f"list_folders(storeName='{store_name[:30]}...') works", isinstance(r, list) and len(r) > 0, str(r)[:80])

    # =========================================================================
    print(f"\n{'=' * 50}")
    print(f"Results: {PASS} passed, {FAIL} failed")
    if FAILURES:
        print(f"Failures:")
        for f in FAILURES:
            print(f"  - {f}")
    print()

    sys.exit(1 if FAIL > 0 else 0)


if __name__ == "__main__":
    main()
