# BiWeekly View Log

Tracked link for the GCX Bi-weekly Report deck. Google Slides has no per-viewer open/duration log,
so instead of the deck link we share an Apps Script web app that embeds the deck and logs every open
— who, when, last activity and minutes on page — to a Google Sheet. Only opens through the tracked
link are logged; opening the deck directly is not.

## Screenshots

![`Summary` tab (viewer blurred)](docs/summary.jpg)
*`Summary` tab (viewer blurred)*

## How it works

- `doGet(e)` — serves `Page.html` (the deck embedded full-page) for report code `?r=<code>`
  (falls back to `DEFAULT_REPORT`). Page title `<code> GCX Bi-weekly Report`.
- `startVisit(code, ua)` — called once on page load; appends a row to the **Visits** tab
  (visit id, name, email, 시작, 마지막 활동, 체류 시간(분), report code, user agent) under a script lock and
  returns the visit id. The viewer's name is looked up in the Workspace directory (People API);
  falls back to the email local part.
- `beat(id)` — heartbeat every `HEARTBEAT_SEC` (30 s) while the tab is visible, plus on
  `visibilitychange` / `pagehide`; updates 마지막 활동 and 체류 시간(분) for that row.
- `authorize()` — run once from the editor so the deployer grants Sheets + directory scopes.

The **Summary** tab of the log sheet (QUERY formula, not written by the script) shows per-person
open count / total / average minutes / last visit.

## Config (`Code.js` constants)

| Constant | Value |
|---|---|
| `LOG_SHEET_ID` | `1NSMiMwz_4nd6NlOv8rvYeyGvBDmCRZefsWD0Ux0PfDg` (tabs `Visits`, `Summary`) |
| `REPORTS` | report code → deck ID, one line per period (`261002` → `1quCr9Xj-pSsVXKrYuaEOq0LPILZMN2LPUwBkY1f_GFI`) |
| `DEFAULT_REPORT` | `261002` |
| `HEARTBEAT_SEC` | `30` |

No Script Properties and no triggers.

## Deploy

Standalone script `1n85nRmLMbjdQ5M_csISeoaO0uLIg5hbZ_zR7kyAiw4PgIFhFMdjqxrSu`. Web app
`executeAs: USER_DEPLOYING`, `access: DOMAIN` (spigen.com accounts only — viewers need no access to
the log sheet). Advanced service People v1.

New period:
1. Add `'<code>': '<deckId>'` to `REPORTS` (and bump `DEFAULT_REPORT`).
2. `clasp push`
3. `clasp deploy -i <deploymentId>` — redeploys the existing deployment so the URL stays the same.
4. Share `https://script.google.com/a/macros/spigen.com/s/<deploymentId>/exec?r=<code>`.
