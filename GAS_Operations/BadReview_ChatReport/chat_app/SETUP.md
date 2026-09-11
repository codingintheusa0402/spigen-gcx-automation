# BadReview Chat app — setup

A **proof of concept** that runs alongside the webhook broadcast
(`../badreview_chat_report.py`). It does not replace it. Decide after you've seen it
whether to keep it.

What it does: posts a card with a **date picker + product dropdown + [조회]** button.
Pick a date → the app re-reads the `1-3점` tab and re-renders the Pixel 11 / Galaxy Z8
배드리뷰 (1~3점) report for the rows whose `Update 날짜` matches that date (same card
layout as the webhook: colored 대분류 headers, decoratedText Top 5, day breakdown,
배드리뷰 button).

**Why this needs an app and not a webhook:** an incoming webhook is send-only — it
can't receive the "user picked a date" event. A Chat app has an endpoint that does.

## Files

| File | Purpose |
|------|---------|
| `Config.gs` | `APP_PRODUCT` — the one product this deployment reports (`pixel11` here, `glxz8` in `../chat_app_jane/`) |
| `Code.gs` | all handlers + card builder + sheet crunching — **identical copy in `../chat_app_jane/`; re-copy after edits** |
| `appsscript.json` | manifest — `addOns.common` (name, logoUrl) + `addOns.chat: {}`; scope: `spreadsheets` |

### Card behaviour (v9+, 2026-09-11)

- Control card: 시작일 / 종료일 DATE_ONLY pickers (default = earliest `Update 날짜` on the tab →
  today, shown in the card subtitle because the pickers themselves render empty), 국가 dropdown
  (`국가(tag)` values + 전체 국가), 기종 dropdown (`기종명` values + 전체 기종), [조회].
- Report card: one product only; **every number is scoped to the range + filters** (the webhook
  card's Top 5 is cumulative — this one is not). Title `✔️ M/D(요일)~M/D(요일) <product> 배드리뷰
  (1~3점) (총 N건)`; filters echoed in the subtitle / section header.
- Message text also works: `9/1~9/11` → that range, `9/11` → one day, anything else → default.
  Add-on events carry the text at `event.chat.messagePayload.message.text` (`messageText_`).

## Live deployment (2026-09-11)

| | Kevin identity |
|---|---|
| Apps Script | `1mfLtA5elEbfi3ghvUnbEeQX220EOndmMVb2mskkxaFIwkBlEaqdy8-jc` (this folder's `.clasp.json`) |
| GCP project | `tctnotifier` (#769852651633) — **not** gcxbot (gcxbot already hosts the live "T2 Report" Chat app; one Chat app identity per GCP project) |
| Chat app name / avatar | `김지우 Kevin 글로벌CX전략팀` / Kevin's profile photo |
| Deployment ID | `AKfycbyAj5jQ_3NmppD3lrLNaiq68_tGJDBT_Qjs_xn21ncKcrRqegJ-iMAyS20zO6QSTaMx` — roll new versions with `clasp deploy -i <id>`; the Console keeps pointing at it |
| Visibility | specific people: `kjw@spigen.com` (widen in Chat API → Configuration → Visibility) |

| | Jane identity (`../chat_app_jane/`) |
|---|---|
| Apps Script | `1Zhc91kpARwwlNctKsk2THxt8nIvj6SvuJMOFOaKzrJotbWsKobSxbFhI` |
| GCP project | `formats-uaox` (#1009119937520) — Chat API enabled + OAuth consent (Internal) created 2026-09-11 |
| Chat app name / avatar | `나아름 Jane 글로벌CX전략팀` / Jane's profile photo |
| Deployment ID | `AKfycbysje3fhJaLZ-VPezr2GTnZRJNJjpFAp22bHSExro0aFzP2IuNyDIyI1i9ixoxnZgJC` |

Each identity needs its own Apps Script project because an Apps Script project can be attached to
only one GCP project; `chat_app_jane/Code.gs` is a copy of this folder's — re-copy after edits.

### Lessons that cost time (read before repeating this on another project)

1. **`chat.bot` is not a user-consentable scope.** With it in `oauthScopes` every authorization
   attempt dies with `Error 400: invalid_scope — Some requested scopes cannot be shown`. This app
   never calls the Chat API (it only *returns* cards), so the scope is unnecessary. Removed.
2. **`SpreadsheetApp.openById` needs the full `.../auth/spreadsheets` scope**, not `.readonly`
   (the add-on runtime rejects it: "Specified permissions are not sufficient").
3. **The Chat API Configuration checkbox "Build this Chat app as a Workspace add-on" is on by
   default and becomes read-only once saved.** TCTNotifier is therefore locked in add-on mode.
   Consequences:
   - manifest must use `addOns.common` + `addOns.chat: {}` (top-level `chat: {}` is legacy);
   - every response must be wrapped in the add-on envelope —
     `{hostAppDataAction:{chatDataAction:{createMessageAction:{message}}}}` for new messages and
     `updateMessageAction` for in-place updates (`chatCreate_` / `chatUpdate_` in `Code.gs`);
     a bare `{cardsV2: [...]}` is rejected with "Invalid add-on response returned";
   - handler names are set in the Console (Triggers: `onMessage`, `onAddToSpace`,
     `onRemoveFromSpace`; the console defaults `onAddedToSpace`/`onRemovedFromSpace` do not
     match `Code.gs`), and button clicks call `onClick.action.function` (`refreshReport`)
     directly. `onCardClick` is kept only as a shim for the classic model.
   - end users see an **Install app** dialog on first DM and a "requires configuration →
     Configure" card the first time a new scope is needed.
4. **The `Run` → authorize popup in the Apps Script editor is blocked when driven by browser
   automation** — a human has to click it. Same for the Chat "Configure" link.
5. Project quota: `kjw@spigen.com` cannot create new GCP projects (limit reached) — reuse an
   existing unused project, and check its Chat API → Configuration page first: GCXDM-Dialogflow
   ("GCX챗봇") and GCX Zendesk Decision Maker ("GCX Internal Assistant") are live, Spigen Bot Beta is
   pending deletion.

## One-time setup (manual — needs Cloud Console)

1. **Create the Apps Script project.** Either:
   - `clasp create --type standalone --title "BadReview Chat App"` in this folder,
     then `clasp push` (`clasp` must be logged in as the account that can read BOTH
     spreadsheets — see the ID table in `../README.md`), **or**
   - script.google.com → New project → paste `Code.gs`, and in Project Settings tick
     "Show appsscript.json" then paste `appsscript.json`.
2. **Attach a standard GCP project** (Apps Script → Project Settings → Google Cloud
   Platform (GCP) Project → Change project) whose number you control.
3. **Enable the Google Chat API** in that GCP project (APIs & Services → Enable APIs
   → "Google Chat API").
4. **Get a deployment ID:** Apps Script → Deploy → **New deployment** → gear icon →
   there's no "Chat" type in the list; just create the deployment (the default is
   fine) and **copy the Deployment ID**. For quick testing you can instead use
   Deploy → **Test deployments** → copy the **Head deployment** ID.
5. **Configure the Chat app:** GCP Console → Google Chat API → **Configuration**:
   - App name: `BadReview 리포트`  ·  Avatar URL: any 256px HTTPS image  ·  Description
   - Functionality: tick **Receive 1:1 messages** and **Join spaces and group
     conversations**
   - Connection settings: **Apps Script** → paste the **Deployment ID** from step 4
   - Slash commands (optional): add `/badreview` → command id 1
   - Visibility: make available to **specific people** (yourself) first for testing,
     widen later
   - Save
6. **Authorize:** open `Code.gs`, run `onMessage` once from the editor to trigger the
   OAuth consent (grants the sheet + chat scopes). The account must have **read
   access to both spreadsheets**.

## Test

- DM the app, or `@BadReview 리포트` in a space it's added to, or `/badreview`.
- You get the control card + today's Pixel 11 and Z8 report cards.
- Change the date, press **조회** → the message updates in place
  (`actionResponse.type = UPDATE_MESSAGE`).
- Typing a date in the message text also works: `@BadReview 리포트 9/5`.

## Notes / limits

- `긍정 리뷰` 인입사유(tag) is excluded from every count and the card (`EXCLUDED_TAGS`
  in `Code.gs`, mirrors the Python) — user rule 2026-09-08.

- Interaction only works where the **app itself** is present — it won't retrofit onto
  the existing webhook messages in the 12 broadcast rooms.
- `refreshReport` is the button handler (Apps Script routes `CARD_CLICKED` to the
  global named in `onClick.action.function`).
- DATE_ONLY returns a UTC-midnight epoch; `refreshReport` converts it back with
  `getUTC*` so the picked calendar day is preserved regardless of the script TZ.
- Card design is duplicated from `../badreview_chat_report.py` — if the card layout
  changes there, mirror it in `Code.gs`.
- Reads via `SpreadsheetApp.openById`; if the deploying account loses sheet access the
  cards will error.
