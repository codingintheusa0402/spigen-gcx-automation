# BadReview → Google Chat app-card report

`badreview_chat_report.py` builds and sends a Google Chat **cardsV2** app card for the
daily bad-review (★1~3) scrape of **Pixel 11 Series** and **Galaxy Z8 Series**, and can
fan it out to the GCX cross-team Chat rooms.

This is the standalone, version-controlled twin of the Claude Code skills
`pixel11-badreview-chat-report`, `glxz8-badreview-chat-report` and
`badreview-chat-broadcast` (`~/.claude/skills/…`). Behaviour and card layout are
identical; keep the two in sync when either changes.

## What the card shows

| Part | Content |
|------|---------|
| `header.title` | `✔️ M/D(요일) <product> 배드리뷰 (1~3점) (총 N건)` |
| `header.subtitle` | `고객 리뷰 ★1~3점 · <date range>` + a product thumbnail (`imageType: SQUARE`) |
| **Top 5 인입사유(누적)** | a 2-column widget: left = Top 5 `인입사유(tag)` where `대분류` = `휴대폰보호필름`, right = `휴대폰케이스`. Each column is headed `<b><font color>대분류</font></b> · {tot}건` (red `#EA4335` for 휴대폰보호필름, blue `#4285F4` for 휴대폰케이스, to set it apart from the rows below) then **5 `decoratedText` rows** — `topLabel` = `n위`, `text` = `<b>이유</b>`, `bottomLabel` = `c건 · p%` — so rank / 인입사유 / 건수·% each sit at a fixed left edge and the two columns line up (`p` = share of that 대분류). Counted over the whole `1-3점` sheet. |
| **오늘 M/D(요일) 최다 인입사유** | biggest `인입사유(tag)` among rows whose `Update 날짜` is today, then a **fixed 5-line** breakdown (`n. 이유 c건`; 6th+ tags collapse into `…외 N건`). |
| button | **배드리뷰** → the product's `1-3점` sheet |

`N` (총 N건) = number of `1-3점` rows whose `Update 날짜` resolves to today (KST).
Every list is padded to a fixed line count so the Pixel 11 and Z8 cards render the
same height.

## Data source

Reads the `1-3점` tab (`A:S`) of each spreadsheet via the **Sheets API v4**, authorised
with the local gws_shim OAuth token at `~/.config/gws_shim/token.json` (dukso123
account, `drive` scope; the script refreshes the token and writes it back).

| Product | Spreadsheet ID |
|---------|----------------|
| Pixel 11 Series | `12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI` |
| Galaxy Z8 Series | `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4` |

Columns used: `인입사유(tag)`, `Update 날짜` (falls back to `Exported Date`), `대분류`
(values `휴대폰보호필름` / `휴대폰케이스`). `Update 날짜` looks like `2026. 9. 2`.

## Rooms

- **TEST room** — always the first target. `--test` sends here and nowhere else.
- **Broadcast rooms** (`ROOMS` in the script): GCX전략 x SDA / ADS1 / ADS2 / ADS3 /
  ADS5 (CP) / JP Sales / IN Sales / 모바일제품개발팀, 실장님 & GCX,
  GCX x 클리어프로텍션 개발팀, [CQ] SPIGEN 국내&해외 CS. The internal GCX team room is
  deliberately excluded.

**Per-product webhook routing.** Each room has a default `token` (the Pixel 11 card
always uses it). Nine rooms also carry a `glxz8` override token — a *separate*
incoming webhook in the *same* room — used **only for the Galaxy Z8 card**:
ADS1, ADS2, ADS5 (CP), JP Sales, IN Sales, 모바일제품개발팀, 클리어프로텍션 개발팀,
실장님 & GCX, [CQ] SPIGEN 국내&해외 CS. Only SDA and ADS3 send both cards
through the default token. `room_url(room, product_key)` picks the token.

Webhook URLs (space id + token) are inlined in the script. They are Google Chat
incoming-webhook tokens, not account credentials.

## Usage

```bash
# 1. always test first — posts BOTH cards to the TEST room only
python3 badreview_chat_report.py --test

# 2. eyeball the test messages, get a human "yes", THEN broadcast
python3 badreview_chat_report.py --broadcast --yes

# subset of rooms / one product / a past date
python3 badreview_chat_report.py --broadcast --yes --only "ADS1,JP Sales"
python3 badreview_chat_report.py --test --product glxz8
python3 badreview_chat_report.py --test --date 2026-09-02

# build only, send nothing
python3 badreview_chat_report.py --dry-run --print-data
```

`--broadcast` refuses to run without `--yes`. Each room receives **2 messages**
(Z8 then Pixel 11), one second apart — 11 rooms = 22 messages. Webhook messages
cannot be edited or deleted afterwards.

## Requirements

- Python 3.9+
- `google-auth`, `google-auth-oauthlib`, `google-api-python-client`
- A valid `~/.config/gws_shim/token.json`

## Safety flow (matches the skill)

1. Re-read the sheets every run (never reuse stale numbers).
2. `--test` → TEST room only.
3. Show the result + the room list, get an explicit human confirmation.
4. `--broadcast --yes` → all rooms.
5. If a product's `todayCount` is 0 the card still sends ("업로드된 배드리뷰 없음"); the
   script prints a warning first.
