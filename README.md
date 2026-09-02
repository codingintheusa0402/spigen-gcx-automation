# spigen-gcx-automation

Internal automation scripts for the Spigen GCX (Global Customer Experience) team — Amazon review monitoring, Seller Central scraping, MCF order tracking, daily reporting, and CS workflow tooling.

---

## Repository structure

Projects are grouped into category folders. Within each folder, projects sit flat (no further nesting) except `Apify/`, which groups the per-product Apify triggers together with the Apify Actor code they call.

```
spigen-gcx-automation/
│
├── Scrapers/                            # Python — standalone scrapers
│   ├── SC_Review_Scraper/               # Playwright scraper for Amazon Seller Central reviews
│   ├── amazon_dp_scraper/               # Amazon /dp/ product detail scraper (Playwright, async)
│   └── amazon_child_asin_scraper/       # Amazon parent→child ASIN resolver + rating/review scraper
│
├── GAS_ReviewAutomation/                # GAS — review scraping & distribution
│   ├── MasterTrigger/                   # Daily review distribution job (all products)
│   ├── Gemini_DR/                       # Galaxy S26 — Gemini-powered =DR() defect classifier
│   ├── GlxZ8_MondayToSheet/             # Galaxy Z8 — Monday board → Sheet (new items only, daily)
│   └── Apify/
│       ├── APIFY_Axesso/                # Legacy master Apify/Axesso review scrape + sheet distribution
│       ├── Glx26_Apify/                 # Galaxy S26 Apify trigger + Monday.com board sync
│       ├── GlxZ8_Apify/                 # Galaxy Z Fold8/Flip8/Fold8 Ultra Apify trigger + Monday sync
│       ├── Auto_Acc_Apify/              # Auto Accessories per-product Apify trigger
│       ├── Pixel10a_Apify/              # Pixel 10a per-product Apify trigger
│       ├── Power_Acc_Apify/             # Power Accessories per-product Apify trigger
│       ├── SDA_Apify/                   # Screen & Display Accessories per-product Apify trigger
│       ├── iPh17e_Apify/                # iPhone 17e per-product Apify trigger
│       ├── iPh17e_Monday/               # iPhone 17e Monday.com board sync (S26-pattern port)
│       ├── 유지훈P_Apify/                # 유지훈P per-product Apify trigger
│       ├── apify-amazon-dp-scraper/     # Apify Actor — Amazon /dp/ scraper (cloud-run version)
│       ├── apify-axesso-wrapper/        # Apify Actor wrapper around Axesso API
│       └── apify-axesso-wrapper-private/ # Private variant of the Axesso wrapper Actor
│
├── GAS_Operations/                      # GAS — operations & reporting
│   ├── MCF_Tracking/                    # MCF order tracking, SP-API fee/tracking lookup, daily Chat alert
│   ├── CX_Dashboard/                    # SP-API data dashboard (orders, sales, inventory, feedback)
│   ├── Bi-Weekly/                       # Bi-weekly CX report Slides auto-updater with arc charts
│   ├── SheetMirror/                     # Chunk-copy `26년 전체문의` to a read-only dashboard sheet
│   ├── TCTChatLog_GCX/                  # Lazada/Shopee Esc T2 alerts + daily close report to Google Chat
│   ├── TicketDailyReport/               # Zendesk daily ticket report with charts sent to Google Chat
│   ├── TriggerAlert/                    # Monday.com board → Google Sheets sync
│   ├── KPI_Report/                      # Populates KPI result cells on the team's KPI spreadsheet
│   ├── Monday_CX_Board/                 # Generic Monday.com board ↔ Google Sheet sync (modeless UI dialog)
│   ├── ASIN_Master_MondaySync/          # Monday.com board sync + ABM_Relay_Log retention cleanup
│   └── BadReview_ChatReport/            # Pixel 11 / Galaxy Z8 배드리뷰(1~3점) daily Chat app-card + team broadcast
│
├── GAS_Zendesk/                         # GAS — Zendesk / CS ticketing operations
│   ├── ABM_TicketMerge/                 # Merges duplicate Amazon Buyer Message tickets + inbound cleanup
│   ├── PurchaseDate_Sync/               # Syncs Zendesk's custom Purchase Date field to a Monday.com board
│   └── GCXReply_GAS/                    # GCX Reply's SP-API/Sheet-lookup backend + versioned script archive
│
├── Browser_Extensions/                  # Userscripts & extensions
│   ├── tampermonkey_scripts/            # GCX Reply, MCF Autofill (EU + JP), Invoice Automation, GChat Reply Suggest
│   ├── gcx-reply-extension/             # WIP native Chrome extension (MV3) port of GCX Reply
│   └── gchat_reply_suggest_server.py    # Local claude-CLI backend for GChat Reply Suggest.user.js
│
├── Zendesk_Themes/                      # Zendesk Guide Help Center theme exports (version control mirror)
│   ├── sq2gcx_AmazonHelpcenter/         # Amazon EU claim-form Help Center theme
│   └── spigen-eu_ShopifyHelpcenter/     # Shopify EU claim-form Help Center theme
│
└── reference/                           # Internal reference docs
```

---

## Projects

### Python scrapers

| Project | Description | README |
|---------|-------------|--------|
| [Scrapers/SC_Review_Scraper](Scrapers/SC_Review_Scraper/) | Scrapes Amazon Seller Central reviews across US/EU/JP/IN with Playwright. Parallel by top-level domain; EU countries (DE→IT→FR→ES→UK) scrape sequentially on one shared tab. Enriches reviews with reviewer image URLs. | [README](Scrapers/SC_Review_Scraper/README.md) |
| [Scrapers/amazon_dp_scraper](Scrapers/amazon_dp_scraper/) | Async Playwright scraper for Amazon `/dp/` pages — rating, review count, title, spec table. Up to 8 domains simultaneously, dual-sheet Excel output (English + local-language). | [README](Scrapers/amazon_dp_scraper/README.md) |
| [Scrapers/amazon_child_asin_scraper](Scrapers/amazon_child_asin_scraper/) | Selenium scraper that resolves parent ASINs into child variants and extracts per-child rating/review counts. Detects shared variation review pools. | [README](Scrapers/amazon_child_asin_scraper/README.md) |

### Google Apps Script — Review automation

| Project | Product | Description | README |
|---------|---------|-------------|--------|
| [GAS_ReviewAutomation/MasterTrigger](GAS_ReviewAutomation/MasterTrigger/) | All | Daily job that reads the `"finalize"` filter view from each product's source sheet and distributes new reviews into destination spreadsheets. Handles dedup, `=dr()` formula injection, and `tem` sheet refresh. | [README](GAS_ReviewAutomation/MasterTrigger/README.md) |
| [GAS_ReviewAutomation/Gemini_DR](GAS_ReviewAutomation/Gemini_DR/) | Galaxy S26 | Gemini-powered `=DR()` custom Sheets formula that classifies review text into a defect/issue label, bound to the Galaxy S26 review spreadsheet. | [README](GAS_ReviewAutomation/Gemini_DR/README.md) |
| [GAS_ReviewAutomation/GlxZ8_MondayToSheet](GAS_ReviewAutomation/GlxZ8_MondayToSheet/) | Galaxy Z8 | Daily (17:00 KST) append-only sync — pulls new items from the Galaxy Z8 Case+CP Monday board into the sheet, never overwrites existing rows. | [README](GAS_ReviewAutomation/GlxZ8_MondayToSheet/README.md) |
| [GAS_ReviewAutomation/Apify/APIFY_Axesso](GAS_ReviewAutomation/Apify/APIFY_Axesso/) | All | Legacy copy of MasterTrigger's `dailyJob()` logic, plus its own Apify run lifecycle (`Apify.js`) and dedup helper (`Sheet_Automation.js`). Kept for reference — MasterTrigger is canonical. | [README](GAS_ReviewAutomation/Apify/APIFY_Axesso/README.md) |
| [GAS_ReviewAutomation/Apify/Glx26_Apify](GAS_ReviewAutomation/Apify/Glx26_Apify/) | Galaxy S26 | Per-product Apify trigger + Monday.com board sync for Galaxy S26 review sheet. | [README](GAS_ReviewAutomation/Apify/Glx26_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/GlxZ8_Apify](GAS_ReviewAutomation/Apify/GlxZ8_Apify/) | Galaxy Z8 | Per-product Apify trigger + Monday.com board sync (board 18421346787, 📌Galaxy Z8 Case+CP) for the Galaxy Z Fold 8 / Flip 8 / Fold 8 Ultra review sheet. Copy of the Glx26 project with Z8 sheet/board/group config. | [README](GAS_ReviewAutomation/Apify/GlxZ8_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/Auto_Acc_Apify](GAS_ReviewAutomation/Apify/Auto_Acc_Apify/) | Auto Accessories | Per-product Apify trigger for the Auto Accessories review sheet. | [README](GAS_ReviewAutomation/Apify/Auto_Acc_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/Pixel10a_Apify](GAS_ReviewAutomation/Apify/Pixel10a_Apify/) | Pixel 10a | Per-product Apify trigger for Pixel 10a review sheet. | [README](GAS_ReviewAutomation/Apify/Pixel10a_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/iPh17e_Apify](GAS_ReviewAutomation/Apify/iPh17e_Apify/) | iPhone 17e | Per-product Apify trigger for iPhone 17e review sheet. | [README](GAS_ReviewAutomation/Apify/iPh17e_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/iPh17e_Monday](GAS_ReviewAutomation/Apify/iPh17e_Monday/) | iPhone 17e | Monday.com board sync for iPhone 17e (same pattern as Glx26_Apify's board sync). | [README](GAS_ReviewAutomation/Apify/iPh17e_Monday/README.md) |
| [GAS_ReviewAutomation/Apify/SDA_Apify](GAS_ReviewAutomation/Apify/SDA_Apify/) | SDA | Per-product Apify trigger for Screen & Display Accessories review sheet. | [README](GAS_ReviewAutomation/Apify/SDA_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/Power_Acc_Apify](GAS_ReviewAutomation/Apify/Power_Acc_Apify/) | Power Acc. | Per-product Apify trigger for Power Accessories review sheet. | [README](GAS_ReviewAutomation/Apify/Power_Acc_Apify/README.md) |
| [GAS_ReviewAutomation/Apify/유지훈P_Apify](GAS_ReviewAutomation/Apify/유지훈P_Apify/) | 유지훈P | Per-product Apify trigger for 유지훈P review sheet. | [README](GAS_ReviewAutomation/Apify/유지훈P_Apify/README.md) |

### Google Apps Script — Operations & reporting

| Project | Description | README |
|---------|-------------|--------|
| [GAS_Operations/MCF_Tracking](GAS_Operations/MCF_Tracking/) | Multi-Channel Fulfillment order tracking. SP-API custom formulas (`=AMZTK()`, `=MCFFee()`), backfill functions, `onEdit` automation, and daily Google Chat alert for orders missing tracking numbers. | [README](GAS_Operations/MCF_Tracking/README.md) |
| [GAS_Operations/CX_Dashboard](GAS_Operations/CX_Dashboard/) | SP-API data dashboard — refreshes Marketplaces, Orders, Sales Metrics, Customer Feedback, and FBA Inventory into dedicated sheets via menu actions or custom formulas. | [README](GAS_Operations/CX_Dashboard/README.md) |
| [GAS_Operations/Bi-Weekly](GAS_Operations/Bi-Weekly/) | Auto-populates a bi-weekly CX report Google Slides deck with live data — text placeholder substitution and half-donut arc chart image insertion for defect/model breakdowns. | [README](GAS_Operations/Bi-Weekly/README.md) |
| [GAS_Operations/SheetMirror](GAS_Operations/SheetMirror/) | Copies `26년 전체문의` to a read-only dashboard spreadsheet in 1,000-row chunks. | [README](GAS_Operations/SheetMirror/README.md) |
| [GAS_Operations/TCTChatLog_GCX](GAS_Operations/TCTChatLog_GCX/) | Lazada/Shopee escalation alerts — sends Google Chat cards when a row status changes to `Esc T2`, plus a daily close-report card. | [README](GAS_Operations/TCTChatLog_GCX/README.md) |
| [GAS_Operations/TicketDailyReport](GAS_Operations/TicketDailyReport/) | Fetches Zendesk ticket views, updates graph sheets, and sends daily chart images to Google Chat via the `hcti.io` image API. | [README](GAS_Operations/TicketDailyReport/README.md) |
| [GAS_Operations/TriggerAlert](GAS_Operations/TriggerAlert/) | Syncs a Monday.com board into a Google Sheet via the Monday API, with a live-log sidebar UI. | [README](GAS_Operations/TriggerAlert/README.md) |
| [GAS_Operations/KPI_Report](GAS_Operations/KPI_Report/) | Populates KPI result cells on the team's KPI tracking spreadsheet. | [README](GAS_Operations/KPI_Report/README.md) |
| [GAS_Operations/Monday_CX_Board](GAS_Operations/Monday_CX_Board/) | Generic Monday.com board ↔ Google Sheet sync, with a modeless dialog UI (live log, Monday branding). | [README](GAS_Operations/Monday_CX_Board/README.md) |
| [GAS_Operations/ASIN_Master_MondaySync](GAS_Operations/ASIN_Master_MondaySync/) | Same Monday.com board ↔ Sheet sync engine as Monday_CX_Board, plus an independent daily cleanup of the `ABM_Relay_Log` tab (prunes rows older than 15 days) written by GCXReply_GAS. | [README](GAS_Operations/ASIN_Master_MondaySync/README.md) |
| [GAS_Operations/BadReview_ChatReport](GAS_Operations/BadReview_ChatReport/) | Standalone Python: builds the Pixel 11 / Galaxy Z8 배드리뷰(1~3점) Google Chat app-card from each `1-3점` sheet (today's count + Top 5 인입사유 by 대분류) and fans it out to the GCX cross-team rooms — `--test` room first, then `--broadcast --yes`. Reads Sheets via the gws_shim OAuth token. Twin of the `*-badreview-chat-report` Claude skills. | [README](GAS_Operations/BadReview_ChatReport/README.md) |

### Google Apps Script — Zendesk / CS ticketing operations

| Project | Description | README |
|---------|-------------|--------|
| [GAS_Zendesk/ABM_TicketMerge](GAS_Zendesk/ABM_TicketMerge/) | Merges duplicate Zendesk tickets created from consecutive Amazon Buyer Messages by the same buyer into one thread (Zendesk creates one ticket per ABM email; this mirrors Seller Central's own case threading). Also cleans up the raw marketing-template HTML Zendesk creates from each inbound ABM email into a readable message. | [README](GAS_Zendesk/ABM_TicketMerge/README.md) |
| [GAS_Zendesk/PurchaseDate_Sync](GAS_Zendesk/PurchaseDate_Sync/) | Syncs a Zendesk ticket's custom Purchase Date field to the matching item's date column on Monday.com board `18421346787` (native Zendesk↔Monday integration can't map custom fields). | [README](GAS_Zendesk/PurchaseDate_Sync/README.md) |
| [GAS_Zendesk/GCXReply_GAS](GAS_Zendesk/GCXReply_GAS/) | Backend for the GCX Reply Tampermonkey script below — SP-API order lookups (SigV4-signed) and Google Sheet product lookups via a GAS web app. Also holds a versioned reference-copy archive (`v*.gs`) of every past GCX Reply script version. | [README](GAS_Zendesk/GCXReply_GAS/README.md) |

### Browser extensions & userscripts

| Project | Description | README |
|---------|-------------|--------|
| [Browser_Extensions/tampermonkey_scripts](Browser_Extensions/tampermonkey_scripts/) | **GCX Reply** (`v3.5.2`) — Zendesk order/product lookup panel, Auto-Fill, MCF handoff, ABM auto-relay + NRN. **Amazon MCF Autofill** (`v1.4.3`) — EU Seller Central MCF order-page autofill. **Amazon JP MCF Autofill** (`v1.5.2`) — JP variant. **Amazon Invoice Automation** (`v1.5`) — Amazon.de invoice download. **GChat Reply Suggest** (`v3.5.1`) — Alt+G shows a T3 Esc (deterministic, no-AI ticket-forward — bold + real hyperlink, ↑/↓ ticket browsing, @mention + honorific pickers, sourced from recently-visited Zendesk tickets) / Gratitude / Reminder picker in every Google Chat room by default (incl. the Chrome-PWA desktop app); Gratitude/Reminder auto-read the mention + honorific from the T3 Esc message a thread is attached to. Only in designated rooms (matched by space ID) does it instead suggest 3 AI-generated reply sentences, backed by a local server (`Browser_Extensions/gchat_reply_suggest_server.py`) that calls the `claude` CLI directly. Install `.user.js` files via Tampermonkey Dashboard → Import. | [README](Browser_Extensions/tampermonkey_scripts/README.md) |

### Zendesk Guide theme exports

| Project | Description | README |
|---------|-------------|--------|
| [Zendesk_Themes/sq2gcx_AmazonHelpcenter](Zendesk_Themes/sq2gcx_AmazonHelpcenter/) | Help Center theme (20 templates + `script.js` + `style.css`) shown after a customer submits the Amazon EU claim form. Redirects to Amazon Store pages (DE/UK/FR/IT/ES/IN/JP). | [README](Zendesk_Themes/README.md) |
| [Zendesk_Themes/spigen-eu_ShopifyHelpcenter](Zendesk_Themes/spigen-eu_ShopifyHelpcenter/) | Help Center theme shown after a customer submits the Spigen EU (Shopify-run) claim form. Redirects to Spigen's own Shopify storefronts (DE/UK/FR/IT/ES); Community templates are empty (feature not enabled on this brand). | [README](Zendesk_Themes/README.md) |
| [Browser_Extensions/gcx-reply-extension](Browser_Extensions/gcx-reply-extension/) | Native Chrome extension (MV3) port of GCX Reply — in-progress migration off Tampermonkey; not yet feature-complete (MCF-page autofill code is currently dead — not wired into `content_scripts.matches`). Dev/testing only, load unpacked. | [README](Browser_Extensions/gcx-reply-extension/README.md) |

---

## Quick start

### Python scrapers

```bash
pip install playwright openpyxl pynput selenium
playwright install chromium
```

### Google Apps Script (clasp)

```bash
npm install -g @google/clasp
clasp login

# Push any project
cd ~/Desktop/GCX/<CategoryFolder>/<ProjectFolder>
clasp push --force
```

Each GAS project has its own `.clasp.json` (gitignored) pointing to the correct GAS script ID. See each project's README for the script ID, linked spreadsheet, and exact `cd` path.

### Script Properties (all GAS projects that call external APIs)

Set in **Extensions → Apps Script → Project Settings → Script Properties**:

| Key | Used by |
|-----|---------|
| `APIFY_TOKEN` | All per-product Apify triggers, APIFY_Axesso |
| `LWA_CLIENT_ID` / `LWA_CLIENT_SECRET` / `LWA_REFRESH_TOKEN` | MCF_Tracking, CX_Dashboard |
| `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` | MCF_Tracking, CX_Dashboard |
| `MONDAY_API_KEY` | TriggerAlert, Apify/Glx26_Apify, Monday_CX_Board, ASIN_Master_MondaySync |

---

## Branching & commit conventions

| Branch | Use |
|--------|-----|
| `main` | Stable, production-ready |
| `feat/<desc>` | New features |
| `fix/<desc>` | Bug fixes |

Commit message format: `<type>(<project>): <description>`

Examples:
```
feat(sc-scraper): add EU single-country re-run support
fix(master-trigger): guard against missing destRidIdx
feat(mcf-tracking): add MCFFee_JP formula
```
