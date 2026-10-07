# GAS_Operations

GCX 운영 자동화 프로젝트 모음 — Google Apps Script(clasp) 프로젝트와 서버/launchd로 도는 Python 잡. 각 폴더의 README에 트리거, Script Properties, 배포 방법이 정리되어 있습니다.

## Screenshots

![BadReview daily carousel](BadReview_ChatReport/docs/carousel.jpg)
*BadReview daily carousel*

![TicketDailyReport chart](TicketDailyReport/docs/all_graph.jpg)
*TicketDailyReport chart*

| Project | Description |
|---|---|
| [ASIN_Master_MondaySync](ASIN_Master_MondaySync/) | Daily monday board 7606389164 → ASIN_Master `Data` sheet sync (SKU lookup used by GCX Reply), plus 15-day pruning of `ABM_Relay_Log`. |
| [BadReview_ChatReport](BadReview_ChatReport/) | Weekday 10:30 KST carousel of iPhone 18 / Z8 / Pixel 11 ★1–3 bad-review stats to the GCX cross-team Chat rooms, plus an interactive Chat app (`chat_app/`, `chat_app_jane/`) with date-range filters. |
| [Bi-Weekly](Bi-Weekly/) | Deck-bound GAS + Python tooling (`tools/apple_theme`) that builds the GCX Bi-weekly Report: claim/review card slides, Overview TOP3 gauges, sales TOP-7 tables, SIREN badges, and the apple.com-style copy that gets sent. |
| [BiWeeklyViewLog](BiWeeklyViewLog/) | Tracked-link web app that embeds the Bi-weekly deck and logs who opened it, when, and for how long. |
| [CX_Dashboard](CX_Dashboard/) | Amazon SP-API (EU/FE) dashboard sheet: menu refreshes for orders/sales/feedback/inventory plus `=SP*()` custom formulas. |
| [CaspiSalesBackfill](CaspiSalesBackfill/) | Python job that fills col Z (cumulative EU+UK Amazon units since launch) on the iPhone 18 / Pixel 11 / Z8 monitoring sheets from Caspi. |
| [DiscolorationReport](DiscolorationReport/) | Mon/Fri 10:00 Chat report of new 이염/변색 Zendesk claims and Claude-classified Amazon bad reviews by SKU, with a 90-day SIREN flag. |
| [KPI_Report](KPI_Report/) | Writes the half-year KPI result text into the KPI sheet; `../generate_kpi_report.py` builds the matching .docx report. |
| [MCF_Tracking](MCF_Tracking/) | GAS bound to the MCF 발송 로그 sheet: SP-API tracking/fee formulas and backfills, GStore tracking upkeep, the GCX Reply mark-MCF web app, and a daily missing-tracking Chat report. |
| [Monday_CX_Board](Monday_CX_Board/) | Daily monday board 5669388007 → SKU_Master `Data` sheet sync, with a live-log dialog. |
| [SheetMirror](SheetMirror/) | Chunked copy of `26년 전체문의` into a dashboard sheet's `RAW` tab, with cleaned Brand/Product helper formulas. |
| [TCTChatLog_GCX](TCTChatLog_GCX/) | TCT Lazada/Shopee chat-log sheet automation: `Esc T2` alerts, weekday 마감보고 card, Listing Issue → monday sync, voucher dropdown. |
| [TicketDailyReport](TicketDailyReport/) | Weekday 9AM Zendesk T2 report (`All_Graph` chart + `K_시트` pending-ticket card) to Chat, retry/resume runner, `/report` Chat-app date picker. |
| [TicketReporterCard](TicketReporterCard/) | "Ticket Reporter" interactive Chat app: turns TCK reports into Zendesk internal notes (card or thread reply), writes back to the TCT log, handles `/revision`, `/manual`, `/btw`. |
| [TriggerAlert](TriggerAlert/) | ⚠️ Legacy duplicate of the ASIN_Master monday sync on the **same scriptId** — do not `clasp push`. |
