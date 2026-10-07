# GAS_Zendesk

Google Apps Script projects that run behind the `spigenhelp.zendesk.com` agent
workflow. Each folder is its own clasp project with its own README.

## Screenshots

![ABM Ticket Merge trigger](ABM_TicketMerge/docs/zendesk_trigger.jpg)
*ABM Ticket Merge trigger*

| Project | What it does |
|---------|--------------|
| [ABM_TicketMerge](ABM_TicketMerge/) | Zendesk-trigger webhook. Merges duplicate Amazon Buyer Message tickets into one thread per buyer case, posts clean copies of raw Amazon emails, auto-fills order fields, and runs a 30-min self-healing sweep. Deployment @40. |
| [GCXReply_GAS](GCXReply_GAS/) | Web-app backend for the GCX Reply userscript: SP-API order lookup, product index, AI 인입사유, and the `ABM_Relay_Log` relay endpoints. Also holds the `v*.gs` archive of every userscript version. Code.js v2.7.1. |
| [PurchaseDate_Sync](PurchaseDate_Sync/) | Twice-daily (9:00 / 21:00 KST) batch that copies Zendesk's Purchase Date field onto the Z8 / Pixel 11 / iPhone 18 Case+CP monday boards. |

**Deploying.** ABM_TicketMerge and GCXReply_GAS are called through **pinned**
web-app deployments, so a change goes live only after `clasp push --force`
**and** `clasp deploy -i <deploymentId>`. PurchaseDate_Sync runs from a time
trigger, so `clasp push` alone is enough.
