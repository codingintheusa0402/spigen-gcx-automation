---
name: product-recommender
description: Substitute-product recommendation for a Zendesk ticket — paste a spigenhelp.zendesk.com ticket link and get same-device, same-kind Spigen products that are in FBA stock AND actively listed in the customer's Amazon marketplace, as a copy-paste table (ASIN linked to that country's Amazon page | Device | Product | SKU), also placed on the clipboard as a rich table for pasting into Zendesk. Trigger on "recommend a substitute for this ticket", "대체 상품 추천해줘", "이 티켓 대체품 찾아줘", "alternative product for <ticket link>", "재고 있는 대체 제품", a pasted ticket link + "추천", or any close paraphrase. Also works without a ticket (SKU + marketplace).
---

# product-recommender

Same matcher as the Ticket Reporter Chat app's `/products-similar` (it runs
`GAS_Operations/TicketReporterCard/ProductSearch.gs` under Node), plus a per-country stock/listing filter.

## Run

```bash
python3 ~/.claude/skills/product-recommender/recommend.py <ticket link or number>
# overrides / no ticket:
python3 ~/.claude/skills/product-recommender/recommend.py 1000166420 --sku ACS10066 --market DE
python3 ~/.claude/skills/product-recommender/recommend.py --sku ACS10066 --market GB --all --n 15
```

Options: `--sku` / `--asin` (product to replace), `--market` (DE GB/UK FR IT ES NL SE PL BE IE TR AE SA JP; US/CA/MX…
give candidates only), `--all` (every category for the device, default = same kind as the product — case → cases,
SP → screen/camera protectors), `--n` rows (10), `--no-copy`, `--json`.

## What it does
1. **Ticket** (read in the logged-in Zendesk Chrome, CDP :9222 — no Zendesk API token exists on these Macs):
   product = ★문의SKU field (900008676703) → ASIN field (360021934312) → SKU/ASIN in subject/description;
   marketplace = Point of Purchase (20016270875033) → Country (4513936822297). If either is missing it stops and
   says so — ask the user, then rerun with `--sku` / `--market`.
2. **Candidates**: same device (multi-device 기종 and sizes handled), same kind, from the `ProductIndex` tab.
3. **Filter**: FBA sellable stock > 0 in that country (`InventorySnapshot`, hourly) **and** an Active listing there
   (`ListingsSnapshot`, Seller Central all-listings, refreshed daily); ASIN = the one that country's listing uses.
   Most stock first. All three tabs live in the TicketQueue spreadsheet `1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M`,
   written by the Chat app's `refreshInventorySnapshot` / `refreshListingsSnapshot` / `buildProductIndex_`.
4. **Output**: header (ticket · marketplace · base product · device · kind · counts), the markdown table, stock per
   row, and on macOS the same table on the clipboard as HTML (paste straight into the Zendesk reply).

## Reply to the user
Show the header + the table exactly as printed (it's meant to be copied). Mention if stock/listing data is older
than a day (footer times in the Chat card / snapshot `ts`), and for US/CA/MX customers say stock couldn't be
checked (the MCF SP-API credentials only reach EU + JP). Never send anything to the customer or post on the ticket —
this skill only recommends; the agent decides and pastes.

## Gotchas
- `401 Couldn't authenticate` → the :9222 Chrome lost its Zendesk login. Never log in yourself (hard rule) — ask
  the user to sign in via Screen Sharing.
- `node` must be on PATH (brew node). The skill reads ProductSearch.gs from the repo, so a Chat-app matcher change
  applies here too — keep that file free of Apps Script services.
- No rows = nothing of that kind for the device is in stock + active in that country; suggest `--all`.
