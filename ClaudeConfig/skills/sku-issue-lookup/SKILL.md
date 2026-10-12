---
name: sku-issue-lookup
description: SKU issue lookup — when another team (아마존직판팀, 영업, CQ …) asks "how many bad reviews / claims came in for SKU X about issue Y, and what's the Amazon rating?", find every Amazon review (Caspi SQ.TAHOE.AMAZON_REVIEWS) and every Zendesk ticket (Caspi S3.ZENDESK.TICKETS) for those SKUs, tag each row's issue, and write them as ONE tab in the review source spreadsheet in the SKU_ACS11498 layout (+ Type = Review/Claim), then answer with counts + a Korean reply draft. Trigger on "search reviews and zendesk claims on caspi for SKU …", "이 SKU 배드리뷰/클레임 몇 건인지 확인해줘", "find all issues for ACS…", "SKU 이슈 조회", a pasted request listing SKUs + an issue (미끄러움, 황변, 유격 …), or any close paraphrase.
---

# SKU Issue Lookup (Caspi reviews + Zendesk claims → sheet tab)

Built and approved 2026-10-12 (request: ACS11909 / ACS11923 "제품 사용 시 미끄러움",
result tab `SKU_ACS11909_ACS11923`, gid 1729059594). Follow it in this order.

## 0. Inputs
From the request take: **SKUs** (8-char base, e.g. `ACS11909`), the **issue** (e.g. 미끄러움 / slippery),
and the destination. Default destination = review source spreadsheet
`1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`, new tab `SKU_<SKU1>_<SKU2>…`. Only ask if SKUs are missing.

Sheets access: local `~/.config/gws_shim/token.json` (kjw@spigen.com) — see memory `gws_shim_sheets_token`.
`write_tab.py` in this folder handles auth.

## 1. SKU → ASIN
Product master `1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo`, tab `Data`
(cols `ASIN`, `SKU`, `자재내역`, `기종명`, `색상명`, `Amazon Rating (US)`, `Amazon Ratings Count (US)`, `Rating 기준일`).
Keep only `B0…` ASINs (numeric ones are non-Amazon channels). Note the US rating/count + 기준일 for the answer.

## 2. Reviews (Caspi `run_query`)
First a coverage query, grouped by `MAPPED_SKU, CHILD_ASIN, PARENT_ASIN, MARKETPLACE` with counts / ≤3★ counts,
`WHERE CHILD_ASIN IN (<asins>) OR LEFT(MAPPED_SKU,8) IN (<skus>)`. Watch out:
- **A parent listing holds every color.** Filter on CHILD_ASIN/MAPPED_SKU, never on PARENT_ASIN, or sibling colors leak in.
  `ASIN_WEIGHTED_STAR_RATING` is per listing (all colors), so say that when you report the rating.
- **US may be empty.** In 2026-10 Caspi had no US reviews for these child ASINs. Cross-check the master `SC` tab
  (col A = ASIN, col K = Review ID) of the same spreadsheet and add any Review IDs Caspi lacks. If US is still empty, say so
  and quote the product-master US rating instead.

Then the full pull (add the padding column so the result lands in a file; parse it with python, drop `PADDING`; null
columns are omitted from row dicts, so always use `.get`):
```sql
SELECT r.MAPPED_SKU, r.CHILD_ASIN, TO_VARCHAR(r.REVIEW_CREATED_AT,'YYYY-MM-DD') D, IFF(r.HAS_IMAGES,'Y','N') PH,
 r.REVIEW_AUTHOR_PUBLIC_NAME AU, r.REVIEW_RATING RT, r.REVIEW_TITLE TI, r.REVIEW_TEXT TX,
 IFF(r.MARKETPLACE='GB','UK',r.MARKETPLACE) MK, r.REVIEW_DETAIL_URL URL, COALESCE(p.U, r.PRIMARY_PHOTO_URL) IMG,
 r.REVIEW_ID RID, r.ORDER_ID OID, r.ASIN_WEIGHTED_STAR_RATING PR, r.ASIN_TOTAL_REVIEW_COUNT RC, REPEAT('X',3000) PADDING
FROM SQ.TAHOE.AMAZON_REVIEWS r
LEFT JOIN (SELECT REVIEW_ID, LISTAGG(PHOTO_URL,'|') U FROM SQ.TAHOE.AMAZON_REVIEW_PHOTOS GROUP BY 1) p ON p.REVIEW_ID=r.REVIEW_ID
WHERE r.CHILD_ASIN IN (<asins>) OR LEFT(r.MAPPED_SKU,8) IN (<skus>)
ORDER BY r.MAPPED_SKU, r.REVIEW_CREATED_AT DESC
```
If the pull is large (>~300 rows), slice it with `MOD(ABS(HASH(REVIEW_ID)),N)` as in memory `caspilm_review_fetch_workflow`.

**Read every review** (all ratings, all languages) and write a short English **Issue Summary**, adding a Korean tag
for the requested issue, e.g. `Slippery (미끄러움)`, `Extremely Slippery / Bulky / Fingerprints (미끄러움)`,
`No Issue`, `No Issue (states not slippery)`, `Kickstand Loose (스탠드 유격)`. Count separately:
requested-issue reviews rated 1–3★, positive reviews that mention the issue only lightly, and other 1–3★ issues.

## 3. Zendesk tickets (Caspi)
Find tickets by ★문의SKU (900008676703) or ASIN (360021934312) custom field. Flatten in a CTE (Snowflake
lateral-join quirk; see memory `caspi_zendesk_claims_fetch`):
```sql
WITH f AS (SELECT t."id" ID, t."created_at" C, t."subject" SUBJ, t."tags" TAGS,
  cf.value:field_id::string FID, cf.value:value::string VAL
  FROM S3.ZENDESK.TICKETS t, LATERAL FLATTEN(input => t."custom_fields") cf
  WHERE cf.value:field_id::string IN ('900008676703','360021934312'))
SELECT ID, MIN(C) C, MAX(SUBJ) SUBJ, MAX(IFF(FID='900008676703',VAL,NULL)) SKU, MAX(IFF(FID='360021934312',VAL,NULL)) ASIN,
 MAX(IFF(ARRAY_CONTAINS('4._issue'::variant,TAGS),1,0)) IS_CLAIM
FROM f WHERE UPPER(VAL) LIKE '%<SKU>%' OR UPPER(VAL) LIKE '%<ASIN>%'   -- one pair per SKU
GROUP BY ID ORDER BY C
```
Also run a **free-text sweep** so you catch tickets with empty SKU/ASIN fields: `"subject"/"description" ILIKE '%<product name>%'`
plus the issue keywords in EN/DE/FR/IT/ES/JP (slippery → `slip, grip, rutsch, scivol, glis, 滑`), `created_at >=` launch.
For the hits, fetch `description`, `status`, and the decoded fields Country 4513936822297, Order ID 360021934132,
1차 Defect 360022182831, 사진유무 26936618247577 (decode the values with the `TICKET_FIELDS` options CTE).

**Claimer names:** Caspi masks `S3.ZENDESK.USERS."name"`. Get `"requester_id"` from TICKETS, then in Chrome
(claude-in-chrome, logged-in Zendesk) open
`https://spigenhelp.zendesk.com/api/v2/users/show_many.json?ids=<id,id,…>`, read the names with get_page_text, and close the tab.

## 4. Write the tab — `write_tab.py`
Build `rows.json` = a list of 16-value rows in this exact header order (the `SKU_ACS11498` layout + Type):
`ASIN | Created 날짜 | 사진 유무 | Reviewer | Review Ratings | Review Title | 본문 | Marketplace | Link | Image URL | Review ID | Order ID | Product Rating | Ratings Count | Issue Summary | Type`
- Reviews: `Type`=`Review`, `Link` = Amazon review URL, `본문` with `<br />` → newline. Sort by SKU, then date descending.
- Zendesk rows go below the reviews **in the same table**, with no blank or label row: `Type`=`Claim`, `Reviewer` = claimer's name,
  `Review Title` = subject, `본문` = a short English gist of the customer's message, `Link` = `https://spigenhelp.zendesk.com/agent/tickets/<id>`,
  `Review ID` = ticket #, `Review Ratings`/`Product Rating`/`Ratings Count` blank,
  `Issue Summary` = 1차 Defect value (e.g. `(Case)_황변`). Add ` — non-defect` for tickets that aren't 4._issue (Delivery, Invoice …).
```
python3 ~/.claude/skills/sku-issue-lookup/write_tab.py --title SKU_ACS11909_ACS11923 --rows rows.json --highlight "미끄러움,Slipper"
```
The script clones `SKU_ACS11498` (gid 1031735538) for its formatting, writes the header + rows, and highlights
the rows whose Issue Summary matches `--highlight` in yellow. It refuses to overwrite an existing tab unless you pass `--replace`. It prints the gid.

## 5. Answer
Give the tab link (`…/edit#gid=<gid>`) and a per-SKU table: total reviews / 1–3★, **requested-issue bad reviews**
(marketplace, ★, date), light mentions, **requested-issue claims**, other claims, Amazon rating per marketplace.
State the caveats (no US data, rating covers all colors). End with a Korean reply draft for the requester:
```
요청하신 N개 SKU 확인 결과 공유드립니다.
- Amazon 평점: …
- '<issue>' 배드리뷰(1~3점): SKU별 건수 (국가 점수) → 총 n건
- Zendesk 클레임: <issue> 관련 n건 (기타 클레임 …)
- 상세: [시트 링크]
```
Don't send it anywhere. The user forwards it themselves.
