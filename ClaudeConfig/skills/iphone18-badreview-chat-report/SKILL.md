---
name: iphone18-badreview-chat-report
description: >-
  Build and send the "iPhone 18 Series 배드리뷰 (1~3점) 먼데이보드 업로드 완료" report to a
  Google Chat room as a v2 app card (cardsV2). Reads the iPhone 18 Series review
  spreadsheet's '1-3점' sheet only: today's upload count + 인입사유(tag) tally, plus
  Top 5 인입사유(tag) split into two columns by 대분류 (휴대폰보호필름 / 휴대폰케이스).
  Trigger when the user asks to "run iPhone18 BadReview Google Chat Report", "send
  iPhone 18 bad review report to Chat", "아이폰18 배드리뷰 챗 리포트 보내줘", "iPhone18
  배드리뷰 구글챗 리포트", or any close paraphrase. Sibling of pixel11-badreview-chat-report
  and glxz8-badreview-chat-report (identical card, different spreadsheet), added
  2026-09-21.
metadata:
  category: automation
  locale: ko-KR
---

# iphone18-badreview-chat-report

Sends a Google Chat **cardsV2** app card announcing that today's iPhone 18 Series
bad reviews (★1~3) were uploaded to the monday board. Same card layout as
`pixel11-badreview-chat-report` / `glxz8-badreview-chat-report`, just a different
sheet — and a different webhook-routing rule (see below).

- Spreadsheet: `1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU`
  ("iPhone 18 Series_Customer Reviews (★1~5) (26/09/21~26/12/21)_글로벌CX전략팀")
- `1-3점` sheet — `gid 970309432` — **the only sheet read; everything comes from here**.
  ⚠️ The URL the user first shared for this sheet had `gid=957652957` in it, which is
  actually the **`1-5점`** tab, not `1-3점` — always resolve the gid from the sheet's
  own tab name (`gid("1-3점", ...)` in the step-3 JS below), never trust a pasted gid
  at face value. (headers are identical to the Pixel 11 / Z8 sheets)
- Default Google Chat webhook — the **private test room's iPhone18-specific token**,
  hard-coded in `report.py`; override with `--webhook`.
- **Broadcast routing is different from Pixel 11**: in `badreview-chat-broadcast`,
  iPhone 18 does NOT get its own per-room webhook. It explicitly **reuses each room's
  `glxz8` override token** — same webhook Z8 already uses in that room — per the
  user's 2026-09-21 instruction. See `room_url()` / `SHARES_GLXZ8_TOKEN` in
  `../badreview-chat-broadcast/broadcast.py`.
- Helper: `report.py` in this skill folder (builds the card + POSTs it)

The card:

| Part | Content |
|------|---------|
| header.title | `✔️ M/D(요일) iPhone 18 Series 배드리뷰 (1~3점) (총 N건)` |
| header.subtitle | `고객 리뷰 ★1~3점 · 26/9/21~26/12/21` |
| Section "Top 5 인입사유(누적)" | 2-column widget, same layout/coloring rule as the sibling skills |
| Section "오늘 M/D(요일) 최다 인입사유" | same as the sibling skills, incl. the "오늘 총 업로드" row |

Same **significance highlighting** rule as the siblings (>= 1.5x its comparison point
→ bold+red in `.text`; see `report.py`'s module docstring) and the same **permanent
`긍정 리뷰` exclusion** rule.

## Steps

### 1. Load browser tools
If the `mcp__claude-in-chrome__*` tools are deferred, load them in one call:
`ToolSearch("select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__tabs_close_mcp")`

### 2. Open the spreadsheet (needs the user's Google session)
`tabs_context_mcp({createIfEmpty:true})`, then `navigate` that tab to
`https://docs.google.com/spreadsheets/d/1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU/edit`

### 3. Fetch + crunch the numbers in the page (authenticated fetch, small JSON out)
Run this with `javascript_tool` on that tab. Use **top-level await, no IIFE wrapper**
(an `(async()=>{})()` wrapper makes the tool return `{}`). It resolves the `1-3점`
gid by tab name (falls back to the hard-coded gid), pulls it as CSV via `gviz`, and
returns only the compact summary — never dump the full CSV through the tool.

```js
const ID = "1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU";
const WANT13 = "970309432";
const tabs = [...document.querySelectorAll('.docs-sheet-tab')];
async function gid(name, fb){
  const el = tabs.find(t => t.querySelector('.docs-sheet-tab-name')?.textContent === name);
  if (!el) return fb;
  el.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
  el.dispatchEvent(new MouseEvent('mouseup',   {bubbles:true}));
  el.dispatchEvent(new MouseEvent('click',     {bubbles:true}));
  await new Promise(r => setTimeout(r, 800));
  const m = location.href.match(/gid=(\d+)/);
  return m ? m[1] : fb;
}
async function csv(g){ return (await fetch(
  `https://docs.google.com/spreadsheets/d/${ID}/gviz/tq?tqx=out:csv&gid=${g}`,
  {credentials:"include"})).text(); }
function parse(s){ const R=[]; let r=[],c='',q=false;
  for (let i=0;i<s.length;i++){ const ch=s[i];
    if (q){ if(ch==='"'){ if(s[i+1]==='"'){c+='"';i++;} else q=false; } else c+=ch; }
    else { if(ch==='"')q=true; else if(ch===','){r.push(c);c='';}
           else if(ch==='\n'){r.push(c);R.push(r);r=[];c='';}
           else if(ch==='\r'){} else c+=ch; } }
  if(c!==''||r.length){r.push(c);R.push(r);} return R; }

const g13 = await gid("1-3점", WANT13);
const s13 = parse(await csv(g13));
const H = s13[0];
const iU = H.indexOf("Update 날짜"), iT = H.indexOf("인입사유(tag)"), iC = H.indexOf("대분류");

const now = new Date(Date.now() + (new Date().getTimezoneOffset()+540)*60000); // KST
const tk = [now.getFullYear(), now.getMonth()+1, now.getDate()];
const isToday = (v) => { const m=(v||'').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? (+m[1]===tk[0] && +m[2]===tk[1] && +m[3]===tk[2]) : false; };

let todayCount = 0; const tal = {}; const dayCounts = {};
const cat = { "휴대폰보호필름": {}, "휴대폰케이스": {} };
for (let i=1;i<s13.length;i++){
  const tag = (s13[i][iT]||'').trim() || '(빈칸)';
  if (tag === '긍정 리뷰') continue;   // user rule 2026-09-08: exclude from all stats/cards
  const cc  = (s13[i][iC]||'').trim();
  if (cat[cc]) cat[cc][tag] = (cat[cc][tag]||0) + 1;
  const m = (s13[i][iU]||'').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (m) dayCounts[`${m[1]}-${m[2]}-${m[3]}`] = (dayCounts[`${m[1]}-${m[2]}-${m[3]}`]||0) + 1;
  if (isToday(s13[i][iU])) { todayCount++; tal[tag] = (tal[tag]||0)+1; }
}
const todayTags = Object.entries(tal).sort((a,b)=>b[1]-a[1]);
const block = (o) => ({ tot: Object.values(o).reduce((a,b)=>a+b,0),
  top5: Object.entries(o).sort((a,b)=>b[1]-a[1]).slice(0,5) });

// trailing-7-day daily average (excl. today) — significance baseline for 오늘 총 N건
let recentSum = 0;
for (let k=1;k<=7;k++){
  const d = new Date(tk[0], tk[1]-1, tk[2]-k);
  recentSum += dayCounts[`${d.getFullYear()}-${d.getMonth()+1}-${d.getDate()}`] || 0;
}
const recentAvg = recentSum/7;

JSON.stringify({ todayCount, todayTags, recentAvg,
  film: block(cat["휴대폰보호필름"]), case: block(cat["휴대폰케이스"]) });
```

### 4. Send the card
Pass that JSON straight through (single-quote it for the shell):

```
python3 ~/.claude/skills/iphone18-badreview-chat-report/report.py --data '<json from step 3>'
```

Add `--dry-run` first to show the user the payload before it goes out.
`report.py` computes `M/D(요일)` itself (KST/local date); use `--date YYYY-MM-DD` to override.
`--webhook '<url>'` sends to a different Google Chat room instead of the default private test webhook.

### 5. Clean up
Close the tab you created (`tabs_close_mcp`). Report the sent message name and the
key numbers (총 N건, today's top 인입사유).

## Notes / gotchas

- **All three product cards are held to the same height on purpose:** every list (the
  2 Top-5 columns and the 오늘 breakdown) is rendered at a fixed 5 lines, padded with
  blank `&nbsp;` lines. Keep this when editing any of the three `report.py`, and keep
  them in sync.
- **New message each run.** Webhook messages can't be edited or deleted from here.
- `Update 날짜` values look like `2026. 9. 21` (Korean, spaced). Fallback column name: `Exported Date`.
- `대분류` column: only `휴대폰보호필름` and `휴대폰케이스` seen so far.
- Header thumbnail is a Spigen iPhone 18 Pro case product photo (fetched 2026-09-21
  from spigen.com) — swap `HEADER_IMG` in `report.py` if a better/official one shows up.
- If `todayCount` is 0, the card still sends with an "업로드된 배드리뷰 없음" note.
- Sibling skills: `pixel11-badreview-chat-report`, `glxz8-badreview-chat-report` (keep
  all three `report.py` in sync when the card layout changes).
- For the 12-room broadcast, use `badreview-chat-broadcast` (updated 2026-09-21 to
  support all 3 products) — this skill's own `report.py` is for a standalone/test send.
