---
name: glxz8-badreview-chat-report
description: >-
  Build and send the "Galaxy Z8 Series 배드리뷰 (1~3점) 먼데이보드 업로드 완료" report to a
  Google Chat room as a v2 app card (cardsV2). Reads the Galaxy Z Fold8/Flip8/Fold8
  Ultra review spreadsheet's '1-3점' sheet only: today's upload count + 인입사유(tag)
  tally, plus Top 5 인입사유(tag) split into two columns by 대분류 (휴대폰보호필름 /
  휴대폰케이스). Trigger when the user asks to "run GalaxyZ8 BadReview Google Chat
  Report", "send Z8 bad review report to Chat", "Z8 배드리뷰 챗 리포트 보내줘", "갤럭시Z8
  배드리뷰 구글챗 리포트", "GlxZ8 배드리뷰 리포트", or any close paraphrase. Sibling of
  pixel11-badreview-chat-report (identical card, different spreadsheet).
metadata:
  category: automation
  locale: ko-KR
---

# glxz8-badreview-chat-report

Sends a Google Chat **cardsV2** app card announcing that today's Galaxy Z8 Series
(Z Fold8 / Flip8 / Fold8 Ultra) bad reviews (★1~3) were uploaded to the monday board.
Same card layout as `pixel11-badreview-chat-report`, just a different sheet.

- Spreadsheet: `19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4`
  ("Galaxy Z Fold 8 / Flip 8 / Fold 8 Ultra Series_Customer Reviews (★1~5)")
- `1-3점` sheet — `gid 970309432` — **the only sheet read; everything comes from here**
  (headers are identical to the Pixel 11 sheet)
- Default Google Chat webhook — GCX team room (space `AAQAc9NQmJQ`), in `report.py`;
  override with `--webhook`
- Helper: `report.py` in this skill folder (builds the card + POSTs it)

The card:

| Part | Content |
|------|---------|
| header.title | `✔️ M/D(요일) Galaxy Z8 Series 배드리뷰 (1~3점) (총 N건)` |
| header.subtitle | `고객 리뷰 ★1~3점 · 26/7/27~26/10/27` |
| Section "Top 5 인입사유(누적)" | a **2-column** widget. Left = Top 5 `인입사유(tag)` where `대분류` == `휴대폰보호필름`; right = Top 5 where `대분류` == `휴대폰케이스`. Each column headed `<b><font color="#EA4335">휴대폰보호필름</font></b> · {tot}건` / `<b><font color="#4285F4">휴대폰케이스</font></b> · {tot}건` (colored to stand out from the rows below; no emoji), then 5 `decoratedText` rows — `topLabel`=`n위`, `text`=`<b>이유</b>`, `bottomLabel`=`c건 · p%` — so rank / 인입사유 / 건수·% each sit at a fixed left edge and the two columns line up. `p = c / that 대분류's total`. Counted over the **whole `1-3점` sheet**, not just today. |
| Section "오늘 M/D(요일) 최다 인입사유" | biggest `인입사유(tag)` among today's `1-3점` rows, then a **fixed 5-line** ranked breakdown (`n. 이유 c건`; blank `&nbsp;` lines pad a short day; 6th+ tags collapse into `…외 N건` on line 5), then a **배드리뷰** button to the `1-3점` sheet |

`M/D(요일)` = today's date, KST, Korean weekday (월/화/…).
`N` (총 N건) = count of `1-3점` rows whose `Update 날짜` column resolves to today.
`대분류` seen in the data: `휴대폰보호필름`, `휴대폰케이스` (exact strings).
**`긍정 리뷰` 인입사유(tag) is excluded** from every count (총 N건, 오늘 breakdown, cumulative Top 5) and never shown on the card — user rule 2026-09-08, permanent unless the user revokes it.

There is **no chart**.

## Steps

### 1. Load browser tools
If the `mcp__claude-in-chrome__*` tools are deferred, load them in one call:
`ToolSearch("select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__tabs_close_mcp")`

### 2. Open the spreadsheet (needs the user's Google session)
`tabs_context_mcp({createIfEmpty:true})`, then `navigate` that tab to
`https://docs.google.com/spreadsheets/d/19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4/edit`

### 3. Fetch + crunch the numbers in the page (authenticated fetch, small JSON out)
Run this with `javascript_tool` on that tab. Use **top-level await, no IIFE wrapper**
(an `(async()=>{})()` wrapper makes the tool return `{}`). It resolves the `1-3점`
gid by tab name (falls back to the hard-coded gid), pulls it as CSV via `gviz`, and
returns only the compact summary — never dump the full CSV through the tool.

```js
const ID = "19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4";
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
  if (m) dayCounts[`${m[1]}-${m[2]}-${m[3]}`] = (dayCounts[`${m[1]}-${m[2]}-${m[3]}`]||0)+1;
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
python3 ~/.claude/skills/glxz8-badreview-chat-report/report.py --data '<json from step 3>'
```

Add `--dry-run` first to show the user the payload before it goes out.
`report.py` computes `M/D(요일)` itself (KST/local date); use `--date YYYY-MM-DD` to override.
`--webhook '<url>'` sends to a different Google Chat room instead of the default GCX room.

### 5. Clean up
Close the tab you created (`tabs_close_mcp`). Report the sent message name and the
key numbers (총 N건, today's top 인입사유).

## Notes / gotchas

- **Both cards are held to the same height on purpose:** every list (the 2 Top-5 columns and the 오늘 breakdown) is rendered at a fixed 5 lines, padded with blank `&nbsp;` lines. Keep this when editing either `report.py`, and keep the two in sync. Residual: a very long reason name can still wrap inside the narrow Top-5 column and add a line to one card.

- **New message each run.** Webhook messages can't be edited or deleted from here.
- The **long header title truncates with "…"** in some Chat clients — expected.
- `Update 날짜` values look like `2026. 8. 19` (Korean, spaced). The regex in step 3
  handles any `YYYY?M?D` separator. Fallback column name: `Exported Date`.
- `대분류` column: only `휴대폰보호필름` and `휴대폰케이스` seen. A new value is silently
  dropped from the Top 5 columns — add a column / adjust `cat` + `report.py` if needed.
- The Top 5 columns count the **whole `1-3점` sheet** (all dates). Percentages are of
  each 대분류's own subtotal.
- Known webhooks: default GCX room (space `AAQAc9NQmJQ`); user's **test room** = same
  space `AAQAc9NQmJQ` token `Nvngg3UoVU-M7TqqlC48NxP-SXRzXj9zWrIoqd4BJdo`; another room
  space `AAQAb-u6r7s` token `AOJntA_PdElbBaGzQaCQhhr0aBvPAy1k3ImqQK0V9_E`.
- If `todayCount` is 0, the card still sends with an "업로드된 배드리뷰 없음" note.
- Sibling skill: `pixel11-badreview-chat-report` (keep the two `report.py` files in
  sync when the card layout changes).
