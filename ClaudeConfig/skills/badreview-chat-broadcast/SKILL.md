---
name: badreview-chat-broadcast
description: >-
  Share the scraped bad-review data: send Pixel 11 + Galaxy Z8 + iPhone 18 배드리뷰
  (1~3점) as one combined swipeable carousel card for today to the GCX cross-team
  Google Chat rooms. Trigger when the user asks to "share the bad review scraped
  data", "배드리뷰 스크랩 데이터 공유해줘", "send the bad review cards to all the rooms",
  "배드리뷰 앱카드 전체 공유", "share bad reviews with the teams", or any close paraphrase.
  ALWAYS test-sends to the test room first and asks the user to confirm before
  broadcasting to the other rooms.
metadata:
  category: automation
  locale: ko-KR
---

# badreview-chat-broadcast

Sends **all three** product cards (Pixel 11 + Galaxy Z8 + iPhone 18, ★1~3 배드리뷰,
iPhone 18 added 2026-09-21) for **today** to the GCX cross-team Chat rooms. Wraps the
three per-product skills [`pixel11-badreview-chat-report`], [`glxz8-badreview-chat-report`],
[`iphone18-badreview-chat-report`] — same cards, just fanned out to many rooms with a
mandatory safety gate.

**Combined carousel format (2026-09-21, default):** when all 3 products are sent
(the normal daily case), they go out as **ONE message per room** — a swipeable
Cards v2 `carousel`, pages in order iPhone 18 → Galaxy Z8 → Pixel 11 — via each
room's default `token` webhook, instead of 3 separate messages. Layout logic lives
in `carousel.py` (same directory); see its module docstring for the hard-won
rendering facts (`decoratedText` renders blank inside a carousel page — everything
is `textParagraph` instead; images are proxied through `wsrv.nl` to force a matching
square crop since the `image` widget has no native size/crop control). A
`--product` subset (a one-off corrective resend of a single product, e.g. after a
data fix) still sends that product's own normal single card via the old
per-product routing — a 1-page "carousel" isn't meaningful, so that path is
unchanged.

**iPhone 18 routing (2026-09-21, explicit user instruction):** iPhone 18 does NOT get
its own per-room webhook — it reuses each room's existing `glxz8` override token, same
as the Z8 card. `--px-data`/`--z8-data`/`--ip18-data` feed the three products;
`--product` now takes a comma-separated subset of `pixel11,glxz8,iphone18` (default
`all`) instead of the old `both`/single-product choice list — see `broadcast.py`'s
own `--help`/docstring.

## HARD RULES (do not skip)

1. **Test first, every time.** Run `broadcast.py --test-only` → it posts both cards to
   the **test room only** (`AAQAc9NQmJQ` / token `Nvngg3UoVU-M7TqqlC48NxP-SXRzXj9zWrIoqd4BJdo`).
2. **Then STOP and ask.** Show the user what was test-sent (date, 총 N건 for each product,
   the message links) and the **explicit list of rooms** that `--all` would post to, and
   ask: *"Send to these N rooms? (yes/no)"*. Do **not** run `--all` without a clear yes
   in this conversation.
3. **Only on an explicit yes**, run `broadcast.py --all`.
4. Never broadcast stale data — always re-fetch the sheets (step 3) in the same run,
   even if numbers were computed earlier in the session.
5. If the user names a subset ("just ADS rooms"), pass `--only "ADS1,ADS2,..."`.
   To exclude instead of restrict (e.g. "except 리더들방"), pass `--exclude "리더들방"`.
6. **Before every send** (normal daily broadcast, a resend/correction, or the
   unattended `auto_broadcast.py` run — added 2026-09-18 per user request, permanent),
   check that the sheet's `인입사유(tag)` column is actually filled in for every row
   whose `Update 날짜` == today (the AI tagging agents can lag behind new rows — this
   is exactly what triggered the 2026-09-18 Z8 resend). If any are still blank, **wait
   10 minutes and recheck, up to 3 retries** (30 min total) before sending. If still
   blank after 3 retries, send anyway but tell the user which rows/how many were still
   untagged instead of staying silent about it.
7. **Z8-only KR gate** (added 2026-09-18, permanent): before sending the **Z8** card,
   check whether any of today's Z8 rows have `국가(tag)` == `KR` (Z8's single largest
   country segment — Pixel 11 has none at all). KR reviews occasionally land after
   11 AM, so **0 KR rows is a signal today's upload may still be incomplete**, not
   necessarily a real zero day. In a manual/interactive run: surface this plainly in
   the confirmation ask ("⚠️ 0 KR reviews for Z8 today — send anyway?") instead of
   silently proceeding. `auto_broadcast.py`'s unattended run cannot literally wait for
   a chat reply, so there it holds the **entire combined carousel** back (all 3
   products are one message now — Z8 can't be selectively omitted), posts an alert +
   full carousel preview to the private test room only — see `AUTO_BROADCAST.md`.

The broadcast room list lives in `broadcast.py` `ROOMS` (source of truth: memory
`gcx_team_gchat_webhooks.md`). Currently 13 rooms: GCX전략 x SDA / ADS1 / ADS2 / ADS3 /
ADS5 (CP) / JP Sales / IN Sales / 모바일제품개발팀, 실장님 & GCX, GCX x 클리어프로텍션
개발팀, [CQ] SPIGEN 국내&해외 CS, 리더들방, GCX전략 Spigen x TCK (added 2026-09-29) —
the last two share a webhook for all cards (no `glxz8` override). The internal GCX
team room (`gcx_gchat_webhook.md`) is **not** in this list.

**Per-product webhook routing** (2026-09-03): the Pixel 11 card always uses a room's
default `token`. Each of the 11 GCX rooms also has a `glxz8` override token — a separate
incoming webhook in the same room — used only for the **Galaxy Z8** card (complete
2026-09-07). 리더들방 has no override; every card goes through its single `token`.
`room_url(room, product_key)` picks it.

## Steps

### 1. Load browser tools
`ToolSearch("select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__tabs_close_mcp")`

### 2. Open a Google Sheets tab (needs the user's Google session)
`tabs_context_mcp({createIfEmpty:true})`, then `navigate` to any of the two sheets, e.g.
`https://docs.google.com/spreadsheets/d/12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI/edit`
(both sheets are `docs.google.com`, so one tab can fetch both via `gviz`).

### 3. Fetch + crunch both sheets (top-level await, no IIFE wrapper — else it returns `{}`)

```js
async function csv(ID,g){ return (await fetch(`https://docs.google.com/spreadsheets/d/${ID}/gviz/tq?tqx=out:csv&gid=${g}`,{credentials:"include"})).text(); }
function parse(s){ const R=[]; let r=[],c='',q=false;
  for (let i=0;i<s.length;i++){ const ch=s[i];
    if (q){ if(ch==='"'){ if(s[i+1]==='"'){c+='"';i++;} else q=false; } else c+=ch; }
    else { if(ch==='"')q=true; else if(ch===','){r.push(c);c='';}
           else if(ch==='\n'){r.push(c);R.push(r);r=[];c='';}
           else if(ch==='\r'){} else c+=ch; } }
  if(c!==''||r.length){r.push(c);R.push(r);} return R; }
const now = new Date(Date.now() + (new Date().getTimezoneOffset()+540)*60000); // KST
const tk = [now.getFullYear(), now.getMonth()+1, now.getDate()];
const isToday = (v) => { const m=(v||'').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? (+m[1]===tk[0] && +m[2]===tk[1] && +m[3]===tk[2]) : false; };
function crunch(rows){
  const H=rows[0];
  const iU=H.indexOf("Update 날짜"), iT=H.indexOf("인입사유(tag)"), iC=H.indexOf("대분류");
  let todayCount=0; const tal={}; const dayCounts={}; const cat={ "휴대폰보호필름":{}, "휴대폰케이스":{} };
  for(let i=1;i<rows.length;i++){
    const tag=(rows[i][iT]||'').trim()||'(빈칸)';
    if(tag==='긍정 리뷰') continue;   // user rule 2026-09-08: exclude from all stats/cards
    const cc=(rows[i][iC]||'').trim();
    if(cat[cc]) cat[cc][tag]=(cat[cc][tag]||0)+1;
    const m=(rows[i][iU]||'').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if(m) dayCounts[`${m[1]}-${m[2]}-${m[3]}`]=(dayCounts[`${m[1]}-${m[2]}-${m[3]}`]||0)+1;
    if(isToday(rows[i][iU])){ todayCount++; tal[tag]=(tal[tag]||0)+1; }
  }
  const todayTags=Object.entries(tal).sort((a,b)=>b[1]-a[1]);
  const block=(o)=>({tot:Object.values(o).reduce((a,b)=>a+b,0),top5:Object.entries(o).sort((a,b)=>b[1]-a[1]).slice(0,5)});
  // trailing-7-day daily average (excl. today) — significance baseline for 오늘 총 N건
  let recentSum=0;
  for(let k=1;k<=7;k++){
    const d=new Date(tk[0], tk[1]-1, tk[2]-k);
    recentSum += dayCounts[`${d.getFullYear()}-${d.getMonth()+1}-${d.getDate()}`] || 0;
  }
  return { todayCount, todayTags, recentAvg: recentSum/7, film:block(cat["휴대폰보호필름"]), case:block(cat["휴대폰케이스"]) };
}
const PX = parse(await csv("12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI","970309432"));
const Z8 = parse(await csv("19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4","970309432"));
JSON.stringify({ tk, pixel11: crunch(PX), glxz8: crunch(Z8) });
```

### 4. Test send
```
python3 ~/.claude/skills/badreview-chat-broadcast/broadcast.py --test-only \
  --px-data '<pixel11 json from step 3>' \
  --z8-data '<glxz8 json from step 3>'
```
Close the browser tab. Report the test result + the room list, then **ask for confirmation**.

### 5. Broadcast (only after explicit "yes")
```
python3 ~/.claude/skills/badreview-chat-broadcast/broadcast.py --all \
  --px-data '<...>' --z8-data '<...>'
```
`--only "name,name"` restricts to a subset. `--date YYYY-MM-DD` overrides today.
Report per-room OK/ERR.

## Notes

- **Significance highlighting** (2026-09-14): any number that's >= 1.5x its natural
  comparison point renders bold+red in the card — 오늘 총 N건 vs the trailing-7-day daily
  average (`recentAvg`, computed above), each Top-5(누적) column's 1위 vs 2위, and 오늘
  최다 인입사유's top tag vs the day's runner-up. Logic lives in the two `report.py`s
  (`SIG_RATIO`/`SIG_RED`) — keep them in sync. Only `decoratedText.text` renders HTML
  color (verified against the live webhook); `topLabel`/`bottomLabel` show raw tags.
- Each room gets **1 combined carousel message** on a normal `--all` (all 3 products) —
  13 rooms = 13 messages. A `--product` subset resend still sends one message per
  product per room, 1s apart, via the old per-product routing.
- New messages every run — webhook messages can't be edited/deleted from here.
- Card layout / height rules live in the two per-product `report.py` (kept in sync).
- If `todayCount` is 0 for a product, its card still sends ("업로드된 배드리뷰 없음") —
  flag it to the user before broadcasting.
