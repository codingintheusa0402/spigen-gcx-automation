#!/usr/bin/env python3
"""Phase G of sc-review-propagate as a plain script (no claude-in-chrome needed).

For every active product's destination sheet, find rows dated today with
사진 유무 = Y and an empty Image URL, open each row's Review Link in a real
(headed) Chrome via Playwright, pull the customer photo URLs with the same
extraction JS the skill uses, and write them back (joined with "|").

Used on the GCX server (Linux/WSL), where the agent has no browser tool;
works on the Mac too. Rules kept from the skill:
  - only today's rows (or --date), columns found by header text, never hardcoded
  - amazon.es is skipped (needs a customer login) and reported
  - right before each write, Image URL is re-read: must still be empty and the
    Review ID must still match; only that one cell is written
  - dry-run unless --commit

  python3 phase_g.py                 # dry run, today (KST)
  python3 phase_g.py --commit        # write
  python3 phase_g.py --date 2026-10-08 --product GlxZ8 --commit
"""
import argparse, asyncio, os, sys
from collections import defaultdict
from urllib.parse import urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import propagate as P  # noqa: E402  (PRODUCTS, get_service, header_row, find_col, is_today_cell, ...)

PROFILE_DIR = os.path.expanduser("~/.chrome-phaseg-profile")
SKIP_DOMAINS = {"www.amazon.es": "needs amazon.es login"}
EXTRACT_JS = """(rid) => {
  const el = document.getElementById(rid) || document.querySelector('[data-hook="review"]');
  if (!el) return {found: false};
  const imgs = Array.from(el.querySelectorAll('img')).map(i => i.src)
    .filter(s => s.includes('/images/I/') && !s.includes('avatars') && !s.includes('/sash/'));
  const map = {};
  imgs.forEach(u => { const base = u.split('/').pop().split('.')[0];
    if (!map[base] || !u.includes('_SY500_')) map[base] = u; });
  return {found: true, imgs: Object.values(map)};
}"""


def col_letter(i):
    return P.idx_to_a1_col(i)


def candidates(svc, product, cfg, day):
    sid, sheet = cfg["dest_id"], cfg["dest_sheet"]
    hdr = P.header_row(svc, sid, sheet)
    pick = lambda *names: next((P.find_col(hdr, want_exact=n) for n in names if P.find_col(hdr, want_exact=n)), None)
    c = {"사진 유무": pick("사진 유무"),
         "Image URL": pick("Image URL", "Review Image / Video"),     # Apify-layout books (SDA, Auto_Acc, …)
         "Review Link": pick("Review Link", "Review Url"),
         "Review ID": pick("Review ID"),
         "date": pick("Update 날짜", "Exported Date")}
    if not all(c.values()):
        missing = [k for k, v in c.items() if not v]
        print(f"  {product}: not checked — no {', '.join(missing)} column in {sheet}")
        return []
    last = max(c.values())
    rows = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sheet}'!A2:{col_letter(last)}",
        valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    out = []
    for r_i, row in enumerate(rows, start=2):
        g = lambda k: row[c[k] - 1] if len(row) >= c[k] else ""
        if not P.is_today_cell(g("date"), day):
            continue
        if str(g("사진 유무")).strip().upper() != "Y" or str(g("Image URL")).strip():
            continue
        link, rid = str(g("Review Link")).strip(), str(g("Review ID")).strip()
        if link and rid:
            out.append(dict(product=product, sid=sid, sheet=sheet, row=r_i, link=link, rid=rid,
                            img_col=col_letter(c["Image URL"]), rid_col=col_letter(c["Review ID"])))
    return out


async def fetch_images(cands):
    from playwright.async_api import async_playwright
    os.makedirs(PROFILE_DIR, exist_ok=True)
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(PROFILE_DIR, channel="chrome", headless=False)
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        for c in cands:
            try:
                await page.goto(c["link"], wait_until="domcontentloaded", timeout=45000)
                await page.wait_for_timeout(2500)
                res = await page.evaluate(EXTRACT_JS, c["rid"])
                c["imgs"] = res.get("imgs", []) if res.get("found") else None
                if "signin" in page.url or "captcha" in (await page.content())[:20000].lower():
                    c["note"] = "sign-in or captcha page"
            except Exception as e:  # keep going; report per row
                c["imgs"], c["note"] = None, f"error: {str(e)[:80]}"
        await ctx.close()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=P.today_kst_iso(), help="YYYY-MM-DD (default: today KST)")
    ap.add_argument("--product", choices=list(P.PRODUCTS))
    ap.add_argument("--commit", action="store_true")
    a = ap.parse_args()
    if sys.platform.startswith("linux"):
        os.environ.setdefault("DISPLAY", ":0")

    svc = P.get_service()
    prods = [a.product] if a.product else [p for p, c in P.PRODUCTS.items() if not c.get("inactive")]
    cands = []
    for p in prods:
        found = candidates(svc, p, P.PRODUCTS[p], a.date)
        print(f"  {p}: {len(found)} candidate row(s)")
        cands += found
    skipped = [c for c in cands if urlparse(c["link"]).netloc in SKIP_DOMAINS]
    todo = [c for c in cands if c not in skipped]
    print(f"Phase G {a.date}: {len(cands)} candidates · {len(skipped)} skipped (amazon.es) · {len(todo)} to check")
    if todo:
        asyncio.run(fetch_images(todo))

    filled, none_found = 0, []
    for c in todo:
        if not c.get("imgs"):
            none_found.append(c)
            continue
        val = "|".join(c["imgs"])
        rng = f"'{c['sheet']}'!{c['img_col']}{c['row']}"
        if not a.commit:
            print(f"  [dry-run] {c['product']} row {c['row']}: {len(c['imgs'])} image(s) → {rng}")
            filled += 1
            continue
        # safety re-check right before writing
        cur = svc.spreadsheets().values().batchGet(
            spreadsheetId=c["sid"], ranges=[rng, f"'{c['sheet']}'!{c['rid_col']}{c['row']}"]).execute()["valueRanges"]
        img_now = (cur[0].get("values") or [[""]])[0][0]
        rid_now = (cur[1].get("values") or [[""]])[0][0]
        if str(img_now).strip() or str(rid_now).strip() != c["rid"]:
            print(f"  {c['product']} row {c['row']}: changed since read — not written")
            continue
        svc.spreadsheets().values().update(spreadsheetId=c["sid"], range=rng, valueInputOption="RAW",
                                           body={"values": [[val]]}).execute()
        filled += 1
        print(f"  {c['product']} row {c['row']}: wrote {len(c['imgs'])} image(s)")

    by = defaultdict(int)
    for c in none_found:
        by[c.get("note") or "no image on live page"] += 1
    print(f"\nSummary: candidates {len(cands)} · {'filled' if a.commit else 'would fill'} {filled} · "
          f"skipped amazon.es {len(skipped)} · not filled {len(none_found)} {dict(by) if by else ''}")

    # GCX server: Amazon customer session expired → alert the GCX Server room with a re-login button per site
    signin = sorted({urlparse(c["link"]).netloc.split("amazon.")[-1] for c in none_found
                     if "sign-in" in (c.get("note") or "")})
    if signin and os.path.exists(os.path.expanduser("~/.config/gcx_server_webhook.txt")):
        try:
            sys.path.insert(0, os.path.expanduser("~/Desktop/GCX/ServerBootstrap/watchdog"))
            from gcx_alert import alert
            for tld in signin:
                n = sum(1 for c in none_found if c["link"].find(f"amazon.{tld}") >= 0)
                alert(f"Amazon login expired: amazon.{tld}",
                      f"Photo check (Phase G) couldn't open {n} review page(s) on amazon.{tld} — it lands on the sign-in page. "
                      f"Tap the button, sign in, then re-run the photo check.", level="bad", relogin=f"amazon:{tld}")
        except Exception as e:
            print(f"  (login alert failed: {e})")


if __name__ == "__main__":
    main()
