"""
Weekly/daily job: find blank or 'B'-valued ASIN cells in the "US" tab of
https://docs.google.com/spreadsheets/d/1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw
(gid 2033325257), look up each row's SKU in Seller Central inventory --
cascading DE -> JP FBA inventory -> US inventory health -- and fill in the
found ASIN. Posts a completion/changes summary to the GCX Chat webhook.

Run unattended via launchd (see com.spigen.sku-asin-filler.plist), gated by
scheduler_gate.py so it only actually executes once per weekday after 9:30 KST.

Requires a one-time interactive login (see one_time_login.py) into the 3
Seller Central accounts inside PROFILE_DIR; this script reuses those cookies.
"""
import asyncio
import json
import os
import re
import sys
from datetime import datetime

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from playwright.async_api import async_playwright

TOKEN = os.path.expanduser("~/.config/gws_shim/token.json")
PROFILE_DIR = os.path.expanduser("~/.chrome-sc-inventory-profile")
SHEET_ID = "1dlY6q8trbVMVJAjw_OUoxp1cguA2oTB8WlPhHR01xIw"
SHEET_TAB = "US"
NOTIFY_WEBHOOK = (
    "https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages"
    "?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI"
    "&token=t0A95TCkAqXjLGYBoZrgZSG_XuSSB9UDL3GsFJX_YDc"
)
LOG_PATH = os.path.expanduser("~/Desktop/GCX/Scrapers/SKU_ASIN_Filler/run.log")

DE_URL = (
    "https://sellercentral.amazon.de/myinventory/inventory"
    "?fulfilledBy=all&page=1&pageSize=25&searchField=all&searchTerm={sku}"
    "&sort=date_created_desc&status=all"
)
JP_URL = (
    "https://sellercentral-japan.amazon.com/inventoryplanning/manageinventoryhealth"
    "?sort_column=product_details&sort_direction=asc&sort_column_sub=msku&search={sku}"
)
US_URL = "https://sellercentral.amazon.com/inventoryplanning/manageinventoryhealth?search={sku}"

ASIN_RE = re.compile(r"\bB0[A-Z0-9]{8}\b")


def log(msg):
    line = f"[{datetime.now().isoformat()}] {msg}"
    print(line)
    with open(LOG_PATH, "a") as f:
        f.write(line + "\n")


def get_sheets_service():
    with open(TOKEN) as f:
        info = json.load(f)
    creds = Credentials.from_authorized_user_info(info)
    creds.refresh(Request())
    with open(TOKEN, "w") as f:
        f.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


def find_blank_rows(svc):
    res = svc.spreadsheets().values().get(
        spreadsheetId=SHEET_ID, range=f"'{SHEET_TAB}'!A3:F2000",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    rows = res.get("values", [])
    blanks = []
    for i, r in enumerate(rows):
        row_num = i + 3
        sku = r[0] if len(r) > 0 else ""
        asin = r[1] if len(r) > 1 else ""
        color = r[4] if len(r) > 4 else ""
        if not sku:
            continue
        if asin.strip() == "" or asin.strip() == "B":
            blanks.append({"row": row_num, "sku": sku, "color": color})
    return blanks


def parse_candidates(text):
    """Extract (status, asin, sku) tuples from a Seller Central results page's
    visible text. Tolerant of DE/JP/US label format differences."""
    candidates = []
    # Split on "ASIN" occurrences and look at a window around each for SKU + status
    for m in re.finditer(r"ASIN[:\s]*\n?\s*(B0[A-Z0-9]{8})", text):
        asin = m.group(1)
        window = text[max(0, m.start() - 300):m.end() + 300]
        sku_m = re.search(r"SKU[:\s]*\n?\s*(\S+)", window)
        sku = sku_m.group(1) if sku_m else ""
        is_bundle = "Bundle" in window[:320]
        candidates.append({"asin": asin, "sku": sku, "is_bundle": is_bundle, "window": window})
    return candidates


def pick_best(candidates, target_sku, color):
    if not candidates:
        return None
    exact = [c for c in candidates if c["sku"] == target_sku]
    if exact:
        return exact[0]["asin"]
    non_bundle = [c for c in candidates if not c["is_bundle"]]
    pool = non_bundle or candidates
    if color:
        color_match = [c for c in pool if color.lower() in c["window"].lower()]
        if len(color_match) == 1:
            return color_match[0]["asin"]
    if len(pool) == 1:
        return pool[0]["asin"]
    return None  # ambiguous -- don't guess


async def search_marketplace(page, url_template, sku, color, wait_s=2.5):
    url = url_template.format(sku=sku)
    await page.goto(url, wait_until="domcontentloaded")
    await page.wait_for_timeout(int(wait_s * 1000))
    text = await page.inner_text("body")
    if re.search(r"\bsign[\s-]?in\b", text, re.I) and "password" in text.lower():
        return "login_required", None
    candidates = parse_candidates(text)
    asin = pick_best(candidates, sku, color)
    if asin:
        return "found", asin
    if candidates:
        return "ambiguous", None
    return "not_found", None


async def run():
    svc = get_sheets_service()
    blanks = find_blank_rows(svc)
    log(f"Found {len(blanks)} blank/B rows: {[b['sku'] for b in blanks]}")

    if not blanks:
        post_chat("오늘도 공란 없음 -- 확인할 ASIN 없습니다. (US 탭)")
        return

    filled = []
    not_found = []
    login_issues = set()

    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            PROFILE_DIR, channel="chrome", headless=False
        )
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()

        for b in blanks:
            sku, color, row = b["sku"], b["color"], b["row"]
            found_asin = None
            found_on = None
            for label, template in (("DE", DE_URL), ("JP", JP_URL), ("US", US_URL)):
                status, asin = await search_marketplace(page, template, sku, color)
                if status == "login_required":
                    login_issues.add(label)
                    continue
                if status == "found":
                    found_asin = asin
                    found_on = label
                    break
            if found_asin:
                # safety check: re-read the cell before writing
                cur = svc.spreadsheets().values().get(
                    spreadsheetId=SHEET_ID, range=f"'{SHEET_TAB}'!A{row}:B{row}",
                    valueRenderOption="FORMATTED_VALUE",
                ).execute().get("values", [[]])[0]
                cur_sku = cur[0] if len(cur) > 0 else ""
                cur_asin = cur[1] if len(cur) > 1 else ""
                if cur_sku == sku and cur_asin.strip() in ("", "B"):
                    svc.spreadsheets().values().update(
                        spreadsheetId=SHEET_ID, range=f"'{SHEET_TAB}'!B{row}",
                        valueInputOption="RAW", body={"values": [[found_asin]]},
                    ).execute()
                    filled.append((sku, found_asin, found_on))
                    log(f"Filled {sku} -> {found_asin} (found on {found_on})")
            else:
                not_found.append(sku)
                log(f"Not found anywhere: {sku}")

        await ctx.close()

    msg_lines = [f"ASIN 채우기 작업 완료 ({datetime.now().strftime('%Y-%m-%d')})\n"]
    if filled:
        msg_lines.append(f"채워짐 ({len(filled)}건):")
        for sku, asin, mkt in filled:
            msg_lines.append(f"- {sku} -> {asin} ({mkt}에서 발견)")
    else:
        msg_lines.append("채워진 항목 없음.")
    if not_found:
        msg_lines.append(f"\nDE/JP/US 전부 검색했으나 못 찾음 ({len(not_found)}건): {', '.join(not_found)}")
    if login_issues:
        msg_lines.append(
            f"\n주의: 다음 마켓플레이스 로그인 세션이 만료된 것으로 보입니다: {', '.join(sorted(login_issues))}. "
            f"one_time_login.py 를 다시 실행해 재로그인 해주세요."
        )
    post_chat("\n".join(msg_lines))


def post_chat(text):
    import urllib.request
    data = json.dumps({"text": text}).encode()
    req = urllib.request.Request(
        NOTIFY_WEBHOOK, data=data, headers={"Content-Type": "application/json"}
    )
    try:
        urllib.request.urlopen(req, timeout=15)
    except Exception as e:
        log(f"Chat post failed: {e}")


if __name__ == "__main__":
    asyncio.run(run())
