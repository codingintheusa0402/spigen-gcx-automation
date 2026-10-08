#!/usr/bin/env python3
"""
SIREN check — defect-pattern early warning for ticket-reporter Monitor mode.

Run once per NEWLY-sent ticket report, right after send.py succeeds. Combines two counts
for the same (defect, product) combination:

  1. zendesk_count — how many Zendesk tickets ever carry the same defect tag + device tag +
     product tag (passed in already computed by the caller, via a browser-side Search API
     call — see SKILL.md "SIREN check" step for the exact JS, since this needs an
     authenticated Zendesk session that this standalone script does not have).
  2. bad_review_count — how many rows in the "1-3점" bad-review monitoring sheet (same
     spreadsheet as 신제품 라인업) match the same defect term by SKU/ASIN (resolved via the
     신제품 라인업 master table, since 모델명 text can differ slightly between the two
     sheets — SKU/ASIN is the reliable join key) and by 인입사유(tag) (column N) equal to
     the defect's bare Korean term (Zendesk tag "_case__황변" -> bare term "황변" -> exact
     match against column N; confirmed 2026-09-21 that both sides use the identical bare
     Korean term).

total = zendesk_count + bad_review_count. total >= 3 => SIREN triggered.

Every check (triggered or not) is appended to:
  - state/siren_log.json (local, full detail incl. matching ids/rows)
  - "SIREN_Log" tab in the ticket-reporter's own queue spreadsheet (QUEUE_SHEET_ID below)

A SIREN alert (Chat thread reply on the ticket's own report thread, listing the matching
tickets/reviews + their images) is posted only the FIRST time a given (defect_tag,
product_tag) cluster crosses the >=3 threshold — later tickets in the same cluster are
logged but not re-alerted, to avoid spamming the same thread chain. (This default can be
changed — e.g. re-alert every time the count grows by N more — if the user asks.)

Usage:
  python3 siren.py check --ticket-id 1000162858 \
      --defect-tag "_case__황변" --device "Galaxy S26 Ultra" --device-tag galaxy_s26_ultra \
      --product-name "Ultra Hybrid MagFit" --product-tag ultra_hybrid_magfit \
      --zendesk-count 85 --zendesk-ids '[1,2,3]' \
      [--dry-run]

If --defect-tag / --device-tag / --product-tag could not be derived (see SKILL.md), pass
--skip-reason "<why>" instead of the four --zendesk-* / tag args — this logs a
'siren_skip' entry (no counting attempted) rather than a check.
"""
import argparse
import datetime
import json
import os
import re
import sys
import urllib.request
import urllib.error

# Defect terms (bare, no "_type__" prefix) SIREN actually tracks — user-specified allowlist
# (2026-09-21, /revision feedback). Everything NOT in this set is skipped without counting,
# most notably 황변 (yellowing) and other non-manufacturing-defect noise
# (동봉제품상이함/배송상태불만/오배송/미배송/중고품배송/파손된상품수령/기타사항). iPhone
# 17/18 Series compatibility inquiries are separately excluded structurally — SIREN only
# ever runs after a "4. Product Issue" report is sent, and compatibility questions are
# Category "6. Product Inquiry", so they never reach this check at all.
ALLOWED_DEFECT_TERMS = frozenset({
    "게이트홀", "TPU늘어남", "디자인", "자석들뜸", "자석탈락", "자국", "충전호환불가",
    "렌즈호환불가", "부착어려움(MagSafe)", "버튼감", "이염/변색", "백화", "유격", "스크래치",
    "얼룩", "버튼부파손", "외관파손", "분리/이탈", "유막", "유해물질",
    "제품품질불만(사이즈,디자인 외)", "코팅벗겨짐", "킥스탠드이슈", "형합", "냄새", "자력약함",
    "색상", "제품정보상이", "가격", "NFC불가", "카메라커버이슈", "재질", "플래시/빛번짐",
    "인쇄불량", "온도센서간섭(미사용)", "가장자리날카로움", "자사필름간섭", "타사필름간섭",
    "컷아웃", "스트랩장탈착이슈", "도크호환불가", "그립호환불가(미사용)", "지문인식",
    "블루투스연결", "앱사용(미사용)", "오디오작동불량(미사용)", "CC커버이슈", "프레임파손",
    "힌지파손", "글라스깨짐", "디바이스와맞지않음", "레인보우현상", "먼지유입", "밀림",
    "버블", "부착벗겨짐", "부착어려움", "선명도떨어짐", "프레임형합", "오렌지필현상",
    "자사케이스간섭", "타사케이스간섭", "점착액", "찍힘", "측면들뜸", "터치인식",
    "표면스크래치", "표면줄생김", "UV라이트문제", "블랙테두리", "전면센서간섭", "후면센서간섭",
    "결로", "전면카메라화질저하", "후면카메라화질저하", "반사방지저하", "프라이버시불가",
    "구성품누락",
})

GWS_SHIM_TOKEN = "/Users/kevinkim/.config/gws_shim/token.json"
QUEUE_SHEET_ID = "1GUU2CLa60XSGWHI6yZ6Jkx4nrzcGgBJnuxwQAZktH7M"
SIREN_TAB = "SIREN_Log"

PRODUCT_SHEET_ID = "1aYxZRm7pf5Egx6fIoAGpGg8CWzHaZ_zsBRKsvh9U1iU"
LINEUP_TAB = "신제품 라인업"     # A=ASIN B=SKU C=기종명 D=모델명 ...
BADREVIEW_TAB = "1-3점"        # A=ASIN ... J=Image URL ... N=인입사유(tag) P=SKU Q=기종명 R=모델명

LOCAL_LOG = os.path.join(os.path.dirname(__file__), "state", "siren_log.json")

CHAT_KEY = "AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI"
CHAT_BASE = "https://chat.googleapis.com/v1/{thread_or_space}/messages?key=" + CHAT_KEY + "&token={token}"
# Same webhook token send.py posts the original reports to (real chatroom).
REPORT_WEBHOOK_SID = "AAQAdqYt1ro"
REPORT_WEBHOOK_TOKEN = "Sm03r1isd7UGFIuWkdg8whlC9nt5HlPsaxt3dYmQXmY"


def _sheets_client():
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    with open(GWS_SHIM_TOKEN, encoding="utf-8") as f:
        info = json.load(f)
    creds = Credentials(
        token=info.get("token"), refresh_token=info["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=info["client_id"], client_secret=info["client_secret"],
        scopes=info.get("scopes"),
    )
    creds.refresh(Request())
    info["token"] = creds.token
    with open(GWS_SHIM_TOKEN, "w", encoding="utf-8") as f:
        json.dump(info, f)
    return build("sheets", "v4", credentials=creds)


def _bare_term(defect_tag: str) -> str:
    """'_case__황변' -> '황변'"""
    return re.sub(r"^_[^_]+__", "", defect_tag)


def _load_local_log() -> dict:
    if os.path.exists(LOCAL_LOG):
        with open(LOCAL_LOG, encoding="utf-8") as f:
            return json.load(f)
    return {"checks": []}


def _save_local_log(data: dict) -> None:
    os.makedirs(os.path.dirname(LOCAL_LOG), exist_ok=True)
    with open(LOCAL_LOG, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def _already_triggered(log: dict, defect_tag: str, product_tag: str) -> bool:
    return any(
        c.get("defect_tag") == defect_tag and c.get("product_tag") == product_tag
        and c.get("triggered")
        for c in log.get("checks", [])
    )


def lookup_asin_sku_set(sheets, device: str, product_name: str):
    res = sheets.spreadsheets().values().get(
        spreadsheetId=PRODUCT_SHEET_ID, range=f"{LINEUP_TAB}!A2:D",
    ).execute()

    def _norm(s: str) -> str:
        return re.sub(r"[^a-z0-9가-힣]", "", s.lower())

    prod_norm = _norm(product_name)
    asins, skus = set(), set()
    exact_hit = False
    for row in res.get("values", []):
        row = row + [""] * (4 - len(row))
        asin, sku, model_line, model_name = row[:4]
        if model_line.strip() != device.strip():
            continue
        if model_name.strip() == product_name.strip():
            exact_hit = True
            if asin:
                asins.add(asin.strip())
            if sku:
                skus.add(sku.strip())
    if not exact_hit:
        # Fallback: 신제품 라인업's own 모델명 can differ from the ticket's Product Name
        # field (confirmed 2026-09-21, e.g. ticket "SP_Glas.tR EZ Fit Slim" vs sheet
        # "Glas.tR EZ Fit") — try substring containment either direction, normalized.
        for row in res.get("values", []):
            row = row + [""] * (4 - len(row))
            asin, sku, model_line, model_name = row[:4]
            if model_line.strip() != device.strip() or not model_name.strip():
                continue
            model_norm = _norm(model_name)
            if model_norm and (model_norm in prod_norm or prod_norm in model_norm):
                if asin:
                    asins.add(asin.strip())
                if sku:
                    skus.add(sku.strip())
    return asins, skus


def count_bad_reviews(sheets, asins: set, skus: set, bare_term: str):
    res = sheets.spreadsheets().values().get(
        spreadsheetId=PRODUCT_SHEET_ID, range=f"{BADREVIEW_TAB}!A2:R",
    ).execute()
    matches = []
    for row in res.get("values", []):
        row = row + [""] * (18 - len(row))
        asin = row[0].strip()
        image_url = row[9].strip()
        review_id = row[10].strip()
        tag = row[13].strip()
        sku = row[15].strip()
        if tag != bare_term:
            continue
        if (asins and asin in asins) or (skus and sku in skus):
            matches.append({"review_id": review_id, "asin": asin, "sku": sku, "image_url": image_url})
    return matches


def append_sheet_row(sheets, row: list) -> None:
    sheets.spreadsheets().values().append(
        spreadsheetId=QUEUE_SHEET_ID, range=f"{SIREN_TAB}!A:P",
        valueInputOption="RAW", insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


def lookup_thread_name(sheets, ticket_id: str):
    res = sheets.spreadsheets().values().get(
        spreadsheetId=QUEUE_SHEET_ID, range="TicketQueue!A:H",
    ).execute()
    for row in reversed(res.get("values", [])):  # most recent send wins
        if row and row[0].strip() == str(ticket_id):
            if len(row) >= 8 and row[7]:
                return row[7]
    return None


def post_alert(thread_name: str, html: str, images: list):
    from urllib.parse import quote
    widgets = [{"textParagraph": {"text": html}}]
    if images:
        capped = images[:10]
        widgets.append({"grid": {
            "columnCount": 3 if len(capped) >= 3 else max(1, len(capped)),
            "items": [{"id": str(i), "image": {"imageUri": u}} for i, u in enumerate(capped)],
        }})
        if len(images) > 10:
            widgets.append({"textParagraph": {"text": f"…외 {len(images) - 10}건 이미지 생략"}})
    payload = {
        "cardsV2": [{"cardId": "siren-alert", "card": {"sections": [{"widgets": widgets}]}}],
        "thread": {"name": thread_name},
    }
    url = (f"https://chat.googleapis.com/v1/spaces/{REPORT_WEBHOOK_SID}/messages"
           f"?key={CHAT_KEY}&token={REPORT_WEBHOOK_TOKEN}"
           f"&messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD")
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                  headers={"Content-Type": "application/json; charset=UTF-8"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return True, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return False, e.read().decode("utf-8", "replace")


def cmd_check(a):
    log = _load_local_log()
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).isoformat()

    if a.skip_reason:
        entry = {"ts": now, "ticket_id": a.ticket_id, "event": "siren_skip", "reason": a.skip_reason}
        log["checks"].append(entry)
        _save_local_log(log)
        print(f"SIREN skip logged: {a.skip_reason}")
        return 0

    bare_term = _bare_term(a.defect_tag)
    if bare_term not in ALLOWED_DEFECT_TERMS:
        entry = {"ts": now, "ticket_id": a.ticket_id, "event": "siren_skip",
                  "reason": f"defect term '{bare_term}' not in SIREN allowlist (2026-09-21 /revision)"}
        log["checks"].append(entry)
        _save_local_log(log)
        print(f"SIREN skip logged: not tracked ({bare_term})")
        return 0

    sheets = _sheets_client()
    asins, skus = lookup_asin_sku_set(sheets, a.device, a.product_name)
    if not asins and not skus:
        entry = {"ts": now, "ticket_id": a.ticket_id, "event": "siren_skip",
                  "reason": f"product '{a.device} {a.product_name}' not in 신제품 라인업 — "
                            f"SIREN only covers products listed there (2026-09-21 user rule)"}
        log["checks"].append(entry)
        _save_local_log(log)
        print(f"SIREN skip logged: product not in 신제품 라인업 ({a.device} {a.product_name})")
        return 0
    bad_matches = count_bad_reviews(sheets, asins, skus, bare_term)
    bad_review_count = len(bad_matches)
    zendesk_ids = json.loads(a.zendesk_ids) if a.zendesk_ids else []
    total = a.zendesk_count + bad_review_count
    triggered = total >= 3
    first_for_cluster = triggered and not _already_triggered(log, a.defect_tag, a.product_tag)

    entry = {
        "ts": now, "ticket_id": a.ticket_id, "defect_term": bare_term, "defect_tag": a.defect_tag,
        "device": a.device, "device_tag": a.device_tag, "product_name": a.product_name,
        "product_tag": a.product_tag, "zendesk_count": a.zendesk_count, "zendesk_ticket_ids": zendesk_ids,
        "bad_review_count": bad_review_count, "bad_review_matches": bad_matches,
        "total": total, "triggered": triggered, "first_trigger_for_cluster": first_for_cluster,
        "alert_thread": None,
    }

    alert_html = None
    if first_for_cluster and not a.dry_run:
        thread_name = lookup_thread_name(sheets, a.ticket_id)
        if thread_name:
            images = [m["image_url"] for m in bad_matches if m["image_url"]]
            ticket_links = "".join(
                f'<a href="https://spigenhelp.zendesk.com/agent/tickets/{tid}">#{tid}</a> '
                for tid in zendesk_ids[:10]
            )
            more_note = f" 외 {len(zendesk_ids) - 10}건" if len(zendesk_ids) > 10 else ""
            alert_html = (
                f'🚨 <b><font color="#D93025">SIREN 확인 필요</font></b> — '
                f'<b>{a.device} {a.product_name}</b> / 결함: <b>{bare_term}</b><br>'
                f'누적 <b>{total}건</b> (Zendesk 티켓 {a.zendesk_count}건 + 배드리뷰 {bad_review_count}건)<br>'
                f'Zendesk: {ticket_links}{more_note}<br>'
                f'배드리뷰 매칭: {bad_review_count}건'
            )
            ok, resp = post_alert(thread_name, alert_html, images)
            entry["alert_thread"] = thread_name if ok else f"FAILED: {resp[:200]}"
        else:
            entry["alert_thread"] = "FAILED: no thread_name found in TicketQueue"

    log["checks"].append(entry)
    _save_local_log(log)

    append_sheet_row(sheets, [
        now, a.ticket_id, bare_term, a.defect_tag, a.device, a.device_tag,
        a.product_name, a.product_tag, a.zendesk_count, ",".join(map(str, zendesk_ids)),
        bad_review_count, ",".join(m["review_id"] for m in bad_matches),
        total, triggered, first_for_cluster, entry["alert_thread"] or "",
    ])

    print(json.dumps(entry, ensure_ascii=False, indent=1))
    return 0


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("check")
    c.add_argument("--ticket-id", required=True)
    c.add_argument("--defect-tag")
    c.add_argument("--device")
    c.add_argument("--device-tag")
    c.add_argument("--product-name")
    c.add_argument("--product-tag")
    c.add_argument("--zendesk-count", type=int)
    c.add_argument("--zendesk-ids", help="JSON array string")
    c.add_argument("--skip-reason", help="log a skip instead of running the check")
    c.add_argument("--dry-run", action="store_true")
    c.set_defaults(func=cmd_check)

    a = ap.parse_args()
    return a.func(a)


if __name__ == "__main__":
    raise SystemExit(main())
