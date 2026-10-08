#!/usr/bin/env python3
"""
SIREN report Slides generator — builds a Google Slides deck matching the exact template
used in the real '26년 SIREN' tracking sheet's linked decks (reverse-engineered 2026-09-21
from two live examples: presentations 19KWYQpoh-8W6WJGGK30UKl6a5njutrTtNF8yc8QcIGg and
1oziw7rbEL1pPLlfbiu_eg6uFOmJgPlSKcbLLwKuPZqU, both linked from spreadsheet
15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI tab "26년 SIREN").

This generator produces the GCX/CX-side portion only (title -> overview -> one slide per
matching claim/bad-review). The later CQ 품질관리팀 investigation section seen in mature
decks (Field Issue / 제품 검증 / CQ 검토 의견 slides) is added manually by the CQ team
after physical sample testing — it is out of scope here and never auto-generated.

Usage:
  python3 build_siren_slides.py --data siren_case.json [--dry-run]

--data JSON schema — see siren_case.example.json in this folder for a worked example
built from a real ticket-reporter Monitor-mode send.
"""
import argparse
import datetime
import json
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GWS_SHIM_TOKEN = "/Users/kevinkim/.config/gws_shim/token.json"

PAGE_W_PT = 720.0
PAGE_H_PT = 405.0

GOTHIC = "Gothic A1"
ARIAL = "Arial"
NOTO = "Noto Sans"
NOTO_KR = "Noto Sans KR"

BLACK = {"red": 0.0, "green": 0.0, "blue": 0.0}
WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
BLUE = {"red": 0.0, "green": 0.0, "blue": 1.0}   # 클레임 #NN title number
RED = {"red": 1.0, "green": 0.0, "blue": 0.0}    # 배드 리뷰 #NN title number
DARK1 = {"red": 0.1, "green": 0.1, "blue": 0.1}  # approximation of theme DARK1 on a blank deck

# Sampled directly from Spigen's own cover-slide template (spigen-slides skill,
# LIGHT_TEMPLATE_ID 1BBG9PR6ZBsEABbJLhbUUfRMkgGYQtNMOWAmLQgPhr70) — the real cover bg
# color, not the #FF6B1A brand-guide value (close but not identical).
COVER_ORANGE = {"red": 1.0, "green": 0.3529412, "blue": 0.0}
# Spigen brand accent orange (#FF6B1A) — used for the highlighted Global 리뷰 평점 cell.
ACCENT_ORANGE = {"red": 1.0, "green": 0.4196078, "blue": 0.1019608}
# Spigen logo mark image — cropped from the cover template's rendered thumbnail (the
# template's own inline lh7-rt.googleusercontent.com/slidesz/... URL is session-gated and
# not fetchable by the Slides API's server-side image fetcher), re-hosted as a public
# Drive file (2026-09-21) since createImage requires a truly public URL.
SPIGEN_LOGO_URL = "https://drive.google.com/uc?export=view&id=1TOgud-mxQDo1BUIE5UVeU050T8C9W-Pj"

COUNTRY_ORDER = ["UK", "DE", "IT", "FR", "ES", "JP", "IN", "US", "KR"]


def _creds():
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
    return creds


class DeckBuilder:
    def __init__(self, title: str):
        self.creds = _creds()
        self.slides = build("slides", "v1", credentials=self.creds)
        pres = self.slides.presentations().create(body={"title": title}).execute()
        self.pid = pres["presentationId"]
        self._default_slide_id = pres["slides"][0]["objectId"]
        self._counter = 0
        self.reqs = []

    def _oid(self, prefix: str) -> str:
        self._counter += 1
        return f"siren_{prefix}_{self._counter:04d}"

    def flush(self):
        if not self.reqs:
            return
        # Slides API caps batchUpdate at a few thousand requests; our decks are small,
        # but chunk defensively anyway.
        for i in range(0, len(self.reqs), 400):
            chunk = self.reqs[i:i + 400]
            self.slides.presentations().batchUpdate(
                presentationId=self.pid, body={"requests": chunk}
            ).execute()
        self.reqs = []

    def new_slide(self) -> str:
        sid = self._oid("slide")
        self.reqs.append({
            "createSlide": {
                "objectId": sid,
                "slideLayoutReference": {"predefinedLayout": "BLANK"},
            }
        })
        return sid

    def delete_default_slide(self):
        self.reqs.append({"deleteObject": {"objectId": self._default_slide_id}})

    def set_background(self, slide_id: str, color: dict):
        self.reqs.append({
            "updatePageProperties": {
                "objectId": slide_id,
                "pageProperties": {
                    "pageBackgroundFill": {"solidFill": {"color": {"rgbColor": color}, "alpha": 1}}
                },
                "fields": "pageBackgroundFill",
            }
        })

    def add_textbox(self, slide_id: str, x: float, y: float, w: float, h: float,
                     text: str, font: str = GOTHIC, size: float = 12, bold: bool = False,
                     color: dict = None, align: str = "START", runs: list = None):
        """text/font/size/bold/color style the whole box unless `runs` is given —
        runs: list of (text, size, bold) applied as separate paragraphs/sizes in order
        (e.g. a big bold title line followed by a smaller bold subtitle line)."""
        oid = self._oid("tb")
        color = color or BLACK
        self.reqs.append({
            "createShape": {
                "objectId": oid,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "width": {"magnitude": w, "unit": "PT"},
                        "height": {"magnitude": h, "unit": "PT"},
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": x, "translateY": y,
                        "unit": "PT",
                    },
                },
            }
        })
        if runs:
            full_text = "".join(t for t, _, _ in runs)
            self.reqs.append({"insertText": {"objectId": oid, "text": full_text}})
            idx = 0
            for run_text, run_size, run_bold in runs:
                end = idx + len(run_text)
                self.reqs.append({
                    "updateTextStyle": {
                        "objectId": oid,
                        "style": {
                            "fontFamily": font,
                            "fontSize": {"magnitude": run_size, "unit": "PT"},
                            "bold": run_bold,
                            "foregroundColor": {"opaqueColor": {"rgbColor": color}},
                        },
                        "fields": "fontFamily,fontSize,bold,foregroundColor",
                        "textRange": {"type": "FIXED_RANGE", "startIndex": idx, "endIndex": end},
                    }
                })
                idx = end
            self.reqs.append({
                "updateParagraphStyle": {
                    "objectId": oid,
                    "style": {"alignment": align},
                    "fields": "alignment",
                    "textRange": {"type": "ALL"},
                }
            })
        elif text:
            self.reqs.append({"insertText": {"objectId": oid, "text": text}})
            self.reqs.append({
                "updateTextStyle": {
                    "objectId": oid,
                    "style": {
                        "fontFamily": font,
                        "fontSize": {"magnitude": size, "unit": "PT"},
                        "bold": bold,
                        "foregroundColor": {"opaqueColor": {"rgbColor": color}},
                    },
                    "fields": "fontFamily,fontSize,bold,foregroundColor",
                    "textRange": {"type": "ALL"},
                }
            })
            self.reqs.append({
                "updateParagraphStyle": {
                    "objectId": oid,
                    "style": {"alignment": align},
                    "fields": "alignment",
                    "textRange": {"type": "ALL"},
                }
            })
        return oid

    def add_table(self, slide_id: str, x: float, y: float, w: float, h: float,
                   rows: list, header_rows: int = 1, merges: list = None,
                   cell_fills: dict = None, cell_links: dict = None):
        """rows: list[list[str]]. merges: list of (row, col, rowspan, colspan) to merge.
        cell_fills: {(row,col): rgbColor} overrides the default black/white background
        for that cell. cell_links: {(row,col): url} hyperlinks that cell's whole text."""
        oid = self._oid("tbl")
        cell_fills = cell_fills or {}
        cell_links = cell_links or {}
        nrows = len(rows)
        ncols = max(len(r) for r in rows)
        self.reqs.append({
            "createTable": {
                "objectId": oid,
                "rows": nrows,
                "columns": ncols,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "width": {"magnitude": w, "unit": "PT"},
                        "height": {"magnitude": h, "unit": "PT"},
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": x, "translateY": y,
                        "unit": "PT",
                    },
                },
            }
        })
        for ri, row in enumerate(rows):
            is_header = ri < header_rows
            for ci, val in enumerate(row):
                if val is None:
                    continue  # merged-away cell, skip
                loc = {"rowIndex": ri, "columnIndex": ci}
                if val:
                    self.reqs.append({"insertText": {"objectId": oid, "cellLocation": loc, "text": str(val)}})
                text_style = {
                    "fontFamily": ARIAL if is_header else NOTO,
                    "fontSize": {"magnitude": 12 if is_header else 11, "unit": "PT"},
                    "bold": True,
                    "foregroundColor": {"opaqueColor": {"rgbColor": WHITE if is_header else DARK1}},
                }
                fields = "fontFamily,fontSize,bold,foregroundColor"
                if (ri, ci) in cell_links:
                    text_style["link"] = {"url": cell_links[(ri, ci)]}
                    fields += ",link"
                self.reqs.append({
                    "updateTextStyle": {
                        "objectId": oid, "cellLocation": loc,
                        "style": text_style,
                        "fields": fields,
                        "textRange": {"type": "ALL"},
                    }
                })
                fill_color = cell_fills.get((ri, ci), BLACK if is_header else WHITE)
                self.reqs.append({
                    "updateTableCellProperties": {
                        "objectId": oid, "tableRange": {"location": loc, "rowSpan": 1, "columnSpan": 1},
                        "tableCellProperties": {
                            "tableCellBackgroundFill": {
                                "solidFill": {"color": {"rgbColor": fill_color}}
                            },
                            "contentAlignment": "MIDDLE",
                        },
                        "fields": "tableCellBackgroundFill,contentAlignment",
                    }
                })
        for (r, c, rowspan, colspan) in (merges or []):
            self.reqs.append({
                "mergeTableCells": {
                    "objectId": oid,
                    "tableRange": {
                        "location": {"rowIndex": r, "columnIndex": c},
                        "rowSpan": rowspan, "columnSpan": colspan,
                    },
                }
            })
        return oid

    def add_video(self, slide_id: str, drive_file_id: str, x: float, y: float, w: float, h: float):
        """Embeds a Google Drive video (customer video re-uploaded from Zendesk to Drive)."""
        oid = self._oid("vid")
        self.reqs.append({"createVideo": {
            "objectId": oid, "source": "DRIVE", "id": drive_file_id,
            "elementProperties": {
                "pageObjectId": slide_id,
                "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"},
            },
        }})
        return oid

    def add_image(self, slide_id: str, url: str, x: float, y: float, w: float, h: float):
        oid = self._oid("img")
        self.reqs.append({
            "createImage": {
                "objectId": oid,
                "url": url,
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "width": {"magnitude": w, "unit": "PT"},
                        "height": {"magnitude": h, "unit": "PT"},
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": x, "translateY": y,
                        "unit": "PT",
                    },
                },
            }
        })
        return oid


def build_title_slide(b: DeckBuilder, data: dict):
    """Matches Spigen's own cover-slide template (spigen-slides skill's LIGHT_TEMPLATE_ID
    1BBG9PR6ZBsEABbJLhbUUfRMkgGYQtNMOWAmLQgPhr70), read directly via the Slides API
    2026-09-21: orange page background, big bold title line + smaller bold subtitle line,
    Spigen logo mark top-right, meta/date lines bottom, all Noto Sans KR."""
    sid = b.new_slide()
    b.set_background(sid, COVER_ORANGE)
    b.add_textbox(sid, 27.6, 34.3, 594.8, 99.4, None, font=NOTO_KR, color=BLACK, runs=[
        ("제품 클레임 조치 사항 검토(SIREN)\n", 36, True),
        (data["issue_title"], 20, True),
    ])
    b.add_image(sid, SPIGEN_LOGO_URL, 634.5, 35.4, 59.1, 55.4)
    b.add_textbox(sid, 27.6, 324.4, 394.6, 40.3,
                  f"경영지원부문ㅣ사업지원실ㅣ글로벌CX전략팀 {data['author']} 담당",
                  font=NOTO_KR, size=12, bold=False, color=BLACK)
    # real template's date box is narrow and expects the compact no-space format
    # (e.g. "2024.12.11") to avoid wrapping to two lines.
    b.add_textbox(sid, 607.7, 323.3, 76.4, 42.5, data["date"].replace(" ", ""),
                  font=NOTO_KR, size=12, bold=False, color=BLACK)


# Product master — 'Data' tab, 생산업체 column is the source of truth for 제조 업체명
# (user rule 2026-09-28; the 신제품 라인업 tab often says '입고처리 미진행').
PRODUCT_MASTER_ID = "1fx9K4r2T9SeZK076zy9kMHoLzAKDgmlRp-C2VtnTKVo"


def _sku_asin_pairs(p: dict):
    """[(sku, asin)] for the SKU/ASIN cell: product.skus list if given, else the legacy
    sku/asin fields ('/'-separated SKUs are split; the single asin pairs only with one SKU)."""
    if p.get("skus"):
        return [(x.get("sku", ""), x.get("asin", "")) for x in p["skus"]]
    skus = [x.strip() for x in str(p.get("sku", "")).split("/") if x.strip()]
    if len(skus) == 1:
        return [(skus[0], p.get("asin", ""))]
    return [(sk, "") for sk in skus]


def lookup_manufacturer(skus):
    """생산업체 values from the product master 'Data' tab for the given SKUs, unique, joined."""
    skus = [sk for sk in skus if sk]
    if not skus:
        return None
    try:
        sheets = build("sheets", "v4", credentials=_creds())
        rows = sheets.spreadsheets().values().get(
            spreadsheetId=PRODUCT_MASTER_ID, range="'Data'").execute().get("values", [])
    except Exception as e:
        print(f"manufacturer lookup failed: {e}", file=sys.stderr)
        return None
    hdr = rows[0]
    i_sku, i_mf = hdr.index("SKU"), hdr.index("생산업체")
    found = []
    for sk in skus:
        for r in rows[1:]:
            if len(r) > i_mf and r[i_sku] == sk and r[i_mf].strip():
                if r[i_mf] not in found:
                    found.append(r[i_mf])
                break
    return " / ".join(found) or None


def _global_rating(p: dict, ov: dict) -> str:
    """Global 리뷰 평점 (user rule 2026-10-08): the star rating on the Amazon page the 제품명 cell links to
    (amazon.de/dp/<asin>), falling back to .com / .co.uk / .co.jp -> e.g. '4.5(DE)'; 'N/A' if none shows one.
    An explicit overview.global_rating other than '-'/'' is kept as given."""
    given = str(ov.get("global_rating") or "").strip()
    if given and given != "-":
        return given
    from amazon_rating import global_rating
    return global_rating(p.get("asin", ""))


def build_overview_slide(b: DeckBuilder, data: dict):
    sid = b.new_slide()
    p = data["product"]
    ov = data["overview"]
    b.add_textbox(sid, 27.6, 10, 565.6, 46, " 제품 클레임 개요\n",
                  font=GOTHIC, size=23, bold=True)

    product_links = {}
    # 제품명 link (and Global 리뷰 평점 source) must be one of this case's own ASINs — a claim-level ASIN from Zendesk
    # can point at another product (2026-10-08: Flip7 deck linked an S24 Ultra case, iPhone 16 deck an iPhone 15 Plus case)
    own = [x.get("asin", "") for x in p.get("skus", []) if x.get("asin")]
    if own and p.get("asin") not in own:
        p["asin"] = own[0]
    if p.get("asin"):
        product_links[(1, 1)] = f"https://www.amazon.de/dp/{p['asin']}"
    pairs = _sku_asin_pairs(p)
    labels = [f"{sk} / {a}" if a else sk for sk, a in pairs]
    # Multi-SKU line-ups (2026-10-02): one pair per line at 11pt in a 130pt column wrapped every pair
    # and pushed table 1 under table 2. Single SKU keeps the original layout; 2-6 pairs: 9pt one per
    # line; 7-12: 8pt two per line; >12: first 10 + "외 N종". Multi-SKU decks use a wider SKU column
    # and content-height tables 1/2 so table 3 still fits on the page.
    if len(labels) <= 1:
        sku_size, sku_lines = 11, labels
    elif len(labels) <= 6:
        sku_size, sku_lines = 9, labels
    else:
        shown = labels if len(labels) <= 12 else labels[:10]
        sku_size = 8
        sku_lines = [" · ".join(x.replace(" / ", "/") for x in shown[i:i + 2]) for i in range(0, len(shown), 2)]
        if len(labels) > len(shown):
            sku_lines.append(f"외 {len(labels) - len(shown)}종")
    sku_asin = "\n".join(sku_lines) or "-"
    manufacturer = p.get("manufacturer")
    if not manufacturer or manufacturer == "-":
        manufacturer = lookup_manufacturer([sk for sk, _ in pairs]) or "-"
    compact = len(labels) > 1
    widths = [95, 140, 95, 220, 100] if compact else [130] * 5  # sums to 650
    # rendered line height ≈ 1.5 × font size (Noto Sans) + ~14pt cell padding
    body_h = max([_est_lines(v, w, 11) * 16.5 for v, w in zip(
        [p.get("device", ""), p.get("product_name", ""), p.get("defect_type", ""), manufacturer],
        [widths[0], widths[1], widths[2], widths[4]])] + [len(sku_lines) * sku_size * 1.5])
    t1_h = 30 + max(30, body_h + 14) if compact else 90
    t1 = b.add_table(sid, 35, 52, 650, t1_h, [
        ["기종", "제품명", "불량 유형", "SKU/ASIN", "제조 업체명"],
        [p.get("device", ""), p.get("product_name", ""), p.get("defect_type", ""),
         sku_asin, manufacturer],
    ], cell_links=product_links)
    for ci, w in enumerate(widths if compact else []):
        b.reqs.append({"updateTableColumnProperties": {
            "objectId": t1, "columnIndices": [ci],
            "tableColumnProperties": {"columnWidth": {"magnitude": w, "unit": "PT"}},
            "fields": "columnWidth"}})
    if sku_size != 11:
        b.reqs.append({"updateTextStyle": {
            "objectId": t1, "cellLocation": {"rowIndex": 1, "columnIndex": 3},
            "style": {"fontSize": {"magnitude": sku_size, "unit": "PT"}},
            "fields": "fontSize", "textRange": {"type": "ALL"}}})
    # original layout for 1-3 SKUs; compact (content-height tables 1/2) for multi-SKU line-ups
    t2_y, t2_h = (52 + t1_h + 8, 60) if compact else (147, 90)
    t3_y = t2_y + t2_h + 8 if compact else 252

    b.add_table(sid, 35, t2_y, 650, t2_h, [
        ["총 인입건 수", "전체 배드 리뷰 수", "전체 클레임 수(Zendesk)", "Global 리뷰 평점"],
        [str(ov.get("total_count", 0)), str(ov.get("bad_review_count", 0)),
         str(ov.get("zendesk_count", 0)), _global_rating(p, ov)],
    ], cell_fills={(0, 3): ACCENT_ORANGE})

    by_review = ov.get("by_country_bad_review", {})
    by_zendesk = ov.get("by_country_zendesk", {})
    countries = [c for c in COUNTRY_ORDER if by_review.get(c) or by_zendesk.get(c)] or COUNTRY_ORDER
    # Countries outside COUNTRY_ORDER (e.g. NL) are appended rather than silently dropped.
    countries += sorted(c for c in set(by_review) | set(by_zendesk) if c not in COUNTRY_ORDER and (by_review.get(c) or by_zendesk.get(c)))
    header = ["국가별 인입 채널"] + countries
    row_review = ["배드 리뷰 수"] + [str(by_review.get(c, "-") or "-") for c in countries]
    row_zendesk = ["고객 클레임 수\n(Zendesk)"] + [str(by_zendesk.get(c, "-") or "-") for c in countries]
    totals = []
    for c in countries:
        rv = by_review.get(c) or 0
        zv = by_zendesk.get(c) or 0
        totals.append(str(rv + zv) if (rv or zv) else "-")
    row_total = ["Total"] + totals
    t3 = b.add_table(sid, 35, t3_y, 650, 100, [header, row_review, row_zendesk, row_total])
    # wider label column so '국가별 인입 채널' / '고객 클레임 수 (Zendesk)' don't wrap into 3-4 lines
    # (with many countries the equal-width columns pushed table 3 off the page bottom)
    if len(countries) > 5:
        first = 120
        rest = (650 - first) / len(countries)
        for ci, w in enumerate([first] + [rest] * len(countries)):
            b.reqs.append({"updateTableColumnProperties": {
                "objectId": t3, "columnIndices": [ci],
                "tableColumnProperties": {"columnWidth": {"magnitude": w, "unit": "PT"}},
                "fields": "columnWidth"}})


def _est_lines(text, width_pt, size):
    """Rough wrapped-line count for a table cell (CJK ≈ 1em, Latin ≈ 0.62em, 14pt cell padding)."""
    import math
    n = 0
    for para in str(text).split("\n"):
        w = sum(size * (1.0 if ord(ch) > 0x2E80 else 0.62) for ch in para)
        n += max(1, math.ceil(w / max(1, width_pt - 14)))
    return n


AMAZON_DOMAINS = {"US": "amazon.com", "UK": "amazon.co.uk", "DE": "amazon.de", "IT": "amazon.it",
                  "FR": "amazon.fr", "ES": "amazon.es", "JP": "amazon.co.jp", "IN": "amazon.in",
                  "NL": "amazon.nl"}


def _claim_link_url(claim: dict):
    """URL for the 링크 cell (user rule 2026-09-28: always hyperlinked). Explicit link_url
    wins; otherwise a numeric link_label is a Zendesk ticket ID and an R-prefixed one is an
    Amazon Review ID on the claim's country marketplace."""
    if claim.get("link_url"):
        return claim["link_url"]
    label = str(claim.get("link_label", "")).strip()
    if label.isdigit():
        return f"https://spigenhelp.zendesk.com/agent/tickets/{label}"
    if label.startswith("R") and label[1:].isalnum():
        domain = AMAZON_DOMAINS.get(claim.get("country", ""), "amazon.com")
        return f"https://{domain}/gp/customer-reviews/{label}/ref_=bcr_shw_rev_dtl"
    return None


def build_claim_slide(b: DeckBuilder, idx: int, claim: dict):
    sid = b.new_slide()
    label = claim.get("label") or f"클레임 #{idx:02d}"
    tb = b.add_textbox(sid, 27.6, 8.9, 565.6, 42, f"{label} \n", font=GOTHIC, size=23, bold=True)
    # "#NN" in the title: blue for 클레임, red for 배드 리뷰 (user rule 2026-09-28).
    hash_at = label.find("#")
    if hash_at >= 0:
        b.reqs.append({"updateTextStyle": {
            "objectId": tb,
            "style": {"foregroundColor": {"opaqueColor": {"rgbColor":
                      RED if label.startswith("배드") else BLUE}}},
            "fields": "foregroundColor",
            "textRange": {"type": "FIXED_RANGE", "startIndex": hash_at, "endIndex": len(label)},
        }})

    images = claim.get("image_urls") or []
    videos = claim.get("video_drive_ids") or []
    is_review = label.startswith("배드")
    # 배드 리뷰 with no customer media: the bottom row becomes '배드 리뷰 원문' holding the
    # review's original text instead of an empty photo label (user rule 2026-09-28).
    show_original = is_review and not images and not videos and claim.get("original_text")
    if show_original:
        last_row = ["배드 리뷰 원문", claim["original_text"], None, None, None, None]
        last_merge = (3, 1, 1, 5)
    else:
        last_row = ["고객 첨부 사진/영상" if videos else "고객 첨부 사진", None, None, None, None, None]
        last_merge = (3, 0, 1, 6)  # label spans full row (media placed as floating elements below)
    rows = [
        ["구분", "국가", "구매일", "인입일", "주문후 결함 발생 기간", "링크"],
        [None, claim.get("country", "-"), claim.get("purchase_date", "-"),
         claim.get("inflow_date", "-"), claim.get("days_elapsed", "-"),
         claim.get("link_label", "클레임")],
        ["상세 내용", claim.get("detail_text", ""), None, None, None, None],
        last_row,
    ]
    link_url = _claim_link_url(claim)
    b.add_table(sid, 27.6, 56.3, 665, 150, rows, header_rows=1, merges=[
        (2, 1, 1, 5),  # 상세 내용 detail spans remaining 5 cols
        last_merge,
    ], cell_links={(1, 5): link_url} if link_url else None)

    media = [("img", u) for u in images] + [("vid", v) for v in videos]
    if media:
        # lay media out left-to-right below the table, kept inside the 405pt page height
        n = len(media)
        gap = 12
        avail_w = 665 - gap * (n - 1)
        img_w = min(200, avail_w / n)
        y = 215
        img_h = min(img_w * 1.3, PAGE_H_PT - y - 8)
        x = 27.6
        for kind, ref in media:
            if kind == "img":
                b.add_image(sid, ref, x, y, img_w, img_h)
            else:
                b.add_video(sid, ref, x, y, img_w, img_h)
            x += img_w + gap


def build_deck(data: dict) -> str:
    title = f"{data['issue_title']} SIREN 검토_글로벌CX전략팀_{data['date'].replace('.', '.')}"
    b = DeckBuilder(title)
    b.delete_default_slide()
    build_title_slide(b, data)
    build_overview_slide(b, data)
    # 클레임 and 배드 리뷰 are numbered separately, each from #01 (user rule 2026-09-28) —
    # the label's number is always recomputed here, whatever the input JSON says.
    counters = {"클레임": 0, "배드 리뷰": 0}
    for i, claim in enumerate(data.get("claims", []), start=1):
        kind = "배드 리뷰" if str(claim.get("label", "")).startswith("배드") else "클레임"
        counters[kind] += 1
        claim = dict(claim, label=f"{kind} #{counters[kind]:02d}")
        build_claim_slide(b, i, claim)
    b.flush()
    return f"https://docs.google.com/presentation/d/{b.pid}/edit"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="path to SIREN case JSON")
    ap.add_argument("--dry-run", action="store_true", help="print parsed data, don't create anything")
    a = ap.parse_args()

    with open(a.data, encoding="utf-8") as f:
        data = json.load(f)
    if "date" not in data:
        data["date"] = datetime.date.today().strftime("%Y. %m. %d")

    if a.dry_run:
        print(json.dumps(data, ensure_ascii=False, indent=1))
        return 0

    url = build_deck(data)
    print(f"SIREN deck created: {url}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
