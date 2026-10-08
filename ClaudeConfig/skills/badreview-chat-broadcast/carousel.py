#!/usr/bin/env python3
"""Shared carousel-message builder for the combined BadReview broadcast (2026-09-21).

Produces ONE Chat message with all 3 products' cards as swipeable carousel pages —
the default broadcast format as of 2026-09-21, replacing the old 3-separate-message-
per-room format. Used by both `broadcast.py` (interactive skill) and
`auto_broadcast.py` (unattended weekday job) so the layout logic lives in exactly
one place.

Hard-won rendering facts (empirically verified against a live webhook, 2026-09-21 —
official Cards v2 docs don't cover any of this, four iterations to get right):
  - `carousel`/`carouselCards` IS a real, working widget — genuine swipeable
    multi-page UI over a plain incoming webhook, no interactive backend needed.
  - A carousel page has NO native header (no title/image at the page level) —
    only a flat `widgets` list.
  - `decoratedText` renders BLANK as a direct child of a carousel page — ALL of
    it, not just custom-iconUrl cases (confirmed by inspecting the exact JSON that
    was sent: every field was populated correctly, yet nothing showed). Every
    decoratedText must be converted to an equivalent `textParagraph` line instead
    (same bold/color HTML, topLabel/bottomLabel folded into the line).
  - The plain `image` widget DOES render, but has NO size/crop control (only
    imageUrl/onClick/altText per the official schema) — unlike `header.imageType:
    SQUARE` in a normal (non-carousel) card, which auto-crops for free. Since the
    3 products' source images come from different CDNs with different native
    aspect ratios (Google shopping tbn thumbnails vs a Spigen/Shopify product
    photo), each image is proxied through wsrv.nl (`?url=...&w=W&h=H&fit=contain`)
    to force a matching square BEFORE handing the URL to the `image` widget — there
    is no Chat-side control that can do this. `fit=contain` (not `cover`): a center
    crop cut off the sides of the phone on Z8/Pixel11's differently-shaped source
    images (2026-10-08 user report) — `contain` scales+letterboxes instead, never
    cropping the product itself.
  - `columns` support inside a carousel page is unconfirmed (Google's only
    official example uses flat textParagraphs) — flattened into a stacked
    sequential list instead, preserving each category's own bold red/blue header
    + rank rows.
"""
import re
import urllib.parse

KOR_WD = ['월', '화', '수', '목', '금', '토', '일']

# Fixed carousel page order (2026-09-21, explicit user instruction): iPhone 18 first,
# then Galaxy Z8, then Pixel 11 — matches the header subtitle text order below.
PAGE_ORDER = ["iphone18", "glxz8", "pixel11"]


def square_img(url, size=300):
    """Proxy any image URL through wsrv.nl to force a WxH square, regardless of
    source CDN/native aspect ratio. `fit=contain` (not `cover`): the 3 products'
    source images differ enough in aspect ratio (gstatic Z8/Pixel11 thumbnails vs
    the Spigen iPhone18 product photo) that a center-crop cut off the sides of the
    phone on Z8/Pixel11 (2026-10-08 user report, confirmed via screenshots) while
    only iPhone18 happened to still look fine. `contain` scales the whole image to
    fit inside the square and letterboxes the rest — verified against all 3 source
    URLs that the padding blends with their white/light product-shot backgrounds."""
    enc = urllib.parse.quote(url, safe="")
    return f"https://wsrv.nl/?url={enc}&w={size}&h={size}&fit=contain&output=jpg"


def _short_page_title(title):
    """Each product's own card title is '✔️ M/D(요일) <Product> 배드리뷰 (1~3점)
    (총 N건)' — the ✔️/date/(1~3점) parts are redundant once folded into a carousel
    page, since the outer card header already shows '✔️ M/D(요일) 배드리뷰 (1~3점)'.
    Strips down to '<Product> 배드리뷰 (총 N건)'."""
    t = re.sub(r'^✔️\s*\S+\([^)]+\)\s*', '', title)
    return t.replace(' (1~3점)', '')


def _dt_to_line(dt):
    top = (dt.get("topLabel") or "").strip()
    text = (dt.get("text") or "").strip()
    bottom = (dt.get("bottomLabel") or "").strip()
    if not top and not text and not bottom:
        return None
    line = " — ".join(b for b in (top, text) if b)
    if bottom:
        line += f' &nbsp;<font color="#5f6368">({bottom})</font>'
    return line


def _flatten(widgets):
    """textParagraph/decoratedText/columns/... -> carousel-safe textParagraph /
    image / buttonList only. Consecutive decoratedText widgets (including ones
    inside `columns`) are merged into one <br>-joined textParagraph block each."""
    out, dt_buf = [], []

    def flush():
        if dt_buf:
            lines = [l for l in (_dt_to_line(d) for d in dt_buf) if l]
            if lines:
                out.append({"textParagraph": {"text": "<br>".join(lines)}})
            dt_buf.clear()

    for w in widgets:
        if "decoratedText" in w:
            dt_buf.append(w["decoratedText"])
        elif "columns" in w:
            for col in w["columns"]["columnItems"]:
                for cw in col["widgets"]:
                    if "decoratedText" in cw:
                        dt_buf.append(cw["decoratedText"])
                    else:
                        flush()
                        out.append(cw)
        else:
            flush()
            out.append(w)
    flush()
    return out


def make_page(card_payload, squared_img_url):
    """One product's normal build_card() output -> one carousel page."""
    c = card_payload["cardsV2"][0]["card"]
    h = c["header"]
    widgets = [
        {"image": {"imageUrl": squared_img_url, "altText": h["title"]}},
        {"textParagraph": {"text": f'<b>{_short_page_title(h["title"])}</b><br>{h.get("subtitle", "")}'}},
    ]
    for section in c["sections"]:
        if section.get("header"):
            widgets.append({"textParagraph": {"text": f'<b>{section["header"]}</b>'}})
        widgets.extend(_flatten(section["widgets"]))
    return {"widgets": widgets}


def build_carousel_message(cards, header_imgs, today, card_id="badreview-carousel"):
    """cards / header_imgs: {"pixel11": ..., "glxz8": ..., "iphone18": ...} — cards
    are build_card() outputs, header_imgs are each product's raw HEADER_IMG URL
    (squared here, callers pass the un-proxied source URL). Returns the full Chat
    message dict: one message, all 3 products as swipeable pages, in PAGE_ORDER."""
    date_str = f"{today.month}/{today.day}({KOR_WD[today.weekday()]})"
    pages = [make_page(cards[k], square_img(header_imgs[k])) for k in PAGE_ORDER]
    return {"cardsV2": [{
        "cardId": card_id,
        "card": {
            "header": {
                "title": f"✔️ {date_str} 배드리뷰 (1~3점)",
                "subtitle": "iPhone 18·Galaxy Z8·Pixel 11",
            },
            "sections": [{"widgets": [{"carousel": {"carouselCards": pages}}]}],
        },
    }]}
