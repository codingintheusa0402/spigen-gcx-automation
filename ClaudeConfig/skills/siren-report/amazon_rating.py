#!/usr/bin/env python3
"""Global 리뷰 평점 lookup for the SIREN 제품 클레임 개요 slide (user rule 2026-10-08).

The value comes from the Amazon product page the 제품명 cell hyperlinks to (amazon.de/dp/<asin>):
the star average shown next to the title (#acrPopover). If amazon.de has no rating (out of stock,
not sold there, page blocked), fall back to amazon.com -> amazon.co.uk -> amazon.co.jp.
Returns e.g. "4.5(DE)", or "N/A" when no marketplace shows a rating.

Usage: python3 amazon_rating.py B0DJ9VF2M6 [more ASINs...]
"""
import re
import sys
import time
import urllib.request

MARKETS = [("DE", "amazon.de"), ("US", "amazon.com"), ("UK", "amazon.co.uk"), ("JP", "amazon.co.jp")]
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/128.0 Safari/537.36")
# The product's own rating sits on the #acrPopover element; other "x out of 5 stars" on the page
# belong to sponsored/related products, so only this element is trusted.
RX = re.compile(r'id="acrPopover"[^>]*?title="([0-9]+[.,][0-9]) (?:out of 5|von 5|5つ星のうち)', re.S)
RX_JP = re.compile(r'id="acrPopover"[^>]*?title="5つ星のうち([0-9]+[.,][0-9])', re.S)


def _fetch(url, tries=3):
    """Page HTML, or '' if Amazon only returns its short bot/captcha page every time."""
    for k in range(tries):
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9",
                                                   "Accept": "text/html,application/xhtml+xml"})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                html = r.read().decode("utf-8", "replace")
        except Exception:
            html = ""
        if len(html) > 20000 and "captcha" not in html[:5000].lower():
            return html
        time.sleep(2 + 3 * k)
    return ""


def rating_on(domain, asin):
    """Star average on https://www.<domain>/dp/<asin>, or None."""
    html = _fetch(f"https://www.{domain}/dp/{asin}")
    if not html or asin not in html:          # redirected to another product / blocked
        return None
    m = RX.search(html) or RX_JP.search(html)
    return m.group(1).replace(",", ".") if m else None


def global_rating(asin):
    """'4.5(DE)' from the first marketplace (DE, US, UK, JP) that shows a rating, else 'N/A'."""
    if not asin:
        return "N/A"
    for cc, domain in MARKETS:
        r = rating_on(domain, asin)
        if r:
            return f"{r}({cc})"
    return "N/A"


if __name__ == "__main__":
    for a in sys.argv[1:]:
        print(a, global_rating(a))
