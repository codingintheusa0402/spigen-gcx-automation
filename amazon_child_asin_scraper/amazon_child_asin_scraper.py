#!/usr/bin/env python3
"""
Amazon Child ASIN Rating & Review Count Scraper

Approach (per Amazon variation model):
  1. Parent ASIN = discovery only → extract child ASINs
  2. Each child ASIN → open /dp/{child_asin}, verify the active ASIN on page
     matches the queued child, then extract rating + review count
  3. Flag shared_variation_reviews=true when all children in a family return
     identical values (Amazon is still sharing the review pool)

Session: saved to ~/.amazon_cookies.json — log in once, reused on every run.
CSV: saved incrementally after every child page visit.
"""

import os
import re
import json
import time
import csv
import tempfile
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, WebDriverException

# ── Config ────────────────────────────────────────────────────────────────────

DOMAINS = {
    'US':  'https://www.amazon.com',
    'DE':  'https://www.amazon.de',
    'FR':  'https://www.amazon.fr',
    'IT':  'https://www.amazon.it',
    'ES':  'https://www.amazon.es',
    'UK':  'https://www.amazon.co.uk',
    'JP':  'https://www.amazon.co.jp',
    'IN':  'https://www.amazon.in',
}

PARENT_ASINS = [
    'B0G7QZ7RMD',
    'B0G7RYP439',
    'B0G7RJH146',
    'B0G6CFQBJJ',
    'B0G6CQ5JPH',
    'B0G7RD3G98',
]

PAGE_LOAD_WAIT  = 3
CHILD_PAGE_WAIT = 2
LOGIN_WAIT      = 60

COOKIES_FILE = os.path.expanduser('~/.amazon_cookies.json')

# ── Driver ────────────────────────────────────────────────────────────────────

def make_driver():
    tmp = tempfile.mkdtemp(prefix='chrome_amz_')
    opts = Options()
    opts.add_argument(f'--user-data-dir={tmp}')
    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--disable-extensions')
    opts.add_argument('--no-first-run')
    opts.add_argument('--no-default-browser-check')
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_argument('--lang=en-US,en;q=0.9')
    opts.add_argument(
        'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    )
    opts.add_experimental_option('excludeSwitches', ['enable-automation'])
    opts.add_experimental_option('useAutomationExtension', False)
    d = webdriver.Chrome(options=opts)
    d.execute_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
    try:
        d.set_window_size(1400, 900)
    except Exception:
        pass
    return d

def restart_driver(driver):
    try:
        driver.quit()
    except Exception:
        pass
    d = make_driver()
    load_cookies(d, 'US', DOMAINS['US'])
    return d


# ── Cookie persistence ────────────────────────────────────────────────────────

def save_cookies(driver, domain_key):
    try:
        all_c = {}
        if os.path.exists(COOKIES_FILE):
            with open(COOKIES_FILE) as f:
                all_c = json.load(f)
        all_c[domain_key] = driver.get_cookies()
        with open(COOKIES_FILE, 'w') as f:
            json.dump(all_c, f)
    except Exception as e:
        print(f"  Warning: could not save cookies: {e}")

def load_cookies(driver, domain_key, domain_url):
    try:
        if not os.path.exists(COOKIES_FILE):
            return False
        with open(COOKIES_FILE) as f:
            all_c = json.load(f)
        if domain_key not in all_c or not all_c[domain_key]:
            return False
        driver.get(domain_url)
        time.sleep(1)
        for c in all_c[domain_key]:
            try:
                driver.add_cookie(c)
            except Exception:
                pass
        driver.refresh()
        time.sleep(2)
        return True
    except Exception:
        return False

def is_logged_in(driver):
    try:
        url = driver.current_url or ''
        if 'signin' in url or 'ap/signin' in url:
            return False
        el = driver.find_element(By.CSS_SELECTOR,
            '#nav-link-accountList-nav-line-1')
        txt = (el.text or '').strip().lower()
        return bool(txt) and txt != 'sign in'
    except Exception:
        return False

def ensure_logged_in(driver):
    print(f"\n  Checking Amazon login...")
    loaded = load_cookies(driver, 'US', DOMAINS['US'])
    if loaded and is_logged_in(driver):
        print(f"  Session restored from {COOKIES_FILE}")
        save_cookies(driver, 'US')
        return
    driver.get(DOMAINS['US'])
    time.sleep(2)
    print(f"  NOT LOGGED IN — log in to Amazon in the browser.")
    print(f"  Waiting up to {LOGIN_WAIT} seconds...")
    for i in range(LOGIN_WAIT):
        time.sleep(1)
        if is_logged_in(driver):
            print(f"  Logged in! Saving session...")
            save_cookies(driver, 'US')
            return
        if (i + 1) % 10 == 0:
            print(f"  ...{LOGIN_WAIT - i - 1}s remaining")
    print(f"  Warning: login not confirmed — proceeding anyway")


# ── Popup dismissal ───────────────────────────────────────────────────────────

def dismiss_popups(driver):
    for text in ['Accept', 'Accept All', 'Akzeptieren', 'Alle akzeptieren',
                 'Accept Cookies', 'Accept all', '同意する']:
        try:
            btn = driver.find_element(
                By.XPATH,
                f"//*[self::button or self::input]"
                f"[contains(translate(normalize-space(.),"
                f"'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{text.lower()}')]"
            )
            btn.click()
            time.sleep(0.8)
            return
        except Exception:
            pass


# ── Active ASIN verification ──────────────────────────────────────────────────

def get_active_asin(driver):
    """
    Determine which ASIN is currently selected/active on the loaded page.
    Checks canonical link, page-source ASIN markers, and current URL.
    Returns the ASIN string or '' if undetermined.
    """
    # 1. Canonical <link> — most reliable
    try:
        href = driver.find_element(
            By.CSS_SELECTOR, 'link[rel="canonical"]'
        ).get_attribute('href') or ''
        m = re.search(r'/dp/([A-Z0-9]{10})', href)
        if m:
            return m.group(1)
    except Exception:
        pass

    # 2. Page source ASIN field (JSON embedded in page)
    try:
        src = driver.page_source
        m = re.search(r'"ASIN"\s*:\s*"([A-Z0-9]{10})"', src)
        if m:
            return m.group(1)
    except Exception:
        pass

    # 3. Current URL
    try:
        url = driver.current_url or ''
        m = re.search(r'/dp/([A-Z0-9]{10})', url)
        if m:
            return m.group(1)
    except Exception:
        pass

    return ''


# ── Rating / review extraction ────────────────────────────────────────────────

RATING_SELECTORS = [
    'span[data-hook="rating-out-of-text"]',
    '#averageCustomerReviews span.a-icon-alt',
    'i[data-hook="average-star-rating"] span.a-icon-alt',
]

REVIEW_COUNT_SELECTORS = [
    '#acrCustomerReviewText',
    'span[data-hook="total-review-count"]',
    '#acrCustomerReviewLink span',
]

def _text(el):
    if el is None:
        return ''
    try:
        return (el.text or el.get_attribute('innerHTML') or '').strip()
    except Exception:
        return ''

def extract_rating(text):
    m = re.search(r'(\d[.,]\d)', text)
    return m.group(1).replace(',', '.') if m else ''

def extract_count(text):
    # "(14,443)" or "14,443 global ratings"
    m = re.search(r'([\d][,\d\.]+\d|\d+)', text)
    return re.sub(r'[^\d]', '', m.group(1)) if m else ''

def scrape_stats(driver):
    rating = ''
    for sel in RATING_SELECTORS:
        try:
            r = extract_rating(_text(driver.find_element(By.CSS_SELECTOR, sel)))
            if r:
                rating = r
                break
        except Exception:
            pass

    review_count = ''
    for sel in REVIEW_COUNT_SELECTORS:
        try:
            c = extract_count(_text(driver.find_element(By.CSS_SELECTOR, sel)))
            if c:
                review_count = c
                break
        except Exception:
            pass

    return rating, review_count


# ── Child ASIN discovery (parent page = discovery only) ───────────────────────

def discover_children(driver, parent_asin, domain_url):
    """
    Open parent /dp/ page to extract child ASIN list.
    Parent page is used for discovery ONLY — not for rating/review extraction.
    """
    driver.get(f"{domain_url}/dp/{parent_asin}")
    time.sleep(PAGE_LOAD_WAIT)
    dismiss_popups(driver)

    children = set()

    try:
        for item in driver.find_elements(By.CSS_SELECTOR, 'li[data-asin]'):
            val = item.get_attribute('data-asin')
            if val and re.match(r'^[A-Z0-9]{10}$', val):
                children.add(val)
    except Exception:
        pass

    try:
        for item in driver.find_elements(
            By.CSS_SELECTOR,
            '#twister [data-asin], #variation_style_name [data-asin], '
            '#variation_color_name [data-asin], #variation_size_name [data-asin]'
        ):
            val = item.get_attribute('data-asin')
            if val and re.match(r'^[A-Z0-9]{10}$', val):
                children.add(val)
    except Exception:
        pass

    try:
        for script in driver.find_elements(By.TAG_NAME, 'script'):
            src = script.get_attribute('innerHTML') or ''
            if 'variationValues' in src or 'dimensionValuesDisplayData' in src:
                for f in re.findall(r'"([A-Z0-9]{10})"', src):
                    children.add(f)
    except Exception:
        pass

    children.add(parent_asin)
    children = {a for a in children if re.match(r'^[A-Z0-9]{10}$', a)}
    return sorted(children)


# ── Per-child extraction (child page = extraction target) ─────────────────────

def scrape_child(driver, child_asin, domain_url):
    """
    Open /dp/{child_asin} with child variation selected.
    Verify the active ASIN on page matches the queued child.
    Extract rating + review count only after confirmation.
    Returns (rating, review_count, active_asin, verified)
    """
    driver.get(f"{domain_url}/dp/{child_asin}")
    time.sleep(CHILD_PAGE_WAIT)
    dismiss_popups(driver)

    active_asin = get_active_asin(driver)
    verified = (active_asin == child_asin)

    rating, review_count = scrape_stats(driver)
    return rating, review_count, active_asin, verified


# ── Shared-review detection ───────────────────────────────────────────────────

def flag_shared(rows):
    """
    Given a list of row dicts for the same parent+domain,
    set shared_variation_reviews=true if all children returned identical
    non-empty rating AND review_count values.
    """
    counts  = [r['child_review_count'] for r in rows if r['child_review_count']]
    ratings = [r['child_rating']       for r in rows if r['child_rating']]
    shared  = (
        len(rows) > 1
        and len(counts)  == len(rows)
        and len(set(counts))  == 1
        and len(ratings) == len(rows)
        and len(set(ratings)) == 1
    )
    for r in rows:
        r['shared_variation_reviews'] = 'true' if shared else 'false'
    return rows


# ── CSV ───────────────────────────────────────────────────────────────────────

CSV_FIELDS = [
    'domain', 'parent_asin', 'child_asin', 'is_parent',
    'child_rating', 'child_review_count',
    'active_asin', 'asin_verified',
    'shared_variation_reviews',
    'scraped_at',
]

def open_csv(path):
    f = open(path, 'w', newline='', encoding='utf-8')
    w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
    w.writeheader()
    f.flush()
    return f, w

def write_rows(f, w, rows):
    for row in rows:
        w.writerow(row)
    f.flush()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    driver = make_driver()

    ts  = datetime.now().strftime('%Y%m%d_%H%M%S')
    out = f"/Users/kevinkim/Desktop/asin_reviews_{ts}.csv"
    csv_file, csv_writer = open_csv(out)
    print(f"\n  CSV: {out}")
    print(f"  Strategy: parent ASIN = discovery only | child /dp/ page = extraction target")

    row_count = 0

    try:
        ensure_logged_in(driver)

        total = len(DOMAINS) * len(PARENT_ASINS)
        done  = 0

        for domain_key, domain_url in DOMAINS.items():
            for parent_asin in PARENT_ASINS:
                done += 1
                print(f"\n[{done}/{total}] {parent_asin}  [{domain_key}]")

                # Step 1: discover children from parent page
                try:
                    children = discover_children(driver, parent_asin, domain_url)
                except WebDriverException:
                    print(f"  Browser crashed — restarting...")
                    driver = restart_driver(driver)
                    try:
                        children = discover_children(driver, parent_asin, domain_url)
                    except WebDriverException as e2:
                        print(f"  Retry failed — skipping: {e2.msg[:80]}")
                        continue

                print(f"  → {len(children)} child ASIN(s): {children}")

                # Step 2: scrape each child individually
                pending_rows = []
                for i, child_asin in enumerate(children):
                    try:
                        rating, review_count, active_asin, verified = scrape_child(
                            driver, child_asin, domain_url
                        )
                    except WebDriverException:
                        print(f"  [{i+1}] {child_asin}  Browser crashed — restarting...")
                        driver = restart_driver(driver)
                        try:
                            rating, review_count, active_asin, verified = scrape_child(
                                driver, child_asin, domain_url
                            )
                        except WebDriverException as e2:
                            print(f"  [{i+1}] {child_asin}  Retry failed: {e2.msg[:60]}")
                            rating, review_count, active_asin, verified = '', '', '', False

                    is_parent = child_asin == parent_asin
                    verified_str = 'YES' if verified else f'NO (got {active_asin or "?"})'
                    print(f"  [{i+1}/{len(children)}] {child_asin}"
                          f"{'  (parent)' if is_parent else ''}"
                          f"  rating={rating or '-'}  reviews={review_count or '-'}"
                          f"  verified={verified_str}")

                    pending_rows.append({
                        'domain':               domain_key,
                        'parent_asin':          parent_asin,
                        'child_asin':           child_asin,
                        'is_parent':            'YES' if is_parent else 'no',
                        'child_rating':         rating,
                        'child_review_count':   review_count,
                        'active_asin':          active_asin,
                        'asin_verified':        'YES' if verified else 'NO',
                        'shared_variation_reviews': '',   # filled below
                        'scraped_at':           datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    })

                # Step 3: flag shared reviews, then write all rows for this parent/domain
                pending_rows = flag_shared(pending_rows)
                if pending_rows and pending_rows[0]['shared_variation_reviews'] == 'true':
                    print(f"  ⚑  shared_variation_reviews = true (all children identical)")
                write_rows(csv_file, csv_writer, pending_rows)
                row_count += len(pending_rows)

    except KeyboardInterrupt:
        print("\n\nStopped by user.")
    finally:
        csv_file.close()
        print(f"\n  Done — {row_count} rows → {out}")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == '__main__':
    main()
