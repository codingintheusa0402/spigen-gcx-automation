# amazon_child_asin_scraper.py

Selenium-based scraper that resolves Amazon parent ASINs into their child variants and extracts per-child rating and review count. Detects when Amazon is sharing a review pool across all children (shared variation reviews).

## Why child ASINs?

Amazon product pages use a "parent ASIN" as a grouping mechanism. The parent page often shows blended or shared review counts. This scraper:

1. Visits the parent `/dp/` page — **discovery only** — to extract the list of child ASINs
2. Visits each child `/dp/` page individually and verifies the active ASIN on the page matches the queued child
3. Extracts rating + review count per child
4. Flags `shared_variation_reviews=true` when all children in a family return identical values

## Output

```
asin_reviews_YYYYMMDD_HHMMSS.csv
```

| Column | Description |
|--------|-------------|
| `domain` | Country code (US/UK/DE/…) |
| `parent_asin` | Input parent ASIN |
| `child_asin` | Discovered child ASIN |
| `is_parent` | `YES` if this child is also the parent |
| `child_rating` | Star rating for this child |
| `child_review_count` | Number of ratings for this child |
| `active_asin` | ASIN detected as active on page (verification) |
| `asin_verified` | `YES` if active_asin == child_asin |
| `shared_variation_reviews` | `true` if all children have identical rating+count |
| `scraped_at` | Timestamp |

## Usage

Edit `PARENT_ASINS` in the script, then:

```bash
python3 amazon_child_asin_scraper.py
```

A Chrome window opens. If not already logged into Amazon, the script waits up to 60 seconds for manual login, then saves the session to `~/.amazon_cookies.json` for reuse.

## Session persistence

Cookies are saved to `~/.amazon_cookies.json`. On subsequent runs the session is restored automatically — manual login is only required when the saved session expires.

## Configuration

| Constant | Default | Description |
|----------|---------|-------------|
| `PAGE_LOAD_WAIT` | `3 s` | Wait after loading parent page |
| `CHILD_PAGE_WAIT` | `2 s` | Wait after loading each child page |
| `LOGIN_WAIT` | `60 s` | Max seconds to wait for manual login |
| `COOKIES_FILE` | `~/.amazon_cookies.json` | Session cookie storage |

## Dependencies

```
selenium
chromedriver (must match installed Chrome version)
```

Install:
```bash
pip install selenium
```

Note: Uses headful Chrome (visible window) to support manual login and cookie persistence. This is intentional — unlike `amazon_dp_scraper.py`, this script is designed for interactive use.
