# Amazon ASIN Rating & Review Count Scraper

Get the **star rating** and **total ratings count** for any list of Amazon ASINs — fast, cheap, no login needed.

## What you get
| Field | Example |
|---|---|
| `asin` | `B0G7RYP439` |
| `title` | Spigen GlasTR EZ Fit Privacy Tempered Glass Screen Protector for Galaxy S26 |
| `rating` | `4.2` |
| `ratingsCount` | `2948` |
| `marketplace` | `amazon.com` |
| `url`, `scrapedAt` | product link, UTC timestamp |

`found: false` = the ASIN doesn't exist / is delisted. `found: null` + `error` = Amazon blocked every retry (not charged).

## Input
- **ASINs** — plain ASINs or product URLs (duplicates removed)
- **Marketplace** — `com`, `co.uk`, `de`, `fr`, `it`, `es`, `co.jp`, `in`, `ca`, `com.au`, … (19 domains)
- **Max concurrency** — default 5
- **Proxy** — residential proxies recommended (default)

## Use cases
Track your own and competitors' ratings over time, monitor rating drops after a product change, build catalog dashboards, schedule daily/weekly snapshots.

## How it works
Fetches each product page through Apify datacenter proxies (auto-escalating to residential for blocked ASINs) with a real-browser TLS fingerprint, passes Amazon's "continue shopping" interstitial, and retries blocked requests on a fresh IP (up to 6 times).
