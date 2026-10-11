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
- **Max concurrency** — default 40 parallel sessions, max 50 (auto-capped to the run's memory: ~1 session per 12 MB)
- **Proxy** — residential proxies recommended (default)

## Use cases
Track your own and competitors' ratings over time, monitor rating drops after a product change, build catalog dashboards, schedule daily/weekly snapshots.

## Speed
About 3,400 ASINs in ~5.5 minutes at the default 512 MB–1 GB. Extra memory beyond 1 GB doesn't add speed.

## How it works
Each parallel session behaves like one real visitor: its own sticky proxy IP, cookie jar and real-browser TLS fingerprint, reused for up to 40 pages with human-like pacing, then rotated. Starts on Apify datacenter proxies and escalates to residential only for ASINs that keep getting blocked; passes Amazon's "continue shopping" interstitial; partial/degraded pages are retried instead of returned. Workers back off automatically when Amazon pushes back. You're only charged for ASINs that return data.
