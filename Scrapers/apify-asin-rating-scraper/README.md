# Amazon ASIN Rating & Review Count Scraper

**Paste a list of Amazon ASINs or product links → get each product's star rating and total ratings count in a clean table.** 19 Amazon marketplaces, no Amazon account or login, ~3,400 products in about 5 minutes, and you pay only for products that return data.

> 💡 **$2 per 1,000 products.** The Apify free plan's monthly credit covers roughly 2,000 ASINs, so you can try it on your real catalog for free.

---

## ✅ What does it do?

For every ASIN or Amazon product URL you give it, the scraper opens the public product page and returns:

| Field | Example | Meaning |
|---|---|---|
| `asin` | `B0G7RYP439` | The product's ASIN |
| `title` | Spigen GlasTR EZ Fit Privacy Tempered Glass Screen Protector for Galaxy S26 | Product title as shown on Amazon |
| `rating` | `4.2` | Average star rating (out of 5) |
| `ratingsCount` | `2958` | Total number of global ratings |
| `marketplace` | `amazon.com` | Which Amazon store the data came from |
| `url` | `https://www.amazon.com/dp/B0G7RYP439` | Product link |
| `found` | `true` | `true` = product page found · `false` = ASIN not sold on that marketplace · `null` = blocked (see `error`) |
| `scrapedAt` | `2026-10-16T01:03:12Z` | When it was scraped (UTC) |

## 🎯 Who uses it?

- **Amazon sellers & brands**: watch your own listings' ratings every day and catch a drop the morning it happens
- **Competitor tracking**: monitor rival products' review growth and rating changes
- **Agencies**: produce rating and review-count reports for client catalogs across countries
- **Analysts & data teams**: feed ratings into dashboards, spreadsheets, BI tools or pricing models
- **Product launches**: see how fast a new product collects reviews in each marketplace

## 🚀 How to use it (3 steps)

1. **Add your products** to the **ASINs or product URLs** field: one per line, or paste a whole column from a spreadsheet.
   - Bare ASIN, e.g. `B0G7RYP439`, uses the **Marketplace** you pick (default amazon.com)
   - Product URL, e.g. `https://www.amazon.de/dp/B0GDHRZFWH`, is scraped on **its own marketplace**, so one run can cover the US, Germany, Japan and more
2. Click **Start**. Leave everything else at the defaults.
3. **Download the results** as Excel, CSV, JSON or HTML from the **Output** tab, or pull them via API.

Duplicates are removed automatically, and anything that isn't an ASIN or Amazon link (blank cells, `-`, `TBD`) is skipped.

## 📥 Input example

```json
{
  "asins": [
    "B0G7RYP439",
    "B0G7QZ7RMD",
    "https://www.amazon.de/dp/B0GDHRZFWH",
    "https://www.amazon.co.jp/dp/B0FD1YH78G"
  ],
  "marketplace": "com"
}
```

## 📤 Output example

```json
[
  {
    "asin": "B0G7RYP439",
    "found": true,
    "url": "https://www.amazon.com/dp/B0G7RYP439",
    "title": "Spigen GlasTR EZ Fit Privacy Tempered Glass Screen Protector for Galaxy S26",
    "rating": 4.2,
    "ratingsCount": 2958,
    "marketplace": "amazon.com",
    "scrapedAt": "2026-10-16T01:03:12.623131+00:00"
  },
  {
    "asin": "B0GDHRZFWH",
    "found": true,
    "url": "https://www.amazon.de/dp/B0GDHRZFWH",
    "title": "Spigen Thin Fit MagFit for Samsung Galaxy S26 Plus - Black",
    "rating": 3,
    "ratingsCount": 245,
    "marketplace": "amazon.de",
    "scrapedAt": "2026-10-16T01:03:20.026163+00:00"
  },
  {
    "asin": "B0B7MFZ2SH",
    "found": false,
    "marketplace": "amazon.com"
  }
]
```

A product with no reviews yet comes back as `"found": true, "rating": null, "ratingsCount": 0`.

## 💰 How much does it cost?

Pay-per-event pricing: you pay for results, not for compute time.

| Event | Price |
|---|---|
| Product scraped (rating + count returned) | **$0.002** |
| Run start | $0.005 per run |
| ASIN not found on that marketplace | free |
| Blocked after all retries | free |

**Examples:** 100 products ≈ $0.21 · 1,000 products ≈ $2.01 · 10,000 products ≈ $20.01, and a daily check of 500 products comes to about $30 a month.

## ⚡ Speed & reliability

- About **3,400 ASINs in ~5.5 minutes** with default settings
- Each parallel session behaves like one real shopper: its own proxy IP, cookies and real-browser fingerprint, human-like pacing, and rotation after a few dozen pages
- Automatically gets past Amazon's "Continue shopping" check page, retries incomplete pages, and slows down if Amazon pushes back
- Starts on fast datacenter proxies and switches to residential proxies only for products that keep getting blocked
- Checked daily with real runs on amazon.com and amazon.de

## 🌍 Supported marketplaces

amazon.com · amazon.co.uk · amazon.de · amazon.fr · amazon.it · amazon.es · amazon.nl · amazon.se · amazon.pl · amazon.com.tr · amazon.co.jp · amazon.in · amazon.ca · amazon.com.mx · amazon.com.br · amazon.com.au · amazon.ae · amazon.sa · amazon.sg

## 🔁 Track ratings over time

1. Save your product list as a **Task**: click **Save as a new task** after entering the input.
2. Add a **Schedule** (for example, every weekday at 08:00) under *Schedules → Create new*.
3. Connect the output to where you work: **Google Sheets**, **Slack**, **email**, **Zapier**, **Make** or a **webhook** (all under *Integrations*).

Each run's results carry a `scrapedAt` timestamp, so stacking daily runs gives you a rating and review history.

## 🧑‍💻 Use it from code

**Python**
```python
from apify_client import ApifyClient

client = ApifyClient("<YOUR_APIFY_TOKEN>")
run = client.actor("cosmic_meteorite/amazon-asin-rating-scraper").call(
    run_input={"asins": ["B0G7RYP439", "https://www.amazon.de/dp/B0GDHRZFWH"]}
)
for item in client.dataset(run["defaultDatasetId"]).iterate_items():
    print(item["asin"], item.get("rating"), item.get("ratingsCount"))
```

**JavaScript**
```js
import { ApifyClient } from 'apify-client';

const client = new ApifyClient({ token: '<YOUR_APIFY_TOKEN>' });
const run = await client.actor('cosmic_meteorite/amazon-asin-rating-scraper').call({
    asins: ['B0G7RYP439', 'https://www.amazon.de/dp/B0GDHRZFWH'],
});
const { items } = await client.dataset(run.defaultDatasetId).listItems();
console.table(items.map(({ asin, rating, ratingsCount }) => ({ asin, rating, ratingsCount })));
```

**cURL** (run and get the results in one call)
```bash
curl -X POST "https://api.apify.com/v2/acts/cosmic_meteorite~amazon-asin-rating-scraper/run-sync-get-dataset-items?token=<YOUR_APIFY_TOKEN>" \
  -H "Content-Type: application/json" \
  -d '{"asins": ["B0G7RYP439", "B0G7QZ7RMD"]}'
```

## ⚙️ Advanced settings

| Setting | Default | When to change it |
|---|---|---|
| Marketplace | `com` | Your bare ASINs belong to another store (e.g. `de`, `co.jp`) |
| Parallel sessions | 40 | Lower it if you want gentler traffic. It's capped automatically to fit the run's memory |
| Proxy | Apify datacenter, auto-escalating to residential | Force a specific proxy group or your own proxies |
| Save HTML (debug) | off | Turn on to inspect the raw pages when something looks wrong |

Memory: 512 MB (default) handles any list size. More memory doesn't make it faster.

## ❓ FAQ

**Why does a product show `found: false`?** The ASIN isn't sold on that marketplace, or it has been delisted. Check that you used the right marketplace, or pass the full product URL.

**Why are several color variants showing the same numbers?** Amazon pools reviews across variations of the same product, so every variant shows the shared rating and count. That's what customers see too.

**Do I need an Amazon account, cookies or API keys?** No. Only public product pages are read.

**Does it return the review texts?** No. This actor is built to be the fastest and cheapest way to get ratings and review counts. Use a review scraper if you need the review texts.

**Is scraping Amazon legal?** This actor only reads publicly available data and collects no personal data. You're responsible for how you use the results. Check Amazon's terms and the laws that apply to you.

## 💬 Feedback & support

Missing a field, marketplace or integration? Open an issue on the **Issues** tab. Requests are usually answered within a day. If the actor saves you time, a ⭐ review helps other people find it.
