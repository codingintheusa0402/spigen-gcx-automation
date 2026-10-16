# Amazon ASIN Rating & Review Count Scraper – Bulk Rating Checker for 19 Amazon Sites

**Turn a list of ASINs or Amazon product URLs into a clean table of star ratings and review counts.** No Amazon account, no API keys, no browser. Paste your list, click **Start**, and download Excel, CSV or JSON.

Perfect for **rating monitoring, review-growth tracking, competitor benchmarking, catalog health checks and daily KPI dashboards**, from **$1.60 per 1,000 products** (and just **$1 per 1,000 on the Apify Free plan**), charged only for products that return data.

***

### 🔍 What does Amazon ASIN Rating Scraper do?

You give it products, and it returns one row per product:

- ⭐ **Average star rating** (e.g. `4.2`)
- 🔢 **Total global ratings count** (e.g. `2958`)
- 🏷️ **Product title** and 🔗 **product link**
- 🌍 **Marketplace** the data came from
- 🕒 **Scrape timestamp**, so repeated runs build a history
- ✅ **Status** per product: found, not sold on that marketplace, or blocked

It accepts:

- **Bare ASINs**: `B0G7RYP439`
- **Product URLs** from any supported Amazon site: `https://www.amazon.de/dp/B0GDHRZFWH`
- **Mixed lists** pasted straight from a spreadsheet column. Duplicates, blank cells, `-` and `TBD` are skipped automatically.

***

### 🌍 Which Amazon site gets scraped?

| You enter | Scraped on |
|---|---|
| A bare ASIN, e.g. `B0G7RYP439` | The **Marketplace** selected in the input (default amazon.com) |
| A product URL, e.g. `https://www.amazon.co.uk/dp/B0H2BX5FF2` | **That URL's own site** (amazon.co.uk) |

So **one run can cover the US, UK, Germany, Japan and more**. List the same ASIN as URLs on different sites to compare its rating country by country.

> ⚠️ Amazon ratings are **per marketplace**. A product can be 4.6★ with 13,000 ratings on amazon.com and 4.1★ with 400 ratings on amazon.de. Always check the `marketplace` column.

***

### 🌟 Key features

- **19 Amazon marketplaces**: US, UK, DE, FR, IT, ES, NL, SE, PL, TR, JP, IN, CA, MX, BR, AU, AE, SA, SG
- **Fast**: ~3,400 products in about 5½ minutes on default settings
- **Cheap**: $1.60–2 per 1,000 products ($1 on the Free plan), roughly 2× cheaper than typical full Amazon product scrapers
- **Fair billing**: not-found and blocked products are **never charged**
- **Stable**: each parallel session behaves like one real shopper, with its own IP, cookies and real-browser fingerprint, human-like pacing, rotation, and automatic slow-down if Amazon pushes back
- **Smart proxies**: fast datacenter IPs first, residential only for products that keep getting blocked
- **Clean data**: incomplete pages are retried instead of saved, and Amazon's "Continue shopping" check page is passed automatically
- **Small footprint**: runs on 512 MB, and more memory isn't needed
- **Monitored daily** with live test runs on amazon.com and amazon.de

***

### ⚖️ Why choose this scraper?

| | **This actor** | Typical Amazon product scrapers |
|---|---|---|
| Price per 1,000 products | **$1.60–2.00** ($1 on Free plan) | $3–5 |
| Charged for not-found / blocked products | **No** | Often yes (every dataset row) |
| Mix several Amazon sites in one run | **Yes, from the URL** | Usually one country per run |
| Default memory | **512 MB** | 1–4 GB |
| ~3,400 products | **~5½ min** | Often much longer (full page parsing) |
| Output | Focused: rating, count, title, link | 40+ fields you may not need |
| Amazon login / cookies needed | **No** | Sometimes |

If you need **only ratings and review counts**, this is the fastest and cheapest way to get them.

***

### 📥 Input

| Field | Required | Description |
|---|---|---|
| **ASINs or product URLs** | ✅ | One per line. Bare ASINs or Amazon product links, mixed freely |
| **Marketplace for bare ASINs** | | `com` (default), `co.uk`, `de`, `co.jp` … Product URLs always use their own site |
| **Parallel sessions** | | Default 40, max 50. Capped automatically to fit the run's memory |
| **Proxy** | | Default: Apify datacenter, auto-escalating to residential. Leave as is |
| **Save raw pages (debug)** | | Stores fetched HTML for troubleshooting. Off by default |

**Example input**

```json
{
  "asins": [
    "B0G7RYP439",
    "B0G7QZ7RMD",
    "https://www.amazon.de/dp/B0GDHRZFWH",
    "https://www.amazon.co.uk/dp/B0H2BX5FF2"
  ],
  "marketplace": "com"
}
```

***

### 📋 Output fields

| Field | Type | Description |
|---|---|---|
| `asin` | string | The product's ASIN |
| `found` | boolean / null | `true` = product page found · `false` = not sold on this marketplace / delisted · `null` = blocked after all retries |
| `rating` | number / null | Average star rating out of 5. `null` if the product has no ratings yet |
| `ratingsCount` | integer | Total number of global ratings (`0` if none) |
| `title` | string | Product title as shown on Amazon |
| `url` | string | Product page link |
| `marketplace` | string | Amazon site the data came from, e.g. `amazon.de` |
| `scrapedAt` | string | Scrape time, UTC, ISO 8601 |
| `error` | string | Only on blocked items: why it couldn't be scraped |

### 📤 Example output

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
  }
]
```

The **Output** tab shows these as a table (*Ratings overview*), ready to export to Excel, CSV, JSON, XML or HTML.

### 🚧 Not-found and blocked items

Every product you enter gets a row, so nothing silently goes missing:

```json
{ "asin": "B0B7MFZ2SH", "found": false, "marketplace": "amazon.com" }
```
`found: false`: Amazon has no product page for this ASIN on this site. **Not charged.**

```json
{ "asin": "B0XXXXXXXX", "found": null, "error": "Blocked after 6 attempts", "marketplace": "amazon.com" }
```
`found: null`: Amazon blocked all 6 attempts, each on a fresh identity and the last ones on residential IPs. Rare; just run those ASINs again. **Not charged.**

***

### 💰 How much does it cost to scrape Amazon ratings?

**From $1.60 per 1,000 products, with bigger discounts on higher Apify plans.** It's pay-per-event, so platform usage isn't billed on top.

| Your Apify plan | Per product | Per 1,000 products |
|---|---|---|
| 🎁 Free (try it on your monthly free credit) | $0.001 | **$1.00** |
| Starter | $0.002 | $2.00 |
| Scale | $0.0018 | $1.80 |
| Business / Platinum / Diamond | $0.0016 | **$1.60** |

Plus $0.005 per run. **Not-found and blocked products are always free.**

| You check (Starter plan) | Cost |
|---|---|
| 100 products once | ≈ $0.21 |
| 1,000 products once | ≈ $2.01 |
| 10,000 products once | ≈ $20.01 |
| 500 products every weekday (~22 runs/month) | ≈ $22 / month |

🎁 **Free trial:** new to Apify? The Free plan's monthly $5 credit covers **about 5,000 products** at the Free-plan price, every month, enough to test on your real catalog before you pay anything.

***

### 🚀 How to scrape Amazon ratings by ASIN (step by step)

1. Click **Try for free** (or **Start**) on this page and sign in to Apify.
2. Paste your **ASINs or product URLs** into the first field, one per line.
3. Bare ASINs only: pick their **Marketplace** (default amazon.com).
4. Click **Start**. A few hundred products take about a minute.
5. Open the **Output** tab and click **Export** to download Excel, CSV or JSON.
6. Optional: click **Save as a new task**, then **Schedule** it (e.g. every weekday at 08:00) to track ratings over time.

***

### 📊 Popular use cases

- **Brand & seller rating monitoring**
  - Daily snapshot of every listing's rating and review count
  - Spot a rating drop the morning it happens, before it hurts conversion
- **Competitor benchmarking**
  - Track rival products' review velocity after their launch or promotion
  - Compare your rating to the category's top sellers
- **Multi-country catalog health**
  - Same product, many marketplaces: find the countries where it underperforms
- **Agency & client reporting**
  - Scheduled rating reports to Google Sheets, Slack or email, no manual checking
- **Product launch tracking**
  - Measure how fast a new ASIN collects its first 100, 500 and 1,000 ratings
- **Data & BI pipelines**
  - Feed ratings into dashboards, pricing models or ML features via API

***

### 🔌 Integrations & API

Connect results to **Google Sheets, Slack, email, Zapier, Make, Airbyte, GitHub, webhooks** and more under the **Integrations** tab. Or call it from code:

**Python**
```python
from apify_client import ApifyClient

client = ApifyClient("<YOUR_APIFY_TOKEN>")
run = client.actor("k-data-lab/amazon-asin-rating-scraper").call(
    run_input={"asins": ["B0G7RYP439", "https://www.amazon.de/dp/B0GDHRZFWH"]}
)
for item in client.dataset(run["defaultDatasetId"]).iterate_items():
    print(item["asin"], item.get("rating"), item.get("ratingsCount"))
```

**JavaScript**
```js
import { ApifyClient } from 'apify-client';

const client = new ApifyClient({ token: '<YOUR_APIFY_TOKEN>' });
const run = await client.actor('k-data-lab/amazon-asin-rating-scraper').call({
    asins: ['B0G7RYP439', 'https://www.amazon.de/dp/B0GDHRZFWH'],
});
const { items } = await client.dataset(run.defaultDatasetId).listItems();
console.table(items.map(({ asin, rating, ratingsCount }) => ({ asin, rating, ratingsCount })));
```

**cURL** (run and get the results in one request)
```bash
curl -X POST "https://api.apify.com/v2/acts/k-data-lab~amazon-asin-rating-scraper/run-sync-get-dataset-items?token=<YOUR_APIFY_TOKEN>" \
  -H "Content-Type: application/json" \
  -d '{"asins": ["B0G7RYP439", "B0G7QZ7RMD"]}'
```

***

### 💡 Tips for best results

- **Test with 10–20 products first** to check the marketplace setting and the output format.
- **Use full product URLs for non-US products.** That way each one is scraped on the right site, even in a mixed list.
- **Variants share reviews.** Color and size variants of one product usually show the same rating and count, because Amazon pools them.
- **Schedule instead of re-running by hand.** The `scrapedAt` field lets you stack runs into a history in Sheets or your database.
- **Leave the proxy and memory defaults alone.** They're tuned for speed and cost.

### 🤝 When to use this vs a full Amazon product scraper

- **Use this actor** when you need **ratings and review counts** for a known list of products, especially many products, often.
- **Use a full product scraper** when you need prices, sellers, images or descriptions as well.
- **Use a review scraper** when you need the individual review texts.

***

### ❓ FAQ

**How do I check Amazon ratings for many ASINs at once?**
Paste them all into the input, one per line. A list of thousands is fine; ~3,400 products take about 5½ minutes.

**Which Amazon countries are supported?**
19 sites: amazon.com, .co.uk, .de, .fr, .it, .es, .nl, .se, .pl, .com.tr, .co.jp, .in, .ca, .com.mx, .com.br, .com.au, .ae, .sa and .sg.

**Do I need an Amazon account, cookies or the Amazon API?**
No. Only public product pages are read.

**Why does a product show `found: false`?**
It isn't sold on that marketplace or has been delisted. Check the marketplace, or pass the full product URL. You aren't charged for it.

**Why is `rating` null with `ratingsCount: 0`?**
The product has no ratings yet. That's real data, not an error.

**Does it return the 5-star / 1-star breakdown or the review texts?**
No. It focuses on the headline rating and total count, which is what keeps it fast and cheap. Use a review scraper for review texts.

**Can I schedule it and send results to Google Sheets?**
Yes. Save your input as a task, add a schedule, and connect the Google Sheets integration.

**Is it legal to scrape Amazon ratings?**
This actor only collects publicly available product data and no personal data. You're responsible for how you use the results. Check Amazon's terms and the laws that apply to you.

**Something doesn't look right. How do I get help?**
Open an issue on the **Issues** tab with the run link and a few example ASINs.

***

### 💬 Your feedback

Missing a marketplace, field or integration? Tell us on the **Issues** tab. If this actor saves you time, please leave a ⭐ **review**, because it helps other people find it.

🔖 **Tags:** amazon rating scraper · amazon review count · asin lookup · bulk asin checker · amazon star rating · review monitoring · amazon competitor tracking · amazon seller tools · ecommerce data · amazon api alternative
