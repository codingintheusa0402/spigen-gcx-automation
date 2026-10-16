# K-Data Lab — Apify publisher profile kit

Apply in Apify Console (logged in as the publisher account, currently `cosmic_meteorite`).
Profile fields have no API (`PUT /v2/users/me` → 405), so these are manual.

## 1. Settings → Account → Username
`k-data-lab` (checked free on 2026-10-16)

After the rename, tell Claude "username changed": the actor README's code snippets get switched
from `cosmic_meteorite/…` to `k-data-lab/…` and redeployed. Scheduled tasks, the Spigen job and the
health check use the actor ID `2AuZCyLxKZuO2OyNY`, so they keep working.

## 2. Settings → Profile (public profile)

**Name**
K-Data Lab

**Picture**
`avatar.png` (this folder)

**Bio** (short line under the name)
Fast, fairly priced scrapers for Korean and global e-commerce, real-estate and social data — Amazon, Coupang, Naver and more. Pay only for results.

**Website / GitHub / X**: optional. Add one if you have a landing page; a profile with a link looks more established.

**Profile README** (paste into the profile description / README box)

```markdown
## 👋 K-Data Lab

**Reliable data from Korea's and the world's biggest platforms — ready in Excel, CSV, JSON or via API.**

We build focused, fast and fairly priced Apify actors:

- ⚡ **Fast & lightweight:** small memory, parallel sessions, results in minutes
- 💰 **Fair pay-per-result pricing:** not-found and blocked items are never charged
- 🛡️ **Stable:** human-like sessions, automatic retries, daily live health checks
- 🇰🇷 **Korea expertise:** Korean platforms with Korean-aware parsing (prices, addresses, dates)

### 🛒 E-commerce
- [Amazon ASIN Rating & Review Count Scraper](https://apify.com/k-data-lab/amazon-asin-rating-scraper): star ratings and review counts for any ASIN list, 19 Amazon sites, from $2 / 1,000
- [Amazon Seller Central Review Scraper](https://apify.com/k-data-lab/amazon-seller-central-review-scraper): export Brand Customer Reviews with Order IDs and photos for US/EU/JP/IN using your own session, from $4 / 1,000

*Coming soon: Coupang, Naver Smart Store & Naver Shopping, Korean real estate (네이버부동산, 부동산114, KB부동산), Instagram.*

### 💬 Need something specific?
Open an issue on any actor. Feature requests and new-site requests are welcome.
```

## 3. Actor icons (Console only — the API rejects any picture URL not uploaded through Console)
Log in to Console as **cosmic_meteorite**, then for each actor: Actor → Publication → Display information → Icon → upload:
- Amazon ASIN Rating & Review Count Scraper → `icons/amazon-asin-rating-scraper.png`
- Amazon Seller Central Review Scraper → `icons/amazon-seller-central-review-scraper.png`

Brand system: deep-blue gradient tile, white K-Data Lab "K" monogram badge top-left (gold lower arm), one gold accent symbol per actor, white main pictogram. Reuse `icons3.py`-style layout for every new actor.

## Naming convention for future actors (keeps the profile coherent)
- Name pattern: `<platform>-<what>-scraper`, e.g. `coupang-product-scraper`, `naver-land-listings-scraper`
- Title pattern: `<Platform> <What> Scraper – <key benefit>`, e.g. "Coupang Product Scraper – Prices, Ratings & Rocket Delivery"
- The same icon style for every actor (deep-blue tile + one gold accent symbol); only the symbol changes
- Every README follows the Amazon actor's structure (what it does → which site → features → why choose → input → output → errors → pricing → steps → use cases → integrations/API → tips → FAQ → tags)
