# Amazon Seller Central Review Scraper – Brand Customer Reviews Export (US · EU · JP · IN)

**Export every customer review of your brand from Seller Central's _Brand Customer Reviews_ page into one clean table**, with the order ID, reviewer photos, product rating and review count, for **US, Germany, Italy, France, Spain, UK, Japan and India** in a single run.

No password and no OTP are given to the actor. **You paste your own logged-in Seller Central session**, and the actor reads exactly what you would see on the Brand Customer Reviews page.

***

### 🔍 What does it do?

- Opens **Seller Central → Brand Customer Reviews** (`/brand-customer-reviews/`) with your session
- Pages through the newest reviews, filtered by stars and date as you choose
- Switches between the EU marketplaces automatically (DE → IT → FR → ES → UK) on one Europe session
- Adds the **Order ID** of verified purchases from Seller Central's own reviews API
- Optionally adds the **reviewer photo URLs**
- Returns **one clean row per review** with stable column names, ready to append into a spreadsheet, database or BI tool

> Brand Customer Reviews is available to **Brand Registry** brand owners. You need access to that page in Seller Central.

***

### 🔐 How to copy your login session (2 minutes)

You only need a session for each **region** you scrape:

| Region field | Covers marketplaces | Copy it from |
|---|---|---|
| **US session** | 🇺🇸 US | `https://sellercentral.amazon.com/brand-customer-reviews/` |
| **Europe session** | 🇩🇪 DE · 🇮🇹 IT · 🇫🇷 FR · 🇪🇸 ES · 🇬🇧 UK | `https://sellercentral-europe.amazon.com/brand-customer-reviews/` |
| **Japan session** | 🇯🇵 JP | `https://sellercentral.amazon.co.jp/brand-customer-reviews/` |
| **India session** | 🇮🇳 IN | `https://sellercentral.amazon.in/brand-customer-reviews/` |

#### Option A: Cookie-Editor (recommended, no technical steps)

1. Install the free **Cookie-Editor** extension (Chrome / Edge / Firefox).
2. Log in to Seller Central normally, including OTP, and open the region's **Brand Customer Reviews** page from the table above.
3. Click the Cookie-Editor icon → **Export** → **Export as JSON**. It's copied to your clipboard.
4. Paste it into that region's **session** field in the actor input.
5. Repeat for each other region you need. Log in to each region's Seller Central and export from its own page.

#### Option B: Copy the Cookie header from DevTools

1. On the region's Brand Customer Reviews page, press **F12** → **Network** tab → reload the page.
2. Click the first request (`brand-customer-reviews/`) → **Headers** → **Request Headers** → right-click **cookie** → **Copy value**.
3. Paste it into that region's session field.

#### Single vs multiple marketplaces

- **One marketplace (e.g. only US):** select `US` and paste only the **US session**.
- **Several EU countries (e.g. DE + UK):** select `DE` and `UK` and paste **one Europe session**. The actor switches countries itself.
- **Everything:** select all 8 marketplaces and paste **4 sessions** (US, Europe, Japan, India).
- **Multi-account sellers:** if Seller Central shows *Select an account* after login, type the account name in **Account name (region)**, exactly as shown (e.g. `Spigen Inc`, `Spigen EU`).

> 🔒 Session fields are **secret inputs**: Apify stores them encrypted, hides them in the UI and API, and the actor never writes their values to the log or the output. A session gives access to your Seller Central account, so only paste it into actors and accounts you trust, and log out to revoke it if needed.

***

### ♻️ Keep sessions fresh (important)

Seller Central sessions **expire**: after roughly a day of inactivity, when you log out, or when Amazon asks for OTP again. An expired session means a failed region, not wrong data, but the region returns nothing until you refresh it.

- **Refresh (re-copy) your sessions before each scheduled run**, or at least **every few days**. Repeat the 2-minute copy steps above.
- **Keep the same browser logged in.** Use the browser you exported from normally and don't click *Sign out*, because signing out invalidates the copied session immediately.
- **Turn on _Save refreshed sessions_.** Amazon rotates session cookies during the run, and with this on the actor saves the updated cookies to the run's key-value store (record `REFRESHED_SESSIONS`, private to your account). Paste those into the next run to extend the session's life.
- **Check `SESSION_HEALTH` after every run.** The run's key-value store has a per-marketplace report:
  - `OK`: scraped, with the number of reviews
  - `SESSION_EXPIRED`: copy a fresh session for that region and re-run
  - `NO_SESSION`: you selected the marketplace but didn't paste its region's session
  - `ERROR`: details in `message` (e.g. a non-English Seller Central language, or an account name to choose)
- One region failing **doesn't stop the others**. The run status lists what needs attention.

***

### 📥 Input example

```json
{
  "marketplaces": ["US", "DE", "UK", "JP"],
  "sessionUS": "[{\"domain\":\".amazon.com\",\"name\":\"at-main\",\"value\":\"…\"}, …]",
  "sessionEU": "session-id=…; at-acbuk=…; sess-at-acbuk=…; …",
  "sessionJP": "[ …Cookie-Editor export… ]",
  "maxPagesPerMarketplace": 30,
  "stars": [1, 2, 3],
  "minReviewDate": "2026-10-01",
  "fetchOrderId": true,
  "accountEU": "Spigen EU"
}
```

### 📤 Output columns

| Column | Example | Meaning |
|---|---|---|
| `ASIN` | `B0G7RYP439` | Child ASIN |
| `Created Date` | `2026-10-15` | Review date (yyyy-mm-dd) |
| `Has Photo` | `Y` / `N` | Review has photos (when photo fetch is on) |
| `Reviewer` | `Jane D.` | Public reviewer name |
| `Review Ratings` | `2` | Stars given |
| `Review Title` | `Stopped sticking after a week` | Title |
| `Review Text` | `…` | Review text |
| `Country` | `US` | Marketplace: US, DE, IT, FR, ES, UK, JP, IN |
| `Review Link` | `https://www.amazon.com/gp/customer-reviews/R…` | Public review link |
| `Image URL` | `https://m.media-amazon.com/images/I/…jpg\|…` | Photo URLs, `\|`-separated |
| `Review ID` | `R2ABCDEF123456` | Amazon review ID (unique key) |
| `Order ID` | `113-1234567-1234567` | Order ID of verified purchases |
| `Product Rating` | `4.2` | Product average rating |
| `Ratings Count` | `2958` | Product total ratings |

Building an integration? Set **Output field names → camelCase** to get `asin`, `createdDate`, `hasPhoto` … instead.

***

### 💰 How much does it cost?

**From $4 per 1,000 reviews.** It's pay-per-event, so platform usage isn't billed on top, and you're **charged only for review rows delivered**. Expired or missing sessions and failed regions cost nothing beyond the run start.

| Your Apify plan | Price per review | Per 1,000 reviews |
|---|---|---|
| Free | $0.008 | $8.00 |
| Starter | $0.006 | $6.00 |
| Scale | $0.005 | $5.00 |
| Business / Platinum / Diamond | $0.004 | $4.00 |

Plus **$0.01 per run** (browser start).

**Examples (Starter plan):** a daily incremental run of ~150 new reviews ≈ $0.91, a full export of 1,500 reviews × 4 marketplaces ≈ $36.01, and the Free plan's monthly credit covers about 600 reviews to test with.

**Keep costs predictable:** set *Maximum cost per run* in the run options. The actor stops cleanly when it's reached, and it never bills past it.

***

### 📸 Reviewer photos (optional)

Amazon shows review photos only to **signed-in shoppers**, and that's a different login from Seller Central. To fill `Has Photo` / `Image URL`:

1. Turn on **Fetch reviewer photo URLs**.
2. Sign in as a shopper on each store site you scrape (e.g. `www.amazon.com`, `www.amazon.de`), export with Cookie-Editor, and paste the export(s) into **Amazon shopper (storefront) session**. Several exports can be merged into one JSON array.

Without a shopper session, every review shows `Has Photo = N`.

***

### ⚙️ Settings

| Setting | Default | Notes |
|---|---|---|
| Max pages per marketplace | 30 | × page size = up to 1,500 newest reviews per marketplace |
| Reviews per page | 50 | 25 / 50 / 100 |
| Star ratings | 1–5 | `1,2,3` = critical reviews only |
| Only reviews on/after | — | Also stops paging early, so runs are faster |
| ASIN filter | — | Keep only these child ASINs |
| Pace | MEDIUM | LOW = fastest, HIGH = safest for large or frequent runs |

### 💡 Tips

- **Seller Central must be in English** (top-right language menu) when you copy the session. Review dates and reviewer names are read from the English page.
- **Schedule it daily** with *Only reviews on/after* set to a recent date for a fast incremental export, and refresh sessions each time.
- Review IDs (and Order IDs) make the output easy to **deduplicate and append** into an existing sheet.

### ❓ FAQ

**Does it need my password or OTP?** No. Only a session you copy from your own logged-in browser.

**Why did a region return `SESSION_EXPIRED`?** The copied session has expired or was signed out. Copy a fresh one (see *Keep sessions fresh*).

**Can one Europe session scrape all EU countries?** Yes. DE, IT, FR, ES and UK are switched inside the same Seller Central Europe login.

**Is this the same data as Amazon's public review pages?** It's what Seller Central's Brand Customer Reviews shows the brand owner, plus Order IDs, which aren't public.

**Is it allowed?** It reads your own Seller Central data with your own session. Follow Amazon's Seller Central terms and your company's data policies.

🔖 **Tags:** seller central review export · brand customer reviews · amazon brand registry reviews · amazon review order id · amazon review monitoring · seller central scraper · amazon eu reviews
