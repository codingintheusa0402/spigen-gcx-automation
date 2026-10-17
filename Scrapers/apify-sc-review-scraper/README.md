# Amazon Seller Central Review Scraper – Brand Customer Reviews Export (US · EU · JP · IN)

**Export every customer review of your brand from Seller Central's _Brand Customer Reviews_ page into one clean table**, with the order ID, reviewer photos, product rating and review count, for **US, Germany, Italy, France, Spain, UK, Japan and India** in a single run.

**Sign in once, yourself, in a real browser inside Apify's Live View**, and the actor remembers the session for later runs. No passwords in the input, no OTP relay, no copying cookies (cookie paste is still there as a fallback).

***

### ⚠️ Read this first: Amazon's terms and your account

**Amazon's terms don't allow automated access to Seller Central** (robots, scrapers and similar tools). Using this actor may break those terms, and **Amazon can restrict or suspend a seller account it believes is accessed automatically**. You decide whether to use it, and **you run it at your own risk.** The publisher isn't responsible for any action Amazon takes on your account.

To keep the footprint small, the actor:
- reads only the Brand Customer Reviews page and the reviews API your own browser uses on it,
- scrapes **one region at a time at a human pace**, with random delays and scrolling,
- **backs off and retries** on throttling or robot checks (8 s → 25 s → 60 s), then **stops instead of hammering**,
- never changes anything in your account.

Keep runs small and infrequent (e.g. daily, with *Only reviews on/after*), and use Pace = HIGH for big exports. This isn't legal advice. If in doubt, ask Amazon or your counsel.

***

### 🚀 Quick start (2 steps)

| Step | Mode | What you do |
|---|---|---|
| **1 · once** | 🔐 **Login** | Pick marketplaces → **Start** → open the run's **Live View** tab → sign in to Seller Central (email, password, 2-step code). The session is saved for each region. |
| **2 · every time** | 🔎 **Scrape** | Pick marketplaces → **Start**. Reviews come out as a table. Saved sessions are checked first and refreshed after each run. |

When a session eventually expires, the run **stops that region with a clear message: "run Login mode again"**. It takes a minute.

***

### 🔐 Login mode, step by step

1. In the input, set **Mode → 🔐 Login**, select your **Marketplaces**, and click **Start**.
2. Open the run's **Live View** tab. You'll see this:

   ![Login mode in Apify Live View](https://raw.githubusercontent.com/codingintheusa0402/spigen-gcx-automation/main/Scrapers/apify-sc-review-scraper/docs/login-live-view.png)

3. **Click inside the picture** to put the cursor in a field, type into the box above it, and press **Type** (or Enter). The 👁 button shows what you typed. Enter, Tab, ⌫, scroll, Back and Reload buttons are there too.
4. Sign in as usual: **email → password → 2-step code** (from your authenticator app or SMS). Tick **"Don't require OTP on this browser"** if Amazon offers it. Sessions then last longer.
5. If Seller Central shows **Select an account**, click your account and marketplace.
6. Make sure Seller Central is in **English** (language menu, top right). The scraper reads the English page.
7. Once you're on Seller Central, the actor **saves the session automatically** ("✅ US session saved") and opens the next region. Each region waits up to *Login mode: minutes to sign in* (default 10). **Skip region** moves on.
8. When the run finishes, switch **Mode → 🔎 Scrape** and run as often as you like.

**Which logins you need**

| Select any of… | Sign in to (one login each) | Notes |
|---|---|---|
| 🇺🇸 US | `sellercentral.amazon.com` | |
| 🇩🇪 DE · 🇮🇹 IT · 🇫🇷 FR · 🇪🇸 ES · 🇬🇧 UK | `sellercentral-europe.amazon.com` | **One Europe login covers all five.** The actor switches countries itself (DE → IT → FR → ES → UK). |
| 🇯🇵 JP | `sellercentral.amazon.co.jp` | The page often opens in Japanese. Switch to **English** before you finish. |
| 🇮🇳 IN | `sellercentral.amazon.in` | |

> 🔒 **Where sessions live:** in a named key-value store **in your own Apify account** (default `sc-review-sessions`, one record per region: `SESSION_US`, `SESSION_EU` …). Other users of this actor can't see it. **To forget a session**, delete that record or the whole store in Storage → Key-value stores, or sign out of all devices in Amazon. Your typing in Live View goes straight to the browser and is never stored or logged.

***

### 🍪 Fallback: paste cookies instead

Prefer not to sign in through Live View? Paste a session per region into the secret **Cookie fallback** fields. They're stored encrypted by Apify and never logged. A pasted session wins over a saved one for that region, and after a successful run it's saved too, so you only paste once.

**Option A: Cookie-Editor (recommended)**
1. Install the free **Cookie-Editor** extension (Chrome / Edge / Firefox).
2. Sign in to Seller Central normally (in English) and open the region's Brand Customer Reviews page:
   - US: `https://sellercentral.amazon.com/brand-customer-reviews/`
   - Europe: `https://sellercentral-europe.amazon.com/brand-customer-reviews/`
   - Japan: `https://sellercentral.amazon.co.jp/brand-customer-reviews/`
   - India: `https://sellercentral.amazon.in/brand-customer-reviews/`
3. Click the Cookie-Editor icon → **Export** → **Export as JSON**. It's copied to your clipboard.
4. Paste it into that region's field (**US session**, **Europe session**, …).

**Option B: DevTools.** On the same page press **F12** → **Network** → reload → click the first `brand-customer-reviews/` request → **Request Headers** → right-click **cookie** → **Copy value** → paste.

Don't click **Sign out** in the browser you copied from. That ends the copied session immediately.

***

### ♻️ Session health: checked at the start of every run

Before scraping, every region's session is tested with a single page view:

- **Healthy:** the region is scraped, then the **refreshed cookies are saved back** (*Keep saved sessions fresh*, on by default), so sessions live longer.
- **Expired:** that region **stops** with *"The US session has expired. Run the actor once in Login mode with US selected…"*. The other regions continue. If none is usable, the whole run stops with that message.
- Every run writes a **`SESSION_HEALTH`** report to its key-value store:
  - `OK`: scraped, with the review count
  - `SESSION_EXPIRED`: run **Login mode** for that region
  - `NO_SESSION`: no saved or pasted session for that region yet
  - `BLOCKED`: Amazon showed robot checks repeatedly. Wait a few hours and use Pace = HIGH.
  - `ERROR`: details in `message` (e.g. language not English, account name to pick)

Expired or missing sessions cost nothing beyond the run start.

***

### 📥 Input example

```json
{
  "mode": "scrape",
  "marketplaces": ["US", "DE", "UK", "JP"],
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

**From $3 per 1,000 reviews on the Free plan, $4–6 on paid plans.** It's pay-per-event, so platform usage isn't billed on top, and you're **charged only for review rows delivered**. Expired or missing sessions and failed regions cost nothing beyond the run start.

| Your Apify plan | Price per review | Per 1,000 reviews |
|---|---|---|
| 🎁 Free (try it on your monthly free credit) | $0.003 | **$3.00** |
| Starter | $0.006 | $6.00 |
| Scale | $0.005 | $5.00 |
| Business / Platinum / Diamond | $0.004 | $4.00 |

Plus **$0.01 per run** (browser start).

**Examples (Starter plan):** a daily incremental run of ~150 new reviews ≈ $0.91, a full export of 1,500 reviews × 4 marketplaces ≈ $36.01, and 🎁 the Free plan's monthly $5 credit covers **about 1,600 reviews** to test with, every month.

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
| Mode | Scrape | **Login** once to save sessions, then **Scrape** |
| Login mode: minutes to sign in | 10 | Per region |
| Max pages per marketplace | 30 | × page size = up to 1,500 newest reviews per marketplace |
| Reviews per page | 50 | 25 / 50 / 100 |
| Star ratings | 1–5 | `1,2,3` = critical reviews only |
| Only reviews on/after | — | Also stops paging early, so runs are faster |
| ASIN filter | — | Keep only these child ASINs |
| Pace | MEDIUM | LOW = fastest, HIGH = safest for large or frequent runs. Retries with backoff are always on |
| Use sessions saved by Login mode | on | Off = only pasted sessions are used |
| Keep saved sessions fresh | on | Saves Amazon's rotated cookies after each successful region |
| Session store name | `sc-review-sessions` | One per brand / seller account |

### 💡 Tips

- **Seller Central must be in English** (top-right language menu) when you sign in or copy the session. Review dates and reviewer names are read from the English page.
- **Schedule it daily** with *Only reviews on/after* set to a recent date for a fast incremental export, and the saved sessions refresh themselves after every successful run.
- Review IDs (and Order IDs) make the output easy to **deduplicate and append** into an existing sheet.

### ❓ FAQ

**Does the actor get my password?** In Login mode you type it into Seller Central's own page through the Live View, like in any browser. The actor passes your keystrokes straight to the page and never stores or logs them. Only the resulting login cookies are saved, in your own account. With the cookie fallback, no password is involved at all.

**Why did a region return `SESSION_EXPIRED`?** Amazon ended that session (time, sign-out, or a new 2-step check). Run **Login mode** for that region again.

**Can several people / brands use it from one Apify account?** Yes. Give each seller account its own **Session store name** (e.g. `brand-a-sessions`).

**Can one Europe session scrape all EU countries?** Yes. DE, IT, FR, ES and UK are switched inside the same Seller Central Europe login.

**Is this the same data as Amazon's public review pages?** It's what Seller Central's Brand Customer Reviews shows the brand owner, plus Order IDs, which aren't public.

**Is it allowed?** See **Read this first** at the top. Amazon's terms don't allow automated access to Seller Central; you run this at your own risk.

🔖 **Tags:** seller central review export · brand customer reviews · amazon brand registry reviews · amazon review order id · amazon review monitoring · seller central scraper · amazon eu reviews
