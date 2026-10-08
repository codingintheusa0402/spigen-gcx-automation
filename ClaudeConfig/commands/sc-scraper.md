Ask the user for each of the following scraper options, one message at a time (collect all answers before editing anything). Show current defaults in brackets so they can just press Enter to accept.

Options to ask (in this order):

1. **Domains** — which marketplaces to scrape?
   Options: US, EU (=DE+IT+FR+ES+UK combined, sequential), JP, IN — or type individual EU countries (UK/DE/FR/IT/ES)
   Default: `["EU", "JP", "US", "IN"]`

2. **Pages per domain** — how many pages to scrape per marketplace?
   Default: `30` (= 1,500 reviews at page size 50)

3. **Page size** — reviews per page (25 / 50 / 100)?
   Default: `50`

4. **Star filter** — which star ratings to include?
   Options: `"1,2,3"` (critical only) | `"1,2,3,4,5"` (all) | any custom combo like `"1,2"`
   Default: `"1,2,3,4,5"`

5. **Detection avoidance** — LOW / MEDIUM / HIGH?
   Default: `MEDIUM`

6. **ASIN filter** — path to a .txt file (one ASIN per line) to restrict output, or none?
   Default: `None`

7. **Output directory** — where to save the CSV files?
   Default: `~/Desktop`

8. **Fetch images** — fetch reviewer-attached image URLs? (yes/no)
   Default: `yes`

9. **Upload to Google Sheets** — after scraping, combine all domain CSVs into one new sheet on the source spreadsheet? (yes/no)
   All domains (EU+JP+US+IN) are combined into ONE new worksheet named `SC_{yymmdd}` (KST date of the run).
   Target spreadsheet: `1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw`
   Default: `yes`

After collecting all answers, do the following steps IN ORDER:

**Step 1** — Edit the USER CONFIG section of `/Users/kevinkim/Desktop/GCX/Scrapers/SC_Review_Scraper/scrape_sc_reviews.py` to reflect the user's choices. Map answers to these variables:
- Domains → `DOMAINS`
- Pages → `PAGES`
- Page size → `PAGE_SIZE`
- Star filter → `STAR_FILTER`
- Detection avoidance → `DETECTION_AVOIDANCE`
- ASIN filter → `ASIN_FILTER_FILE` (use `None` if not provided)
- Output dir → `OUT_DIR`
- Fetch images → `FETCH_IMAGES` (True/False)
- Upload to Sheets → `UPLOAD_TO_SHEETS` (True/False)
- Leave `HEADERS_TO_INCLUDE` as-is (includes `Product Rating` / `Ratings Count` right after `Order ID`, excludes `Domain Code` — standing preference, see note below). Only change it if the user explicitly asks to customize output columns for this run.

**Step 2** — Kill any existing scraper run:
```bash
kill $(pgrep -f "scrape_sc_reviews.py") 2>/dev/null; echo ok
```

**Step 3** — Run the scraper in the background:
```bash
cd /Users/kevinkim/Desktop/GCX/Scrapers/SC_Review_Scraper && python3 scrape_sc_reviews.py > /tmp/sc_scraper.log 2>&1
```

**Step 4** — After 5 seconds, tail the output and show it to the user so they can confirm it started correctly.

**Step 5** — Open a new Terminal window showing the live log:
```bash
osascript -e 'tell application "Terminal" to do script "tail -f /tmp/sc_scraper.log"'
```

**Important notes:**
- The script auto-launches Chrome with a persistent scraper profile (`~/.chrome-scraper-profile`). If sessions are still valid from a previous run, it skips the login step entirely.
- If login is required, Chrome opens one tab per endpoint. In interactive mode (TTY) the script waits for Enter; in background mode it uses the `LOGIN_WAIT_SECONDS` countdown. Tell the user to complete OTP on all tabs and press Enter.
- The script scrapes all top-level domains (US, EU, JP, IN) in parallel. Within EU: sub-countries from `EU_COUNTRIES` (default: DE → IT → FR → ES → UK) scrape **sequentially** on one shared tab — parallel tabs would race each other on the shared SC Europe session cookie. DE scrapes first (Phase 1); remaining countries reuse the same tab, switching marketplace via the SC Europe two-level account-switcher dropdown (Phase 2). This keeps the session alive throughout the EU run. If the session expires between countries anyway, the script detects the login redirect, pauses `MID_RUN_LOGIN_WAIT_SECONDS` (default 120 s) for OTP, then retries automatically — no pages are skipped. All EU reviews flush into one `EU_seller_central_reviews.csv`.
- **Single-country EU re-run** (e.g. append Italy to an existing EU CSV): set `DOMAINS=["EU"]`, `EU_COUNTRIES=["IT"]`, `APPEND_CSV=True`, `PAGES=50`. DE Phase 1 is skipped automatically when DE is not in `EU_COUNTRIES`.
- **EU image fetch**: only DE gets image URLs — the scraper Chrome profile has a customer session on amazon.de only. IT, FR, ES, UK are skipped for image fetch (printed as SKIP in logs) until those Amazon customer accounts are logged into the scraper profile.
- Output CSVs are saved as `<OUT_DIR>/<DOMAIN>_seller_central_reviews.csv`.
- **`Created 날짜` format**: always output as `yyyy-mm-dd` (e.g. `2026-05-08`). The script normalizes SC's raw English date strings ("May 8, 2026" / "8 May 2026") to ISO format at extraction time.
- **Google Sheets upload**: uses credentials at `~/.config/gws_shim/token.json` (Drive scope). All 4 domain CSVs (EU+JP+US+IN) are combined into ONE new worksheet named `SC_{yymmdd}` (KST date of the run). If that name is already taken (scraper ran more than once today), the next free name is used instead: `SC_{yymmdd}_1`, `_2`, etc. — an existing dated sheet is never overwritten automatically. Requires `gspread` and `google-auth` Python packages (already installed).
- **Default output columns (since 2026-08-18)**: `HEADERS_TO_INCLUDE` = ASIN, Created 날짜, 사진 유무, Reviewer, Review Ratings, Review Title, 본문, 국가, Review Link, Image URL, Review ID, Order ID, Product Rating, Ratings Count — in that order (`Product Rating`/`Ratings Count` sit right after `Order ID`, i.e. cols M/N). Only `Domain Code` stays excluded. Do not change this or reset `HEADERS_TO_INCLUDE` to `None` unless the user explicitly asks to customize columns for a given run.

---

## When running on the GCX server (Linux / WSL `gcx-server`, e.g. started from the phone)

Detect with `uname` = Linux. Then, instead of Steps 3–5 above:

- **Ask the 9 options exactly as above.** The user may be on the phone, so accept "defaults" / "기본값" as an answer for all of them.
- **Edit USER CONFIG the same way** (the path is the same: `/Users/kevinkim/...` is symlinked to the server home).
- **Run it in its own tmux window, so it shows live on the laptop's "GCX Live" window and from the Mac**:
  ```bash
  tmux kill-window -t gcx:sc-scraper 2>/dev/null
  tmux new-window -d -t gcx -n sc-scraper "cd ~/Desktop/GCX/Scrapers/SC_Review_Scraper && export DISPLAY=:0 WAYLAND_DISPLAY=wayland-0 PATH=\$HOME/.local/bin:\$PATH && python3 -u scrape_sc_reviews.py < /dev/null 2>&1 | tee /tmp/sc_scraper.log; echo '[scraper exited]'; exec bash"
  ```
  `< /dev/null` = background mode: no "press Enter" prompt (nobody may be at the laptop); it uses the countdowns instead.
- **Never use osascript / Terminal.app** (doesn't exist on the server).
- **Login / OTP:** if `~/.config/sc_scraper/credentials.txt` has `DOMAIN|EMAIL|PASSWORD` lines, the scraper logs in by itself. When Amazon asks for a one-time code, it posts a link to the user's private Chat room (`~/.config/sc_scraper/otp_chat_webhook.txt`). Tell the user: "open the link in your private Chat room and type the code from Google Authenticator". If the file has no login lines, the login must be done at the laptop. Tell the user that plainly, and don't wait silently.
  Never print, read aloud or edit `credentials.txt`. Only the user types into it.
- **Progress:** check `/tmp/sc_scraper.log` every minute or so (grep `Page N/30`, `✓`, `Traceback`, `session expired`, `OTP`) and give the user a short update per marketplace. Finish with the uploaded `SC_yymmdd[_n]` tab name and row counts. Then offer sc-review-propagate, as on the Mac.
