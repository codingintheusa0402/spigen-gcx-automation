# Tampermonkey_GCX

Tampermonkey userscripts for the Spigen GCX Amazon operations workflow. Install via [Tampermonkey](https://www.tampermonkey.net/) → Dashboard → Utilities → Import.

---

## Scripts

### Amazon MCF Autofill (`v0.8.1`)
**Matches:** `sellercentral.amazon.*` and `sellercentral-europe.amazon.*` — MCF create-order pages

Injects a floating panel on the MCF order creation page (EU marketplaces: UK, DE, FR, IT, ES). Autofills recipient name, address, and line items from a GCX order ID, reducing manual data entry when placing Multi-Channel Fulfillment orders.

---

### Amazon JP MCF Autofill (`v1.4.4`)
**Matches:** `sellercentral-japan.amazon.com` — MCF create-order pages

JP-specific variant of the MCF Autofill script. Pulls order data from a Google Apps Script endpoint, maps Japanese prefecture names to their romanized equivalents, and autofills the JP MCF order form.

---

### Amazon Invoice Automation (`v1.5`)
**Matches:** `sellercentral.amazon.de` — individual order pages

Adds a "Run Now" button on Amazon.de Seller Central order pages. On click, attempts to download the deemed resale/supply invoice first, falling back to the Amazon-generated invoice. Copies the result to clipboard via `GM_setClipboard`.

---

## Installation

1. Install the [Tampermonkey extension](https://www.tampermonkey.net/) in Chrome.
2. Open Tampermonkey Dashboard → Utilities → Import from file.
3. Select the `.user.js` file for the script you want to install.
4. Click "Install" when prompted.
