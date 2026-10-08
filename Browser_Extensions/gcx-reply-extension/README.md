# GCX Reply — Chrome Extension (MV3)

Native-extension port of `tampermonkey_scripts/GCX Reply.user.js`. See
`~/.claude/plans/elegant-sleeping-locket.md` for the full migration plan.

> **Status: paused prototype. Not in use, not tracked in git.** The port was
> taken from userscript **v3.0.1** (manifest `version` 3.0.1, `content.js`
> last touched 2026-07-21). The live tool is still the Tampermonkey script,
> now at **v3.7.2**, so this folder is missing everything since v3.0.1: Seller
> Notes, the ABM relay log and retry sweep, raw-ABM collapse, MCF 담당자,
> product caching, the language auto-correct, the dock mode fixes, and more.
> Re-port from the current `.user.js` before resuming. Don't patch this copy.

## Status

- **Phase 0** (network shim go/no-go): passed. Background-worker `fetch()`
  carries Zendesk and Seller Central session cookies, including redirects
  through Amazon's SSO checkpoint, once `*.amazon.com` is in `host_permissions`.
- **Phase 1** (Zendesk ticket panel, full port): paused at v3.0.1. `content.js`
  is a verbatim port of the original script except `initMcfPage_()`, which is
  dead code here because the SC-MCF URLs aren't in `content_scripts.matches`.
- **Phase 2** (SC MCF page + cross-tab bridge to the companion Tampermonkey
  MCF Autofill scripts): not started.
- **Phase 3** (Web Store submission readiness): not started.

## How it differs from the userscript

- `background.js` is a stateless fetch relay: one `GCX_XHR` message in, one
  `fetch(..., {credentials:'include'})` out. It stands in for
  `GM_xmlhttpRequest`.
- `content.js` defines a `GM_xmlhttpRequest`-compatible shim that messages the
  worker, so no call site in the ported body had to change. `SCRIPT_VER` comes
  from `chrome.runtime.getManifest().version`.
- Content script matches: `spigenhelp.zendesk.com/agent/tickets/*` and
  `/agent/filters*`. Host permissions cover Zendesk, `script.google.com`, and
  the Seller Central / amazon.* storefront domains. The JP entry in
  `manifest.json` still uses `sellercentral.amazon.co.jp`; the userscript
  switched to `sellercentral-japan.amazon.com` in 2026-08.
- Backend: the same GCXReply_GAS web app (`GAS_URL`) as the userscript.

## Load unpacked (dev/testing)

1. `chrome://extensions` → enable **Developer mode**
2. **Load unpacked** → select this folder
3. After any edit to `manifest.json`/`content.js`/`background.js`, click the
   reload icon on this extension's card, then refresh any open ticket tabs

Don't run it in the same profile as the Tampermonkey GCX Reply, or both will
inject a panel on the same ticket. This port has no ABM relay code, so it can't
send anything to Seller Central.

## Files

- `manifest.json`: MV3 manifest, `host_permissions`, content script matches
- `background.js`: stateless fetch relay (`GM_xmlhttpRequest` replacement)
- `content.js`: ported panel logic plus the shim that calls into `background.js`
- `icons/`: placeholders; replace before Web Store submission
