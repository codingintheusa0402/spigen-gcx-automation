// Stateless single-shot fetch relay — the MV3 replacement for GM_xmlhttpRequest.
// One message in, one fetch(), one response out. No multi-step orchestration
// here; that stays in the content script, same as it does today under Tampermonkey.
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg?.type !== 'GCX_XHR') return false;
  (async () => {
    const controller = new AbortController();
    const t = msg.timeout ? setTimeout(() => controller.abort(), msg.timeout) : null;
    try {
      const resp = await fetch(msg.url, {
        method: msg.method,
        headers: msg.headers,
        body: msg.data,
        credentials: 'include', // attach Zendesk / Seller Central session cookies
        redirect: 'follow',
        signal: controller.signal,
      });
      const responseText = await resp.text();
      const responseHeaders = [...resp.headers.entries()]
        .map(([k, v]) => `${k}: ${v}`).join('\r\n'); // CRLF-joined, matches GM_xmlhttpRequest's shape
      sendResponse({
        status: resp.status,
        statusText: resp.statusText,
        responseText,
        responseHeaders,
        finalUrl: resp.url,
      });
    } catch (e) {
      sendResponse({ error: e?.name === 'AbortError' ? 'timeout' : 'network', message: String(e) });
    } finally {
      if (t) clearTimeout(t);
    }
  })();
  return true; // keep the message channel open for the async sendResponse
});
