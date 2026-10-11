/**
 * GCX KakaoRelay — copies every bot/webhook alert posted in the private GCX room
 * (spaces/AAQAc9NQmJQ: the "GCX Server" webhook + the private-room webhook both post there)
 * to KakaoTalk "나에게 보내기" (memo to self), once a minute, from Google's cloud — so it keeps
 * working when the server or the Mac is down.
 *
 * Setup (once): open the web app URL → paste the Kakao REST API key (+ client secret if enabled)
 * → "Connect KakaoTalk" → agree → "Send test". installTrigger() runs automatically on connect.
 * Kakao access tokens (6h) and refresh tokens (2 months) renew themselves on every run.
 * If Kakao ever needs a reconnect, an email + Chat message with the setup link is sent once.
 */
const SPACE = 'spaces/AAQAc9NQmJQ';
const ROOM_URL = 'https://chat.google.com/room/AAQAc9NQmJQ';
const RELAY_TAG = '[KakaoRelay]';             // our own notices to the room — never relayed back
const MAX_PER_RUN = 15;                       // burst guard; the rest goes out next minute
const P = PropertiesService.getScriptProperties();

/* ---------------- web app: setup page + Kakao OAuth redirect ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.code) {
    try { exchangeCode_(p.code); installTrigger(); return page_('✅ KakaoTalk connected. Alerts from the GCX room will now arrive in KakaoTalk (나와의 채팅).'); }
    catch (err) { return page_('❌ ' + err.message); }
  }
  if (p.error) return page_('❌ Kakao: ' + p.error + ' ' + (p.error_description || ''));
  return page_('');
}

function page_(msg) {
  const t = HtmlService.createTemplate(
    `<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">
    <style>body{font:15px -apple-system,system-ui,sans-serif;background:#000;color:#ede6d6;max-width:620px;margin:32px auto;padding:0 16px}
    input{width:100%;padding:9px;margin:4px 0 12px;background:#111;color:#ede6d6;border:1px solid #444;border-radius:6px}
    button,a.b{display:inline-block;padding:9px 16px;border-radius:6px;border:0;background:#ede6d6;color:#14110b;font-weight:600;text-decoration:none;cursor:pointer;margin-right:8px}
    code{background:#111;padding:2px 6px;border-radius:4px;word-break:break-all}.m{margin:14px 0;padding:10px;border:1px solid #444;border-radius:6px}</style></head><body>
    <h2>GCX → KakaoTalk alerts</h2>
    <? if (msg) { ?><div class="m"><?= msg ?></div><? } ?>
    <p>Status: <b><?= status.connected ? 'connected' : 'not connected' ?></b> · trigger <?= status.trigger ? 'on' : 'off' ?> · last relayed <?= status.last || '—' ?></p>
    <p>Kakao Redirect URI (register this exactly): <code><?= redirect ?></code></p>
    <label>Kakao REST API key</label><input id="k" value="<?= status.keySet ? '••••••••(saved)' : '' ?>">
    <label>Client secret (only if enabled in the Kakao app)</label><input id="s" value="<?= status.secretSet ? '••••••••(saved)' : '' ?>">
    <button onclick="save()">Save keys</button>
    <? if (status.keySet) { ?><a class="b" href="<?= authUrl ?>">Connect KakaoTalk</a><? } ?>
    <? if (status.connected) { ?><button onclick="google.script.run.withSuccessHandler(r=>alert(r)).withFailureHandler(e=>alert(e.message)).sendTest()">Send test</button><? } ?>
    <script>function save(){const k=document.getElementById('k').value,s=document.getElementById('s').value;
      google.script.run.withSuccessHandler(()=>location.reload()).withFailureHandler(e=>alert(e.message)).saveKeys(k,s);}</script>
    </body></html>`);
  t.msg = msg; t.status = status(); t.redirect = redirectUri_(); t.authUrl = authUrl_();
  return t.evaluate().setTitle('GCX KakaoRelay').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function saveKeys(key, secret) {
  if (key && key.indexOf('•') < 0) P.setProperty('KAKAO_REST_KEY', key.trim());
  if (secret && secret.indexOf('•') < 0) P.setProperty('KAKAO_CLIENT_SECRET', secret.trim());
  if (secret === '') P.deleteProperty('KAKAO_CLIENT_SECRET');
  return true;
}

function status() {
  return {
    keySet: !!P.getProperty('KAKAO_REST_KEY'), secretSet: !!P.getProperty('KAKAO_CLIENT_SECRET'),
    connected: !!P.getProperty('KAKAO_REFRESH'), last: P.getProperty('LAST_RELAYED_AT'),
    trigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'relay'),
  };
}

const redirectUri_ = () => P.getProperty('REDIRECT_URI') || ScriptApp.getService().getUrl();
const authUrl_ = () => 'https://kauth.kakao.com/oauth/authorize?response_type=code&scope=talk_message&prompt=consent'   // always show the consent screen (talk_message is optional — must be ticked)
  + '&client_id=' + encodeURIComponent(P.getProperty('KAKAO_REST_KEY') || '')
  + '&redirect_uri=' + encodeURIComponent(redirectUri_());

/* ---------------- Kakao tokens ---------------- */
function tokenCall_(form) {
  form.client_id = P.getProperty('KAKAO_REST_KEY');
  const sec = P.getProperty('KAKAO_CLIENT_SECRET'); if (sec) form.client_secret = sec;
  const r = UrlFetchApp.fetch('https://kauth.kakao.com/oauth/token', { method: 'post', payload: form, muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText() || '{}');
  if (r.getResponseCode() !== 200) throw new Error('Kakao token ' + r.getResponseCode() + ': ' + (j.error_description || j.error || r.getContentText()));
  P.setProperty('KAKAO_ACCESS', j.access_token);
  P.setProperty('KAKAO_ACCESS_EXP', String(Date.now() + (j.expires_in - 300) * 1000));
  if (j.refresh_token) P.setProperty('KAKAO_REFRESH', j.refresh_token);    // Kakao rotates it when < 1 month left
  P.deleteProperty('REAUTH_NOTIFIED');
  return j.access_token;
}
function exchangeCode_(code) { return tokenCall_({ grant_type: 'authorization_code', redirect_uri: redirectUri_(), code }); }
function accessToken_() {
  if (P.getProperty('KAKAO_ACCESS') && Date.now() < Number(P.getProperty('KAKAO_ACCESS_EXP') || 0)) return P.getProperty('KAKAO_ACCESS');
  if (!P.getProperty('KAKAO_REFRESH')) throw new Error('not connected');
  return tokenCall_({ grant_type: 'refresh_token', refresh_token: P.getProperty('KAKAO_REFRESH') });
}

/* ---------------- send to KakaoTalk (나에게 보내기) ---------------- */
function kakaoSend_(text, url, title) {
  const tpl = { object_type: 'text', text: text.length > 200 ? text.slice(0, 199) + '…' : text,
    link: { web_url: url || ROOM_URL, mobile_web_url: url || ROOM_URL }, button_title: title || 'Chat에서 보기' };
  const go = () => UrlFetchApp.fetch('https://kapi.kakao.com/v2/api/talk/memo/default/send', {
    method: 'post', headers: { Authorization: 'Bearer ' + accessToken_() },
    payload: { template_object: JSON.stringify(tpl) }, muteHttpExceptions: true });
  let r = go();
  if (r.getResponseCode() === 401) { P.deleteProperty('KAKAO_ACCESS'); r = go(); }
  if (r.getResponseCode() !== 200) throw new Error('Kakao send ' + r.getResponseCode() + ': ' + r.getContentText());
}

function sendTest() { kakaoSend_('🔵 GCX KakaoRelay 테스트 — 연결 완료. GCX 알림이 여기로 옵니다.'); return 'Sent — check KakaoTalk (나와의 채팅).'; }

/* ---------------- relay loop (every minute) ---------------- */
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'relay').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('relay').timeBased().everyMinutes(1).create();
  if (!P.getProperty('LAST_CREATE')) P.setProperty('LAST_CREATE', new Date().toISOString());   // start from now, no backlog flood
}

function relay() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return;
  try {
    if (!P.getProperty('KAKAO_REFRESH')) return;
    let since = P.getProperty('LAST_CREATE') || new Date().toISOString();
    const q = 'filter=' + encodeURIComponent('createTime > "' + since + '"') + '&pageSize=' + MAX_PER_RUN;
    const r = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + SPACE + '/messages?' + q,
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Chat list ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
    const msgs = (JSON.parse(r.getContentText()).messages || []);
    for (const m of msgs) {
      const isBot = m.sender && m.sender.type === 'BOT';
      const text = messageText_(m);
      if (isBot && text && text.indexOf(RELAY_TAG) < 0) {
        // a message that asks for something (OTP page, re-login page) → the Kakao button opens THAT page
        const act = actionUrl_(m);
        try { act ? kakaoSend_(text, act, /otp/i.test(act) ? 'OTP 입력하기' : '로그인 / 열기') : kakaoSend_(text, threadUrl_(m)); }
        catch (err) { if (/token|401|invalid_grant|not connected/i.test(err.message)) { needReauth_(err); return; } throw err; }
      }
      since = m.createTime; P.setProperty('LAST_CREATE', since);
      P.setProperty('LAST_RELAYED_AT', Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm'));
    }
  } finally { lock.releaseLock(); }
}

// card/text → plain lines: header title, subtitle, then every text-ish widget field
function messageText_(m) {
  const out = [];
  const clean = s => String(s).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").trim();
  const walk = o => {
    if (!o || typeof o !== 'object') return;
    if (Array.isArray(o)) return o.forEach(walk);
    for (const k of Object.keys(o)) {
      const v = o[k];
      if (typeof v === 'string') {
        if (/^(title|subtitle|text|topLabel|bottomLabel)$/.test(k)) { const c = clean(v); if (c && out.indexOf(c) < 0) out.push(c); }
      } else if (k !== 'onClick' && k !== 'button' && k !== 'buttonList') walk(v);
    }
  };
  (m.cardsV2 || []).forEach(c => walk(c.card));
  (m.cards || []).forEach(walk);
  if (m.text) { const c = clean(m.text); if (c && out.indexOf(c) < 0) out.unshift(c); }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

// first openLink button in a card (e.g. "Re-login amazon:de"), else the first URL in the text (e.g. the OTP page)
function actionUrl_(m) {
  let found = null;
  const walk = o => {
    if (found || !o || typeof o !== 'object') return;
    if (Array.isArray(o)) return o.forEach(walk);
    if (o.openLink && o.openLink.url && !/chat\.google\.com/.test(o.openLink.url)) { found = o.openLink.url; return; }
    Object.keys(o).forEach(k => walk(o[k]));
  };
  (m.cardsV2 || []).forEach(c => walk(c.card));
  if (!found && m.text) { const u = String(m.text).match(/https?:\/\/[^\s<>|]+/); if (u) found = u[0]; }
  return found;
}

const threadUrl_ = m => m.thread && m.thread.name
  ? ROOM_URL + '/' + m.thread.name.split('/').pop() : ROOM_URL;

function needReauth_(err) {
  if (P.getProperty('REAUTH_NOTIFIED')) return;
  P.setProperty('REAUTH_NOTIFIED', '1');
  const url = ScriptApp.getService().getUrl();
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'GCX KakaoRelay: reconnect KakaoTalk',
    'KakaoTalk alerts stopped: ' + err.message + '\n\nReconnect (one click): ' + url);
}
