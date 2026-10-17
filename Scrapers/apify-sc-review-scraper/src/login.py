"""Login mode: the user signs in to Seller Central THEMSELVES in a real browser shown in Apify's Live View tab.

How it works
  · The actor opens Chromium on <region host>/brand-customer-reviews/ (Amazon redirects to its sign-in page).
  · A tiny web page is served on the actor's container web server, which Apify shows in the run's **Live View**
    tab: a live picture of the browser. Clicking the picture clicks the page; the text box types into it
    (email, password, the two-factor code); buttons send Enter / Tab / Backspace, scroll, back and reload.
  · As soon as the browser is on Seller Central (not a sign-in, 2-step or account-picker page), the session
    cookies are saved to the user's own named key-value store (see session_store.py), one record per region,
    and the next region opens. Later scrape runs reuse them until Amazon expires them.

Nothing the user types is stored or logged — it goes straight to the browser as keystrokes.
"""
import asyncio
import json
import os
import time

from aiohttp import web
from apify import Actor

from . import session_store
from .sessions import fingerprint

LOGIN_MARKERS = ('/ap/signin', '/ap/mfa', '/ap/cvf', '/ap/challenge', 'signin', '/mfa', 'validateCaptcha')
VIEW = {'width': 1280, 'height': 820}

PAGE = """<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Seller Central login</title><style>
 body{margin:0;font:14px/1.4 -apple-system,Segoe UI,Roboto,sans-serif;background:#f4f5f7;color:#1d1f23}
 header{padding:10px 14px;background:#232f3e;color:#fff;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
 header b{font-size:15px} #st{opacity:.9}
 .ok{color:#7ee2a8}.warn{color:#ffd479}
 main{padding:10px 14px} #wrap{position:relative;max-width:1280px}
 #s{width:100%;max-width:1280px;border:1px solid #c9ccd1;border-radius:6px;cursor:crosshair;background:#fff;display:block}
 .bar{display:flex;gap:6px;flex-wrap:wrap;margin:8px 0;max-width:1280px}
 input[type=text],input[type=password]{flex:1;min-width:220px;padding:8px 10px;font-size:15px;border:1px solid #9aa0a6;border-radius:6px}
 button{padding:8px 12px;font-size:14px;border:1px solid #9aa0a6;background:#fff;border-radius:6px;cursor:pointer}
 button.p{background:#ff9900;border-color:#e88b00;font-weight:600}
 ol{margin:6px 0 0 18px;padding:0;max-width:1280px} small{color:#5f6368}
</style></head><body>
<header><b>🔐 Seller Central login · <span id="rg">…</span></b><span id="st">starting…</span></header>
<main>
<ol>
 <li>Click inside the picture to focus a field, type in the box below and press <b>Type</b> (or Enter).</li>
 <li>Sign in as usual — email, password, then the two-factor code. Tick <i>Don't require OTP on this browser</i> if offered.</li>
 <li>If Seller Central asks <i>Select an account</i>, click your account and marketplace. Set the language to <b>English</b> (top right) if it isn't.</li>
 <li>When you're on Seller Central, the session is saved automatically and the next region opens.</li>
</ol>
<div class="bar"><input id="t" type="password" autocomplete="off" placeholder="Text to type into the browser (hidden)">
 <button class="p" onclick="typeIt()">Type</button><button onclick="vis()">👁</button>
 <button onclick="key('Enter')">Enter ⏎</button><button onclick="key('Tab')">Tab</button><button onclick="key('Backspace')">⌫</button>
 <button onclick="act('scroll',{dy:500})">Scroll ↓</button><button onclick="act('scroll',{dy:-500})">Scroll ↑</button>
 <button onclick="act('back')">◀ Back</button><button onclick="act('reload')">⟳ Reload</button>
 <button onclick="if(confirm('Skip this region?'))act('skip')">Skip region</button></div>
<div id="wrap"><img id="s" alt="browser"></div>
<small>Your typing goes straight to the browser as keystrokes and is never stored or logged. Only the resulting
login cookies are saved, privately, in your own Apify account (key-value store <code>%STORE%</code>).</small>
</main><script>
const s=document.getElementById('s'),t=document.getElementById('t');
async function act(a,b={}){await fetch('/'+a,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});tick(true)}
function key(k){act('key',{key:k})}
function typeIt(){if(!t.value)return;act('type',{text:t.value});t.value=''}
function vis(){t.type=t.type==='password'?'text':'password'}
t.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();if(t.value)typeIt();else key('Enter')}});
s.addEventListener('click',e=>{const r=s.getBoundingClientRect();
 act('click',{x:(e.clientX-r.left)*s.naturalWidth/r.width,y:(e.clientY-r.top)*s.naturalHeight/r.height})});
s.addEventListener('wheel',e=>{e.preventDefault();act('scroll',{dy:e.deltaY})},{passive:false});
let busy=false;
async function tick(now){if(busy&&!now)return;busy=true;
 try{const st=await (await fetch('/state')).json();
  document.getElementById('rg').textContent=st.region||'done';
  const el=document.getElementById('st');el.textContent=st.message;el.className=st.level||'';
  if(!st.finished){const img=new Image();img.onload=()=>{s.src=img.src};img.src='/shot.jpg?'+Date.now()}}
 catch(e){} busy=false}
setInterval(tick,900);tick(true);
</script></body></html>"""


class RemoteBrowser:
    def __init__(self, ctx, inp):
        self.ctx, self.inp = ctx, inp
        self.page = None
        self.lock = asyncio.Lock()
        self.shot, self.shot_t = b'', 0.0
        self.state = {'region': '', 'message': 'starting…', 'level': '', 'finished': False}
        self.skip = False
        ctx.on('page', self._on_page)   # sign-in popups / new tabs: follow the newest

    def _on_page(self, p):
        self.page = p

    # ── HTTP handlers ───────────────────────────────────────────────
    async def index(self, _):
        return web.Response(text=PAGE.replace('%STORE%', session_store.store_name(self.inp)), content_type='text/html')

    async def state_h(self, _):
        return web.json_response(self.state)

    async def shot_h(self, _):
        async with self.lock:
            if time.time() - self.shot_t > 0.5 and self.page and not self.page.is_closed():
                try:
                    self.shot = await self.page.screenshot(type='jpeg', quality=60, timeout=8000)
                    self.shot_t = time.time()
                except Exception:
                    pass
        return web.Response(body=self.shot, content_type='image/jpeg', headers={'Cache-Control': 'no-store'})

    async def action(self, req):
        name = req.match_info['a']
        try:
            body = await req.json()
        except Exception:
            body = {}
        p = self.page
        if not p or p.is_closed():
            return web.json_response({'ok': False})
        try:
            async with self.lock:
                if name == 'click':
                    await p.mouse.click(float(body['x']), float(body['y']))
                elif name == 'type':
                    await p.keyboard.type(str(body.get('text', ''))[:500], delay=40)   # never logged
                elif name == 'key' and body.get('key') in ('Enter', 'Tab', 'Backspace', 'Escape'):
                    await p.keyboard.press(body['key'])
                elif name == 'scroll':
                    await p.mouse.wheel(0, max(-2000, min(2000, float(body.get('dy', 0)))))
                elif name == 'back':
                    await p.go_back(timeout=15000)
                elif name == 'reload':
                    await p.reload(timeout=30000)
                elif name == 'skip':
                    self.skip = True
                self.shot_t = 0
        except Exception as e:
            Actor.log.debug(f'remote action {name}: {e.__class__.__name__}')
        return web.json_response({'ok': True})

    def app(self):
        a = web.Application()
        a.router.add_get('/', self.index)
        a.router.add_get('/state', self.state_h)
        a.router.add_get('/shot.jpg', self.shot_h)
        a.router.add_post('/{a}', self.action)
        return a

    def say(self, msg, level=''):
        self.state.update(message=msg, level=level)


def on_login_page(url: str) -> bool:
    return any(m in url for m in LOGIN_MARKERS)


async def signed_in(page, hosts) -> bool:
    """On the region's Seller Central (any of its hosts), past sign-in / 2-step / account picker."""
    url = page.url
    host = url.split('/')[2] if '://' in url else ''
    if host not in hosts or on_login_page(url):
        return False
    try:
        if await page.get_by_text('Select an account', exact=False).count():
            return False
    except Exception:
        return False
    return True


async def run(inp: dict, regions: dict[str, str], launch_browser) -> dict:
    """regions: {region: host}. launch_browser(): coroutine returning (browser, ctx). Returns {region: result}."""
    port = int(os.environ.get('ACTOR_WEB_SERVER_PORT') or os.environ.get('APIFY_CONTAINER_PORT') or 4321)
    url = os.environ.get('ACTOR_WEB_SERVER_URL') or os.environ.get('APIFY_CONTAINER_URL') or f'http://localhost:{port}'
    timeout = max(2, min(30, int(inp.get('loginTimeoutMinutes') or 10))) * 60
    browser, ctx = await launch_browser()
    rb = RemoteBrowser(ctx, inp)
    runner = web.AppRunner(rb.app())
    await runner.setup()
    await web.TCPSite(runner, '0.0.0.0', port).start()
    Actor.log.info(f'Login mode: open the run\'s "Live View" tab (or {url}) and sign in.')
    results = {}
    try:
        for region, host in regions.items():
            rb.skip = False
            rb.state['region'] = region
            rb.page = await ctx.new_page()
            await Actor.set_status_message(f'👉 Open the Live View tab and sign in to Seller Central {region}')
            rb.say(f'Sign in to Seller Central {region}…')
            try:
                await rb.page.goto(f'https://{host}/brand-customer-reviews/', wait_until='domcontentloaded', timeout=60000)
            except Exception as e:
                Actor.log.warning(f'[{region}] could not open Seller Central ({e.__class__.__name__}) — use Reload in Live View')
            hosts = {host} | {d for d in session_store.REGION_DOMAINS[region] if d.startswith('sellercentral')}
            t0, ok = time.time(), False
            while time.time() - t0 < timeout and not rb.skip:
                await asyncio.sleep(1.5)
                if await signed_in(rb.page, hosts):
                    await asyncio.sleep(2)   # let Amazon finish setting cookies / redirects
                    if await signed_in(rb.page, hosts):
                        ok = True
                        break
                left = int(timeout - (time.time() - t0))
                if left % 60 < 2:
                    await Actor.set_status_message(f'👉 Waiting for you to sign in to {region} in Live View · {left // 60} min left')
            if not ok:
                results[region] = {'status': 'SKIPPED' if rb.skip else 'TIMED_OUT'}
                rb.say(f'{region}: {"skipped" if rb.skip else "timed out"} — no session saved', 'warn')
                Actor.log.warning(f'[{region}] login {results[region]["status"].lower().replace("_", " ")} — no session saved')
                await rb.page.close()
                continue
            # Confirm the landing page works and is in English (the scraper reads the English page)
            try:
                await rb.page.goto(f'https://{host}/brand-customer-reviews/', wait_until='domcontentloaded', timeout=45000)
            except Exception:
                pass
            lang = ''
            try:
                lang = await rb.page.evaluate("document.documentElement.lang || ''")
            except Exception:
                pass
            cookies = await ctx.cookies()
            n = await session_store.save(inp, region, cookies, 'login-mode')
            results[region] = {'status': 'SAVED', 'cookies': n, 'language': lang or '?'}
            note = '' if not lang or lang.lower().startswith('en') else \
                f' ⚠️ Seller Central language is "{lang}" — switch it to English before scraping'
            rb.say(f'✅ {region} session saved ({n} cookies).{note}', 'ok' if not note else 'warn')
            Actor.log.info(f'[{region}] session saved to store "{session_store.store_name(inp)}": '
                           f'{fingerprint(session_store.region_cookies(cookies, region))}{note}')
            await asyncio.sleep(2.5)
            await rb.page.close()
        saved = [r for r, v in results.items() if v['status'] == 'SAVED']
        rb.state.update(region='', finished=True, level='ok' if saved else 'warn',
                        message=(f'Done — saved: {", ".join(saved)}. Start the actor in Scrape mode now; it reuses these '
                                 f'sessions until Amazon expires them.' if saved else 'Done — no session was saved.'))
        await asyncio.sleep(4)   # let the Live View show the final message
    finally:
        await runner.cleanup()
        await browser.close()
    await Actor.set_value('LOGIN_RESULT', results)
    return results


def summary(results: dict) -> str:
    saved = [r for r, v in results.items() if v['status'] == 'SAVED']
    other = {r: v['status'] for r, v in results.items() if v['status'] != 'SAVED'}
    if not saved:
        return f'⚠️ No session saved ({json.dumps(other)}). Run Login mode again and sign in within the time limit.'
    return (f'✅ Saved sessions: {", ".join(saved)}. Now run in Scrape mode — no session pasting needed.'
            + (f' Not saved: {other}' if other else ''))
