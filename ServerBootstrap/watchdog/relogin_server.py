#!/usr/bin/env python3
"""GCX Server re-login page — systemd service gcx-relogin on gcx-server.

Listens ONLY on the server's Tailscale address (port 8787), so it is reachable from the user's
iPhone/Mac on Tailscale and nowhere else. Links come from gcx_alert.new_relogin_link(): one per
alert, single-use, 24 h expiry. The user types email + password (and the authenticator code if
Amazon asks); the server signs in on the right Chrome profile. Nothing is stored: the password
lives only in memory for the duration of the sign-in.

Targets:  amazon:<tld>          Amazon customer account (photo check / scraper image fetch)
          sellercentral:<US|EU|JP|IN>   Seller Central (SC scraper)
"""
import asyncio, html, json, os, subprocess, time
from aiohttp import web

TOKENS = os.path.expanduser("~/.gcx-relogin/tokens.json")
HOME = os.path.expanduser("~")
BASE_PROFILE = f"{HOME}/.chrome-phaseg-profile"
SC_URL = {"US": "https://sellercentral.amazon.com/home", "EU": "https://sellercentral-europe.amazon.com/home",
          "JP": "https://sellercentral-japan.amazon.com/home", "IN": "https://sellercentral.amazon.in/home"}
jobs = {}   # token -> {"state", "msg", "otp_event", "otp"}

CSS = ("body{font-family:-apple-system,system-ui,sans-serif;background:#0b0d17;color:#e9eaff;max-width:440px;margin:40px auto;padding:0 18px}"
       "input,button{font-size:18px;width:100%;padding:12px;margin:6px 0 14px;border-radius:10px;border:1px solid #2a3060;"
       "background:#161a33;color:#eef0ff;box-sizing:border-box}button{background:#5b6cff;border:0;font-weight:600}"
       ".m{color:#8c93b8;font-size:14px}.ok{color:#3ddc97}.bad{color:#ff5c7a}")


def page(body, refresh=False):
    r = '<meta http-equiv="refresh" content="3">' if refresh else ""
    return web.Response(text=f'<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">{r}'
                             f"<style>{CSS}</style><h2>GCX Server · re-login</h2>{body}", content_type="text/html")


def token_target(tok):
    try:
        t = json.load(open(TOKENS)).get(tok)
    except Exception:
        return None
    return t["target"] if t and t.get("exp", 0) > time.time() and not t.get("used") else None


def mark_used(tok):
    d = json.load(open(TOKENS)); d.setdefault(tok, {})["used"] = True
    json.dump(d, open(TOKENS, "w")); os.chmod(TOKENS, 0o600)


def profile_and_url(target):
    kind, _, arg = target.partition(":")
    if kind == "amazon":
        return BASE_PROFILE, f"https://www.amazon.{arg}/gp/css/homepage.html"
    if kind == "sellercentral":
        per = f"{BASE_PROFILE}_{arg}"
        return (per if os.path.isdir(per) else BASE_PROFILE), SC_URL[arg]
    raise ValueError(target)


async def do_login(tok, target, email, password):
    j = jobs[tok]
    from playwright.async_api import async_playwright
    prof, url = profile_and_url(target)
    if subprocess.run(["pgrep", "-f", f"user-data-dir={prof}"], capture_output=True).returncode == 0:
        j.update(state="failed", msg="That Chrome profile is in use right now (a scrape or photo check is running, "
                                      "or the login window is open). Close it or wait a few minutes, then reopen this link.")
        return
    os.environ.setdefault("DISPLAY", ":0")
    try:
        async with async_playwright() as pw:
            ctx = await pw.chromium.launch_persistent_context(prof, channel="chrome", headless=False)
            p = ctx.pages[0] if ctx.pages else await ctx.new_page()
            await p.goto(url, wait_until="domcontentloaded", timeout=45000)
            j["msg"] = "entering email…"
            if await p.locator("#ap_email").count():
                await p.fill("#ap_email", email); await p.click("#continue"); await p.wait_for_load_state("domcontentloaded")
            j["msg"] = "entering password…"
            if await p.locator("#ap_password").count():
                await p.fill("#ap_password", password)
                if await p.locator("input[name='rememberMe']").count():
                    await p.check("input[name='rememberMe']")
                await p.click("#signInSubmit"); await p.wait_for_load_state("domcontentloaded")
            otp = "#auth-mfa-otpcode, input[name='otpCode']"
            try:
                await p.wait_for_selector(otp, timeout=12000)
                j.update(state="otp", msg="Amazon is asking for the authenticator code.")
                await asyncio.wait_for(j["otp_event"].wait(), timeout=300)
                await p.fill(otp, j["otp"])
                for sel in ("#auth-mfa-remember-device", "input[name='rememberDevice']"):
                    if await p.locator(sel).count():
                        await p.check(sel); break
                await p.click("#auth-signin-button, input[type='submit']"); await p.wait_for_load_state("domcontentloaded")
                j.update(state="working", msg="checking…")
            except asyncio.TimeoutError:
                j.update(state="failed", msg="No code entered within 5 minutes."); await ctx.close(); return
            except Exception:
                pass                                            # no OTP step on this device
            await p.wait_for_timeout(4000)
            if any(x in p.url for x in ("/ap/signin", "/ap/mfa", "/ap/cvf")):
                j.update(state="failed", msg=f"Still on Amazon's sign-in page ({html.escape(p.url[:60])}…). "
                                              "Wrong password, a captcha, or an extra check — use VNC to finish it on the server screen.")
            else:
                j.update(state="done", msg=f"Signed in — {html.escape(target)} is ready."); mark_used(tok)
            await ctx.close()
    except Exception as e:
        j.update(state="failed", msg=html.escape(str(e)[:200]))


async def get_form(req):
    tok = req.match_info["tok"]; target = token_target(tok)
    if tok in jobs:
        raise web.HTTPFound(f"/r/{tok}/status")
    if not target:
        return page('<p class="bad">This link has expired or was already used.</p>')
    return page(f'<p>Sign in again: <b>{html.escape(target)}</b></p>'
                f'<form method="post"><label class="m">Email</label><input name="email" type="email" autocomplete="username" required>'
                f'<label class="m">Password</label><input name="password" type="password" autocomplete="current-password" required>'
                f'<button>Sign in on the server</button></form>'
                f'<p class="m">Used once for this sign-in, never saved. If Amazon asks for a code, this page will ask you for it.</p>')


async def post_form(req):
    tok = req.match_info["tok"]; target = token_target(tok)
    if not target or tok in jobs:
        raise web.HTTPFound(f"/r/{tok}/status" if tok in jobs else f"/r/{tok}")
    d = await req.post()
    jobs[tok] = {"state": "working", "msg": "opening Chrome on the server…", "otp_event": asyncio.Event(), "otp": None}
    asyncio.create_task(do_login(tok, target, str(d.get("email", "")).strip(), str(d.get("password", ""))))
    raise web.HTTPFound(f"/r/{tok}/status")


async def status(req):
    tok = req.match_info["tok"]; j = jobs.get(tok)
    if not j:
        raise web.HTTPFound(f"/r/{tok}")
    if j["state"] == "otp":
        return page(f'<p>{j["msg"]}</p><form method="post" action="/r/{tok}/otp"><label class="m">6-digit code from Google Authenticator</label>'
                    f'<input name="code" inputmode="numeric" pattern="[0-9]{{6}}" maxlength="6" autofocus required><button>Submit code</button></form>')
    cls = {"done": "ok", "failed": "bad"}.get(j["state"], "m")
    return page(f'<p class="{cls}">{j["msg"]}</p>', refresh=j["state"] == "working")


async def post_otp(req):
    tok = req.match_info["tok"]; j = jobs.get(tok); d = await req.post()
    code = str(d.get("code", "")).strip()
    if j and j["state"] == "otp" and code.isdigit() and len(code) == 6:
        j.update(otp=code, state="working", msg="submitting code…"); j["otp_event"].set()
    raise web.HTTPFound(f"/r/{tok}/status")


def tailscale_ip():
    return subprocess.run(["tailscale", "ip", "-4"], capture_output=True, text=True).stdout.split()[0]


if __name__ == "__main__":
    app = web.Application()
    app.add_routes([web.get("/r/{tok}", get_form), web.post("/r/{tok}", post_form),
                    web.get("/r/{tok}/status", status), web.post("/r/{tok}/otp", post_otp)])
    web.run_app(app, host=tailscale_ip(), port=8787, print=lambda *_: None)
