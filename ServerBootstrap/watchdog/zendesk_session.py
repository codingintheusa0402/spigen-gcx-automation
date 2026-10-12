#!/usr/bin/env python3
"""Zendesk session of the server's ticket-monitor Chrome (~/.chrome-zendesk-profile, CDP 9222).

  zendesk_session.py            print whether the Zendesk window is signed in
  zendesk_session.py --alert    cron, weekdays 08:20 (# mesh:zendesk-session): if signed out, post one
                                🔴 "Zendesk login needed" card with a re-login button (→ KakaoTalk via KakaoRelay)

Signing in is done ONLY by the user, on the Tailscale-only re-login page (relogin_server.py, target
"zendesk"): they type email + password (+ code); login() below drives the already-open Zendesk window
over CDP with what they typed. Nothing is stored. Claude never types the password (hard rule).
"""
import asyncio, json, os, sys, time

CDP = "http://127.0.0.1:9222"
HOST = "https://spigenhelp.zendesk.com"
VIEW = f"{HOST}/agent/filters/360103290632"
STATE = os.path.expanduser("~/.gcx-zendesk-session.json")


async def _page(pw):
    browser = await pw.chromium.connect_over_cdp(CDP)
    ctx = browser.contexts[0]
    pages = [p for p in ctx.pages if "zendesk.com" in p.url]
    page = pages[0] if pages else await ctx.new_page()
    if not pages:
        await page.goto(VIEW, wait_until="domcontentloaded", timeout=45000)
    return browser, page


async def check():
    """(signed_in, who) — asks Zendesk itself (/api/v2/users/me) from inside the Zendesk window."""
    from playwright.async_api import async_playwright
    async with async_playwright() as pw:
        browser, page = await _page(pw)
        try:
            me = await page.evaluate("fetch('/api/v2/users/me.json',{credentials:'include'}).then(r=>r.json())")
            u = (me or {}).get("user") or {}
            return bool(u.get("id")) and u.get("role") != "end-user" and u.get("role") is not None, u.get("email") or ""
        finally:
            await browser.close()          # disconnects only — the Chrome window stays open


async def login(email, password, ask_code, say=print):
    """Sign the open Zendesk window in with what the user typed. ask_code() → awaitable returning the 2FA code."""
    from playwright.async_api import async_playwright
    async with async_playwright() as pw:
        browser, page = await _page(pw)
        try:
            await page.goto(f"{HOST}/access/login", wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(2500)

            async def find(selectors):
                for fr in [page.main_frame] + page.frames:
                    for sel in selectors:
                        try:
                            el = fr.locator(sel).first
                            if await el.count() and await el.is_visible():
                                return el
                        except Exception:
                            pass
                return None
            say("entering email…")
            e = await find(["input[name='user[email]']", "#user_email", "input[type='email']", "input[name='email']"])
            if not e:
                if (await check_page(page)):
                    return True, "Already signed in."
                return False, "Couldn't find Zendesk's sign-in form — use VNC to finish on the server screen."
            await e.fill(email)
            p = await find(["input[name='user[password]']", "#user_password", "input[type='password']"])
            if not p:                                            # two-step form: email first
                b = await find(["button[type='submit']", "input[type='submit']"])
                if b:
                    await b.click(); await page.wait_for_timeout(2500)
                p = await find(["input[name='user[password]']", "#user_password", "input[type='password']"])
            if not p:
                return False, "Zendesk didn't show a password field (SSO?) — use VNC to finish on the server screen."
            say("entering password…")
            await p.fill(password)
            b = await find(["#sign-in-submit-button", "button[type='submit']", "input[type='submit']"])
            await (b.click() if b else p.press("Enter"))
            await page.wait_for_timeout(4000)
            c = await find(["input[name='code']", "input[name='otp']", "input[autocomplete='one-time-code']", "#code",
                            "input[name*='verification']", "input[name*='two_factor']"])
            if c:
                say("Zendesk is asking for the verification code.")
                code = await ask_code()
                if not code:
                    return False, "No code entered within 5 minutes."
                await c.fill(code)
                b = await find(["button[type='submit']", "input[type='submit']"])
                await (b.click() if b else c.press("Enter"))
                await page.wait_for_timeout(4000)
            await page.goto(VIEW, wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(3000)
            ok = await check_page(page)
            return ok, ("Signed in — the ticket monitor resumes on its next run." if ok else
                        f"Still not signed in ({page.url[:60]}…). Wrong password, a captcha or an extra check — use VNC.")
        finally:
            await browser.close()


async def check_page(page):
    me = await page.evaluate("fetch('/api/v2/users/me.json',{credentials:'include'}).then(r=>r.json()).catch(()=>({}))")
    u = (me or {}).get("user") or {}
    return bool(u.get("id")) and u.get("role") not in (None, "end-user")


def main():
    try:
        ok, who = asyncio.run(check())
    except Exception as e:
        print(f"check failed: {e}")                           # Chrome down → self-heal / watchdog handle that
        sys.exit(2)
    print(f"zendesk: {'signed in as ' + who if ok else 'SIGNED OUT'}")
    if "--alert" in sys.argv and not ok:
        try:
            st = json.load(open(STATE))
        except Exception:
            st = {}
        today = time.strftime("%Y-%m-%d")
        if st.get("alerted") != today:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from gcx_alert import alert
            alert("Zendesk login needed (ticket monitor)",
                  "The server's Zendesk window is signed out, so the ticket monitor can't read the pending views. "
                  "Tap the re-login button (iPhone on Tailscale), type the Zendesk email + password (and the code if asked) — "
                  "the server signs the window back in and the monitor resumes by itself.", level="bad", relogin="zendesk")
            json.dump({"alerted": today}, open(STATE, "w"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
