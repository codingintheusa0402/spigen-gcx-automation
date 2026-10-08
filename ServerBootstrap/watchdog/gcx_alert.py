#!/usr/bin/env python3
"""Post a GCX Server alert card to the GCX Server Chat room.

Webhook lives outside git: ~/.config/gcx_server_webhook.txt (one line).
Optional re-login link: pass --relogin <target> and a one-time link to the re-login page
(served by relogin_server.py, Tailscale-only) is attached.

  gcx_alert.py "Title" "body text" [--level bad|warn|ok] [--relogin amazon:de|sellercentral:US]
  python: from gcx_alert import alert; alert("Title", "text", level="bad", relogin="amazon:com")
"""
import argparse, json, os, secrets, socket, time, urllib.request

HOOK_FILE = os.path.expanduser("~/.config/gcx_server_webhook.txt")
TOKENS = os.path.expanduser("~/.gcx-relogin/tokens.json")
RELOGIN_BASE = "http://gcx-server:8787"          # Tailscale MagicDNS name of the server's Ubuntu
ICON = {"bad": "🔴", "warn": "🟡", "ok": "🟢", "info": "🔵"}


def new_relogin_link(target, ttl_hours=24):
    os.makedirs(os.path.dirname(TOKENS), mode=0o700, exist_ok=True)
    try:
        data = json.load(open(TOKENS))
    except Exception:
        data = {}
    now = time.time()
    data = {k: v for k, v in data.items() if v.get("exp", 0) > now}      # drop expired
    tok = secrets.token_urlsafe(18)
    data[tok] = {"target": target, "exp": now + ttl_hours * 3600}
    with open(TOKENS, "w") as f:
        json.dump(data, f)
    os.chmod(TOKENS, 0o600)
    return f"{RELOGIN_BASE}/r/{tok}"


def alert(title, text, level="bad", relogin=None, host=None):
    hook = open(HOOK_FILE).read().strip()
    host = host or ("gcx-server" if os.path.exists("/proc/version") and "microsoft" in open("/proc/version").read().lower()
                    else socket.gethostname())
    widgets = [{"textParagraph": {"text": text}}]
    if relogin:
        link = new_relogin_link(relogin)
        widgets.append({"textParagraph": {"text": "<i>Opens a page on the server (Tailscale only — iPhone/Mac must be on Tailscale). "
                                                  "Your password is used once to sign in and is not saved.</i>"}})
        widgets.append({"buttonList": {"buttons": [{"text": f"Re-login {relogin}", "onClick": {"openLink": {"url": link}}}]}})
    card = {"cardsV2": [{"cardId": f"gcx-{int(time.time())}", "card": {
        "header": {"title": f"{ICON.get(level, '🔵')} {title}",
                   "subtitle": f"GCX Server · {host} · {time.strftime('%m-%d %H:%M')}"},
        "sections": [{"widgets": widgets}]}}]}
    req = urllib.request.Request(hook, json.dumps(card).encode(), {"Content-Type": "application/json; charset=UTF-8"})
    return urllib.request.urlopen(req, timeout=15).status


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("title"); ap.add_argument("text")
    ap.add_argument("--level", default="bad"); ap.add_argument("--relogin")
    a = ap.parse_args()
    print(alert(a.title, a.text, a.level, a.relogin))
