#!/usr/bin/env python3
"""Daily health check for the public Apify actor cosmic_meteorite/amazon-asin-rating-scraper.

Runs on the GCX server (cron, 10:00 KST). Checks that Apify hasn't flagged the actor, that the
latest build succeeded, and that real test runs (amazon.com default input + amazon.de) return
valid ratings. On failure it alerts the GCX Server Chat room and, with --autofix, starts a
headless Claude session that fixes the code and deploys it as a new actor version, then re-checks.

  health_check.py [--autofix] [--no-alert]
"""
import argparse, json, os, subprocess, sys, time, urllib.request
from datetime import datetime

ACTOR = '2AuZCyLxKZuO2OyNY'
CONF = os.path.expanduser('~/.config/apify_asin_scraper')
TOKEN = open(f'{CONF}/token').read().strip()
LOG = f'{CONF}/health.log'
HERE = os.path.dirname(os.path.abspath(__file__))
API = 'https://api.apify.com/v2'
CONSOLE = f'https://console.apify.com/actors/{ACTOR}'
TESTS = [  # (input, min valid items)
    ({}, 2),  # default/prefill input — the same input Apify's daily automated test uses
    ({'asins': ['B0GDHRZFWH'], 'marketplace': 'de'}, 1),
]

sys.path.insert(0, os.path.join(HERE, '..', '..', 'ServerBootstrap', 'watchdog'))


def api(path, data=None, timeout=60):
    url = f'{API}{path}{"&" if "?" in path else "?"}token={TOKEN}'
    req = urllib.request.Request(url, data=json.dumps(data).encode() if data is not None else None,
                                 headers={'Content-Type': 'application/json'}, method='POST' if data is not None else 'GET')
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


def log(msg):
    line = f'{datetime.now():%F %T} {msg}'
    print(line)
    with open(LOG, 'a') as f:
        f.write(line + '\n')


def check():
    problems = []
    try:
        d = api(f'/acts/{ACTOR}')['data']
        if not d.get('isPublic'):
            problems.append('Actor is no longer public')
        if d.get('isDeprecated'):
            problems.append('Actor is marked deprecated')
        if d.get('isUnderMaintenance') or d.get('notice') not in (None, 'NONE'):
            problems.append(f"Apify flagged the actor (isUnderMaintenance={d.get('isUnderMaintenance')}, notice={d.get('notice')})")
        b = api(f'/acts/{ACTOR}/builds?desc=1&limit=1')['data']['items']
        if b and b[0]['status'] != 'SUCCEEDED':
            problems.append(f"Latest build {b[0].get('buildNumber')} is {b[0]['status']}")
    except Exception as e:
        problems.append(f'Apify API error: {e}')
    for inp, need in TESTS:
        label = inp.get('marketplace', 'com') + (' (default input)' if not inp else '')
        try:
            items = api(f'/acts/{ACTOR}/run-sync-get-dataset-items?timeout=280', inp, timeout=320)
            ok = [i for i in items if i.get('found') and i.get('rating') and i.get('ratingsCount')]
            if len(ok) < need:
                problems.append(f'Test run amazon.{label}: {len(ok)}/{need} valid items — {json.dumps(items, ensure_ascii=False)[:300]}')
        except Exception as e:
            problems.append(f'Test run amazon.{label} failed: {e}')
    return problems


def alert(title, text, level):
    try:
        from gcx_alert import alert as send
        send(title, text, level=level)
    except Exception as e:
        log(f'alert failed: {e}')


FIX_PROMPT = """The daily health check of the public Apify actor in this folder (cosmic_meteorite/amazon-asin-rating-scraper, \
actor id {actor}) failed:

{problems}

Fix it and deploy the fix as a NEW actor version:
1. Diagnose: read src/main.py, the latest run logs (Apify API, token in ~/.config/apify_asin_scraper/token), and if useful \
run the actor with {{"debugHtml": true}} to inspect what Amazon returns. If the cause is outside the code (Apify outage, \
account/billing issue), change nothing, and explain.
2. Fix only files in this folder. Keep the output fields and input schema backward compatible.
3. Bump "version" in .actor/actor.json by 0.1 (e.g. 0.2 -> 0.3).
4. Deploy ONLY with: HOME=$HOME/.config/apify_asin_scraper/clihome apify push --force </dev/null
   (that login is cosmic_meteorite; never use any other Apify account or token).
5. Verify with real runs of the new build (default input, and {{"asins":["B0GDHRZFWH"],"marketplace":"de"}}) that ratings come back.
6. git commit the change in this folder with a message starting "apify-asin-rating-scraper vX.Y:" (do not push; autosync does).
Never change pricing, never make the actor private, never run it from any account other than cosmic_meteorite.
End with a 3-line summary: cause, fix, new version."""


def autofix(problems):
    stamp = f'{CONF}/autofix_{datetime.now():%F}'
    if os.path.exists(stamp):
        log('autofix already attempted today — skipping')
        return None
    open(stamp, 'w').close()
    log('starting autofix Claude session')
    env = dict(os.environ, PATH=f"{os.path.expanduser('~/.venvs/apify-asin/bin')}:{os.path.expanduser('~/.local/bin')}:/usr/local/bin:/usr/bin:/bin")
    try:
        r = subprocess.run(['claude', '-p', FIX_PROMPT.format(actor=ACTOR, problems='\n'.join(f'- {p}' for p in problems)),
                            '--dangerously-skip-permissions'], cwd=HERE, env=env, stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=2700)
        out = (r.stdout or r.stderr).strip()
    except subprocess.TimeoutExpired:
        out = 'Claude autofix session timed out after 45 min'
    log(f'autofix output: {out[-1500:]}')
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--autofix', action='store_true')
    ap.add_argument('--no-alert', action='store_true')
    a = ap.parse_args()
    problems = check()
    if not problems:
        log('OK — public, not flagged, build ok, test runs valid')
        return
    log('FAIL — ' + ' | '.join(problems))
    if not a.no_alert:
        alert('Apify ASIN scraper: health check FAILED',
              '\n'.join(f'• {p}' for p in problems) + f'\n{"Starting auto-fix…" if a.autofix else ""}\n{CONSOLE}', 'bad')
    if not a.autofix:
        sys.exit(1)
    summary = autofix(problems)
    if summary is None:
        sys.exit(1)
    time.sleep(30)
    after = check()
    log('after autofix: ' + ('OK' if not after else 'STILL FAILING — ' + ' | '.join(after)))
    if not a.no_alert:
        alert('Apify ASIN scraper: ' + ('fixed & redeployed' if not after else 'auto-fix did NOT resolve it'),
              (summary or '')[-1200:] + ('' if not after else '\nStill failing:\n' + '\n'.join(f'• {p}' for p in after)) + f'\n{CONSOLE}',
              'ok' if not after else 'bad')
    sys.exit(0 if not after else 1)


if __name__ == '__main__':
    main()
