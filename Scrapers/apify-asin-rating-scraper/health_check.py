#!/usr/bin/env python3
"""Daily health check for K-Data Lab's public Apify actors (owner account cosmic_meteorite).

Covers every actor in ACTORS: checks that Apify hasn't flagged it, that the latest build succeeded,
that real test runs behave, and that the Store failure rate stays low (Apify tests public actors
daily and ranks/flags them on it). On failure it alerts the GCX Server Chat room and, with
--autofix, starts a headless Claude session in that actor's folder that fixes the code and deploys
a new actor version, then re-checks.

Runs from the GCX server cron / Mac fallback via health_check_gate.py (10:00 KST, catch-up).

  health_check.py [--autofix] [--no-alert] [--only <slug>]
"""
import argparse, json, os, subprocess, sys, time, urllib.request
from datetime import datetime

CONF = os.path.expanduser('~/.config/apify_asin_scraper')
TOKEN = open(f'{CONF}/token').read().strip()
LOG = f'{CONF}/health.log'
HERE = os.path.dirname(os.path.abspath(__file__))
SCRAPERS = os.path.dirname(HERE)
API = 'https://api.apify.com/v2'
MAX_STORE_FAIL_RATE = 0.10  # of public runs in the last 30 days

sys.path.insert(0, os.path.join(HERE, '..', '..', 'ServerBootstrap', 'watchdog'))


def api(path, data=None, timeout=60, method=None):
    url = f'{API}{path}{"&" if "?" in path else "?"}token={TOKEN}'
    req = urllib.request.Request(url, data=json.dumps(data).encode() if data is not None else None,
                                 headers={'Content-Type': 'application/json'},
                                 method=method or ('POST' if data is not None else 'GET'))
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.load(r)


def run_and_wait(actor, inp, limit=300):
    run = api(f'/acts/{actor}/runs', inp)['data']
    t0 = time.time()
    while run['status'] in ('READY', 'RUNNING') and time.time() - t0 < limit:
        time.sleep(10)
        run = api(f"/actor-runs/{run['id']}")['data']
    return run


def kv_record(store_id, key):
    try:
        return api(f'/key-value-stores/{store_id}/records/{key}')
    except Exception:
        return None


# ── per-actor test suites ──────────────────────────────────────────────────────────────────

def tests_asin(actor):
    problems = []
    for inp, need in (({}, 2), ({'asins': ['B0GDHRZFWH'], 'marketplace': 'de'}, 1)):
        label = inp.get('marketplace', 'com') + (' (default input)' if not inp else '')
        try:
            items = api(f'/acts/{actor}/run-sync-get-dataset-items?timeout=280', inp, timeout=320)
            ok = [i for i in items if i.get('found') and i.get('rating') and i.get('ratingsCount')]
            if len(ok) < need:
                problems.append(f'Test run amazon.{label}: {len(ok)}/{need} valid items — {json.dumps(items, ensure_ascii=False)[:300]}')
        except Exception as e:
            problems.append(f'Test run amazon.{label} failed: {e}')
    return problems


def tests_sc(actor):
    """No real session is available to the checker, so verify what can be verified:
    the prefill run Apify's daily test uses succeeds, and a canary run with a fake US session reaches
    Seller Central and correctly reports SESSION_EXPIRED (browser + proxy + login detection work)."""
    problems = []
    try:
        run = run_and_wait(actor, {})
        if run['status'] != 'SUCCEEDED':
            problems.append(f"Prefill run (Apify's daily test input) ended {run['status']}: {run.get('statusMessage')}")
    except Exception as e:
        problems.append(f'Prefill run failed to start: {e}')
    try:
        run = run_and_wait(actor, {'marketplaces': ['US'], 'sessionUS': 'session-id=000-0000000-0000000; at-main=HEALTHCHECK',
                                   'maxPagesPerMarketplace': 1})
        health = kv_record(run['defaultKeyValueStoreId'], 'SESSION_HEALTH') or {}
        status = ((health.get('marketplaces') or {}).get('US') or {}).get('status')
        if run['status'] != 'SUCCEEDED' or status != 'SESSION_EXPIRED':
            problems.append(f"Canary run (fake US session) ended {run['status']} with US={status!r} "
                            f"(expected SUCCEEDED + SESSION_EXPIRED): {run.get('statusMessage')}")
    except Exception as e:
        problems.append(f'Canary run failed: {e}')
    return problems


ACTORS = [
    {'slug': 'amazon-asin-rating-scraper', 'id': '2AuZCyLxKZuO2OyNY', 'folder': 'apify-asin-rating-scraper',
     'tests': tests_asin,
     'verify': 'real runs of the new build (default input, and {"asins":["B0GDHRZFWH"],"marketplace":"de"}) return ratings',
     'debug': 'run the actor with {"debugHtml": true} to inspect what Amazon returns'},
    {'slug': 'amazon-seller-central-review-scraper', 'id': '94QvEGIwZP7pyczTt', 'folder': 'apify-sc-review-scraper',
     'tests': tests_sc,
     'verify': 'the prefill run SUCCEEDS and a run with a fake US session SUCCEEDS with SESSION_HEALTH US = SESSION_EXPIRED '
               '(no real Seller Central session is available — never ask for or use one)',
     'debug': 'read the run log and the SESSION_HEALTH record of the failing run'},
]


def check(a):
    actor, problems = a['id'], []
    try:
        d = api(f'/acts/{actor}')['data']
        if not d.get('isPublic'):
            problems.append('Actor is no longer public')
        if d.get('isDeprecated'):
            problems.append('Actor is marked deprecated')
        if d.get('isUnderMaintenance') or d.get('notice') not in (None, 'NONE'):
            problems.append(f"Apify flagged the actor (isUnderMaintenance={d.get('isUnderMaintenance')}, notice={d.get('notice')})")
        st = (d.get('stats') or {}).get('publicActorRunStats30Days') or {}
        total = st.get('TOTAL') or 0
        bad = (st.get('FAILED') or 0) + (st.get('TIMED-OUT') or 0)
        if total >= 10 and bad / total > MAX_STORE_FAIL_RATE:
            problems.append(f'Store failure rate {bad}/{total} public runs in 30 days (> {MAX_STORE_FAIL_RATE:.0%}) — '
                            f'inspect recent failed runs of other users')
        b = api(f'/acts/{actor}/builds?desc=1&limit=1')['data']['items']
        if b and b[0]['status'] != 'SUCCEEDED':
            problems.append(f"Latest build {b[0].get('buildNumber')} is {b[0]['status']}")
    except Exception as e:
        problems.append(f'Apify API error: {e}')
    return problems + a['tests'](actor)


def alert(title, text, level):
    try:
        from gcx_alert import alert as send
        send(title, text, level=level)
    except Exception as e:
        log(f'alert failed: {e}')


def log(msg):
    line = f'{datetime.now():%F %T} {msg}'
    print(line)
    with open(LOG, 'a') as f:
        f.write(line + '\n')


FIX_PROMPT = """The daily health check of the public Apify actor in this folder (cosmic_meteorite/{slug}, \
actor id {actor}) failed:

{problems}

Fix it and deploy the fix as a NEW actor version:
1. Diagnose: read src/, the latest run logs (Apify API, token in ~/.config/apify_asin_scraper/token), and if useful \
{debug}. If the cause is outside the code (Apify outage, account/billing issue, Amazon-wide outage), change nothing, and explain.
2. Fix only files in this folder. Keep the output fields and input schema backward compatible.
3. Bump "version" in .actor/actor.json by 0.1 (e.g. 0.4 -> 0.5).
4. Deploy ONLY with: HOME=$HOME/.config/apify_asin_scraper/clihome apify push --force </dev/null
   (that login is cosmic_meteorite; never use any other Apify account or token).
5. Verify that {verify}.
6. git commit the change in this folder with a message starting "{slug} vX.Y:" (do not push; autosync does).
Never change pricing, never make the actor private, never run it from any account other than cosmic_meteorite.
End with a 3-line summary: cause, fix, new version."""


def autofix(a, problems):
    day = f'{datetime.now():%F}'
    stamp = f"{CONF}/autofix_{a['slug']}_{day}"
    if os.path.exists(stamp):
        log(f"[{a['slug']}] autofix already attempted today — skipping")
        return None
    sys.path.insert(0, os.path.join(HERE, '..', '..', 'ServerBootstrap', 'fallback'))
    try:
        import gcx_ledger                              # server ⇄ Mac fallback: one autofix per actor per day
        key = f"apify-autofix:{a['slug']}:{day}"
        if not gcx_ledger.claim(key):
            log(f"[{a['slug']}] autofix already attempted today on the other machine — skipping")
            return None
        gcx_ledger.put(key, 'done')
    except Exception as e:
        log(f'ledger unavailable ({e}) — continuing with the local once-a-day stamp only')
    open(stamp, 'w').close()
    log(f"[{a['slug']}] starting autofix Claude session")
    env = dict(os.environ, PATH=f"{os.path.expanduser('~/.venvs/apify-asin/bin')}:{os.path.expanduser('~/.local/bin')}:/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin")
    prompt = FIX_PROMPT.format(slug=a['slug'], actor=a['id'], problems='\n'.join(f'- {p}' for p in problems),
                               debug=a['debug'], verify=a['verify'])
    try:
        r = subprocess.run(['claude', '-p', prompt, '--dangerously-skip-permissions'], cwd=os.path.join(SCRAPERS, a['folder']),
                           env=env, stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=2700)
        out = (r.stdout or r.stderr).strip()
    except subprocess.TimeoutExpired:
        out = 'Claude autofix session timed out after 45 min'
    log(f"[{a['slug']}] autofix output: {out[-1500:]}")
    return out


def handle(a, args) -> bool:
    console = f"https://console.apify.com/actors/{a['id']}"
    problems = check(a)
    if not problems:
        log(f"[{a['slug']}] OK — public, not flagged, build ok, store failure rate ok, test runs valid")
        return True
    log(f"[{a['slug']}] FAIL — " + ' | '.join(problems))
    if not args.no_alert:
        alert(f"Apify {a['slug']}: health check FAILED",
              '\n'.join(f'• {p}' for p in problems) + f'\n{"Starting auto-fix…" if args.autofix else ""}\n{console}', 'bad')
    if not args.autofix:
        return False
    summary = autofix(a, problems)
    if summary is None:
        return False
    time.sleep(30)
    after = check(a)
    log(f"[{a['slug']}] after autofix: " + ('OK' if not after else 'STILL FAILING — ' + ' | '.join(after)))
    if not args.no_alert:
        alert(f"Apify {a['slug']}: " + ('fixed & redeployed' if not after else 'auto-fix did NOT resolve it'),
              (summary or '')[-1200:] + ('' if not after else '\nStill failing:\n' + '\n'.join(f'• {p}' for p in after)) + f'\n{console}',
              'ok' if not after else 'bad')
    return not after


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--autofix', action='store_true')
    ap.add_argument('--no-alert', action='store_true')
    ap.add_argument('--only', help='slug of a single actor to check')
    args = ap.parse_args()
    results = [handle(a, args) for a in ACTORS if not args.only or a['slug'] == args.only]
    sys.exit(0 if all(results) else 1)


if __name__ == '__main__':
    main()
