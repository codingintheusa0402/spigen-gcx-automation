#!/usr/bin/env python3
"""Server test: wait for the SC scraper run to finish, then post a result card to the TEST room only."""
import json, os, re, subprocess, sys, time, urllib.request
LOG = "/tmp/sc_scraper.log"
WEBHOOK = os.environ["TEST_WEBHOOK"]          # passed in by caller, never hard-coded here
SHEET_ID = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw"
while subprocess.run(["pgrep", "-f", "scrape_sc_reviews.py"], capture_output=True).returncode == 0:
    time.sleep(20)
log = open(LOG, encoding="utf-8", errors="replace").read()
rows = re.findall(r"^\s+(US|EU|JP|IN)\s+(\d+) reviews\s+(\d+) with images\s+\[(.+?)\]", log, re.M)
m = re.search(r"✓ (?:created|appended \d+ new rows →) '([^']+)'(?: with (\d+) rows)?", log)
sheet = m.group(1) if m else "(upload not found in log)"
ok = bool(rows) and all(s.strip().upper().startswith("OK") or "✓" in s for *_, s in rows) and m
tail = "\n".join(log.strip().splitlines()[-6:])
lines = [f"<b>{d}</b>  {n} reviews · {i} with images · {s}" for d, n, i, s in rows] or [f"<i>no summary table — last log lines:</i>\n{tail}"]
card = {"cardsV2": [{"cardId": "sctest", "card": {
  "header": {"title": "[SERVER TEST] SC Scraper — " + time.strftime("%y%m%d %H:%M KST"),
             "subtitle": "gcx-server (Windows/WSL2) · " + ("✅ success" if ok else "⚠️ check")},
  "sections": [{"widgets": [{"textParagraph": {"text": "<br>".join(lines)}},
                            {"decoratedText": {"topLabel": "Sheet tab", "text": sheet}},
                            {"buttonList": {"buttons": [{"text": "Open sheet", "onClick": {"openLink": {
                              "url": f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"}}}]}}]}]}}]}
req = urllib.request.Request(WEBHOOK, json.dumps(card).encode(), {"Content-Type": "application/json; charset=UTF-8"})
print("posted:", urllib.request.urlopen(req).status, "| sheet:", sheet, "| ok:", bool(ok))
