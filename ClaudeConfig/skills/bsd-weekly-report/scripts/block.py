#!/usr/bin/env python3
"""Read / prepare / write my column block in the weekly tab.
  block.py read  TAB RANGE                 -> {cell: text} (+ which cells carry smart chips)
  block.py prev  TAB                       -> my block in TAB (rows from the 글로벌CX전략팀 row in col B down to the next team), col D
  block.py seed  NEWTAB RANGE PREVTAB PREVRANGE
        -> if NEWTAB!RANGE is empty: insert blank rows right above the next team (only if needed), then copy
           PREVTAB!PREVRANGE into it with copyPaste (keeps smart chips / formatting). Refuses if RANGE has text.
  block.py write TAB edits.json            -> backup + write {cell: text}. Cells holding smart chips are rewritten
           with the chip kept: the chip's display text must still appear verbatim in the new text.
Never writes outside the given cells; other teams/teammates (col E etc.) are untouched."""
import sys, json, os, time, re
from _g import *
cmd = sys.argv[1]; sid = WEEKLY_SID
props = {t['title']: t for t in tabs(sid)}
def cells(tab, rng):
    col, r1, r2 = a1(rng)
    g = S.get(spreadsheetId=sid, ranges=[f"'{tab}'!{rng}"], fields='sheets.data.rowData.values(userEnteredValue,chipRuns,textFormatRuns)').execute()
    rd = g['sheets'][0]['data'][0].get('rowData', [])
    out = {}
    for i in range(r2 - r1 + 1):
        v = (rd[i].get('values', [{}]) if i < len(rd) else [{}])[0] if i < len(rd) and rd[i].get('values') else {}
        out[f"{col}{r1+i}"] = v
    return out
txt = lambda v: v.get('userEnteredValue', {}).get('stringValue', v.get('userEnteredValue', {}).get('formulaValue', ''))
if cmd == 'read':
    for k, v in cells(sys.argv[2], sys.argv[3]).items(): print(k, ('[chip] ' if v.get('chipRuns') else '') + repr(txt(v)))
elif cmd == 'prev':
    tab = sys.argv[2]; b = vals(sid, f"'{tab}'!B1:B1500")
    s = next(i for i, x in enumerate(b) if x and x[0].strip() == TEAM) + 1
    e = next((i for i, x in enumerate(b) if i >= s and x and x[0].strip()), len(b))
    print(f"{tab} block: D{s}:D{e}")
    for k, v in cells(tab, f"D{s}:D{e}").items(): print(k, ('[chip] ' if v.get('chipRuns') else '') + repr(txt(v)))
elif cmd == 'seed':
    new, rng, prev, prng = sys.argv[2:6]
    col, r1, r2 = a1(rng); _, p1, p2 = a1(prng); n = p2 - p1 + 1
    assert not any(txt(v).strip() for v in cells(new, rng).values()), f'{new}!{rng} is not empty - refusing to seed'
    b = vals(sid, f"'{new}'!B1:B1500")
    nxt = next(i for i, x in enumerate(b) if i >= r1 and x and x[0].strip()) + 1   # next team's first row
    need = max(0, r1 + n - nxt)
    reqs = []
    if need: reqs.append({'insertDimension': {'range': {'sheetId': props[new]['sheetId'], 'dimension': 'ROWS', 'startIndex': nxt - 2, 'endIndex': nxt - 2 + need}, 'inheritFromBefore': True}})
    ci = ord(col) - 65
    reqs.append({'copyPaste': {'source': {'sheetId': props[prev]['sheetId'], 'startRowIndex': p1 - 1, 'endRowIndex': p2, 'startColumnIndex': ci, 'endColumnIndex': ci + 1},
                 'destination': {'sheetId': props[new]['sheetId'], 'startRowIndex': r1 - 1, 'endRowIndex': r1 - 1 + n, 'startColumnIndex': ci, 'endColumnIndex': ci + 1}, 'pasteType': 'PASTE_NORMAL'}})
    S.batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()
    print(f"inserted {need} rows above row {nxt}; copied {prev}!{prng} -> {new}!{col}{r1}:{col}{r1+n-1}")
elif cmd == 'write':
    tab, ej = sys.argv[2], json.load(open(sys.argv[3]))
    rows = sorted(int(re.sub(r'\D', '', k)) for k in ej); col = re.sub(r'\d', '', next(iter(ej)))
    cur = cells(tab, f"{col}{rows[0]}:{col}{rows[-1]}")
    bk = os.path.expanduser(f"~/.config/bsd_weekly_report/backup_{tab}_{time.strftime('%Y%m%d_%H%M%S')}.json")
    os.makedirs(os.path.dirname(bk), exist_ok=True); json.dump({k: txt(v) for k, v in cur.items()}, open(bk, 'w'), ensure_ascii=False, indent=1)
    plain, reqs = [], []
    for k, new in ej.items():
        v = cur[k]
        if not v.get('chipRuns'): plain.append({'range': f"'{tab}'!{k}", 'values': [[new]]}); continue
        old = txt(v); runs = v['chipRuns']; s2, nr, shift = new, [], 0
        bounds = [r.get('startIndex', 0) for r in runs] + [len(old)]
        for i, r in enumerate(runs):
            if 'chip' not in r: continue
            a, b = bounds[i], bounds[i + 1]; disp = old[a:b]
            p = s2.find(disp, a - shift if a - shift >= 0 else 0)
            if p < 0: p = s2.find(disp)
            assert p >= 0, f"{k}: chip text {disp!r} missing from new text"
            s2 = s2[:p] + '@' + s2[p + len(disp):]; nr.append({'startIndex': p, 'chip': r['chip']}); shift += len(disp) - 1
        r0 = int(re.sub(r'\D', '', k)) - 1
        reqs.append({'updateCells': {'range': {'sheetId': props[tab]['sheetId'], 'startRowIndex': r0, 'endRowIndex': r0 + 1, 'startColumnIndex': ord(col) - 65, 'endColumnIndex': ord(col) - 64},
                     'rows': [{'values': [{'userEnteredValue': {'stringValue': s2}, 'chipRuns': nr}]}], 'fields': 'userEnteredValue,chipRuns'}})
    if plain: S.values().batchUpdate(spreadsheetId=sid, body={'valueInputOption': 'RAW', 'data': plain}).execute()
    if reqs: S.batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()
    print('backup:', bk); print('written:', list(ej))
