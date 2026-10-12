"""Write a SKU issue-lookup tab (reviews + Zendesk claims) into the review source spreadsheet.

Clones the SKU_ACS11498 tab (for its formatting), writes the 16-column header + rows,
and highlights rows whose Issue Summary matches any --highlight keyword.

  python3 write_tab.py --title SKU_ACS11909_ACS11923 --rows rows.json --highlight "미끄러움,Slipper"
"""
import argparse, json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN = "/Users/kevinkim/.config/gws_shim/token.json"
SID = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw"
TEMPLATE_GID = 1031735538  # SKU_ACS11498
HEADER = ["ASIN", "Created 날짜", "사진 유무", "Reviewer", "Review Ratings", "Review Title", "본문", "Marketplace",
          "Link", "Image URL", "Review ID", "Order ID", "Product Rating", "Ratings Count", "Issue Summary", "Type"]
YELLOW = {"red": 1, "green": 0.95, "blue": 0.6}


def sheets():
    info = json.load(open(TOKEN))
    creds = Credentials.from_authorized_user_info(info)
    creds.refresh(Request())
    info["token"] = creds.token
    json.dump(info, open(TOKEN, "w"))
    return build("sheets", "v4", credentials=creds).spreadsheets()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--title", required=True)
    ap.add_argument("--rows", required=True, help="JSON list of 16-value rows in HEADER order")
    ap.add_argument("--highlight", default="", help="comma-separated keywords matched against Issue Summary")
    ap.add_argument("--replace", action="store_true", help="overwrite the tab if it already exists")
    a = ap.parse_args()

    rows = [["" if v is None else str(v) for v in r] for r in json.load(open(a.rows))]
    bad = [i for i, r in enumerate(rows) if len(r) != len(HEADER)]
    if bad:
        raise SystemExit(f"rows {bad[:5]} do not have {len(HEADER)} values")

    S = sheets()
    meta = S.get(spreadsheetId=SID, fields="sheets(properties(sheetId,title,index))").execute()["sheets"]
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta}
    if a.title in existing:
        if not a.replace:
            raise SystemExit(f"tab {a.title} already exists (gid {existing[a.title]}); pass --replace to overwrite")
        gid = existing[a.title]
    else:
        tpl_index = next(s["properties"]["index"] for s in meta if s["properties"]["sheetId"] == TEMPLATE_GID)
        r = S.batchUpdate(spreadsheetId=SID, body={"requests": [{"duplicateSheet": {
            "sourceSheetId": TEMPLATE_GID, "newSheetName": a.title, "insertSheetIndex": tpl_index + 1}}]}).execute()
        gid = r["replies"][0]["duplicateSheet"]["properties"]["sheetId"]

    end = max(len(rows) + 50, 1000)
    S.values().clear(spreadsheetId=SID, range=f"'{a.title}'!A1:Z{end}").execute()
    S.values().update(spreadsheetId=SID, range=f"'{a.title}'!A1", valueInputOption="RAW",
                      body={"values": [HEADER] + rows}).execute()

    reqs = [
        {"repeatCell": {"range": {"sheetId": gid, "startRowIndex": 1, "endRowIndex": end,
                                  "startColumnIndex": 0, "endColumnIndex": len(HEADER)},
                        "cell": {"userEnteredFormat": {"backgroundColor": {"red": 1, "green": 1, "blue": 1},
                                                       "textFormat": {"bold": False}}},
                        "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold"}},
        # header formatting for the extra Type column, copied from Issue Summary
        {"copyPaste": {"source": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 14, "endColumnIndex": 15},
                       "destination": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1,
                                       "startColumnIndex": 15, "endColumnIndex": 16},
                       "pasteType": "PASTE_FORMAT"}},
    ]
    keys = [k.strip() for k in a.highlight.split(",") if k.strip()]
    for i, r in enumerate(rows):
        if any(k in r[14] for k in keys):
            reqs.append({"repeatCell": {"range": {"sheetId": gid, "startRowIndex": i + 1, "endRowIndex": i + 2,
                                                  "startColumnIndex": 0, "endColumnIndex": len(HEADER)},
                                        "cell": {"userEnteredFormat": {"backgroundColor": YELLOW}},
                                        "fields": "userEnteredFormat.backgroundColor"}})
    S.batchUpdate(spreadsheetId=SID, body={"requests": reqs}).execute()

    n_rev = sum(r[15] == "Review" for r in rows)
    print(json.dumps({"gid": gid, "reviews": n_rev, "claims": len(rows) - n_rev,
                      "highlighted": len(reqs) - 2,
                      "url": f"https://docs.google.com/spreadsheets/d/{SID}/edit#gid={gid}"}, ensure_ascii=False))


if __name__ == "__main__":
    main()
