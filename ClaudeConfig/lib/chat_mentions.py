"""Real Google Chat @mentions by email, for webhook / app senders.

Google Chat only notifies a person when the message's plain-text part contains `<users/USER_ID>`
(the same form as `<users/all>`). Cards, "@name" text and email addresses never ping anyone.
USER_ID is the person's numeric Google ID, so we keep an email → "users/ID" map:

    ~/.config/gcx_chat_user_ids.json   {"kjw@spigen.com": "users/1234567890", ...}

Build/refresh it with the Apps Script GAS_Operations/ChatMentions (buildMentionMap), which lists the
GCX rooms' members and looks emails up in the Workspace directory, then saves the same JSON to Drive
as "gcx_chat_user_ids.json". Copy that file to ~/.config/ on each Mac (or point CHAT_IDS_FILE at it).

Usage:
    from chat_mentions import mention, add_mentions
    payload = add_mentions(payload, ["kjw@spigen.com", "yangsr@spigen.com"])   # adds/extends payload["text"]
    text = f"{mention('kjw@spigen.com')} 확인 부탁드립니다"
The person must be a member of the room, and the mention must be in "text", not inside cardsV2.
"""
import json, os, sys

MAP_FILE = os.environ.get("CHAT_IDS_FILE", os.path.expanduser("~/.config/gcx_chat_user_ids.json"))
_cache = None


def load_map():
    global _cache
    if _cache is None:
        try:
            with open(MAP_FILE, encoding="utf-8") as f:
                _cache = {k.lower(): v for k, v in json.load(f).items()}
        except FileNotFoundError:
            _cache = {}
    return _cache


def mention(email):
    """'<users/123>' for a known email; '@email' (no alert) and a stderr warning for an unknown one."""
    if email.strip().lower() == "all":
        return "<users/all>"
    uid = load_map().get(email.strip().lower())
    if not uid:
        print(f"chat_mentions: no Chat user id for {email} — run buildMentionMap()", file=sys.stderr)
        return "@" + email.split("@")[0]
    return f"<{uid}>"


def add_mentions(payload, emails, lead=""):
    """Prepend mentions to payload['text'] (created if missing — a card stays as is below it)."""
    if not emails:
        return payload
    tags = " ".join(mention(e) for e in emails)
    text = payload.get("text", "")
    payload["text"] = " ".join(x for x in (tags, lead, text) if x).strip()
    return payload


if __name__ == "__main__":   # quick check:  python3 chat_mentions.py kjw@spigen.com
    for e in sys.argv[1:]:
        print(e, "→", mention(e))
