#!/usr/bin/env python3
"""
Local backend for the "GChat Reply Suggest" Tampermonkey userscript.

Listens on 127.0.0.1:8765. The userscript (running inside the Google Chat
Chrome PWA) POSTs the tail of the visible conversation transcript to
/suggest via GM_xmlhttpRequest, this server shells out to the Claude Code
CLI (reusing the existing Claude subscription, same pattern as
~/.screenpipe-zendesk-*.py) to generate 3 short candidate reply sentences,
and returns them as JSON.

Run manually:
    python3 ~/Desktop/GCX/Browser_Extensions/gchat_reply_suggest_server.py

Or keep it running via a LaunchAgent if you want it always available.
"""

import json
import os
import re
import shutil
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOST = "127.0.0.1"
PORT = 8765


def _resolve_claude_bin() -> str:
    """Locate the Claude CLI on whichever machine this is running on.

    Was hardcoded to the Mac's path, which made the server unusable on the
    Windows box. Order: explicit override -> PATH -> the per-user native
    install location used by both platforms.
    """
    override = os.environ.get("CLAUDE_BIN")
    if override:
        return override
    found = shutil.which("claude")
    if found:
        return found
    local = os.path.expanduser(
        "~/.local/bin/claude.exe" if os.name == "nt" else "~/.local/bin/claude"
    )
    return local


CLAUDE_BIN = _resolve_claude_bin()
TRANSCRIPT_CHAR_LIMIT = 2000

PROMPT_TEMPLATE = """You are helping a Spigen GCX (Global CX) team member reply quickly in \
Google Chat. Below is the tail end of a chat conversation (may be in Korean or English).

--- TRANSCRIPT ---
{transcript}
--- END TRANSCRIPT ---

Suggest 3 short, natural candidate replies to the most recent message, written from the \
perspective of the person about to reply (not the last speaker). Match the language and tone \
already used in the conversation (workplace-casual Korean is the default for this team). Keep \
each suggestion to one sentence, no greetings/sign-offs, no markdown, no explanations.

Respond with ONLY a JSON array of exactly 3 strings, nothing else. Example:
["알겠습니다, 확인해볼게요", "네 지금 바로 처리하겠습니다", "혹시 추가로 필요한 정보 있을까요?"]
"""


def ask_claude(transcript: str) -> list[str]:
    prompt = PROMPT_TEMPLATE.format(transcript=transcript[-TRANSCRIPT_CHAR_LIMIT:])
    result = subprocess.run(
        [CLAUDE_BIN, "-p", prompt],
        capture_output=True,
        text=True,
        timeout=60,
    )
    output = result.stdout.strip()

    try:
        parsed = json.loads(output)
        if isinstance(parsed, list):
            return [str(s) for s in parsed][:3]
    except json.JSONDecodeError:
        pass

    match = re.search(r"\[.*\]", output, re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group(0))
            if isinstance(parsed, list):
                return [str(s) for s in parsed][:3]
        except json.JSONDecodeError:
            pass

    raise RuntimeError(f"Could not parse Claude output as JSON array: {output[:300]!r}")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[gchat-reply-suggest] {self.address_string()} - {format % args}")

    def _send_json(self, status: int, payload: dict):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path != "/suggest":
            self._send_json(404, {"error": "not found"})
            return

        length = int(self.headers.get("Content-Length", 0))
        try:
            data = json.loads(self.rfile.read(length) or b"{}")
        except json.JSONDecodeError:
            self._send_json(400, {"error": "invalid JSON body"})
            return

        transcript = (data.get("transcript") or "").strip()
        if not transcript:
            self._send_json(400, {"error": "empty transcript"})
            return

        try:
            suggestions = ask_claude(transcript)
        except subprocess.TimeoutExpired:
            self._send_json(504, {"error": "claude CLI timed out"})
            return
        except Exception as e:
            self._send_json(500, {"error": str(e)})
            return

        self._send_json(200, {"suggestions": suggestions})

    def do_GET(self):
        if self.path == "/health":
            self._send_json(200, {"status": "ok"})
        else:
            self._send_json(404, {"error": "not found"})


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"gchat-reply-suggest-server listening on http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
