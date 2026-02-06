"""Multi-platform Claude messenger server.

Receives messages via WhatsApp webhooks and Telegram long-polling, routes
them to persistent Claude loops (one per conversation), and relays responses
back. Each conversation gets a ChatActor with an asyncio queue.

Uses loop daemon for Claude sessions — policy-based tool approval
(--permission-mode acceptEdits) instead of interactive buttons.
"""

import asyncio
import contextlib
import hashlib
import hmac
import json
import os
import sys
import threading
import time
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Protocol
from urllib.error import HTTPError
from urllib.parse import parse_qs, urlparse
from urllib.request import Request, urlopen

from mcp_handley_lab.loop.client import kill, run, spawn

# ---------------------------------------------------------------------------
# Environment (set via systemd EnvironmentFile or shell exports)
# ---------------------------------------------------------------------------

VERIFY_TOKEN = os.environ.get("WHATSAPP_VERIFY_TOKEN", "")
ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN", "")
PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID", "")
APP_SECRET = os.environ.get("WHATSAPP_APP_SECRET", "")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
_tg_allowed_raw = os.environ.get("TELEGRAM_ALLOWED_CHAT_IDS", "")
TELEGRAM_ALLOWED_CHAT_IDS: set[int] | None = (
    {int(x.strip()) for x in _tg_allowed_raw.split(",") if x.strip()}
    if _tg_allowed_raw
    else None
)

CLAUDE_PERMISSION_MODE = os.environ.get("CLAUDE_PERMISSION_MODE", "acceptEdits")
CLAUDE_SYSTEM_PROMPT = os.environ.get(
    "CLAUDE_SYSTEM_PROMPT",
    "You are a personal assistant. Keep responses concise for mobile.",
)

MESSENGER_DIR = Path.home() / "messenger"
GRAPH_API = f"https://graph.facebook.com/v21.0/{PHONE_NUMBER_ID}/messages"
TELEGRAM_API = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _truncate(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------


class Platform(Protocol):
    """Messaging platform abstraction."""

    def send_text(self, conversation_id: str, text: str) -> None: ...


@dataclass
class IncomingEvent:
    conversation_id: str
    kind: str  # "text", "command"
    text: str
    platform: Platform


# ---------------------------------------------------------------------------
# WhatsApp platform
# ---------------------------------------------------------------------------

_WA_TEXT_MAX = 4096


class WhatsAppPlatform:
    def send_text(self, conversation_id: str, text: str) -> None:
        text = _truncate(text, _WA_TEXT_MAX)
        _send_whatsapp(conversation_id, {"text": {"body": text}})


def _send_whatsapp(to: str, payload: dict) -> None:
    payload["messaging_product"] = "whatsapp"
    payload["to"] = to
    req = Request(
        GRAPH_API,
        data=json.dumps(payload).encode(),
        headers={
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json",
        },
    )
    try:
        with urlopen(req) as resp:
            print(f"WA sent to {to}: {resp.status}", flush=True)
    except HTTPError as e:
        body = e.read().decode(errors="replace")
        print(f"WhatsApp API error {e.code}: {body}", flush=True)
        raise


# ---------------------------------------------------------------------------
# Telegram platform
# ---------------------------------------------------------------------------

_TG_TEXT_MAX = 4096


class TelegramPlatform:
    @staticmethod
    def _parse_conversation_id(conversation_id: str) -> tuple[str, str | None]:
        parts = conversation_id.split(":", 2)
        chat_id = parts[1] if len(parts) > 1 else conversation_id
        topic_id = parts[2] if len(parts) > 2 else None
        return chat_id, topic_id

    def _call(self, method: str, payload: dict) -> dict:
        url = f"{TELEGRAM_API}/{method}"
        req = Request(
            url,
            data=json.dumps(payload).encode(),
            headers={"Content-Type": "application/json"},
        )
        try:
            with urlopen(req) as resp:
                return json.loads(resp.read())
        except HTTPError as e:
            body = e.read().decode(errors="replace")
            print(f"Telegram API error {e.code} ({method}): {body}", flush=True)
            raise

    def send_text(self, conversation_id: str, text: str) -> None:
        chat_id, topic_id = self._parse_conversation_id(conversation_id)
        text = _truncate(text, _TG_TEXT_MAX)
        payload = {"chat_id": chat_id, "text": text}
        if topic_id:
            payload["message_thread_id"] = int(topic_id)
        self._call("sendMessage", payload)


# ---------------------------------------------------------------------------
# Directory mapping
# ---------------------------------------------------------------------------


def _cwd_for_conversation(conversation_id: str) -> Path:
    parts = conversation_id.split(":", 2)
    if len(parts) == 3:
        return MESSENGER_DIR / parts[0] / parts[1] / parts[2]
    if len(parts) == 2:
        return MESSENGER_DIR / parts[0] / parts[1]
    return MESSENGER_DIR / conversation_id


def _migrate_old_dirs():
    MESSENGER_DIR.mkdir(parents=True, exist_ok=True)
    renames = {"wa": "whatsapp", "tg": "telegram"}

    old_wa = Path.home() / "whatsapp"
    if old_wa.is_dir():
        new_wa = MESSENGER_DIR / "whatsapp"
        new_wa.mkdir(parents=True, exist_ok=True)
        for child in old_wa.iterdir():
            if child.is_dir() and not (new_wa / child.name).exists():
                child.rename(new_wa / child.name)
                print(f"Migrated {child} → {new_wa / child.name}", flush=True)

    old_bridge = Path.home() / "bridge"
    if old_bridge.is_dir():
        for child in old_bridge.iterdir():
            dest_name = renames.get(child.name, child.name)
            dest = MESSENGER_DIR / dest_name
            if not dest.exists():
                child.rename(dest)
                print(f"Migrated {child} → {dest}", flush=True)


# ---------------------------------------------------------------------------
# ChatActor — one per conversation, owns a persistent loop
# ---------------------------------------------------------------------------


class ChatActor:
    def __init__(self, conversation_id: str, platform: Platform):
        self.conversation_id = conversation_id
        self.platform = platform
        self.queue: asyncio.Queue[IncomingEvent] = asyncio.Queue(maxsize=50)
        self.cwd = _cwd_for_conversation(conversation_id)
        self.loop_id: str | None = None
        self._state_file = self.cwd / "loop_state.json"
        self._task: asyncio.Task | None = None

    async def start(self):
        self.cwd.mkdir(parents=True, exist_ok=True)
        self._load_state()
        self._task = asyncio.create_task(self._run())

    async def _run(self):
        while True:
            event = await self.queue.get()
            try:
                await self._handle(event.text)
            except Exception as e:
                print(f"Chat {self.conversation_id} error: {e}", flush=True)
                self._send(f"Error: {e}")

    def _send(self, text: str) -> None:
        print(f"Reply to {self.conversation_id}: {text[:200]}", flush=True)
        self.platform.send_text(self.conversation_id, text)

    async def _handle(self, text: str) -> None:
        for attempt in (1, 2):
            try:
                output = await asyncio.to_thread(self._query, text)
                self._send(output)
                return
            except RuntimeError as e:
                # Stale loop: "not_found: loop not found: {id}"
                # Dead tmux pane: "backend_error: Claude session not found: {id}"
                if attempt == 1 and "not found" in str(e):
                    self._clear_state()
                    continue
                raise

    def _query(self, text: str) -> str:
        """Ensure loop exists and run text. Called via to_thread."""
        if not self.loop_id:
            self.loop_id = spawn(
                "claude",
                label=f"msg-{self.conversation_id[:20]}",
                cwd=str(self.cwd),
                prompt=CLAUDE_SYSTEM_PROMPT,
                args=f"--permission-mode {CLAUDE_PERMISSION_MODE}",
            )
            self._save_state()
        return run(self.loop_id, text, sync_timeout=-1)

    def reset(self):
        if self.loop_id:
            with contextlib.suppress(RuntimeError):
                kill(self.loop_id)
        self._clear_state()

    def _load_state(self):
        with contextlib.suppress(FileNotFoundError, json.JSONDecodeError):
            self.loop_id = json.loads(self._state_file.read_text()).get("loop_id")

    def _save_state(self):
        self._state_file.write_text(json.dumps({"loop_id": self.loop_id}))

    def _clear_state(self):
        self.loop_id = None
        self._state_file.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Event loop dispatch
# ---------------------------------------------------------------------------

_loop: asyncio.AbstractEventLoop | None = None
_actors: dict[str, ChatActor] = {}


def _get_or_create_actor(conversation_id: str, platform: Platform) -> ChatActor:
    if conversation_id not in _actors:
        actor = ChatActor(conversation_id, platform)
        _actors[conversation_id] = actor
        _loop.create_task(actor.start())
    return _actors[conversation_id]


async def _dispatch(event: IncomingEvent):
    if event.kind == "command":
        cmd = event.text.strip().lower().split("@")[0]
        if cmd in ("/reset", "/new"):
            actor = _actors.get(event.conversation_id)
            if actor:
                actor.reset()
                del _actors[event.conversation_id]
            event.platform.send_text(
                event.conversation_id,
                "Session reset. Send a new message to start fresh.",
            )
            return

    actor = _get_or_create_actor(event.conversation_id, event.platform)
    try:
        actor.queue.put_nowait(event)
    except asyncio.QueueFull:
        event.platform.send_text(
            event.conversation_id, "Too many pending messages. Please wait."
        )


def _post_to_loop(event: IncomingEvent):
    if _loop is None:
        print(
            f"Event loop not ready, dropping {event.kind} from {event.conversation_id}",
            flush=True,
        )
        return
    fut = asyncio.run_coroutine_threadsafe(_dispatch(event), _loop)
    fut.add_done_callback(
        lambda f: print(f"Dispatch error: {f.exception()}", flush=True)
        if f.exception()
        else None
    )


# ---------------------------------------------------------------------------
# WhatsApp webhook → IncomingEvent
# ---------------------------------------------------------------------------

_wa_platform: WhatsAppPlatform | None = None


def _classify_wa_event(sender: str, text: str) -> IncomingEvent:
    conversation_id = f"whatsapp:{sender}"
    if text.strip().lower().split("@")[0] in ("/reset", "/new"):
        return IncomingEvent(
            conversation_id, kind="command", text=text, platform=_wa_platform
        )
    return IncomingEvent(conversation_id, kind="text", text=text, platform=_wa_platform)


def verify_signature(payload: bytes, signature_header: str) -> bool:
    if not APP_SECRET or not signature_header:
        return False
    if not signature_header.startswith("sha256="):
        return False
    expected = hmac.new(APP_SECRET.encode(), payload, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, signature_header[7:])


def extract_messages(data: dict) -> list[tuple[str, str]]:
    """Extract (sender, text) pairs from webhook payload."""
    messages = []
    for entry in data.get("entry", []):
        for change in entry.get("changes", []):
            value = change.get("value", {})
            for msg in value.get("messages", []):
                if msg.get("type") == "text":
                    messages.append((msg["from"], msg["text"]["body"]))
    return messages


# ---------------------------------------------------------------------------
# Telegram long-polling → IncomingEvent
# ---------------------------------------------------------------------------

_tg_platform: TelegramPlatform | None = None
_tg_offset_file = MESSENGER_DIR / "tg_offset.txt"


def _load_tg_offset() -> int:
    try:
        return int(_tg_offset_file.read_text().strip())
    except (FileNotFoundError, ValueError):
        return 0


def _save_tg_offset(offset: int):
    _tg_offset_file.parent.mkdir(parents=True, exist_ok=True)
    tmp = _tg_offset_file.with_suffix(".tmp")
    tmp.write_text(str(offset))
    tmp.rename(_tg_offset_file)


def _tg_conversation_id(chat_id: int, thread_id: int | None) -> str:
    if thread_id:
        return f"telegram:{chat_id}:{thread_id}"
    return f"telegram:{chat_id}"


def _telegram_poll():
    """Long-polling loop for Telegram updates. Runs on a daemon thread."""
    offset = _load_tg_offset()
    backoff = 1

    while True:
        try:
            payload = json.dumps(
                {
                    "offset": offset + 1,
                    "timeout": 30,
                    "allowed_updates": ["message"],
                }
            ).encode()
            req = Request(
                f"{TELEGRAM_API}/getUpdates",
                data=payload,
                headers={"Content-Type": "application/json"},
            )
            with urlopen(req, timeout=35) as resp:
                data = json.loads(resp.read())

            backoff = 1

            for update in data.get("result", []):
                update_id = update["update_id"]
                if update_id > offset:
                    offset = update_id
                    _save_tg_offset(offset)

                if "message" in update:
                    _handle_tg_message(update["message"])

        except KeyboardInterrupt:
            break
        except Exception as e:
            print(f"Telegram poll error: {e}", flush=True)
            time.sleep(min(backoff, 30))
            backoff = min(backoff * 2, 30)


def _handle_tg_message(msg: dict):
    chat_id = msg["chat"]["id"]
    if (
        TELEGRAM_ALLOWED_CHAT_IDS is not None
        and chat_id not in TELEGRAM_ALLOWED_CHAT_IDS
    ):
        print(f"[TG blocked] chat_id={chat_id}", flush=True)
        return

    text = msg.get("text")
    if not text:
        return
    thread_id = msg.get("message_thread_id")
    conversation_id = _tg_conversation_id(chat_id, thread_id)

    cmd = text.strip().lower().split("@")[0]
    if cmd in ("/reset", "/new"):
        event = IncomingEvent(
            conversation_id, kind="command", text=text, platform=_tg_platform
        )
    else:
        event = IncomingEvent(
            conversation_id, kind="text", text=text, platform=_tg_platform
        )

    print(f"[TG {event.kind}] {chat_id}: {text[:100]}", flush=True)
    _post_to_loop(event)


# ---------------------------------------------------------------------------
# HTTP handler (WhatsApp webhooks)
# ---------------------------------------------------------------------------


class WebhookHandler(BaseHTTPRequestHandler):
    PRIVACY_HTML = b"""<!DOCTYPE html>
<html><head><title>Privacy Policy - handley-lab</title></head>
<body><h1>Privacy Policy</h1>
<p>This WhatsApp integration is a personal project by Handley Lab.
No user data is stored, shared, or sold. Messages are processed
in real time and not retained.</p>
<p>Contact: handleylab@gmail.com</p>
</body></html>"""

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/privacy":
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(self.PRIVACY_HTML)
            return
        if parsed.path != "/webhook":
            self.send_error(404)
            return

        params = parse_qs(parsed.query)
        mode = params.get("hub.mode", [None])[0]
        token = params.get("hub.verify_token", [None])[0]
        challenge = params.get("hub.challenge", [None])[0]

        if mode == "subscribe" and token == VERIFY_TOKEN:
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(challenge.encode() if challenge else b"")
            print("Webhook verified", flush=True)
        else:
            self.send_error(403)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/webhook":
            self.send_error(404)
            return

        content_length = int(self.headers.get("Content-Length", 0))
        payload = self.rfile.read(content_length)

        signature = self.headers.get("X-Hub-Signature-256", "")
        if not verify_signature(payload, signature):
            print("Invalid signature", flush=True)
            self.send_error(403)
            return

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(b'{"status":"ok"}')

        try:
            data = json.loads(payload)
        except json.JSONDecodeError:
            return

        for sender, text in extract_messages(data):
            event = _classify_wa_event(sender, text)
            print(f"[WA {event.kind}] {sender}: {text}", flush=True)
            _post_to_loop(event)

    def log_message(self, format, *args):
        print(f"{self.client_address[0]} - {format % args}", flush=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    global _loop, _wa_platform, _tg_platform
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080

    MESSENGER_DIR.mkdir(parents=True, exist_ok=True)
    _migrate_old_dirs()

    _wa_platform = WhatsAppPlatform()

    _loop = asyncio.new_event_loop()

    def _run_loop():
        asyncio.set_event_loop(_loop)
        _loop.run_forever()

    threading.Thread(target=_run_loop, daemon=True).start()

    if TELEGRAM_BOT_TOKEN:
        _tg_platform = TelegramPlatform()
        threading.Thread(target=_telegram_poll, daemon=True).start()
        print("Telegram polling started", flush=True)

    server = ThreadingHTTPServer(("127.0.0.1", port), WebhookHandler)
    print(f"Listening on 127.0.0.1:{port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down", flush=True)
        server.server_close()


if __name__ == "__main__":
    main()
