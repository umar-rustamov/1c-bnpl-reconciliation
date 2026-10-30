# -*- coding: utf-8 -*-
"""Relays this project's console output to a Telegram group, so a batch
run can be watched/reviewed without staying attached to the terminal.

Credentials (logs_bot_token, telegram_group_id) come from credentials\\.env
(see env_config.py) - same "never hardcoded, never committed" pattern as
every other secret in this project.

How it works: install() replaces sys.stdout/sys.stderr with a small Tee
that writes through to the real console EXACTLY as before (nothing about
what you see in the terminal changes) and ALSO buffers everything into
memory, sending it to Telegram as soon as a complete line is available
(confirmed live: batching only at "per order/application finished"
checkpoints - the first design here - meant a batch of quick, no-1C-work
skips could finish and report before ever notifying Telegram mid-run,
which isn't what "as soon as it's printed" means). The explicit flush()
calls still scattered at those checkpoints (and via atexit at interpreter
exit) are now just a safety net for any trailing partial line with no
newline yet - the auto-flush-per-line below is what actually delivers
real-time-ish updates. Wrap a block of related output that should arrive
as ONE message instead (e.g. a final batch report) in `with batch():`.

Sending is deliberately best-effort: a network hiccup, a Telegram-side
error response (including a 429 rate limit - retried once after the
delay Telegram itself reports, since a burst of rapid prints, e.g. a
final report listing many skipped rows, can plausibly trigger one), or a
misconfigured token prints a warning (to the REAL stdout, bypassing the
Tee, so that warning doesn't itself get buffered for a retry that would
just fail again) and otherwise doesn't interrupt the actual reconciliation
work - this is a monitoring side-channel, not something any part of the
reconciliation flow should ever depend on succeeding.
"""
import atexit
import queue
import sys
import threading
import time
from contextlib import contextmanager

import requests

from env_config import load_env

_env = load_env()
BOT_TOKEN = _env.get("logs_bot_token")
GROUP_ID = _env.get("telegram_group_id")
MENTION_USERNAME = _env.get("mention_username")  # Telegram @username to tag on human-input pauses
_API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage" if BOT_TOKEN else None
_API_BASE = f"https://api.telegram.org/bot{BOT_TOKEN}" if BOT_TOKEN else None

_CHUNK_SIZE = 3500  # comfortably under Telegram's ~4096-character cap
_buffer = []
_installed = False
_auto_flush = True


def _post(chunk):
    return requests.post(_API_URL, data={"chat_id": GROUP_ID, "text": chunk}, timeout=10)


def _send(text):
    if not _API_URL or not GROUP_ID:
        return
    for i in range(0, len(text), _CHUNK_SIZE):
        chunk = text[i:i + _CHUNK_SIZE]
        if not chunk.strip():
            continue
        try:
            resp = _post(chunk)
            if resp.status_code == 429:
                # Telegram's own rate-limit response - it tells us how long
                # to wait via retry_after (seconds); one retry is enough for
                # the odd burst this project's prints could plausibly cause,
                # not worth a longer retry loop for a best-effort side-channel.
                wait = resp.json().get("parameters", {}).get("retry_after", 2)
                time.sleep(wait)
                resp = _post(chunk)
            if not resp.ok or not resp.json().get("ok"):
                print(f"   (Telegram log relay rejected a message: HTTP {resp.status_code} {resp.text[:200]!r} "
                      "- continuing without it)", file=sys.__stdout__)
        except requests.exceptions.RequestException as e:
            print(f"   (Telegram log relay failed: {e!r} - continuing without it)", file=sys.__stdout__)


class _Tee:
    """Wraps a real stream (stdout or stderr): every write() goes through
    to it unchanged, and is also appended to the shared buffer - a write()
    containing a newline (i.e. essentially every print() call) triggers an
    immediate flush(), so lines reach Telegram as they're printed rather
    than waiting for some later checkpoint."""
    def __init__(self, real_stream):
        self._real = real_stream

    def write(self, text):
        self._real.write(text)
        _buffer.append(text)
        if _auto_flush and "\n" in text:
            flush()
        return len(text)

    def flush(self):
        self._real.flush()

    def __getattr__(self, name):
        return getattr(self._real, name)


def install():
    """Start relaying console output to Telegram - call this once, near
    the top of main(). Safe to call more than once (a no-op after the
    first). Registers flush() via atexit so anything buffered still gets
    sent even if the script exits via an unhandled exception."""
    global _installed
    if _installed:
        return
    if not _API_URL or not GROUP_ID:
        print("   (credentials\\.env has no logs_bot_token/telegram_group_id - Telegram log relay disabled)")
        _installed = True
        return
    sys.stdout = _Tee(sys.stdout)
    sys.stderr = _Tee(sys.stderr)
    atexit.register(flush)
    _installed = True


def flush():
    """Send everything written since the last flush() as one (possibly
    chunked) Telegram message, then clear the buffer. Mostly self-triggered
    per line by _Tee.write() now - still called explicitly at a few
    checkpoints (and via atexit) to catch any trailing partial line with no
    newline yet."""
    global _buffer
    if not _buffer:
        return
    text = "".join(_buffer)
    _buffer = []
    _send(text)


@contextmanager
def batch():
    """Group everything printed inside this 'with' block into ONE Telegram
    message (still chunked if it's long enough to need it) instead of the
    usual one-message-per-line - use around a block of related output that
    reads better arriving together (e.g. a final batch report: the
    "=== Batch report ===" header, the done/skipped/partial/errored
    counts, and the per-row breakdown all belong in one message, not
    scattered across a dozen)."""
    global _auto_flush
    previous = _auto_flush
    _auto_flush = False
    try:
        yield
    finally:
        _auto_flush = previous
        flush()


def mention(text):
    """Tag the configured user (mention_username in credentials\\.env) in
    the Telegram group with `text`, so they get a phone notification that
    something needs their attention - a pause waiting on input (see
    automate_order_ekvayring.confirm / guard_no_surprise_dialog) or an
    order that was auto-skipped for a reason worth flagging (a Доставка
    line item). The caller provides the wording. Flushes pending output
    first so the tag arrives after the context that explains why.
    Best-effort like everything else here: never raises, never blocks. A
    no-op if the relay or mention_username isn't configured."""
    flush()  # get whatever led up to this out first
    if not _API_URL or not GROUP_ID or not MENTION_USERNAME:
        return
    try:
        _post(f"@{MENTION_USERNAME} {text}")
    except requests.exceptions.RequestException as e:
        print(f"   (Telegram mention failed: {e!r} - continuing)", file=sys.__stdout__)


# ------------------------------------------------- two-way control (Telegram)
# A human-input pause (a below-threshold product-name match, an unexpected
# 1C dialog) can be answered from the Telegram group instead of only the
# terminal - so an unattended batch can be steered from a phone. The user
# sends a slash command in the group (this bot has privacy mode ON, but
# Telegram still delivers messages starting with "/" to it - CONFIRMED
# live). Commands, all also accepted without the leading slash and via the
# usual single-letter terminal shortcuts:
_CMD_MAP = {
    "/continue": "continue", "continue": "continue", "/c": "continue", "c": "continue",
    "/skip": "skip", "skip": "skip", "/s": "skip", "s": "skip",
    "/quit": "quit", "quit": "quit", "/q": "quit", "q": "quit", "/stop": "quit", "stop": "quit",
}

_stdin_q = queue.Queue()
_stdin_thread = None


def _stdin_reader():
    # One persistent daemon reader for the whole process: feeds every typed
    # line into _stdin_q so prompt_and_wait can poll it alongside Telegram
    # without a blocking input() that we couldn't cancel when the Telegram
    # side answers first. Daemon, so it never holds up interpreter exit.
    try:
        for line in sys.stdin:
            _stdin_q.put(line)
    except Exception:  # noqa: BLE001 - stdin closed/redirected; just stop reading
        pass


def _ensure_stdin_thread():
    global _stdin_thread
    if _stdin_thread is None:
        _stdin_thread = threading.Thread(target=_stdin_reader, daemon=True)
        _stdin_thread.start()


def _parse_command(text, terminal=False):
    if text is None:
        return None
    t = text.strip().lower()
    if terminal and t == "":
        return "continue"  # bare Enter at the terminal = continue (unchanged behaviour)
    if t.startswith("/") and "@" in t:
        t = t.split("@", 1)[0]  # strip @botname from "/continue@your_logs_bot"
    return _CMD_MAP.get(t)


def _get_updates(offset=None, poll=0):
    params = {"timeout": poll}
    if offset is not None:
        params["offset"] = offset
    # HTTP timeout must exceed the long-poll timeout or requests aborts it.
    resp = requests.get(f"{_API_BASE}/getUpdates", params=params, timeout=poll + 15)
    return resp.json().get("result", [])


def _command_from_update(u):
    """Return a command word if update u is a slash command in OUR group from
    the allowed user (mention_username, if set), else None."""
    msg = u.get("message") or {}
    chat = msg.get("chat") or {}
    if str(chat.get("id")) != str(GROUP_ID):
        return None
    if MENTION_USERNAME:
        sender = (msg.get("from") or {}).get("username") or ""
        if sender.lower() != MENTION_USERNAME.lower():
            return None
    return _parse_command(msg.get("text"))


def prompt_and_wait(message, terminal_prompt, timeout=3600, poll=8):
    """Ask the user to make a decision and block until they answer, from
    EITHER the Telegram group (a slash command - /continue, /skip, /quit)
    OR the terminal - whichever comes first. Returns 'continue', 'skip', or
    'quit'.

    If the relay isn't configured, falls back to a plain blocking input()
    on terminal_prompt (unchanged old behaviour). If it IS configured but
    nobody answers within `timeout` seconds, returns 'skip' - the safe
    default (skipping changes nothing), so an unattended run can't hang
    forever on one order.

    Best-effort on the Telegram side: any API error just means that source
    is temporarily unavailable this poll; the terminal still works."""
    if not _API_BASE or not GROUP_ID:
        # No relay - old terminal-only behaviour.
        return _parse_command(input(terminal_prompt), terminal=True) or "continue"

    _ensure_stdin_thread()
    # Discard any terminal line typed before this prompt so a stale keystroke
    # doesn't auto-answer it.
    while not _stdin_q.empty():
        try:
            _stdin_q.get_nowait()
        except queue.Empty:
            break

    # Baseline the Telegram update cursor so we only react to messages sent
    # AFTER this prompt: drain+confirm everything currently pending.
    offset = None
    try:
        pending = _get_updates(poll=0)
        if pending:
            offset = pending[-1]["update_id"] + 1
            _get_updates(offset=offset, poll=0)  # confirm/delete them
    except requests.exceptions.RequestException:
        pass

    mention(message)  # tag + show the prompt in the group
    print(terminal_prompt, end="", flush=True)

    deadline = time.time() + timeout
    while time.time() < deadline:
        # 1) terminal (non-blocking - the daemon reader feeds the queue)
        try:
            line = _stdin_q.get_nowait()
            cmd = _parse_command(line, terminal=True)
            if cmd:
                return cmd
            print(terminal_prompt, end="", flush=True)  # unrecognised - re-prompt
        except queue.Empty:
            pass
        # 2) Telegram (long-poll up to `poll` seconds - also paces the loop)
        try:
            updates = _get_updates(offset=offset, poll=poll)
            for u in updates:
                offset = u["update_id"] + 1
                cmd = _command_from_update(u)
                if cmd:
                    print(f"\n-> received /{cmd} from Telegram", file=sys.__stdout__)
                    return cmd
        except requests.exceptions.RequestException:
            time.sleep(poll)  # transient - back off, keep the terminal responsive-ish

    print("\n-> no answer within the time limit - defaulting to skip (nothing changed).", file=sys.__stdout__)
    return "skip"


def send_file(path, caption=""):
    """Send a file (e.g. the per-run Excel report) to the Telegram group as
    a document. Best-effort - never raises. A no-op if the relay isn't
    configured. Any buffered console output is flushed first so the file
    arrives after the run's log/report text, not interleaved with it."""
    flush()
    if not _API_BASE or not GROUP_ID:
        return
    try:
        with open(path, "rb") as f:
            requests.post(
                f"{_API_BASE}/sendDocument",
                data={"chat_id": GROUP_ID, "caption": caption[:1024]},
                files={"document": f},
                timeout=60,
            )
    except (requests.exceptions.RequestException, OSError) as e:
        print(f"   (Telegram file send failed: {e!r} - continuing)", file=sys.__stdout__)
