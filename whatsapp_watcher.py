"""
whatsapp_watcher.py
────────────────────────────────────────────────────────────
Watches the local WhatsApp bridge SQLite database for a new
PPTX document received from a specific contact.

When found:
  1. Downloads the PPTX via the Go bridge REST API
  2. Builds a config JSON
  3. Launches guru_auto_form.py which fills the form,
     uploads the PPTX, and auto-submits.

Prerequisites:
  - Go bridge must be running: cd whatsapp-mcp/whatsapp-bridge && go run main.go
  - Set SENDER_PHONE below to the contact's phone (country code, no +)
"""

import sqlite3
import subprocess
import threading
import time
import json
import os
import requests
from datetime import datetime, timezone

# ── Load central user configuration ──────────────────────────────────────────
from config import (
    ALLOWED_SENDERS,
    FORM_LINK,
    DEFAULT_MENTEE_NAME,
    DEFAULT_REGISTER_NUMBER,
    DEFAULT_MENTOR_NAME,
    DEFAULT_DEPARTMENT,
    NOTIFY_PHONE,
    AUTO_SUBMIT,
    HEADLESS,
    POLL_INTERVAL,
    BRIDGE_API,
    EVENT_SERVER_PORT,
    CHROME_PROFILE_DIR,
)

# ─────────────────────────── RUNTIME CONSTANTS ───────────────────────────────
# (All user-configurable values are in config.py — edit that file, not here)

# Path to the Go bridge's SQLite message database (relative to THIS file)
DB_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "whatsapp-mcp", "whatsapp-bridge", "store", "messages.db"
)

# Path to your form-filling script
FORM_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guru_auto_form.py")

# Default form config assembled from config.py values
DEFAULT_FORM_CONFIG = {
    "form_link":        FORM_LINK,
    "department":       DEFAULT_DEPARTMENT,
    "mentee_name":      DEFAULT_MENTEE_NAME,
    "register_number":  DEFAULT_REGISTER_NUMBER,
    "mentor_name":      DEFAULT_MENTOR_NAME,
    "auto_submit":      AUTO_SUBMIT,
    "headless":         HEADLESS,
    "notify_phone":     NOTIFY_PHONE,
}

# ─────────────────────────────────────────────────────────────────────────────

# Track already-processed message IDs so we never trigger the same file twice.
PROCESSED_IDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".processed_wa_ids.json")

# Global lock — ensures the event server thread and poll loop
# can never process the same message_id concurrently.
_trigger_lock = threading.Lock()

def load_processed_ids() -> set:
    if os.path.exists(PROCESSED_IDS_FILE):
        with open(PROCESSED_IDS_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_processed_ids(ids: set):
    with open(PROCESSED_IDS_FILE, "w") as f:
        json.dump(list(ids), f)

def get_new_pptx_messages(processed_ids: set) -> list:
    """Query messages.db for unprocessed PPTX documents from allowed senders only."""
    if not os.path.exists(DB_PATH):
        print(f"[ERROR] messages.db not found at: {DB_PATH}")
        print("        Is the Go bridge running? (go run main.go inside whatsapp-bridge/)")
        return []

    try:
        conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
        cursor = conn.cursor()

        # Build sender conditions — only match the SENDER field (who sent us the file),
        # NOT chat_jid (which would also match files WE sent TO that person).
        # Also enforce is_from_me = 0 so only received messages trigger.
        sender_conditions = []
        params = []
        for sender_id in ALLOWED_SENDERS:
            sender_conditions.append("m.sender LIKE ?")
            params.append(f"%{sender_id}%")

        sender_where = " OR ".join(sender_conditions)

        # Only look at messages received today (local date)
        today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        params.append(today_start)

        cursor.execute(f"""
            SELECT
                m.id,
                m.chat_jid,
                m.sender,
                m.filename,
                m.media_type,
                m.timestamp,
                c.name
            FROM messages m
            JOIN chats c ON m.chat_jid = c.jid
            WHERE
                ({sender_where})
                AND m.is_from_me = 0
                AND m.media_type = 'document'
                AND (
                    LOWER(m.filename) LIKE '%.pptx'
                    OR LOWER(m.filename) LIKE '%.ppt'
                )
                AND m.timestamp >= ?
            ORDER BY m.timestamp DESC
            LIMIT 20
        """, params)
        rows = cursor.fetchall()
        conn.close()
        # Filter out already-processed message IDs only (not filenames —
        # same filename can legitimately arrive in a new message next week)
        return [r for r in rows if r[0] not in processed_ids]
    except sqlite3.Error as e:
        print(f"[DB ERROR] {e}")
        return []

def download_pptx(message_id: str, chat_jid: str) -> str | None:
    """Call the Go bridge /api/download endpoint to download the file locally."""
    print(f"[DOWNLOAD] Requesting download of message {message_id} in chat {chat_jid}...")
    try:
        resp = requests.post(
            f"{BRIDGE_API}/download",
            json={"message_id": message_id, "chat_jid": chat_jid},
            timeout=60
        )
        if resp.status_code == 200:
            data = resp.json()
            if data.get("success"):
                path = data.get("path")
                print(f"[DOWNLOAD] ✅ File saved to: {path}")
                return path
            else:
                print(f"[DOWNLOAD] ❌ Bridge error: {data.get('message')}")
        else:
            print(f"[DOWNLOAD] ❌ HTTP {resp.status_code}: {resp.text}")
    except requests.RequestException as e:
        print(f"[DOWNLOAD] ❌ Request failed: {e}")
    return None

def trigger_form_script(pptx_path: str):
    """Write a temp config JSON and launch guru_auto_form.py."""
    config = dict(DEFAULT_FORM_CONFIG)
    config["pptx_path"] = pptx_path

    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_config.json")
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)

    print(f"[TRIGGER] Launching guru_auto_form.py with PPTX: {pptx_path}")

    # ── Kill Chrome and release the Selenium profile lock ──────────────────
    # When launched from start_automation.bat, Chrome may already be open.
    # taskkill alone isn't enough — Chrome leaves a SingletonLock file in the
    # profile directory that prevents re-launch. We delete it explicitly.
    print("[TRIGGER] Ensuring Chrome is closed and profile lock is cleared...")
    subprocess.call(
        ["taskkill", "/F", "/IM", "chrome.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    time.sleep(3)  # Give OS time to release file handles

    # Delete SingletonLock (and SingletonCookie) left by Chrome
    import glob
    for lock_pattern in [
        os.path.join(CHROME_PROFILE_DIR, "SingletonLock"),
        os.path.join(CHROME_PROFILE_DIR, "SingletonCookie"),
        os.path.join(CHROME_PROFILE_DIR, "SingletonSocket"),
    ]:
        for lock_file in glob.glob(lock_pattern):
            try:
                os.remove(lock_file)
                print(f"[TRIGGER] Removed lock file: {lock_file}")
            except OSError:
                pass  # Already gone — that's fine

    # Popen (non-blocking) — the browser will open and handle itself
    subprocess.Popen(
        ["python", FORM_SCRIPT, config_path],
        creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0)  # Opens in a new window on Windows
    )

def handle_pptx_event(msg_id: str, chat_jid: str, sender: str, filename: str):
    """
    Shared handler called by both the Flask endpoint and the poll loop.
    Uses a lock so the exact same message_id is NEVER processed twice.

    Dedup strategy (two keys):
      1. message ID          - prevents race between Flask push + poll for the SAME send.
      2. file::DATE::name    - blocks resending the same file on the SAME DAY,
                               but allows the same filename next week
                               (different date = different key).
    """
    today = datetime.now().strftime("%Y-%m-%d")
    filename_day_key = f"file::{today}::{filename.strip().lower()}"

    # Fast pre-check without the lock
    processed = load_processed_ids()
    if msg_id in processed:
        print(f"[SKIP] Already processed (message ID): {msg_id}")
        return
    if filename_day_key in processed:
        print(f"[SKIP] Already submitted today (same filename): {filename}")
        print(f"       Same filename next week will trigger normally.")
        return

    # Acquire lock — only ONE thread can be inside this block at a time
    with _trigger_lock:
        # Re-check inside the lock (another thread may have just marked it)
        processed = load_processed_ids()
        if msg_id in processed:
            print(f"[SKIP] Already processed (lock - msg ID): {msg_id}")
            return
        if filename_day_key in processed:
            print(f"[SKIP] Already submitted today (lock - filename): {filename}")
            return

        print(f"\n{'='*60}")
        print(f"[MATCH] 📎 PPTX received!")
        print(f"        From    : {sender}")
        print(f"        File    : {filename}")
        print(f"        Date key: {filename_day_key}")
        print(f"{'='*60}")

        # Mark BOTH keys immediately (before download starts)
        processed.add(msg_id)
        processed.add(filename_day_key)
        save_processed_ids(processed)
        print(f"[INFO] Marked as processed: msg={msg_id}")
        print(f"[INFO] Marked as processed: {filename_day_key}")

    # Download and trigger OUTSIDE the lock (can take time; no need to block)
    pptx_path = download_pptx(msg_id, chat_jid)
    if pptx_path:
        trigger_form_script(pptx_path)
    else:
        print(f"[WARN] Download failed for {filename}. Will NOT retry (already marked processed).")


def start_event_server():
    """
    Run a tiny Flask HTTP server on port 9999.
    The Go bridge POSTs here instantly when a PPTX arrives —
    no need to wait for the poll cycle.
    """
    try:
        from flask import Flask, request, jsonify
    except ImportError:
        print("[EVENT SERVER] Flask not installed. Run: pip install flask")
        print("[EVENT SERVER] Falling back to poll-only mode.")
        return

    app = Flask(__name__)

    @app.route("/pptx-received", methods=["POST"])
    def on_pptx():
        data = request.get_json(force=True)
        msg_id   = data.get("message_id", "")
        chat_jid = data.get("chat_jid", "")
        sender   = data.get("sender", "")
        filename = data.get("filename", "")
        print(f"\n[EVENT] ⚡ Instant push from Go bridge: {filename}")
        # Run in background thread so Flask returns immediately
        import threading
        threading.Thread(
            target=handle_pptx_event,
            args=(msg_id, chat_jid, sender, filename),
            daemon=True
        ).start()
        return jsonify({"ok": True})

    @app.route("/health", methods=["GET"])
    def health():
        return jsonify({"status": "running"})

    print(f"[EVENT SERVER] ⚡ Listening on http://localhost:{EVENT_SERVER_PORT} for instant bridge events")
    # Use threaded=True so it doesn't block; werkzeug handles requests alongside poll loop
    app.run(host="127.0.0.1", port=EVENT_SERVER_PORT, threaded=True, use_reloader=False)


def main():
    processed_ids = load_processed_ids()

    print("=" * 60)
    print("  WhatsApp PPTX Watcher")
    print("=" * 60)
    print(f"  Allowed senders ({len(ALLOWED_SENDERS)}):")
    for s in ALLOWED_SENDERS:
        print(f"    • {s}")

    print(f"  Database     : {DB_PATH}")
    print(f"  Poll interval: {POLL_INTERVAL}s (backup)")
    print(f"  Event server : http://localhost:{EVENT_SERVER_PORT} (instant trigger)")
    print(f"  Already seen : {len(processed_ids)} message(s)")
    print("=" * 60)

    if not os.path.exists(DB_PATH):
        print("\n[WAIT] messages.db doesn't exist yet.")
        print("       Start the Go bridge first, then re-run this watcher.\n")

    # Start Flask event server in a background thread
    import threading
    server_thread = threading.Thread(target=start_event_server, daemon=True)
    server_thread.start()

    # Poll loop runs in main thread as backup (catches anything the event server missed)
    while True:
        try:
            new_messages = get_new_pptx_messages(processed_ids)

            if new_messages:
                for (msg_id, chat_jid, sender, filename, media_type, timestamp, chat_name) in new_messages:
                    handle_pptx_event(msg_id, chat_jid, sender, filename)
                    processed_ids = load_processed_ids()  # Reload after handler updates it
            else:
                now = datetime.now().strftime("%H:%M:%S")
                print(f"[{now}] Polling... (next check in {POLL_INTERVAL}s)", end="\r")

        except Exception as e:
            print(f"\n[FATAL ERROR] {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()

