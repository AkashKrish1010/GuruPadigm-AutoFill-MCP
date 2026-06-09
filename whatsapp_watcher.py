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
import sys

# Force UTF-8 encoding for stdout/stderr to support emojis on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

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

# Enforce that headless mode only runs if auto_submit is also enabled
actual_headless = HEADLESS
if HEADLESS and not AUTO_SUBMIT:
    print("[WARNING] HEADLESS is set to True but AUTO_SUBMIT is False. Overriding HEADLESS to False to allow manual review/submission.")
    actual_headless = False

# Default form config assembled from config.py values
DEFAULT_FORM_CONFIG = {
    "form_link":        FORM_LINK,
    "department":       DEFAULT_DEPARTMENT,
    "mentee_name":      DEFAULT_MENTEE_NAME,
    "register_number":  DEFAULT_REGISTER_NUMBER,
    "mentor_name":      DEFAULT_MENTOR_NAME,
    "auto_submit":      AUTO_SUBMIT,
    "headless":         actual_headless,
    "notify_phone":     NOTIFY_PHONE,
}

# ─────────────────────────────────────────────────────────────────────────────

# Track already-processed message IDs so we never trigger the same file twice.
PROCESSED_IDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".processed_wa_ids.json")

# Global lock — ensures the event server thread and poll loop
# can never process the same message_id concurrently.
_trigger_lock = threading.Lock()

# In-memory list of message IDs and file keys currently being processed
IN_PROGRESS_IDS = set()

def load_processed_ids() -> set:
    if os.path.exists(PROCESSED_IDS_FILE):
        with open(PROCESSED_IDS_FILE, "r") as f:
            return set(json.load(f))
    return set()

def save_processed_ids(ids: set):
    with open(PROCESSED_IDS_FILE, "w") as f:
        json.dump(list(ids), f)


def perform_maintenance():
    """
    Delete any files in the bridge store/ directory older than 30 days,
    and clean up SQLite messages database entries older than 30 days.
    """
    print("[MAINTENANCE] Running database and file cleanup...")
    db_path = DB_PATH
    store_dir = os.path.join(os.path.dirname(db_path)) # store/ folder
    
    # 1. Clean up old files in store/
    now = time.time()
    cutoff_time = now - (30 * 24 * 60 * 60) # 30 days in seconds
    files_deleted = 0
    
    if os.path.exists(store_dir):
        for root, dirs, files in os.walk(store_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Don't delete database files!
                if file.endswith((".db", ".db-journal", ".db-shm", ".db-wal")):
                    continue
                try:
                    file_stat = os.stat(file_path)
                    if file_stat.st_mtime < cutoff_time:
                        os.remove(file_path)
                        files_deleted += 1
                except Exception as e:
                    print(f"[MAINTENANCE] Error deleting file {file_path}: {e}")
                    
    if files_deleted > 0:
        print(f"[MAINTENANCE] Deleted {files_deleted} files older than 30 days.")
    else:
        print("[MAINTENANCE] No old files to clean up.")

    # 2. Clean up database entries older than 30 days
    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            from datetime import datetime, timedelta
            cutoff_date = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
            
            # Delete messages older than 30 days
            cursor.execute("DELETE FROM messages WHERE timestamp < ?", (cutoff_date,))
            deleted_messages = cursor.rowcount
            
            # Delete chats that haven't had messages in 30 days
            cursor.execute("DELETE FROM chats WHERE last_message_time < ?", (cutoff_date,))
            deleted_chats = cursor.rowcount
            
            conn.commit()
            conn.close()
            
            if deleted_messages > 0 or deleted_chats > 0:
                print(f"[MAINTENANCE] Deleted {deleted_messages} messages and {deleted_chats} chats older than 30 days.")
            else:
                print("[MAINTENANCE] Database is already clean.")
        except Exception as e:
            print(f"[MAINTENANCE] Database cleanup failed: {e}")

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

        # Only look at messages received recently (e.g. in the last 1 hour)
        # to avoid processing old historical messages on startup.
        from datetime import timedelta
        recent_start = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
        params.append(recent_start)

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

        # Filter out already-processed message IDs and enforce the FILENAME_PREFIX filter
        from config import FILENAME_PREFIX
        filtered_rows = []
        for r in rows:
            msg_id = r[0]
            filename = r[3]
            today = datetime.now().strftime("%Y-%m-%d")
            filename_day_key = f"file::{today}::{filename.strip().lower()}"
            if msg_id in processed_ids or msg_id in IN_PROGRESS_IDS:
                continue
            if filename_day_key in processed_ids or filename_day_key in IN_PROGRESS_IDS:
                continue
            if FILENAME_PREFIX and not filename.strip().lower().startswith(FILENAME_PREFIX.strip().lower()):
                continue
            filtered_rows.append(r)
        return filtered_rows
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

def run_and_monitor_form_script(msg_id: str, filename_day_key: str, config_path: str):
    print(f"[MONITOR] Starting form-filler process for message {msg_id}...")
    proc = subprocess.Popen(
        [sys.executable, FORM_SCRIPT, config_path],
        creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0)  # Opens in a new window on Windows
    )
    # Wait for the process to finish
    return_code = proc.wait()
    
    if return_code != 0:
        print(f"[MONITOR] [WARNING] Form-filler failed with exit code {return_code} for message {msg_id}.")
        # Remove from in-progress list so it can be retried!
        with _trigger_lock:
            IN_PROGRESS_IDS.discard(msg_id)
            IN_PROGRESS_IDS.discard(filename_day_key)
            print(f"[MONITOR] Removed from in-progress cache. Ready for retry.")
    else:
        print(f"[MONITOR] [SUCCESS] Form-filler finished successfully for message {msg_id}.")
        # Succeeded. The form script itself has written to .processed_wa_ids.json.
        # Clean up files immediately since they are fully processed and submitted
        try:
            if os.path.exists(config_path):
                with open(config_path, "r") as f:
                    cfg = json.load(f)
                pptx_path = cfg.get("pptx_path")
                if pptx_path and os.path.exists(pptx_path):
                    os.remove(pptx_path)
                    print(f"[MONITOR] Cleaned up processed PPTX file: {pptx_path}")
                    # Also delete PDF if it exists
                    pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
                    if os.path.exists(pdf_path):
                        os.remove(pdf_path)
                        print(f"[MONITOR] Cleaned up processed PDF file: {pdf_path}")
        except Exception as e:
            print(f"[MONITOR] [WARNING] Error cleaning up files: {e}")

        # Remove config file
        try:
            if os.path.exists(config_path):
                os.remove(config_path)
        except OSError:
            pass

        # But we remove from in-progress list since it is now permanently in the processed list.
        with _trigger_lock:
            IN_PROGRESS_IDS.discard(msg_id)
            IN_PROGRESS_IDS.discard(filename_day_key)


def trigger_form_script(msg_id: str, filename_day_key: str, pptx_path: str):
    """Write a temp config JSON and launch guru_auto_form.py."""
    config = dict(DEFAULT_FORM_CONFIG)
    config["pptx_path"] = pptx_path
    config["msg_id"] = msg_id
    config["filename_day_key"] = filename_day_key

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

    # Start the monitoring thread
    threading.Thread(
        target=run_and_monitor_form_script,
        args=(msg_id, filename_day_key, config_path),
        daemon=True
    ).start()


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
    # Filename prefix validation
    from config import FILENAME_PREFIX
    if FILENAME_PREFIX and not filename.strip().lower().startswith(FILENAME_PREFIX.strip().lower()):
        print(f"[SKIP] Filename '{filename}' does not start with configured prefix '{FILENAME_PREFIX}'")
        return

    today = datetime.now().strftime("%Y-%m-%d")
    filename_day_key = f"file::{today}::{filename.strip().lower()}"

    # Fast pre-check without the lock
    processed = load_processed_ids()
    if msg_id in processed or msg_id in IN_PROGRESS_IDS:
        print(f"[SKIP] Already processed or in progress (message ID): {msg_id}")
        return
    if filename_day_key in processed or filename_day_key in IN_PROGRESS_IDS:
        print(f"[SKIP] Already submitted today or in progress (same filename): {filename}")
        print(f"       Same filename next week will trigger normally.")
        return

    # Acquire lock — only ONE thread can be inside this block at a time
    with _trigger_lock:
        # Re-check inside the lock (another thread may have just marked it)
        processed = load_processed_ids()
        if msg_id in processed or msg_id in IN_PROGRESS_IDS:
            print(f"[SKIP] Already processed or in progress (lock - msg ID): {msg_id}")
            return
        if filename_day_key in processed or filename_day_key in IN_PROGRESS_IDS:
            print(f"[SKIP] Already submitted today or in progress (lock - filename): {filename}")
            return

        print(f"\n{'='*60}")
        print(f"[MATCH] [INFO] PPTX received!")
        print(f"        From    : {sender}")
        print(f"        File    : {filename}")
        print(f"        Date key: {filename_day_key}")
        print(f"{'='*60}")

        # Mark in progress (in-memory only)
        IN_PROGRESS_IDS.add(msg_id)
        IN_PROGRESS_IDS.add(filename_day_key)
        print(f"[INFO] Marked as in progress: msg={msg_id}")
        print(f"[INFO] Marked as in progress: {filename_day_key}")

    # Download and trigger OUTSIDE the lock (can take time; no need to block)
    pptx_path = download_pptx(msg_id, chat_jid)
    if pptx_path:
        trigger_form_script(msg_id, filename_day_key, pptx_path)
    else:
        print(f"[WARN] Download failed for {filename}. Removing from in-progress list.")
        with _trigger_lock:
            IN_PROGRESS_IDS.discard(msg_id)
            IN_PROGRESS_IDS.discard(filename_day_key)


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
        print(f"\n[EVENT] [INFO] Instant push from Go bridge: {filename}")
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

    print(f"[EVENT SERVER] [INFO] Listening on http://localhost:{EVENT_SERVER_PORT} for instant bridge events")
    # Use threaded=True so it doesn't block; werkzeug handles requests alongside poll loop
    app.run(host="127.0.0.1", port=EVENT_SERVER_PORT, threaded=True, use_reloader=False)


def main():
    # Check if Chrome profile directory exists
    if not os.path.exists(CHROME_PROFILE_DIR):
        print("[WARNING] Chrome profile directory does not exist.")
        print("          Automatically running login helper to configure your Google Account...")
        subprocess.call([sys.executable, "guru_auto_form.py", "--login"])

    processed_ids = load_processed_ids()

    print("=" * 60)
    print("  WhatsApp PPTX Watcher")
    print("=" * 60)
    print(f"  Allowed senders ({len(ALLOWED_SENDERS)}):")
    for s in ALLOWED_SENDERS:
        print(f"    - {s}")

    print(f"  Database     : {DB_PATH}")
    print(f"  Poll interval: {POLL_INTERVAL}s (backup)")
    print(f"  Event server : http://localhost:{EVENT_SERVER_PORT} (instant trigger)")
    print(f"  Already seen : {len(processed_ids)} message(s)")
    if DEFAULT_FORM_CONFIG["headless"]:
        print("  Browser Mode : Headless (Invisible background runs)")
    else:
        print("  Browser Mode : Headed (Visible window runs)")
    print("=" * 60)

    if not os.path.exists(DB_PATH):
        print("\n[WAIT] messages.db doesn't exist yet.")
        print("       Start the Go bridge first, then re-run this watcher.\n")

    # Run database and file maintenance
    try:
        perform_maintenance()
    except Exception as e:
        print(f"[MAINTENANCE] [WARNING] Maintenance task failed: {e}")

    # Start Flask event server in a background thread
    import threading
    server_thread = threading.Thread(target=start_event_server, daemon=True)
    server_thread.start()

    # Poll loop runs in main thread as backup (catches anything the event server missed)
    while True:
        try:
            processed_ids = load_processed_ids()
            new_messages = get_new_pptx_messages(processed_ids)

            if new_messages:
                for (msg_id, chat_jid, sender, filename, media_type, timestamp, chat_name) in new_messages:
                    handle_pptx_event(msg_id, chat_jid, sender, filename)
            else:
                now = datetime.now().strftime("%H:%M:%S")
                print(f"[{now}] Polling... (next check in {POLL_INTERVAL}s)", end="\r")

        except Exception as e:
            print(f"\n[FATAL ERROR] {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()

