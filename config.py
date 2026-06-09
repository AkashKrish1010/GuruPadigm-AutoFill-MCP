"""
config.py  ──  GuruPadigm-MCP-AutoFill  ·  Central Configuration
═══════════════════════════════════════════════════════════════════
THIS IS THE ONLY FILE YOU NEED TO EDIT.

Fill in your details below, then run:
    start_automation.bat
or:
    python whatsapp_watcher.py
"""

# ─────────────────────────────────────────────────────────────────
#  1. ALLOWED WHATSAPP SENDERS
#     Phone numbers (country code, no +) whose PPTX messages should
#     trigger the form.  Example: +91 98765 43210  →  "919876543210"
#     You can add WhatsApp LIDs too (visible in the DB as long numbers).
# ─────────────────────────────────────────────────────────────────
ALLOWED_SENDERS = [
    "919876543210",  # Replace with sender's number (e.g. "919876543210")
]

# ─────────────────────────────────────────────────────────────────
#  2. GOOGLE FORM LINK
#     Paste the full URL of your Google Form here.
# ─────────────────────────────────────────────────────────────────
FORM_LINK = "https://forms.gle/YOUR_GOOGLE_FORM_LINK"

# ─────────────────────────────────────────────────────────────────
#  3. FORM FIELD DEFAULTS
#     These are used as fallback values if the PPTX data can't be
#     parsed automatically.
# ─────────────────────────────────────────────────────────────────
DEFAULT_MENTEE_NAME     = "Your Student/Mentee Name"
DEFAULT_REGISTER_NUMBER = "19XXXXXXXX"
DEFAULT_MENTOR_NAME     = "Dr. Mentor Name"
DEFAULT_DEPARTMENT      = "Your Department Dropdown Option"

# ─────────────────────────────────────────────────────────────────
#  4. WHATSAPP NOTIFICATION NUMBER
#     After a successful form submission, a WhatsApp confirmation
#     is sent to this number. Same format: country code + number.
#     Leave as "" to disable notifications.
# ─────────────────────────────────────────────────────────────────
NOTIFY_PHONE = "919876543210"   # e.g. "919876543210"

# ─────────────────────────────────────────────────────────────────
#  5. CHROME PROFILE PATH & NAME  (usually no need to change)
#     Playwright uses a dedicated Chrome profile to stay logged in
#     to Google. Point this at wherever you created the profile.
#
#     CHROME_PROFILE_DIR  → the folder passed to --user-data-dir
#     CHROME_PROFILE_NAME → the profile sub-folder inside that dir
#                           (almost always "Default" unless you set
#                           up multiple Chrome profiles — check
#                           chrome://version to confirm yours)
# ─────────────────────────────────────────────────────────────────
CHROME_PROFILE_DIR  = r"C:\PlaywrightProfiles\GuruProfile"
CHROME_PROFILE_NAME = "Default"  # Change to "Profile 1", "Profile 2", etc. if needed

# ─────────────────────────────────────────────────────────────────
#  6. BEHAVIOUR FLAGS
# ─────────────────────────────────────────────────────────────────
AUTO_SUBMIT    = True    # Set True to submit the form automatically and Must be True for headless mode to
HEADLESS       = True   # Set True to run Chrome without a visible window
POLL_INTERVAL  = 300     # Seconds between DB polls (default: 5 minutes)

# ─────────────────────────────────────────────────────────────────
#  8. FILENAME PREFIX FILTER (Optional)
#     Only process PPTX files whose filename starts with this prefix.
#     Case-insensitive. Set to "" or None to disable filtering and
#     process all PPTX files from allowed senders.
# ─────────────────────────────────────────────────────────────────
FILENAME_PREFIX = "19XXXXXXXX Student Name SSE Gurupadigam"

# ─────────────────────────────────────────────────────────────────
#  7. SERVICE ENDPOINTS  (change only if you moved the defaults)
# ─────────────────────────────────────────────────────────────────
BRIDGE_API     = "http://localhost:8080/api"   # Go WhatsApp bridge REST API
EVENT_SERVER_PORT = 9999                        # Flask webhook listener port
