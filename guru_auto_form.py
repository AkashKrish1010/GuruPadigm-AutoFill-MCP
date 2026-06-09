import os
import shutil
import subprocess
import time
import sys
import json
import requests
from datetime import datetime
from playwright.sync_api import sync_playwright

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

# ── Load central user configuration (edit config.py, not this file) ──────────
from config import (
    FORM_LINK,
    DEFAULT_MENTEE_NAME,
    DEFAULT_REGISTER_NUMBER,
    DEFAULT_MENTOR_NAME,
    DEFAULT_DEPARTMENT,
    NOTIFY_PHONE as CFG_NOTIFY_PHONE,
    CHROME_PROFILE_DIR,
    CHROME_PROFILE_NAME,
    AUTO_SUBMIT as CFG_AUTO_SUBMIT,
    HEADLESS as CFG_HEADLESS,
    BRIDGE_API,
)

custom_profile_dir = CHROME_PROFILE_DIR

# ─── Load configuration from file (passed by whatsapp_watcher.py) ────────────
PPTX_PATH = None   # Will be set below if config has it

if len(sys.argv) > 1 and sys.argv[1] != "--login":
    config_file = sys.argv[1]
    with open(config_file, 'r') as f:
        config = json.load(f)
    form_link       = config.get('form_link', FORM_LINK)
    DEPARTMENT_NAME = config.get('department', DEFAULT_DEPARTMENT)
    PPTX_PATH       = config.get('pptx_path', None)
    AUTO_SUBMIT     = config.get('auto_submit', CFG_AUTO_SUBMIT)
    HEADLESS        = config.get('headless', CFG_HEADLESS)
    NOTIFY_PHONE    = config.get('notify_phone', CFG_NOTIFY_PHONE)
    current_date    = datetime.now().strftime("%Y-%m-%d")
    person_data = {
        "mentee_name":     config.get('mentee_name', DEFAULT_MENTEE_NAME),
        "date":            current_date,
        "register_number": config.get('register_number', DEFAULT_REGISTER_NUMBER),
        "mentor_name":     config.get('mentor_name', DEFAULT_MENTOR_NAME),
    }
else:
    # ── Standalone / manual execution defaults (from config.py) ─────────────
    form_link       = FORM_LINK
    DEPARTMENT_NAME = DEFAULT_DEPARTMENT
    PPTX_PATH       = None
    AUTO_SUBMIT     = CFG_AUTO_SUBMIT
    HEADLESS        = CFG_HEADLESS
    NOTIFY_PHONE    = CFG_NOTIFY_PHONE
    current_date    = datetime.now().strftime("%Y-%m-%d")
    person_data = {
        "mentee_name":     DEFAULT_MENTEE_NAME,
        "date":            current_date,
        "register_number": DEFAULT_REGISTER_NUMBER,
        "mentor_name":     DEFAULT_MENTOR_NAME,
    }
    config = {}

# Enforce that headless mode only runs if auto_submit is also enabled
if HEADLESS and not AUTO_SUBMIT:
    print("[WARNING] HEADLESS is set to True but AUTO_SUBMIT is False. Overriding HEADLESS to False to allow manual review/submission.")
    HEADLESS = False

# ─────────────────────────────────────────────────────────────────────────────


def send_whatsapp_notification(phone: str, message: str):
    """
    Send a WhatsApp message via the local Go bridge.
    The bridge must be running: cd whatsapp-mcp/whatsapp-bridge && go run main.go
    """
    try:
        resp = requests.post(
            f"{BRIDGE_API}/send",
            json={"recipient": phone, "message": message},
            timeout=10
        )
        if resp.status_code == 200 and resp.json().get("success"):
            print(f"[NOTIFY] [OK] WhatsApp notification sent to {phone}")
        else:
            print(f"[NOTIFY] [WARNING] Bridge responded: {resp.text}")
    except requests.RequestException as e:
        print(f"[NOTIFY] [WARNING] Could not reach WhatsApp bridge: {e}")
        print("[NOTIFY]    Is the Go bridge running? (go run main.go inside whatsapp-bridge/)")


def mark_processed(msg_id: str | None, filename_day_key: str | None):
    """
    Add the message ID and filename day key to the processed registry.
    This prevents the watcher from triggering the same file/message again.
    """
    if not msg_id and not filename_day_key:
        return
    processed_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".processed_wa_ids.json")
    processed_ids = set()
    if os.path.exists(processed_file):
        try:
            with open(processed_file, "r") as f:
                processed_ids = set(json.load(f))
        except Exception as e:
            print(f"[WARNING] Could not read processed IDs file: {e}")
    
    if msg_id:
        processed_ids.add(msg_id)
    if filename_day_key:
        processed_ids.add(filename_day_key)
        
    try:
        with open(processed_file, "w") as f:
            json.dump(list(processed_ids), f)
        print(f"[INFO] Marked as processed in .processed_wa_ids.json: msg_id={msg_id}, key={filename_day_key}")
    except Exception as e:
        print(f"[ERROR] Could not write to processed IDs file: {e}")


def launch_chrome_for_login():
    """
    Launches a real, non-automated Google Chrome instance using the configured
    profile directory. This allows the user to sign in to Google normally
    without being blocked by Google's automation/bot detection.
    """
    print("[LOGIN SETUP] Launching real Google Chrome for authentication...")
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    chrome_path = None
    for p in paths:
        if os.path.exists(p):
            chrome_path = p
            break
            
    if not chrome_path:
        chrome_path = shutil.which("chrome")
        
    if not chrome_path:
        print("[ERROR] Could not find chrome.exe in standard paths or PATH.")
        print(f"Please run it manually: chrome.exe --user-data-dir=\"{custom_profile_dir}\"")
        return

    print(f"[INFO] Using Chrome at: {chrome_path}")
    print(f"[INFO] Opening profile directory: {custom_profile_dir}")
    
    # Launch Chrome as a real, non-automated process
    proc = subprocess.Popen([chrome_path, f"--user-data-dir={custom_profile_dir}"])
    
    print("\n" + "=" * 60)
    print("  GOOGLE SIGN-IN HELPER")
    print("=" * 60)
    print("  1. A real Google Chrome window has been opened.")
    print("  2. Sign in to your Google Account (for the form submission).")
    print("  3. Once signed in, close the Chrome browser window.")
    print("=" * 60)
    input("\nPress Enter here once you have finished signing in and closed Chrome...")
    try:
        proc.terminate()
    except Exception:
        pass
    print("[LOGIN SETUP] Setup complete! You can now run the automation normally.")


def copy_chrome_profile():
    r"""
    One-time setup: ensures a dedicated Chrome profile directory exists
    for browser automation to use with persistent Google login.

    IMPORTANT - create a DEDICATED profile:
      1. Run:  python guru_auto_form.py --login
      2. Log in to the Google account used for the form
      3. Close Chrome
      4. Never open that profile in regular Chrome (only via automation)
    This avoids the "logged out" problem caused by session conflicts.
    """
    print("[STEP 1] Checking Chrome profile...")
    if os.path.exists(custom_profile_dir):
        print("[INFO] Custom profile already exists. Skipping setup.")
        return
    print(f"[INFO] Creating Chrome profile directory: {custom_profile_dir}")
    os.makedirs(custom_profile_dir, exist_ok=True)
    print(f"[DONE] Profile directory created.")
    print("[INFO] Automatically launching login helper...")
    launch_chrome_for_login()


def convert_pptx_to_pdf(pptx_path: str) -> str:
    """
    Convert a PPTX file to PDF using Microsoft PowerPoint COM (Windows).
    Falls back to LibreOffice if PowerPoint is not installed.
    Returns the absolute path to the generated PDF.
    """
    pptx_path = os.path.abspath(pptx_path)
    pdf_path  = os.path.splitext(pptx_path)[0] + ".pdf"

    # ── Try PowerPoint COM (requires MS Office) ───────────────────────────
    try:
        import comtypes.client
        print("[CONVERT] Converting PPTX → PDF via PowerPoint COM...")
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1   # must be visible for COM to work reliably
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)
        presentation.SaveAs(pdf_path, 32)   # 32 = ppSaveAsPDF
        presentation.Close()
        powerpoint.Quit()
        print(f"[CONVERT] [OK] PDF saved: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"[CONVERT] PowerPoint COM failed: {e}. Trying LibreOffice...")

    # ── Fallback: LibreOffice (free, must be installed) ───────────────────
    try:
        out_dir = os.path.dirname(pptx_path)
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             pptx_path, "--outdir", out_dir],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0 and os.path.isfile(pdf_path):
            print(f"[CONVERT] [OK] PDF saved via LibreOffice: {pdf_path}")
            return pdf_path
        else:
            raise RuntimeError(result.stderr)
    except Exception as e:
        raise RuntimeError(
            f"[CONVERT] Could not convert PPTX to PDF. "
            f"Install Microsoft Office or LibreOffice.\nError: {e}"
        )


def upload_pptx(page, pptx_path):
    """
    Upload a file to the Google Form file-upload field.
    Automatically converts PPTX to PDF before uploading.

    Strategy:
      1. Scroll to the file-upload section so it fully renders
      2. Click the form's "Add file" button → opens the Google Drive picker
      3. Locate the picker iframe (lives inside an iframe)
      4. Set input files on the <input type="file"> inside the iframe
      5. Wait for Google to finish uploading to Drive
    """
    print("[STEP 8] Preparing file for upload...")

    if not pptx_path or not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"[ERROR] File not found: {pptx_path}")

    # ── Convert PPTX → PDF ────────────────────────────────────────────────
    if pptx_path.lower().endswith((".pptx", ".ppt")):
        print("[STEP 8a] Converting PPTX to PDF...")
        upload_path = convert_pptx_to_pdf(pptx_path)
    else:
        upload_path = pptx_path

    abs_path = os.path.abspath(upload_path)
    print(f"[INFO] File to upload: {abs_path}")

    # ── Scroll to the file-upload section so it fully renders ───────────────
    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
    page.wait_for_timeout(1500)

    # ── STEP 8b: Click the form's upload button to open the Drive picker ──
    print("[STEP 8b] Clicking file upload button to open Drive picker...")
    upload_btn = None
    for selector in [
        'xpath=//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[3]/span/span[2]',
        'xpath=//span[contains(text(), "Add file") or contains(text(), "Choose file") or contains(text(), "Upload")]'
    ]:
        try:
            btn = page.wait_for_selector(selector, timeout=5000)
            if btn:
                btn.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
                btn.click()
                upload_btn = btn
                print("[INFO] Upload button clicked — Drive picker should open.")
                break
        except Exception as e:
            print(f"[DEBUG] Selector {selector} failed: {e}")

    if not upload_btn:
        raise RuntimeError("[ERROR] Could not click any upload button")

    # Wait for the Google Drive picker dialog to appear
    page.wait_for_timeout(3000)

    # ── STEP 8c: Switch into the Google Drive picker iframe ──────────────
    print("[STEP 8c] Looking for Google Drive picker iframe...")
    picker_frame = None
    
    # Check current frames for the Google Drive picker
    for frame in page.frames:
        if 'docs.google.com/picker' in frame.url:
            picker_frame = frame
            print(f"[INFO] Found Drive picker frame: {frame.url[:80]}...")
            break
            
    # Retry if it is still loading
    if not picker_frame:
        page.wait_for_timeout(2000)
        for frame in page.frames:
            if 'docs.google.com/picker' in frame.url:
                picker_frame = frame
                print(f"[INFO] Found Drive picker frame after retry: {frame.url[:80]}...")
                break

    file_input = None
    if picker_frame:
        try:
            file_input = picker_frame.wait_for_selector('input[type="file"]', state="attached", timeout=10000)
            if file_input:
                print("[INFO] [OK] Found file input inside the picker iframe")
        except Exception as e:
            print(f"[WARN] Failed to locate file input in iframe: {e}")

    # Fallback to main page if frame not detected
    if not file_input:
        try:
            file_input = page.wait_for_selector('input[type="file"]', state="attached", timeout=5000)
            if file_input:
                print("[INFO] Found file input in main page context")
        except Exception:
            pass

    if not file_input:
        raise RuntimeError(
            "[ERROR] Could not find the file upload input in any iframe or the main document."
        )

    # ── STEP 8d: Make target input file path selection ──
    print("[STEP 8d] Sending file path to input element...")
    try:
        if picker_frame:
            picker_frame.locator('input[type="file"]').set_input_files(abs_path)
        else:
            page.locator('input[type="file"]').set_input_files(abs_path)
        print(f"[INFO] [OK] File path sent to input: {abs_path}")
    except Exception as e:
        print(f"[WARN] Direct file set failed: {e}. Attempting styling override...")
        override_js = """
            var el = document.querySelector('input[type="file"]');
            if (el) {
                el.style.display    = 'block';
                el.style.visibility = 'visible';
                el.style.opacity    = '1';
                el.style.width      = '200px';
                el.style.height     = '20px';
            }
        """
        if picker_frame:
            picker_frame.evaluate(override_js)
            page.wait_for_timeout(500)
            picker_frame.locator('input[type="file"]').set_input_files(abs_path)
        else:
            page.evaluate(override_js)
            page.wait_for_timeout(500)
            page.locator('input[type="file"]').set_input_files(abs_path)
        print(f"[INFO] [OK] File path sent to input after style override: {abs_path}")

    # Wait for Google to finish uploading the file to Drive
    print("[INFO] Waiting for file upload to complete...")
    _wait_for_upload_complete(page)


def _wait_for_upload_complete(page, timeout=120):
    """
    Poll until the Google Forms upload spinner is gone and
    the uploaded filename appears, confirming the upload finished.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            # A successfully uploaded file shows a 'remove' or 'delete' button
            remove_btn = page.locator(
                'xpath=//*[@aria-label="Remove file" or @aria-label="Delete file" or contains(@class,"freebirdFormviewerViewItemsFileRemoveFile")]'
            )
            if remove_btn.count() > 0:
                print("[INFO] [OK] File upload confirmed (remove button appeared).")
                return
        except Exception:
            pass
        page.wait_for_timeout(2000)
    print("[WARN] Upload confirmation timed out — the file may still have uploaded. Proceeding.")


def fill_form():
    print("[STEP 4] Launching Chrome with copied profile via Playwright...")
    
    launch_args = [
        "--disable-blink-features=AutomationControlled",
        "--disable-infobars",
        "--disable-notifications",
        "--ignore-certificate-errors",
        "--window-size=1920,1080",
        f"--profile-directory={CHROME_PROFILE_NAME}"
    ]

    with sync_playwright() as p:
        try:
            print("[INFO] Starting Chrome browser context...")
            # Try launching using system Chrome channel to locate profiles correctly
            context = p.chromium.launch_persistent_context(
                user_data_dir=custom_profile_dir,
                headless=HEADLESS,
                channel="chrome",
                args=launch_args,
                no_viewport=True
            )
        except Exception as e:
            print(f"[WARN] Launching with local Chrome channel failed: {e}")
            print("[INFO] Retrying launch using default Playwright Chromium browser...")
            context = p.chromium.launch_persistent_context(
                user_data_dir=custom_profile_dir,
                headless=HEADLESS,
                args=launch_args,
                no_viewport=True
            )

        page = context.pages[0] if context.pages else context.new_page()
        page.goto(form_link)
        
        # Give Chrome a moment to settle page load
        page.wait_for_timeout(3000)

        try:
            # ── LOGIN GUARD: detect if Google logged us out ───────────────────────
            print("[STEP 4b] Checking login state...")
            page.evaluate("window.scrollTo(0, 0);")
            current_url = page.url
            if "accounts.google.com" in current_url or "signin" in current_url.lower():
                print("[WARNING] Chrome profile session expired or not logged in.")
                print("[INFO] Closing automated browser to free profile directory lock...")
                try:
                    context.close()
                except Exception:
                    pass
                
                # Automatically run the login helper inline
                launch_chrome_for_login()
                
                # Re-launch the persistent context
                print("[INFO] Re-opening automated Chrome browser context...")
                try:
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=custom_profile_dir,
                        headless=HEADLESS,
                        channel="chrome",
                        args=launch_args,
                        no_viewport=True
                    )
                except Exception as e:
                    print(f"[WARN] Launching with local Chrome channel failed: {e}")
                    print("[INFO] Retrying launch using default Playwright Chromium browser...")
                    context = p.chromium.launch_persistent_context(
                        user_data_dir=custom_profile_dir,
                        headless=HEADLESS,
                        args=launch_args,
                        no_viewport=True
                    )
                
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(form_link)
                page.wait_for_timeout(3000)
                
                # Recheck login state
                page.evaluate("window.scrollTo(0, 0);")
                current_url = page.url
                if "accounts.google.com" in current_url or "signin" in current_url.lower():
                    alert_msg = (
                        "[WARNING] *GuruAuto Login Alert*\n"
                        "Re-authentication failed — Google is still asking for sign-in.\n"
                        "Form was NOT submitted."
                    )
                    print(f"[ERROR] {alert_msg}")
                    if NOTIFY_PHONE:
                        send_whatsapp_notification(NOTIFY_PHONE, alert_msg)
                    sys.exit(1)
                print("[INFO] Re-authentication successful!")

            print("[INFO] Login state OK — Google account is active.")

            # ── STEP 5: Email checkbox (optional — only shown on first visit) ────
            print("[STEP 5] Checking for email consent checkbox...")
            try:
                email_checkbox = page.wait_for_selector('//div[@role="checkbox"]', timeout=5000)
                if email_checkbox:
                    email_checkbox.click()
                    print("[INFO] Email checkbox clicked.")
            except Exception:
                print("[INFO] No email checkbox found (already consented or not required). Continuing.")

            # ── STEP 5b: Click 'Next' if form split into multiple pages ──────────
            try:
                next_btn = page.wait_for_selector(
                    'xpath=//span[text()="Next"]/ancestor::div[@role="button"] | //div[@role="button"]//span[text()="Next"]',
                    timeout=5000
                )
                if next_btn:
                    next_btn.click()
                    print("[INFO] 'Next' button clicked — moving to form page 2.")
                    page.wait_for_timeout(2000)
            except Exception:
                print("[INFO] No 'Next' button — form is single-page, continuing.")

            # ── STEP 6: Fill text fields ──────────────────────────────────────────
            print("[STEP 6] Filling text input fields...")
            page.wait_for_function("document.querySelectorAll('input[type=\"text\"]').length >= 3", timeout=30000)

            # Date — fill using direct JS injection for locale independence
            iso_date = datetime.now().strftime("%Y-%m-%d")
            page.evaluate(f"""
                var el = document.querySelector('input[type="date"]');
                if (el) {{
                    el.value = "{iso_date}";
                    el.dispatchEvent(new Event('input', {{bubbles:true}}));
                    el.dispatchEvent(new Event('change', {{bubbles:true}}));
                }}
            """)
            print(f"[INFO] Date filled: {iso_date}")

            # Fill text inputs by nth index
            text_inputs = page.locator('//input[@type="text"]')
            text_inputs.nth(0).fill(person_data["mentee_name"])
            print(f"[INFO] Mentee name filled: {person_data['mentee_name']}")

            text_inputs.nth(1).fill(person_data["register_number"])
            print(f"[INFO] Register number filled: {person_data['register_number']}")

            text_inputs.nth(2).fill(person_data["mentor_name"])
            print(f"[INFO] Mentor name filled: {person_data['mentor_name']}")

            # ── STEP 7: Department dropdown ───────────────────────────────────────
            print("[STEP 7] Selecting department...")
            dropdown = page.wait_for_selector('//div[@role="listbox"]', timeout=30000)
            dropdown.click()
            page.wait_for_timeout(1000)

            # Retrieve option list items
            page.wait_for_selector('//div[@role="option"]', timeout=10000)
            options = page.locator('//div[@role="option"]')
            count = options.count()
            option_found = False
            for i in range(count):
                opt = options.nth(i)
                txt = opt.inner_text().strip()
                if txt == DEPARTMENT_NAME:
                    opt.click()
                    option_found = True
                    print(f"[INFO] Selected department: {DEPARTMENT_NAME}")
                    break

            if not option_found:
                raise Exception(f"Department '{DEPARTMENT_NAME}' not found in dropdown options")

            # ── STEP 8: Upload PPTX (if provided) ────────────────────────────────
            if PPTX_PATH:
                upload_pptx(page, PPTX_PATH)
            else:
                print("[SKIP] No PPTX path provided. Skipping file upload step.")

            # ── STEP 9: Submit ────────────────────────────────────────────────────
            if AUTO_SUBMIT:
                print("[STEP 9] Auto-submitting form...")
                page.wait_for_timeout(1000)

                submit_btn = page.wait_for_selector(
                    'xpath=//div[@role="button"]/span[text()="Submit"] | //span[text()="Submit"]/ancestor::div[@role="button"]',
                    timeout=15000
                )
                submit_btn.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
                submit_btn.click()
                print("[INFO] Submit button clicked.")

                # Wait for confirmation page
                page.wait_for_selector(
                    'xpath=//*[contains(text(),"Your response has been recorded") or contains(text(),"response has been recorded") or contains(text(),"Thanks")]',
                    timeout=30000
                )
                print("[SUCCESS] Form submitted successfully!")
                page.wait_for_timeout(2000)

                # Send WhatsApp confirmation
                if NOTIFY_PHONE:
                    notify_msg = (
                        f"[SUCCESS] *Gurupadigam Form Submitted!*\n"
                        f"Mentee : {person_data['mentee_name']}\n"
                        f"Reg No : {person_data['register_number']}\n"
                        f"Mentor : {person_data['mentor_name']}\n"
                        f"Dept   : {DEPARTMENT_NAME}\n"
                        f"Date   : {person_data['date']}"
                    )
                    send_whatsapp_notification(NOTIFY_PHONE, notify_msg)

                # Mark as processed in the database/JSON
                mark_processed(config.get('msg_id'), config.get('filename_day_key'))
            else:
                # Manual review mode
                print("\n[INFO] All fields filled. Please review and click 'Submit' manually.")
                input("Press Enter to close the browser...")

                # Mark as processed in the database/JSON
                mark_processed(config.get('msg_id'), config.get('filename_day_key'))

        except Exception as e:
            print(f"[ERROR] Something went wrong: {str(e)}")
            print("[DEBUG] URL:", page.url)
            try:
                print("[DEBUG] Page Source (partial):", page.content()[:500])
            except Exception:
                pass
            if AUTO_SUBMIT:
                time.sleep(10)   # Read error message
            else:
                input("Press Enter to close the browser after error...")
            sys.exit(1)
        finally:
            try:
                context.close()
            except Exception:
                pass
            print("[INFO] Browser closed.")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--login":
        launch_chrome_for_login()
        sys.exit(0)

    copy_chrome_profile()
    fill_form()