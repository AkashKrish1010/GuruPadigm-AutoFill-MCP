import os
import shutil
import subprocess
import time
import sys
import json
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime

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

if len(sys.argv) > 1:
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
            print(f"[NOTIFY] ✅ WhatsApp notification sent to {phone}")
        else:
            print(f"[NOTIFY] ⚠️  Bridge responded: {resp.text}")
    except requests.RequestException as e:
        print(f"[NOTIFY] ⚠️  Could not reach WhatsApp bridge: {e}")
        print("[NOTIFY]    Is the Go bridge running? (go run main.go inside whatsapp-bridge/)")


def copy_chrome_profile():
    r"""
    One-time setup: ensures a dedicated Chrome profile directory exists
    for Selenium to use with persistent Google login.

    IMPORTANT - create a DEDICATED profile:
      1. Run:  chrome --user-data-dir="<your CHROME_PROFILE_DIR from config.py>"
      2. Log in to the Google account used for the form
      3. Close Chrome
      4. Never open that profile in regular Chrome (only via Selenium)
    This avoids the "logged out" problem caused by session conflicts.
    """
    print("[STEP 1] Checking Chrome profile...")
    if os.path.exists(custom_profile_dir):
        print("[INFO] Custom profile already exists. Skipping setup.")
        return
    print(f"[INFO] Creating Chrome profile directory: {custom_profile_dir}")
    os.makedirs(custom_profile_dir, exist_ok=True)
    print(f"[DONE] Profile directory created.")
    print(f"       Please sign in manually by running:")
    print(f'       chrome --user-data-dir="{custom_profile_dir}"')
    print(f"       Then log in to Google and close Chrome.\n")


def convert_pptx_to_pdf(pptx_path: str) -> str:
    """
    Convert a PPTX file to PDF using Microsoft PowerPoint COM (Windows).
    Falls back to LibreOffice if PowerPoint is not installed.
    Returns the absolute path to the generated PDF.
    """
    import os
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
        print(f"[CONVERT] ✅ PDF saved: {pdf_path}")
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
            print(f"[CONVERT] ✅ PDF saved via LibreOffice: {pdf_path}")
            return pdf_path
        else:
            raise RuntimeError(result.stderr)
    except Exception as e:
        raise RuntimeError(
            f"[CONVERT] Could not convert PPTX to PDF. "
            f"Install Microsoft Office or LibreOffice.\nError: {e}"
        )


def upload_pptx(browser, wait, pptx_path):
    """
    Upload a file to the Google Form file-upload field.
    Automatically converts PPTX to PDF before uploading.

    Strategy:
      1. Scroll to the file-upload section so it fully renders
      2. Click the form's "Add file" button → opens the Google Drive picker
      3. The picker lives inside an IFRAME — switch into it
      4. Find the hidden <input type="file"> near the "Browse" button
      5. send_keys the file path, switch back to main content
      6. Wait for Google to finish uploading to Drive
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
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1.5)  # Let the widget render

    # ── STEP 8b: Click the form's upload button to open the Drive picker ──
    UPLOAD_BUTTON_XPATH = '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[3]/span/span[2]'
    print("[STEP 8b] Clicking file upload button to open Drive picker...")
    try:
        upload_btn = wait.until(EC.element_to_be_clickable((By.XPATH, UPLOAD_BUTTON_XPATH)))
        browser.execute_script("arguments[0].scrollIntoView(true);", upload_btn)
        time.sleep(0.5)
        upload_btn.click()
        print("[INFO] Upload button clicked — Drive picker should open.")
    except Exception as e:
        print(f"[WARN] Could not click upload button via primary XPath: {e}")
        # Fallback: try text-based button search
        try:
            fallback_btn = browser.find_element(By.XPATH,
                '//span[contains(text(), "Add file") or contains(text(), "Choose file") or contains(text(), "Upload")]')
            browser.execute_script("arguments[0].scrollIntoView(true);", fallback_btn)
            fallback_btn.click()
            print("[INFO] Upload button clicked via fallback text XPath.")
        except Exception as e2:
            raise RuntimeError(f"[ERROR] Could not click any upload button: {e2}")

    # Wait for the Google Drive picker dialog to appear
    time.sleep(3)

    # ── STEP 8c: Switch into the Google Drive picker iframe ──────────────
    print("[STEP 8c] Looking for Google Drive picker iframe...")
    file_input = None
    picker_iframe_switched = False

    # The picker is rendered inside an iframe. Find and switch into it.
    iframes = browser.find_elements(By.TAG_NAME, "iframe")
    print(f"[DEBUG] Found {len(iframes)} iframe(s) on the page")

    for idx, iframe in enumerate(iframes):
        try:
            src = iframe.get_attribute("src") or ""
            iframe_id = iframe.get_attribute("id") or ""
            print(f"[DEBUG] iframe[{idx}]: id='{iframe_id}', src='{src[:80]}...'")

            # Google Drive picker iframes typically have 'docs.google.com/picker' in src
            browser.switch_to.frame(iframe)
            picker_iframe_switched = True
            print(f"[INFO] Switched into iframe[{idx}]")

            # Try to find input[type="file"] inside this iframe
            try:
                file_input = browser.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                if file_input:
                    print(f"[INFO] ✅ Found file input inside iframe[{idx}]")
                    break
            except:
                pass

            # Also try the uploadButtonId — Browse button's sibling/child input
            try:
                file_input = browser.execute_script("""
                    // Look near the Browse/upload button for a file input
                    var uploadBtn = document.getElementById('uploadButtonId');
                    if (uploadBtn) {
                        var parent = uploadBtn.parentElement;
                        while (parent) {
                            var inp = parent.querySelector('input[type="file"]');
                            if (inp) return inp;
                            parent = parent.parentElement;
                        }
                    }
                    // Fallback: any input[type="file"] anywhere in this frame
                    return document.querySelector('input[type="file"]');
                """)
                if file_input:
                    print(f"[INFO] ✅ Found file input near uploadButtonId in iframe[{idx}]")
                    break
            except:
                pass

            # Didn't find it here, switch back and try next iframe
            browser.switch_to.default_content()
            picker_iframe_switched = False

        except Exception as e:
            print(f"[DEBUG] Could not process iframe[{idx}]: {e}")
            try:
                browser.switch_to.default_content()
            except:
                pass
            picker_iframe_switched = False

    # ── If no iframe worked, try nested iframes (picker inside picker) ──
    if not file_input:
        print("[INFO] Trying nested iframes...")
        browser.switch_to.default_content()
        for idx, iframe in enumerate(iframes):
            try:
                browser.switch_to.frame(iframe)
                nested_iframes = browser.find_elements(By.TAG_NAME, "iframe")
                for nidx, nested in enumerate(nested_iframes):
                    try:
                        browser.switch_to.frame(nested)
                        file_input = browser.find_element(By.CSS_SELECTOR, 'input[type="file"]')
                        if file_input:
                            print(f"[INFO] ✅ Found file input in nested iframe[{idx}][{nidx}]")
                            picker_iframe_switched = True
                            break
                    except:
                        browser.switch_to.parent_frame()
                if file_input:
                    break
                browser.switch_to.default_content()
            except:
                try:
                    browser.switch_to.default_content()
                except:
                    pass

    # ── Last resort: check main document (in case picker is not in iframe) ──
    if not file_input:
        browser.switch_to.default_content()
        picker_iframe_switched = False
        try:
            file_input = browser.find_element(By.CSS_SELECTOR, 'input[type="file"]')
            if file_input:
                print("[INFO] Found file input in main document (no iframe needed)")
        except:
            pass

    if not file_input:
        # Switch back to main content before raising
        try:
            browser.switch_to.default_content()
        except:
            pass
        raise RuntimeError(
            "[ERROR] Could not find the file upload input in any iframe or the main document. "
            "Google may have changed the picker structure."
        )

    # ── STEP 8d: Make the hidden input interactable and send the file path ──
    print("[STEP 8d] Sending file path to input element...")
    browser.execute_script(
        "arguments[0].style.display    = 'block';"
        "arguments[0].style.visibility = 'visible';"
        "arguments[0].style.opacity    = '1';"
        "arguments[0].style.width      = '200px';"
        "arguments[0].style.height     = '20px';",
        file_input
    )
    time.sleep(0.5)
    file_input.send_keys(abs_path)
    print(f"[INFO] ✅ File path sent to input: {abs_path}")

    # Switch back to the main content
    browser.switch_to.default_content()
    print("[INFO] Switched back to main document.")

    # Wait for Google to finish uploading the file to Drive
    print("[INFO] Waiting for file upload to complete...")
    _wait_for_upload_complete(browser, wait)


def _upload_via_button_click(browser, wait, abs_path):
    """
    Click the 'Add file' button in Google Forms, then use keyboard to
    type the file path into the Windows file picker dialog.
    Requires pywinauto: pip install pywinauto
    """
    try:
        import pywinauto
        from pywinauto.keyboard import send_keys as pw_send_keys
    except ImportError:
        raise RuntimeError(
            "pywinauto is not installed. Run: pip install pywinauto\n"
            "This is needed to control the Windows file picker dialog."
        )

    # Click the 'Choose files from your device' / 'Add file' button
    upload_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH,
         '//span[contains(text(),"Add file") or contains(text(),"Choose files")'
         ' or contains(text(),"Browse")]'
         '/ancestor::div[@role="button"]')
    ))
    upload_button.click()
    print("[INFO] Clicked upload button. Waiting for file dialog...")
    time.sleep(2)  # Let the OS file picker open

    # Type the path into the Windows Open File dialog
    pw_send_keys(abs_path, with_spaces=True)
    time.sleep(0.5)
    pw_send_keys("{ENTER}")
    time.sleep(1)

    print("[INFO] File path entered into OS dialog.")


def _wait_for_upload_complete(browser, wait, timeout=120):
    """
    Poll until the Google Forms upload spinner is gone and
    the uploaded filename appears, confirming the upload finished.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            # A successfully uploaded file shows a 'remove' or 'delete' button
            remove_btn = browser.find_elements(
                By.XPATH,
                '//*[@aria-label="Remove file" or @aria-label="Delete file"'
                ' or contains(@class,"freebirdFormviewerViewItemsFileRemoveFile")]'
            )
            if remove_btn:
                print("[INFO] ✅ File upload confirmed (remove button appeared).")
                return
        except Exception:
            pass
        time.sleep(2)
    print("[WARN] Upload confirmation timed out — the file may still have uploaded. Proceeding.")


def fill_form():
    print("[STEP 4] Launching Chrome with copied profile...")
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={custom_profile_dir}")
    options.add_argument(f"--profile-directory={CHROME_PROFILE_NAME}")
    # NOTE: --disable-extensions removed — Google treats it as suspicious
    # and kills the authenticated session. Use stealth flags instead:
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    # Suppress the "Chrome is being controlled by automated software" bar
    options.add_argument("--disable-notifications")

    # Always set window size — when launched via subprocess (watcher),
    # Chrome can get an undefined/tiny window, breaking form rendering.
    options.add_argument("--window-size=1920,1080")

    if HEADLESS:
        options.add_argument("--headless=new")   # Modern headless — supports file upload
        print("[INFO] Running in HEADLESS mode (no browser window).")
    else:
        print("[INFO] Running in VISIBLE mode (browser window will open).")

    print("[INFO] Starting Chrome browser...")
    browser = webdriver.Chrome(options=options)

    # Give Chrome a moment to settle its window handle before we interact with it.
    # When launched via start_automation.bat → watcher → subprocess, Chrome spawns in
    # a nested process context where the window isn't ready immediately.
    time.sleep(1)

    # Try to maximize only in visible mode — headless has no real window.
    # Silently ignore failure; --window-size=1920,1080 already ensures correct size.
    if not HEADLESS:
        try:
            browser.maximize_window()
        except Exception:
            pass  # Window size already set via --window-size option

    browser.get(form_link)
    wait = WebDriverWait(browser, 30)

    try:
        # ── LOGIN GUARD: detect if Google logged us out ───────────────────────
        # If the session expired, Google redirects to accounts.google.com.
        # We catch that early and alert via WhatsApp instead of silently failing.
        print("[STEP 4b] Checking login state...")
        time.sleep(3)  # Let page fully render + redirect happen if needed
        browser.execute_script("window.scrollTo(0, 0);")  # Start at top
        current_url = browser.current_url
        if "accounts.google.com" in current_url or "signin" in current_url.lower():
            alert_msg = (
                "⚠️ *GuruAuto Login Alert*\n"
                "Chrome profile session expired — Google is asking for sign-in.\n"
                "Please re-authenticate the Selenium profile:\n"
                "1. Open Chrome with: chrome --user-data-dir=C:\\SeleniumProfiles\\GuruProfile\n"
                "2. Sign in to Google\n"
                "3. Close Chrome\n"
                "Form was NOT submitted."
            )
            print(f"[ERROR] {alert_msg}")
            if NOTIFY_PHONE:
                send_whatsapp_notification(NOTIFY_PHONE, alert_msg)
            return   # Exit gracefully — don't crash
        print("[INFO] Login state OK — Google account is active.")

        # ── STEP 5: Email checkbox (optional — only shown on first visit) ────
        print("[STEP 5] Checking for email consent checkbox...")
        try:
            short_wait = WebDriverWait(browser, 5)
            email_checkbox = short_wait.until(
                EC.element_to_be_clickable((By.XPATH, '//div[@role="checkbox"]'))
            )
            email_checkbox.click()
            print("[INFO] Email checkbox clicked.")
        except Exception:
            print("[INFO] No email checkbox found (already consented or not required). Continuing.")

        # ── STEP 5b: Click 'Next' if form split into multiple pages ──────────
        # Google Forms inserts a page break when a file-upload field is added.
        try:
            short_wait = WebDriverWait(browser, 5)
            next_btn = short_wait.until(EC.element_to_be_clickable(
                (By.XPATH,
                 '//span[text()="Next"]/ancestor::div[@role="button"]'
                 ' | //div[@role="button"]//span[text()="Next"]')
            ))
            next_btn.click()
            print("[INFO] 'Next' button clicked — moving to form page 2.")
            time.sleep(2)  # Wait for next page to render
        except Exception:
            print("[INFO] No 'Next' button — form is single-page, continuing.")

        # ── STEP 6: Fill text fields ──────────────────────────────────────────
        # Re-fetch each element immediately before use to avoid stale refs.
        # Google Forms re-renders parts of the DOM after consent loads.
        print("[STEP 6] Filling text input fields...")

        # Wait until at least 3 text inputs are present
        wait.until(lambda d: len(d.find_elements(By.XPATH, '//input[@type="text"]')) >= 3)

        # Date — use JS to set value (locale-independent, stale-safe)
        date_el = browser.find_element(By.XPATH, '//input[@type="date"]')
        iso_date = datetime.now().strftime("%Y-%m-%d")
        browser.execute_script("arguments[0].value = arguments[1];", date_el, iso_date)
        # Trigger change event so Google Forms registers the value
        browser.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
            date_el
        )
        print(f"[INFO] Date filled: {iso_date}")

        # Mentee name — re-fetch fresh
        browser.find_elements(By.XPATH, '//input[@type="text"]')[0].send_keys(person_data["mentee_name"])
        print(f"[INFO] Mentee name filled: {person_data['mentee_name']}")

        # Register number — re-fetch fresh
        browser.find_elements(By.XPATH, '//input[@type="text"]')[1].send_keys(person_data["register_number"])
        print(f"[INFO] Register number filled: {person_data['register_number']}")

        # Mentor name — re-fetch fresh
        browser.find_elements(By.XPATH, '//input[@type="text"]')[2].send_keys(person_data["mentor_name"])
        print(f"[INFO] Mentor name filled: {person_data['mentor_name']}")

        # ── STEP 7: Department dropdown ───────────────────────────────────────
        print("[STEP 7] Selecting department...")
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@role="listbox"]')))
        dropdown.click()
        time.sleep(1)

        # Re-fetch options fresh after dropdown opens
        dropdown_options = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, '//div[@role="option"]')))

        option_found = False
        for option in dropdown_options:
            if option.text.strip() == DEPARTMENT_NAME:
                option.click()
                option_found = True
                print(f"[INFO] Selected department: {DEPARTMENT_NAME}")
                break

        if not option_found:
            raise Exception(f"Department '{DEPARTMENT_NAME}' not found in dropdown options")

        # ── STEP 8: Upload PPTX (if provided) ────────────────────────────────
        if PPTX_PATH:
            upload_pptx(browser, wait, PPTX_PATH)
        else:
            print("[SKIP] No PPTX path provided. Skipping file upload step.")

        # ── STEP 9: Submit ────────────────────────────────────────────────────
        if AUTO_SUBMIT:
            print("[STEP 9] Auto-submitting form...")
            time.sleep(1)  # Small buffer after upload

            # Google Forms submit button text can be "Submit" or "Next"
            submit_button = wait.until(EC.element_to_be_clickable(
                (By.XPATH,
                 '//div[@role="button"]/span[text()="Submit"]'
                 ' | //span[text()="Submit"]/ancestor::div[@role="button"]')
            ))
            browser.execute_script("arguments[0].scrollIntoView(true);", submit_button)
            time.sleep(0.5)
            submit_button.click()
            print("[INFO] Submit button clicked.")

            # Wait for the confirmation page
            wait.until(EC.presence_of_element_located(
                (By.XPATH,
                 '//*[contains(text(),"Your response has been recorded")'
                 ' or contains(text(),"response has been recorded")'
                 ' or contains(text(),"Thanks")]')
            ))
            print("[SUCCESS] ✅ Form submitted successfully!")
            time.sleep(2)

            # ── Notify via WhatsApp ───────────────────────────────────────────
            if NOTIFY_PHONE:
                notify_msg = (
                    f"✅ *Gurupadigam Form Submitted!*\n"
                    f"👤 Mentee : {person_data['mentee_name']}\n"
                    f"🔢 Reg No : {person_data['register_number']}\n"
                    f"👨‍🏫 Mentor : {person_data['mentor_name']}\n"
                    f"🏫 Dept   : {DEPARTMENT_NAME}\n"
                    f"📅 Date   : {person_data['date']}"
                )
                send_whatsapp_notification(NOTIFY_PHONE, notify_msg)

        else:
            # Manual mode — prompt user to review before submitting
            print("\n[INFO] All fields filled. Please review and click 'Submit' manually.")
            input("Press Enter to close the browser...")

    except Exception as e:
        print(f"[ERROR] Something went wrong: {str(e)}")
        print("[DEBUG] URL:", browser.current_url)
        print("[DEBUG] Page Source (partial):", browser.page_source[:500])
        if AUTO_SUBMIT:
            time.sleep(10)   # Give time to read the error before closing
        else:
            input("Press Enter to close the browser after error...")
    finally:
        browser.quit()
        print("[INFO] Browser closed.")


if __name__ == "__main__":
    copy_chrome_profile()
    fill_form()