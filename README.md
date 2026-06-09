# 🎓 GuruPadigm-MCP-AutoFill

> **Send a PPTX on WhatsApp. Walk away. The form is submitted in under 30 seconds.**

GuruPadigm-MCP-AutoFill is a Windows automation pipeline that watches for incoming **PPTX files on WhatsApp**, extracts student data, converts the file to PDF, and **auto-fills & submits a Google Form** — completely hands-free.

Built with **Python · Playwright · WhatsApp MCP · Go**

---

## ⚡ Quick Start (Recommended)

> **This is the only command most users need.**

1. Complete the [one-time setup](#️-one-time-setup) below.
2. Double-click **`start_automation.bat`** from the project root.

That's it. The script will:

| Step | What happens |
|------|-------------|
| 🔴 | Closes any open Chrome windows (frees the Chrome profile lock) |
| 🟡 | Verifies Go is installed and accessible in PATH |
| ⚙️ | Enables **CGO** (`go env -w CGO_ENABLED=1`) — requires MSYS2 on PATH |
| 🟡 | Starts the **Go WhatsApp bridge** via `go run main.go` in a new terminal window |
| ⏳ | Waits 10 seconds for the bridge to initialize |
| 🟢 | Starts the **Python PPTX watcher** in a new terminal window |

Once running, **send a PPTX from a whitelisted WhatsApp number** and the form fills itself.

```
📱 WhatsApp  ──►  🌉 Go Bridge  ──►  🐍 Python Watcher  ──►  🌐 Playwright  ──►  ✅ Form Submitted
```

---

## ✨ How It Works

```
Mentor sends PPTX on WhatsApp
        │
        ▼
Go Bridge  (whatsapp-mcp/whatsapp-bridge)
  ↓ detects new PPTX document in WhatsApp
  ↓ POSTs instantly to localhost:9999
        │
        ▼
whatsapp_watcher.py
  ↓ validates sender is in ALLOWED_SENDERS whitelist
  ↓ downloads PPTX via bridge REST API
  ↓ writes temp_config.json
        │
        ▼
guru_auto_form.py  (Playwright)
  ↓ launches Chrome with a dedicated profile
  ↓ converts PPTX → PDF  (via MS Office COM / LibreOffice)
  ↓ fills all Google Form fields
  ↓ uploads PDF attachment
  ↓ submits the form
        │
        ▼
WhatsApp confirmation message sent back ✅
```

> **Fallback:** If the instant push is missed, the watcher also polls the SQLite database every 5 minutes as a safety net.

---

## 📁 Project Structure

```
guruauto/
├── 🚀 start_automation.bat        ← START HERE — launches everything
│
├── ⚙️ config.py                    ← EDIT THIS — all your settings in one place
├── guru_auto_form.py              # Core Playwright form-filler
├── whatsapp_watcher.py            # WhatsApp PPTX watcher + event server
│
├── whatsapp-mcp/
│   └── whatsapp-bridge/           # Go WebSocket bridge for WhatsApp Web
│       └── store/
│           └── messages.db        # SQLite message store (auto-created on first run)
│
├── temp_config.json               # Auto-generated per-run config (git-ignored)
└── .processed_wa_ids.json         # Dedup tracker — prevents double submissions (git-ignored)
```

---

## 🛠️ One-Time Setup

### Prerequisites

| Requirement | Version / Notes |
|---|---|
| **Windows 10 / 11** | Required (uses COM, `.bat` scripts, `taskkill`) |
| **Python** | 3.10 or newer — [download](https://www.python.org/downloads/) |
| **Go** | 1.21 or newer — **[download installer from go.dev/dl](https://go.dev/dl/)** |
| **MSYS2 (C compiler)** | Required for CGO (whatsapp-mcp bridge). Install from [msys2.org](https://www.msys2.org/), then add `ucrt64\bin` to your PATH. [Step-by-step guide ↗](https://www.msys2.org/docs/environments/) |
| **Google Chrome** | Latest stable |
| **Microsoft Office** | PowerPoint (for PPTX → PDF). Falls back to LibreOffice if not installed. |
| **WhatsApp** | Must be logged in on your phone (bridge uses WhatsApp Web protocol) |

---

### Step 1 — Clone the repo

```bash
git clone https://github.com/AkashKrish1010/GuruPadigm-MCP-AutoFill.git
cd GuruPadigm-MCP-AutoFill
```

---

### Step 2 — Install Python dependencies & Playwright

```bash
pip install playwright requests flask flask-cors comtypes
playwright install chromium
```

---

### Step 3 — Install Go (required for the WhatsApp bridge)

Download and install Go 1.21+ from **[go.dev/dl](https://go.dev/dl/)**.

---

### Step 3b — Install a C compiler (MSYS2) for the WhatsApp bridge

The whatsapp-mcp bridge requires **CGO**, which needs a C compiler on Windows.

We recommend using **MSYS2**:

1. Download and install MSYS2 from **[msys2.org](https://www.msys2.org/)**.
2. Open the **MSYS2 UCRT64** shell (search "UCRT64" in Start Menu — **not** the regular MSYS2 shell) and run:
   ```bash
   pacman -S mingw-w64-ucrt-x86_64-gcc
   ```
3. Add `C:\msys64\ucrt64\bin` to your **Windows PATH** environment variable.
4. **Open a brand-new `cmd` or PowerShell window** (existing terminals won't see the PATH change).
5. Verify gcc is found — run this in the new terminal:
   ```bat
   gcc --version
   ```
   You should see output like `gcc (Rev1, Built by MSYS2 project) 13.x.x`. If you get `not found`, the PATH is not set correctly — go back to step 3.
6. Now enable CGO and run the bridge (run these **in order**, in the same new terminal):
   ```bat
   cd whatsapp-mcp\whatsapp-bridge
   go env -w CGO_ENABLED=1
   go run main.go
   ```

> ⚠️ **`go env -w CGO_ENABLED=1` alone is not enough.** Go silently falls back to `CGO_ENABLED=0` if it cannot find `gcc` in PATH, even if the env var is set.

> 📖 A full step-by-step guide is available at the [MSYS2 environments docs](https://www.msys2.org/docs/environments/).

---

### Step 4 — Configure your settings  ⭐

> **This is the most important step.** Open **`config.py`** and fill in your details.  
> This is the **only file you need to edit.**

```python
# ── config.py ──────────────────────────────────────────────

# 1. Allowed WhatsApp senders (country code + number, no '+')
ALLOWED_SENDERS = [
    "919876543210",        # Example: +91 98765 43210
]

# 2. Google Form link
FORM_LINK = "https://forms.gle/YOUR_FORM_LINK"

# 3. Form field defaults
DEFAULT_MENTEE_NAME     = "Your Name"
DEFAULT_REGISTER_NUMBER = "XXXXXXXXX"
DEFAULT_MENTOR_NAME     = "Dr. Mentor Name"
DEFAULT_DEPARTMENT      = "Your Department Name"

# 4. WhatsApp notification number (leave "" to disable)
NOTIFY_PHONE = "919876543210"

# 5. Chrome profile path
CHROME_PROFILE_DIR = r"C:\PlaywrightProfiles\GuruProfile"

# 6. Behaviour flags
AUTO_SUBMIT   = True     # True = submit automatically
HEADLESS      = True     # True = run Chrome without a visible window
POLL_INTERVAL = 300      # Seconds between DB polls
```

> **💡 Important behavior flags:**
>
> | Flag | Default | What it does |
> |---|---|---|
> | `HEADLESS = True` | `True` | Chrome runs invisibly in the background. Set to `False` to **watch the browser fill the form** — great for first-time setup or debugging. |
> | `AUTO_SUBMIT = True` | `True` | Form is submitted automatically. Set to `False` to **pause before submit** (overriding headed mode) so you can review the filled form manually. |

---

### Step 5 — Google Account Login Setup & Session Recovery

The automation uses a dedicated Chrome profile directory (`C:\PlaywrightProfiles\GuruProfile`) to persist your Google Account session.

#### Initial Setup
The first time you double-click `start_automation.bat` (or run `python whatsapp_watcher.py`), the script will detect if the profile is missing and **automatically launch a login setup window**.

Alternatively, you can run the login setup manually at any time:
```bash
python guru_auto_form.py --login
```
1. A real, non-automated Google Chrome window will open.
2. Sign in to the Google Account used to submit the Google Form.
3. Close the Chrome browser window and press **Enter** in the console.

#### Automatic Runtime Recovery
If your Google login session expires or gets logged out mid-run, the script will gracefully:
1. Temporarily pause automation and close the automated browser.
2. Open a real Chrome window prompting you to sign back in.
3. Once you sign in and press Enter in the terminal, it will automatically resume the form execution and submit it.

---

### Step 6 — First run & QR code scan

1. Double-click **`start_automation.bat`**.
2. In the **WhatsApp Bridge** window that opens, a **QR code** will appear.
3. On your phone, open WhatsApp → **Linked Devices → Link a Device** → scan the QR code.
4. The bridge will say `Connected`. The watcher window will start polling.

> You only need to scan the QR code once. The session is saved in `whatsapp-mcp/whatsapp-bridge/store/`.

---

## 🚀 Running the Automation

### ✅ Option A — One-click (recommended)

Double-click **`start_automation.bat`**.

Two terminal windows will open — leave them running. Close them to stop the automation.

---

### Option B — Manual (two terminals)

**Terminal 1 — Go bridge:**
```bat
cd whatsapp-mcp\whatsapp-bridge
go env -w CGO_ENABLED=1
go run main.go
```
Scan the QR code on first run.

**Terminal 2 — Python watcher:**
```bash
python whatsapp_watcher.py
```

---

## 🧪 Testing Without WhatsApp

Run `guru_auto_form.py` directly to test form-filling with the defaults from `config.py`:

```bash
python guru_auto_form.py
```

The script runs with the `HEADLESS` and `AUTO_SUBMIT` values from `config.py` so you can watch it fill the form and manually review before clicking Submit.

**To pass a custom config:**
```bash
python guru_auto_form.py path/to/your_config.json
```

**Config JSON schema:**
```json
{
  "form_link":       "https://forms.gle/...",
  "department":      "Your Department Name",
  "mentee_name":     "Student Name",
  "register_number": "XXXXXXXXX",
  "mentor_name":     "Dr. Mentor",
  "pptx_path":       "C:\\path\\to\\file.pptx",
  "auto_submit":     true,
  "headless":        true,
  "notify_phone":    "91XXXXXXXXXX"
}
```

---

## 🔒 Security Notes

- **`config.py`** contains YOUR personal settings — **review it before committing.** The repo ships with placeholder values only.
- **`temp_config.json`** and **`.processed_wa_ids.json`** are in `.gitignore` — they contain personal data and must never be committed.
- The `ALLOWED_SENDERS` whitelist ensures only trusted contacts can trigger form submissions.
- The Playwright Chrome profile (set via `CHROME_PROFILE_DIR` in `config.py`) stores your Google session cookies — **do not share this folder**.

---

## ❗ Common Issues

| Problem | Fix |
|---|---|
| `messages.db not found` | Start the Go bridge first — the DB is created after the first WhatsApp connection. |
| `Binary was compiled with 'CGO_ENABLED=0', go-sqlite3 requires cgo` | `gcc` is not in your PATH. Install MSYS2 UCRT64 GCC, add `C:\msys64\ucrt64\bin` to PATH, open a **new terminal**, verify with `gcc --version`, then run `go env -w CGO_ENABLED=1` and `go run main.go`. |
| QR code doesn't appear | Make sure Go and gcc are installed correctly and `go run main.go` started without errors. |
| `Chrome profile session expired` | Run `python guru_auto_form.py --login` to re-login, or wait for the automatic runtime setup prompt. |
| `PPTX → PDF conversion failed` | Ensure Microsoft Office or LibreOffice is installed. |
| `Bridge not reachable` | Make sure the Go bridge is running at `http://localhost:8080`. |
| Form submits to wrong department | Update `DEFAULT_DEPARTMENT` in `config.py` to exactly match the dropdown text. |
| PPTX triggers form twice | Check `.processed_wa_ids.json` — the dedup tracker prevents this; if corrupted, delete it. |

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first.

1. Fork the repository
2. Create your branch: `git checkout -b feature/my-feature`
3. Commit changes: `git commit -m 'Add my feature'`
4. Push: `git push origin feature/my-feature`
5. Open a Pull Request

---

## 📄 License

This project is open-source under the [MIT License](LICENSE).

---

## 👤 Author

**Akash S** — built to automate the weekly Gurupadigam mentoring form at college, cutting submission time from ~10 minutes to under 30 seconds (30× speedup).

> *"Automate the boring stuff so you can focus on what matters."*

---

## 🙏 Acknowledgments

- **[whatsapp-mcp](https://github.com/lharries/whatsapp-mcp)** by [lharries](https://github.com/lharries) — the Go-based WhatsApp Web bridge that powers the messaging layer of this project.
