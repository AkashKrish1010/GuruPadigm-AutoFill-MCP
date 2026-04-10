# 🎓 GuruPadigm-MCP-AutoFill

> **Send a PPTX on WhatsApp. Walk away. The form is submitted in under 30 seconds.**

GuruPadigm-MCP-AutoFill is a Windows automation pipeline that watches for incoming **PPTX files on WhatsApp**, extracts student data, converts the file to PDF, and **auto-fills & submits a Google Form** — completely hands-free.

Built with **Python · Selenium · WhatsApp MCP · Go**

---

## ⚡ Quick Start (Recommended)

> **This is the only command most users need.**

1. Complete the [one-time setup](#️-one-time-setup) below.
2. Double-click **`start_automation.bat`** from the project root.

That's it. The script will:

| Step | What happens |
|------|-------------|
| 🔴 | Closes any open Chrome windows (frees the Selenium profile lock) |
| 🔨 | **Auto-builds** `whatsapp-bridge.exe` if it doesn't exist yet (first run only, ~1-2 min) |
| 🟡 | Starts the **Go WhatsApp bridge** in a new terminal window |
| ⏳ | Waits 10 seconds for the bridge to initialize |
| 🟢 | Starts the **Python PPTX watcher** in a new terminal window |

Once running, **send a PPTX from a whitelisted WhatsApp number** and the form fills itself.

```
📱 WhatsApp  ──►  🌉 Go Bridge  ──►  🐍 Python Watcher  ──►  🌐 Selenium  ──►  ✅ Form Submitted
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
guru_auto_form.py  (Selenium)
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
├── guru_auto_form.py              # Core Selenium form-filler
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
| **Python** | 3.10 or newer |
| **Go** | 1.21 or newer — [download](https://go.dev/dl/) |
| **Google Chrome** | Latest stable |
| **ChromeDriver** | Must match your Chrome version — [download](https://chromedriver.chromium.org/downloads) |
| **Microsoft Office** | PowerPoint (for PPTX → PDF). Falls back to LibreOffice if not installed. |
| **WhatsApp** | Must be logged in on your phone (bridge uses WhatsApp Web protocol) |

---

### Step 1 — Clone the repo

```bash
git clone https://github.com/AkashKrish1010/GuruPadigm-MCP-AutoFill.git
cd GuruPadigm-MCP-AutoFill
```

---

### Step 2 — Install Python dependencies

```bash
pip install selenium requests flask flask-cors comtypes
```

> **Optional** (Windows file-dialog fallback):
> ```bash
> pip install pywinauto
> ```

---

### Step 3 — Install Go (required for the WhatsApp bridge)

Download and install Go 1.21+ from **[go.dev/dl](https://go.dev/dl/)**.

> ✅ **That's all.** `start_automation.bat` will automatically build `whatsapp-bridge.exe` on the first run. You don't need to run any build commands manually.

> If you ever want to rebuild manually (e.g. after editing `main.go`):
> ```bash
> cd whatsapp-mcp/whatsapp-bridge
> go mod tidy
> go build -o whatsapp-bridge.exe .
> ```

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
CHROME_PROFILE_DIR = r"C:\SeleniumProfiles\GuruProfile"

# 6. Behaviour flags
AUTO_SUBMIT   = False    # True = submit automatically
HEADLESS      = False    # True = run Chrome without a visible window
POLL_INTERVAL = 300      # Seconds between DB polls
```

> **💡 Two important flags to know:**
>
> | Flag | Default | What it does |
> |---|---|---|
> | `HEADLESS = True` | `False` | Chrome runs invisibly in the background. Set to `False` to **watch the browser fill the form** — great for first-time setup or debugging. |
> | `AUTO_SUBMIT = True` | `False` | Form is submitted automatically. Set to `False` to **pause before submit** so you can review the filled form manually. |

---

### Step 5 — Create a dedicated Selenium Chrome profile

This profile stores your Google login so the form can be accessed without re-authenticating on every run.

1. Open a terminal and run (use the same path as `CHROME_PROFILE_DIR` in `config.py`):
   ```
   chrome --user-data-dir="C:\SeleniumProfiles\GuruProfile"
   ```
2. Sign in to the **Google account** used to submit the form.
3. Close Chrome.

> ⚠️ **Important:** Never open this profile manually again — Selenium needs exclusive access to it.

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
```bash
cd whatsapp-mcp/whatsapp-bridge
.\whatsapp-bridge.exe
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
- The Selenium Chrome profile (set via `CHROME_PROFILE_DIR` in `config.py`) stores your Google session cookies — **do not share this folder**.

---

## ❗ Common Issues

| Problem | Fix |
|---|---|
| `messages.db not found` | Start the Go bridge first — the DB is created after the first WhatsApp connection. |
| QR code doesn't appear | Make sure `whatsapp-bridge.exe` was built correctly (run `go build -o whatsapp-bridge.exe .` inside `whatsapp-mcp/whatsapp-bridge`). |
| `Chrome profile session expired` | Re-open Chrome with the Selenium profile, sign into Google, close it, retry. |
| `PPTX → PDF conversion failed` | Ensure Microsoft Office or LibreOffice is installed. |
| `ChromeDriver version mismatch` | Download ChromeDriver matching your installed Chrome version from [here](https://chromedriver.chromium.org/downloads). |
| `Bridge not reachable` | Make sure the Go bridge is running at `http://localhost:8080`. |
| Form submits to wrong department | Update `DEFAULT_DEPARTMENT` in `config.py` to exactly match the dropdown text. |
| PPTX triggers form twice | Check `.processed_wa_ids.json` — the dedup tracker prevents this; if corrupted, delete it. |

---

## 🗺️ Roadmap

- [ ] Parse mentee name & register number directly from PPTX slide content
- [ ] Schedule automated runs (cron-style trigger)
- [ ] Export submission history to Excel
- [ ] Support multiple Google Form templates
- [ ] Cloud deployment (headless VPS)
- [ ] Desktop app via Electron

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
