# setup.ps1 ── GuruPadigm-MCP-AutoFill Dependency Installer
# Auto-installs Python, Go, MSYS2/GCC, configures PATH, installs Python packages, and compiles the Go WhatsApp Bridge.

# Self-elevate to Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[INFO] Re-running setup script as Administrator..." -ForegroundColor Yellow
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# Clear the screen
Clear-Host

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  GuruPadigm-MCP-AutoFill Dependency Installer" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "This script will install all required runtimes and libraries" -ForegroundColor Gray
Write-Host "and compile the WhatsApp bridge. Please wait...`n" -ForegroundColor Gray

# Check if winget is available
$hasWinget = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)

# ── 1. Check & Install Python 3.12 ────────────────────────────────────────
Write-Host "[1/6] Checking Python installation..." -ForegroundColor Blue
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    if ($hasWinget) {
        Write-Host "      Python not found. Installing Python 3.12 via winget..." -ForegroundColor Yellow
        winget install Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    } else {
        Write-Host "      [WARN] winget not found. Please install Python 3.12 manually from https://www.python.org/downloads/" -ForegroundColor Red
    }
} else {
    $pyVer = python --version
    Write-Host "      [OK] Python found: $pyVer" -ForegroundColor Green
}

# ── 2. Check & Install Go ──────────────────────────────────────────────────
Write-Host "`n[2/6] Checking Go installation..." -ForegroundColor Blue
if (-not (Get-Command go -ErrorAction SilentlyContinue)) {
    if ($hasWinget) {
        Write-Host "      Go not found. Installing GoLang via winget..." -ForegroundColor Yellow
        winget install GoLang.Go --silent --accept-package-agreements --accept-source-agreements
    } else {
        Write-Host "      [WARN] winget not found. Please install Go manually from https://go.dev/dl/" -ForegroundColor Red
    }
} else {
    $goVer = go version
    Write-Host "      [OK] Go found: $goVer" -ForegroundColor Green
}

# ── 3. Check & Install MSYS2 & GCC Compiler ─────────────────────────────────
Write-Host "`n[3/6] Checking MSYS2 & C compiler (GCC) installation..." -ForegroundColor Blue
$hasGcc = $null -ne (Get-Command gcc -ErrorAction SilentlyContinue)
$hasMsys = Test-Path "C:\msys64"

if (-not $hasMsys) {
    if ($hasWinget) {
        Write-Host "      MSYS2 not found. Installing MSYS2 via winget..." -ForegroundColor Yellow
        winget install MSYS2.MSYS2 --silent --accept-package-agreements --accept-source-agreements
        $hasMsys = $true
    } else {
        Write-Host "      [WARN] winget not found. Please download and install MSYS2 manually from https://www.msys2.org/" -ForegroundColor Red
    }
} else {
    Write-Host "      [OK] MSYS2 directory C:\msys64 exists." -ForegroundColor Green
}

if ($hasMsys -and -not $hasGcc) {
    Write-Host "      Installing GCC compiler inside MSYS2. This can take a minute..." -ForegroundColor Yellow
    if (Test-Path "C:\msys64\usr\bin\bash.exe") {
        # Run pacman to install mingw-w64-ucrt-x86_64-gcc
        Start-Process "C:\msys64\usr\bin\bash.exe" -ArgumentList "-lc", '"pacman -S --noconfirm mingw-w64-ucrt-x86_64-gcc"' -NoNewWindow -Wait
        Write-Host "      [OK] GCC installed successfully." -ForegroundColor Green
    } else {
        Write-Host "      [ERROR] MSYS2 bash not found at C:\msys64\usr\bin\bash.exe. Please install GCC manually." -ForegroundColor Red
    }
} elseif ($hasGcc) {
    Write-Host "      [OK] GCC compiler found in PATH." -ForegroundColor Green
}

# ── 4. Set Environment variables / PATH ──────────────────────────────────────
Write-Host "`n[4/6] Configuring system PATH environment variables..." -ForegroundColor Blue
$pathsToAdd = @("C:\msys64\ucrt64\bin", "C:\Program Files\Go\bin")
$currentMachinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
$pathUpdated = $false

foreach ($path in $pathsToAdd) {
    if (Test-Path $path) {
        if ($currentMachinePath -notlike "*$path*") {
            Write-Host "      Adding $path to System PATH..." -ForegroundColor Yellow
            $currentMachinePath = "$currentMachinePath;$path"
            $pathUpdated = $true
        }
        # Update current process PATH so we can compile in the next step immediately
        if ($env:Path -notlike "*$path*") {
            $env:Path = "$env:Path;$path"
        }
    }
}

if ($pathUpdated) {
    [Environment]::SetEnvironmentVariable("Path", $currentMachinePath, "Machine")
    Write-Host "      [OK] System environment variables updated successfully." -ForegroundColor Green
} else {
    Write-Host "      [OK] System PATH is already up to date." -ForegroundColor Green
}

# ── 5. Install Python Libraries & Playwright Browser ─────────────────────────
Write-Host "`n[5/6] Installing Python libraries and browser dependencies..." -ForegroundColor Blue
Write-Host "      Upgrading pip..." -ForegroundColor Gray
python -m pip install --upgrade pip --quiet

Write-Host "      Installing requirements (playwright, requests, flask, comtypes)..." -ForegroundColor Gray
python -m pip install playwright requests flask flask-cors comtypes --quiet

Write-Host "      Installing Playwright Chromium browser binaries..." -ForegroundColor Gray
python -m playwright install chromium

Write-Host "      [OK] Python dependencies installed successfully." -ForegroundColor Green

# ── 6. Compile Go WhatsApp Bridge ─────────────────────────────────────────────
Write-Host "`n[6/6] Compiling Go WhatsApp bridge binary..." -ForegroundColor Blue
$bridgeDir = Join-Path $PSScriptRoot "whatsapp-mcp\whatsapp-bridge"

if (Test-Path $bridgeDir) {
    Push-Location $bridgeDir
    Write-Host "      Entering $bridgeDir" -ForegroundColor Gray
    
    # Configure CGO environment
    $env:CGO_ENABLED = "1"
    go env -w CGO_ENABLED=1
    
    Write-Host "      Running go mod tidy to fetch dependencies..." -ForegroundColor Gray
    go mod tidy
    
    Write-Host "      Compiling main.go into whatsapp-bridge.exe..." -ForegroundColor Yellow
    go build -o whatsapp-bridge.exe main.go
    
    if (Test-Path "whatsapp-bridge.exe") {
        Write-Host "      [OK] Pre-compiled executable built successfully at whatsapp-bridge.exe" -ForegroundColor Green
    } else {
        Write-Host "      [ERROR] Compilation failed. Could not build whatsapp-bridge.exe" -ForegroundColor Red
    }
    Pop-Location
} else {
    Write-Host "      [ERROR] WhatsApp bridge directory not found at $bridgeDir." -ForegroundColor Red
}

# ── Final Status ─────────────────────────────────────────────────────────────
Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  SETUP COMPLETE! All dependencies have been installed." -ForegroundColor Green
Write-Host "  You can now start the automation by double-clicking:" -ForegroundColor Green
Write-Host "  start_automation.bat" -ForegroundColor Green
Write-Host "============================================================`n" -ForegroundColor Green

Read-Host "Press Enter to exit this script..."
