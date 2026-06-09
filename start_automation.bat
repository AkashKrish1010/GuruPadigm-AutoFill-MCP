@echo off
title GuruAuto - WhatsApp PPTX Automation Launcher

echo ============================================================
echo   GuruAuto - WhatsApp PPTX to Google Form Automation
echo ============================================================
echo.

:: ── Step 0: Close all Chrome instances to free the Selenium profile lock ──────
echo [0/3] Checking for open Chrome windows...
tasklist /FI "IMAGENAME eq chrome.exe" 2>nul | find /I "chrome.exe" >nul
if %ERRORLEVEL% EQU 0 (
    echo [WARN] Chrome is currently open. Closing all Chrome windows...
    taskkill /F /IM chrome.exe >nul 2>&1
    timeout /t 2 /nobreak >nul
    :: Confirm Chrome is gone
    tasklist /FI "IMAGENAME eq chrome.exe" 2>nul | find /I "chrome.exe" >nul
    if %ERRORLEVEL% EQU 0 (
        echo [ERROR] Could not close Chrome. Please close it manually and re-run.
        pause
        exit /b 1
    )
    echo [OK] Chrome closed successfully.
) else (
    echo [OK] No Chrome windows open. Good to go.
)
echo.

:: ── Step 1: Verify Go is installed (only if running from source) ───────────
echo [1/3] Checking Go installation...
set BRIDGE_DIR=%~dp0whatsapp-mcp\whatsapp-bridge

if exist "%BRIDGE_DIR%\whatsapp-bridge.exe" (
    echo [OK] Pre-compiled WhatsApp bridge binary found. Skipping Go runtime check.
) else (
    where go >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Go is not installed or not in PATH.
        echo         Please run setup.ps1 first, or install Go manually.
        pause
        exit /b 1
    )
    echo [OK] Go found.
)


:: ── Step 2: Start Go bridge in a new terminal window ────────────────────────
echo [2/3] Starting WhatsApp bridge...
echo.
if exist "%BRIDGE_DIR%\whatsapp-bridge.exe" (
    start "WhatsApp Bridge" cmd /k "cd /d %BRIDGE_DIR% & echo Starting WhatsApp Bridge from binary... & whatsapp-bridge.exe"
) else (
    start "WhatsApp Bridge" cmd /k "cd /d %BRIDGE_DIR% & echo Enabling CGO... & go env -w CGO_ENABLED=1 & echo Starting WhatsApp Bridge from source... & go run main.go"
)

:: Wait for the bridge to initialize before starting the watcher
echo [INFO] Waiting 10 seconds for bridge to initialize...
timeout /t 10 /nobreak >nul

:: ── Step 3: Start the Python watcher in a new terminal window ────────────────
echo [3/3] Starting WhatsApp PPTX watcher (Python)...
start "PPTX Watcher" cmd /k "cd /d %~dp0 && python whatsapp_watcher.py"

echo.
echo ============================================================
echo   Both services are now running in separate windows.
echo   Close those windows to stop the automation.
echo ============================================================
pause
