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

:: ── Step 1: Build the Go bridge if it doesn't exist yet ──────────────────────
echo [1/3] Checking WhatsApp bridge binary...
set BRIDGE_DIR=%~dp0whatsapp-mcp\whatsapp-bridge
set BRIDGE_EXE=%BRIDGE_DIR%\whatsapp-bridge.exe

if not exist "%BRIDGE_EXE%" (
    echo [INFO] whatsapp-bridge.exe not found. Building from source...
    echo [INFO] This only happens once and may take 1-2 minutes...
    echo.

    :: Check that Go is installed
    where go >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Go is not installed or not in PATH.
        echo         Install Go from https://go.dev/dl/ and re-run.
        pause
        exit /b 1
    )

    pushd "%BRIDGE_DIR%"
    echo [BUILD] Running: go mod tidy ...
    go mod tidy
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] go mod tidy failed. Check errors above.
        popd
        pause
        exit /b 1
    )

    echo [BUILD] Running: go build -o whatsapp-bridge.exe . ...
    go build -o whatsapp-bridge.exe .
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Build failed. Check errors above.
        popd
        pause
        exit /b 1
    )
    popd

    echo [OK] whatsapp-bridge.exe built successfully!
    echo.
) else (
    echo [OK] whatsapp-bridge.exe found.
)
echo.

:: ── Step 2: Start Go bridge in a new terminal window ────────────────────────
echo [2/3] Starting WhatsApp bridge (Go)...
echo.
start "WhatsApp Bridge" cmd /k "cd /d %BRIDGE_DIR% && echo Starting WhatsApp Bridge... && whatsapp-bridge.exe"

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
