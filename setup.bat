@echo off
title GuruAuto - Setup Launcher
echo.
echo ============================================================
echo   Starting GuruAuto Dependency Setup...
echo ============================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
