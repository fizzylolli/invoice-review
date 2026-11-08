@echo off
REM Double-click to open the Invoice Review app. No black window stays open.
cd /d "%~dp0"
where pythonw >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install it from python.org and tick "Add python.exe to PATH", then try again.
  pause
  exit /b
)
start "" pythonw rules_editor.py
