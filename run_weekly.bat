@echo off
REM Weekly runner for Windows Task Scheduler. Edit the three paths below once.
REM   INVOICES = the Drive folder as Google Drive for Desktop shows it on this PC
REM   OUTROOT  = where each week's results go (also inside Drive, so it syncs back up)
set INVOICES=G:\My Drive\Invoice Inbox
set OUTROOT=G:\My Drive\Weekly Review
set HISTORY=%~dp0history.json

cd /d "%~dp0"
python invoice_pipeline.py "%INVOICES%" --out "%OUTROOT%\{date}" --history "%HISTORY%" >> "%~dp0run_log.txt" 2>&1
