# Weekly invoice review

A folder of invoices goes in. A spreadsheet comes out with one row per file:
`file, vendor, invoice_number, due_date, amount, flag, reason`.
Anything over $5,000, a duplicate, or worth a second look gets flag = Y and a plain-English reason.
Nothing is ever paid by the script. A person reviews the Y rows.

## The easy way: the app window
Double-click **Start Invoice Review.bat**. A page opens in your browser (it runs only on your own computer):
1. **Browse** to the folder of invoices and click **Run**.
2. See what needs a person, with the reason and a suggested next step for each.
3. Click **Open the spreadsheet** or **Open the Monday email**.
4. Open **Rules** to switch rules on and off or add your own, then **Save rules**.
5. Click **Quit the app** when you are done.
The command line below does the same thing and is what a weekly scheduled job uses.

## Review and send to AP
After a run the app shows every invoice, flag first. Click a column heading to sort.
- **Colors and tags.** Each flagged row has a colored tag under the Y: Duplicate, Accounting error, Billed entity
  (a portfolio company is named but the invoice is billed to Waverock Holdings), Currency, Over limit, or Review
  (your own custom rules). The key above the table shows the count for each. Click a key item, or the
  "duplicates" and "overbilling" chips, to highlight those rows.
- **Open flags only** hides invoices that are clear or already cleared, so only open flags remain.
- **File names are links.** Click one to open the original invoice. "Open the invoice folder" opens the folder.
- **Clear** (it asks you to confirm) records that a person checked a flag. Cleared rows are saved in
  `out/approvals.csv` and `out/decisions.json`, and are remembered when you run the same folder again. Undo reverses it.
  Clearing does not change the spreadsheet or the Monday email, and it does not pay anything.
- **Send to AP** opens a Gmail draft to the AP address listing every invoice that was never flagged or that you cleared
  (earliest due first), and saves the same list as `out/ap_queue.csv` to attach. Invoices still flagged are left out.

## Getting the summary by email
- **In the app window:** after a run, type your email address and click **Email me this summary**. A ready-to-send draft opens in Gmail and you press Send. Attach the spreadsheet if you want the full list. The app never stores a password and never sends anything by itself.
- **Automatically every Monday:** `emailLatestDigest` in `collect_invoices.gs` finds the newest `monday_digest.txt` in Drive and emails it to you (or to `DIGEST_TO`). Add it as a weekly trigger in Apps Script. This also needs no password because it runs inside your own Google account.

## What you get after a run (in the `out` folder)
- `invoice_review.xlsx`
  - **Invoices**: the seven columns requested, Y rows shaded.
  - **Review Queue**: only the Y rows, worst first, a suggested next step, and a dropdown for the reviewer's decision.
  - **Summary**: counts, dollars, "double payments avoided", and a chart.
  - **Rules**: exactly how Y/N is decided.
- `invoice_review.csv`: same as the Invoices tab.
- `monday_digest.txt`: the email a reviewer would get.

## One-time setup on Windows
1. Install Python from python.org. On the first screen tick **"Add python.exe to PATH"**.
2. Open **Command Prompt** (Windows key, type `cmd`, Enter) and run:

       pip install pdfplumber pypdfium2 pillow openpyxl

3. Install Tesseract (the program that reads scans and photos). In the same window:

       winget install UB-Mannheim.TesseractOCR

   Close the window and open a new one afterward. (The script also looks in the default install folder, so it usually works even if PATH is not updated.)
4. Optional, for the AI extractor: `pip install anthropic`, then `setx ANTHROPIC_API_KEY "your-key"` and reopen the window.

## Run it
Put the script and the unzipped invoice folder somewhere simple, like `C:\invoices-demo`. In Command Prompt:

    cd C:\invoices-demo
    python invoice_pipeline.py packet --out out

Options:

    --threshold 10000          change the dollar limit
    --as-of 2026-11-08         date used for "overdue" checks
    --history history.json     weekly mode: remembers past runs
    --out "Weekly Review\{date}"   {date} becomes the run date

Check the rules still behave: `python test_rules.py` (31 checks, should say OK).

## Changing the rules (no code)
Open the app window (above), or run `python rules_editor.py`, and open the **Rules** section.
- Switch each rule on or off and change its setting (dollar limit, days past due, company names).
- Add your own rules: "if the invoice mentions these words, flag it, and tell the reviewer this."
- Click **Save rules**. The settings go into `rules.json`, which the script reads on every run.
- The **Run now** box runs the script on a folder so you can see the effect right away.
Every run also saves the settings it used (`rules_used.json`) and shows them on the Rules tab.
Leave `rules.json` out and the built-in defaults apply. Delete a rule by switching it off or removing it in the window.

## How it works
1. **Read.** PDFs are read directly. Scans and photos go through OCR (turning a picture of text into text). Emails give up their attachment.
2. **Extract.** Vendor, invoice number, dates, total and line items are pulled out. With an API key set, Claude does this. Without one, a local pattern-matching parser does it.
3. **Flag.** Plain Python rules decide Y or N, so every flag can be explained.
   - over the dollar threshold
   - duplicate (same vendor + invoice number, or same vendor + amount + date, including invoices from past runs)
   - line items do not add up to the total
   - billed in a foreign currency
   - names a portfolio company (Grasp, PlanIt, WorthPoint) but is billed to the holding company
   - a field could not be read
4. **Write.** The files above.

## How the invoices get into the folder
The script only looks at one folder. Getting invoices into it is a separate step:

    Vendors email ap@ address
       -> Google Apps Script (collect_invoices.gs, runs daily) saves PDF/photo attachments to Drive folder "Invoice Inbox"
       -> Google Drive for Desktop shows that folder on the PC as G:\My Drive\Invoice Inbox
       -> Windows Task Scheduler runs run_weekly.bat every Monday 7am (this script, weekly mode)
       -> results land in Drive folder "Weekly Review\<date>" (open the xlsx as a Google Sheet)
       -> Apps Script (emailLatestDigest, Monday 8am) emails the digest to the reviewer
       -> a person works the Review Queue tab. Nothing is paid automatically.

Invoices that arrive by other routes (a vendor portal, a scan from the copier, a text photo) get dropped into the same Drive folder by hand.

Create the Monday task with one command (edit the path):

    schtasks /create /tn "Invoice Review" /tr "\"C:\path\to\invoice-review\run_weekly.bat\"" /sc weekly /d MON /st 07:00

Use the real folder that holds run_weekly.bat (keep the quotes if the path has spaces). Try the .bat by double-clicking
it first and read run_log.txt next to it. Then run the task once by hand to prove the scheduler can start it:

    schtasks /run /tn "Invoice Review"

The PC has to be on and signed in at that time. A cloud runner (GitHub Actions or a small server) removes that limit.

## Test the email part (Gmail + Drive)
Use a mailbox you do not mind experimenting with. The script only touches emails that carry the label `Invoices-Inbox`.
1. In Drive, create folders **Invoice Inbox** and **Weekly Review**.
2. In Gmail, create a label named `Invoices-Inbox`. Then Settings > See all settings > Filters > Create a new filter,
   put `yourname+invoices@gmail.com` in the **To** box (your own address with `+invoices` added), choose "Apply the label: Invoices-Inbox".
3. Go to script.google.com > New project, paste in `collect_invoices.gs`, save, choose `collectInvoices` and click Run.
   Google asks for permission. Click Advanced > Go to project (unsafe). It is your own script.
4. From another email account, send `yourname+invoices@gmail.com` a message with 2 or 3 invoice PDFs attached.
5. Run `collectInvoices` again. The files should appear in **Invoice Inbox**, and the email gets the label `Invoices/Collected`.
6. Triggers (clock icon in Apps Script): `collectInvoices` daily, `emailLatestDigest` weekly on Monday.

## Status
- Script, tests, spreadsheet, digest, weekly history: built and tested on the 30 sample invoices.
- `collect_invoices.gs`, `run_weekly.bat` and the Task Scheduler step: written but not yet run against a real mailbox or PC.
- AI extractor: written, not yet run with a live API key.
