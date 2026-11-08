/**
 * Invoice collector - runs inside Google (script.google.com) on the account that receives invoices.
 * NOT YET TESTED against a live mailbox. Test with the steps in README.md before relying on it.
 *
 * collectInvoices()     Finds new emails with PDF/image attachments, saves them to the Drive
 *                       folder "Invoice Inbox", and labels the email so it is not collected twice.
 * emailLatestDigest()   Emails the newest monday_digest.txt (found anywhere in Drive) to DIGEST_TO, or to yourself.
 *
 * Triggers (Apps Script > Triggers > Add Trigger):
 *   collectInvoices      time-driven, every day, 6-7am
 *   emailLatestDigest    time-driven, weekly, Monday, 8-9am
 */
const INBOX_FOLDER  = 'Invoice Inbox';
const REVIEW_FOLDER = 'Weekly Review';
const DONE_LABEL    = 'Invoices/Collected';
// ONLY emails carrying this Gmail label are collected, so nothing else in the mailbox is touched.
// Create the label "Invoices-Inbox" and a filter that applies it (see README, "Test the email part").
const INBOX_LABEL = 'Invoices-Inbox';
const SEARCH = 'label:' + INBOX_LABEL + ' has:attachment -label:Invoices-Collected newer_than:30d';
// Who gets the Monday digest. Leave blank to email the account running this script.
const DIGEST_TO = '';
// Who the "Send for approval" buttons address. Leave blank and the button opens an email with no To address.
const APPROVER_TO = '';
const MIN_BYTES = 15000;   // skips tiny IMAGES such as email-signature logos (PDFs are never skipped)

function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function collectInvoices() {
  const inbox = folder_(INBOX_FOLDER);
  const label = GmailApp.getUserLabelByName(DONE_LABEL) || GmailApp.createLabel(DONE_LABEL);
  const threads = GmailApp.search(SEARCH, 0, 50);
  let saved = 0;
  threads.forEach(function (thread) {
    let handled = 0;
    thread.getMessages().forEach(function (msg) {
      msg.getAttachments({ includeInlineImages: false }).forEach(function (att) {
        const ctype = att.getContentType(), aname = att.getName();
        Logger.log('Saw attachment: ' + aname + ' (' + ctype + ', ' + att.getSize() + ' bytes)');
        const isPdf = /pdf/i.test(ctype) || /\.pdf$/i.test(aname);
        const isImg = /jpeg|png/i.test(ctype) || /\.(jpe?g|png)$/i.test(aname);
        if (!isPdf && !isImg) return;
        if (!isPdf && att.getSize() < MIN_BYTES) return;
        const stamp = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyyMMdd');
        const name = stamp + '_' + att.getName();
        handled++;
        if (inbox.getFilesByName(name).hasNext()) return;      // already saved
        inbox.createFile(att.copyBlob().setName(name));
        saved++;
      });
    });
    if (handled > 0) thread.addLabel(label);   // only mark emails we actually collected from
  });
  Logger.log('Saved ' + saved + ' attachments from ' + threads.length + ' emails.');
}

function emailLatestDigest() {
  // The pipeline writes monday_digest.txt into the "out" folder (inside Drive). Find the newest copy anywhere in Drive.
  const files = DriveApp.searchFiles('title = "monday_digest.txt" and trashed = false');
  let newest = null;
  while (files.hasNext()) {
    const f = files.next();
    if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
  }
  if (!newest) { Logger.log('No digest found yet.'); return; }
  const lines = newest.getBlob().getDataAsString().split('\n');
  const subject = lines[0].replace(/^Subject:\s*/, '');
  const body = lines.slice(1).join('\n').trim();
  const to = DIGEST_TO || Session.getActiveUser().getEmail();    // blank = email yourself
  // Send the table version if the pipeline wrote one next to the text digest; otherwise plain text.
  let html = null;
  const parents = newest.getParents();
  if (parents.hasNext()) {
    const h = parents.next().getFilesByName('monday_digest.html');
    if (h.hasNext()) html = h.next().getBlob().getDataAsString().split('href="mailto:?').join('href="mailto:' + APPROVER_TO + '?');
  }
  if (html) MailApp.sendEmail(to, subject, body, { htmlBody: html });
  else MailApp.sendEmail(to, subject, body);
  Logger.log('Sent digest from ' + newest.getName() + ' (updated ' + newest.getLastUpdated() + ') to ' + to);
}
