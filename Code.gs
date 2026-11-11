/**
 * BTM 4.0 registration receiver.
 * SETUP (5 steps):
 * 1. Create a Google Sheet, then open Extensions > Apps Script and paste this file over Code.gs.
 * 2. Edit FEST_NAME and REPLY_TO below, then click Save.
 * 3. Click Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 * 4. Approve the permissions (Sheets + sending email), then copy the Web app URL.
 * 5. Paste that URL into CONFIG.endpoint in index.html. After any later edit to this file, use Deploy > Manage deployments > Edit > New version.
 */
const FEST_NAME = 'Beyond the Metrics 4.0';
const REPLY_TO  = 'btm@example.com';
const SHEET     = 'Registrations';
const HEAD = ['Timestamp','Reference ID','Event','Event ID','Team name','Name','Email','Phone',
              'University','Student ID','Department','Extra members','Payment method','Transaction ID','Fee'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // stops two simultaneous submits from clashing
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
    if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); }

    // Duplicate check: same email + same event (columns G and D)
    const email = String(d.email).trim().toLowerCase();
    const dup = sh.getDataRange().getValues().slice(1)
      .some(r => String(r[6]).toLowerCase() === email && r[3] === d.event);
    if (dup) return out('duplicate');

    const members = (d.members || []).map(m => m.name + (m.email ? ' <' + m.email + '>' : '')).join('; ');
    // Leading apostrophe keeps phone/ID values as text (no lost zeros)
    sh.appendRow([new Date(), d.ref, d.eventName, d.event, d.team, d.name, d.email, "'" + d.phone,
                  d.university, "'" + d.sid, d.dept, members, d.method, d.txn, d.fee]);

    MailApp.sendEmail({
      to: d.email,
      replyTo: REPLY_TO,
      subject: 'Registration received: ' + d.eventName + ' | ' + FEST_NAME,
      body: 'Hi ' + d.name + ',\n\nYour registration for ' + d.eventName + ' is received.\n' +
            'Reference ID: ' + d.ref + '\n' +
            (d.fee > 0 ? 'Payment: ' + d.method + ', transaction ' + d.txn + ' (we will verify it).\n' : '') +
            '\nQuestions? Reply to this email.\n\n' + FEST_NAME
    });
    return out('ok');
  } catch (err) {
    return out('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

function out(msg) { return ContentService.createTextOutput(msg); }
