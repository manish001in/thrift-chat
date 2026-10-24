// Paste into the Google Sheet: Extensions → Apps Script. Deploy → New deployment → Web app,
// Execute as: Me, Who has access: Anyone. Put the web-app URL into ENDPOINT in index.html.
// One row per person; later saves from the same person update their row. New questions add columns.

const SHEET = "Responses";
const FIRST = ["updated", "event", "stage", "interested", "chat", "name", "instagram", "phone", "roles"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || typeof d.id !== "string" || d.id.length > 64) return out("bad");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);

    let head = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (!head.length) { head = ["id"].concat(FIRST); sh.appendRow(head); sh.setFrozenRows(1); }
    const keys = Object.keys(d).filter(k => /^[a-z_]{1,40}$/.test(k)).slice(0, 60);
    const missing = keys.filter(k => head.indexOf(k) < 0);
    if (missing.length) {
      sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
      head = head.concat(missing);
    }

    // A later save from the same person must not overwrite a newer one.
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    const at = ids.indexOf(d.id);
    if (at >= 0) {
      const prev = sh.getRange(at + 2, head.indexOf("updated") + 1).getValue();
      if (prev && new Date(prev).getTime() > new Date(d.updated).getTime()) return out("stale");
    }
    const row = head.map(h => clean(d[h]));
    if (at >= 0) sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out("ok");
  } catch (err) {
    return out("error");
  } finally {
    lock.releaseLock();
  }
}

// Stops typed answers being read as spreadsheet formulas.
function clean(v) {
  if (v === undefined || v === null) return "";
  const s = String(v).slice(0, 2000);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function out(s) { return ContentService.createTextOutput(s); }
