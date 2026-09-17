// ─────────────────────────────────────────────────────────────
//  RSVP collector for the roof-deck invite
//  Paste this into Apps Script, then Deploy → New deployment →
//  Web app → Execute as: Me → Who has access: Anyone.
//  Copy the /exec URL it gives you.
// ─────────────────────────────────────────────────────────────

var SHEET_NAME = 'RSVPs';

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['When', 'Name', 'Status', 'Extra guests', 'Bringing']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// The page reads the running tally from here. Counts only — no names leave the sheet.
function doGet() {
  var rows = sheet_().getDataRange().getValues().slice(1);
  var latest = {};                      // one reply per name; last one wins
  rows.forEach(function (r) {
    var key = String(r[1] || '').trim().toLowerCase();
    if (key) latest[key] = r;
  });

  var heads = 0, maybe = 0, no = 0;
  Object.keys(latest).forEach(function (k) {
    var r = latest[k];
    var status = String(r[2] || '').toLowerCase();
    var extra = Number(r[3]) || 0;
    if (status === 'yes') heads += 1 + extra;
    else if (status === 'maybe') maybe += 1;
    else no += 1;
  });

  return json_({ heads: heads, maybe: maybe, no: no, replies: Object.keys(latest).length });
}

// The page posts one RSVP here.
function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) {}
  var name = String(d.name || '').trim().slice(0, 80);
  if (name) {
    sheet_().appendRow([
      new Date(),
      name,
      String(d.status || 'yes'),
      Number(d.guests) || 0,
      String(d.bring || '').slice(0, 200)
    ]);
  }
  return doGet();
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
