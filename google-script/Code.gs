/**
 * Winter with the Segal's in the Blue Ridge — "Who's coming when" storage.
 *
 * How to install (once):
 *   1. In your Google Sheet: Extensions → Apps Script. Replace everything with this file. Save.
 *   2. Choose the function "setup" at the top and press Run. Allow the permissions.
 *   3. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone. Deploy.
 *   4. Copy the Web app URL (ends in /exec) and give it to the website.
 *
 * The sheet gets two tabs:
 *   Households — one row per household, with its private code and personal link.
 *                Edit names / usual sizes here. Never share the codes publicly.
 *   Answers    — one row per household per date (written by the website).
 *
 * Nobody can read or change anything without a valid household code.
 * A household can change only its own answers; an admin household can change everyone's.
 */

const SITE_URL = 'https://mvaiser-cyber.github.io/winterride/';

const HOUSEHOLDS = [
  // id,          display name,      usual size, admin?
  ['segal',       'Segal',            4, true],
  ['vaiser',      'Vaiser',           2, false],
  ['vaisiboni',   'Vaisiboni',        3, false],
  ['casaverde',   'Casaverde',        2, false],
  ['batisr',      'Batikoff Sr',      2, false],
  ['across',      'Across the road',  4, false],
  ['batijr',      'Bati Jr',          4, false],
];

const DATES = [
  '2026-12-22', '2026-12-23', '2026-12-24', '2026-12-25', '2026-12-26', '2026-12-27', '2026-12-28',
  '2026-12-29', '2026-12-30', '2026-12-31', '2027-01-01', '2027-01-02', '2027-01-03',
];

/** Run once from the Apps Script editor. Safe to run again: it never overwrites existing data. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let h = ss.getSheetByName('Households');
  if (!h) h = ss.insertSheet('Households');
  if (h.getLastRow() === 0) {
    h.getRange('A:F').setNumberFormat('@'); // keep codes as plain text
    h.getRange('C:C').setNumberFormat('0');
    h.appendRow(['id', 'name', 'usual_size', 'admin', 'code', 'personal_link']);
    HOUSEHOLDS.forEach(([id, name, size, admin]) => {
      const code = newCode_();
      h.appendRow([id, name, size, admin ? 'yes' : 'no', code, SITE_URL + '?h=' + code]);
    });
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, 6).setFontWeight('bold');
    h.autoResizeColumns(1, 6);
  }

  let a = ss.getSheetByName('Answers');
  if (!a) a = ss.insertSheet('Answers');
  if (a.getLastRow() === 0) {
    a.getRange('A:D').setNumberFormat('@'); // keep dates as plain text
    a.appendRow(['household_id', 'date', 'answer', 'people', 'updated_at', 'updated_by']);
    a.setFrozenRows(1);
    a.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

/** Website reads the plan: GET ?h=CODE */
function doGet(e) {
  const hs = households_();
  const me = findByCode_(hs, e && e.parameter && e.parameter.h);
  if (!me) return json_({ ok: false, error: 'unknown_link' });
  return json_(state_(me, hs));
}

/** Website saves answers: POST {h: CODE, changes: [{hid, date, value}]}; value is '', 'out' or a number */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hs = households_();
    const me = findByCode_(hs, body.h);
    if (!me) return json_({ ok: false, error: 'unknown_link' });

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Answers');
    const data = sh.getDataRange().getValues();
    const index = {};
    for (let i = 1; i < data.length; i++) index[String(data[i][0]) + '|' + dateStr_(data[i][1])] = i + 1;

    const changes = Array.isArray(body.changes) ? body.changes.slice(0, 200) : [];
    let refused = 0;
    changes.forEach(c => {
      const target = hs.find(x => x.id === c.hid);
      if (!target || (target.id !== me.id && !me.admin) || DATES.indexOf(c.date) < 0) { refused++; return; }
      let answer = '', people = '';
      if (c.value === 'out') answer = 'out';
      else if (c.value !== '' && c.value != null) {
        const n = Math.round(Number(c.value));
        if (!(n >= 1 && n <= 30)) { refused++; return; }
        answer = 'in'; people = String(n);
      }
      const row = [target.id, c.date, answer, people, new Date(), me.name];
      const key = target.id + '|' + c.date;
      if (index[key]) sh.getRange(index[key], 1, 1, 6).setValues([row]);
      else { sh.appendRow(row); index[key] = sh.getLastRow(); }
    });

    const out = state_(me, hs);
    out.refused = refused;
    return json_(out);
  } finally {
    lock.releaseLock();
  }
}

function households_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Households');
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({
      id: String(r[0]).trim(),
      name: String(r[1]).trim(),
      size: Math.max(1, Number(r[2]) || 1),
      admin: String(r[3]).trim().toLowerCase() === 'yes',
      code: String(r[4]).trim(),
    }));
}

function state_(me, hs) {
  const resp = {};
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Answers').getDataRange().getValues().slice(1);
  rows.forEach(r => {
    const hid = String(r[0]), d = dateStr_(r[1]), ans = String(r[2]);
    if (!hid || DATES.indexOf(d) < 0) return;
    if (ans === 'in') (resp[hid] = resp[hid] || {})[d] = { in: true, n: Math.max(1, Number(r[3]) || 1) };
    else if (ans === 'out') (resp[hid] = resp[hid] || {})[d] = { in: false };
  });
  return {
    ok: true,
    me: me.id,
    households: hs.map(x => ({ id: x.id, name: x.name, size: x.size, admin: x.admin })),
    resp: resp,
  };
}

function findByCode_(hs, code) {
  code = String(code || '').trim();
  if (code.length < 8) return null;
  return hs.find(x => x.code === code) || null;
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

function newCode_() {
  return 'k' + Utilities.getUuid().replace(/-/g, '').slice(0, 15);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
