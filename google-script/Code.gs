/**
 * Winter with the Segal's in the Blue Ridge — "Who's coming when" storage.
 *
 * How to install or update:
 *   1. In your Google Sheet: Extensions → Apps Script. Replace everything with this file. Save.
 *   2. Choose the function "setup" at the top and press Run (safe to run again; keeps your data).
 *   3. First time: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *      Updating: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (same URL).
 *
 * Tabs:
 *   Households — one row per household. Edit names / usual sizes here.
 *                "hidden" = yes hides a household (e.g. a surprise) from everyone except the admin.
 *                (Old "code" / "personal_link" columns from the first version are no longer used.)
 *   Answers    — one row per household per date, written by the website.
 *   Checklist  — the shared planning checklist (ticked on the site; edit item text here).
 *   History    — every change ever made on the site: when, which household, device type, what.
 *   Settings   — admin_pin: the PIN the admin household types on the site to edit every row
 *                and see hidden households. Change it here any time.
 *
 * Everyone picks their household on the site (honour system) and can change only that row.
 * The admin household needs the PIN.
 */


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

// Hidden from other households when the "hidden" column is first added. Change it later in the sheet.
const HIDDEN_AT_START = [];

const CHECKLIST = [
  'Confirm guest list and bedroom needs', 'Shortlist cabins with free cancellation', 'Verify dog policy at each stay',
  'Check accessible bedrooms and stairs', 'Confirm winter driveway access', 'Arrange holiday-eve groceries',
  'Plan family holiday dinner and games', "Confirm Savannah New Year's event", 'Arrange quiet dog space for fireworks',
  'Check SeaWorld tickets and hours', 'Confirm Orlando unattended-pet policy', 'Finalize meal and rest stops',
];
const HISTORY_SHOWN = 30;

const DATES = [
  '2026-12-22', '2026-12-23', '2026-12-24', '2026-12-25', '2026-12-26', '2026-12-27', '2026-12-28',
  '2026-12-29', '2026-12-30', '2026-12-31', '2027-01-01', '2027-01-02', '2027-01-03',
];

/** Run from the Apps Script editor. Safe to run again: it never overwrites existing data. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let h = ss.getSheetByName('Households');
  if (!h) h = ss.insertSheet('Households');
  if (h.getLastRow() === 0) {
    h.getRange('A:E').setNumberFormat('@');
    h.getRange('C:C').setNumberFormat('0');
    h.appendRow(['id', 'name', 'usual_size', 'admin', 'hidden']);
    HOUSEHOLDS.forEach(([id, name, size, admin]) => {
      h.appendRow([id, name, size, admin ? 'yes' : 'no', HIDDEN_AT_START.indexOf(id) >= 0 ? 'yes' : 'no']);
    });
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, 5).setFontWeight('bold');
    h.autoResizeColumns(1, 5);
  }
  // Older sheets: add the "hidden" column without touching anything else.
  const header = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(String);
  if (header.indexOf('hidden') < 0) {
    const col = header.length + 1;
    h.getRange(1, col).setValue('hidden');
    h.getRange(1, col).setFontWeight('bold');
    for (let r = 2; r <= h.getLastRow(); r++) {
      const id = String(h.getRange(r, 1).getValue()).trim();
      if (id) h.getRange(r, col).setValue(HIDDEN_AT_START.indexOf(id) >= 0 ? 'yes' : 'no');
    }
  }

  let a = ss.getSheetByName('Answers');
  if (!a) a = ss.insertSheet('Answers');
  if (a.getLastRow() === 0) {
    a.getRange('A:D').setNumberFormat('@'); // keep dates as plain text
    a.appendRow(['household_id', 'date', 'answer', 'people', 'updated_at', 'updated_by']);
    a.setFrozenRows(1);
    a.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  let c = ss.getSheetByName('Checklist');
  if (!c) c = ss.insertSheet('Checklist');
  if (c.getLastRow() === 0) {
    c.getRange('A:C').setNumberFormat('@');
    c.appendRow(['id', 'item', 'done', 'updated_at', 'updated_by']);
    CHECKLIST.forEach((item, i) => c.appendRow(['c' + (i + 1), item, 'no', '', '']));
    c.setFrozenRows(1);
    c.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  let hi = ss.getSheetByName('History');
  if (!hi) hi = ss.insertSheet('History');
  if (hi.getLastRow() === 0) {
    hi.getRange('B:H').setNumberFormat('@');
    hi.appendRow(['when', 'by_household', 'device', 'change', 'household', 'date', 'old', 'new']);
    hi.setFrozenRows(1);
    hi.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  let s = ss.getSheetByName('Settings');
  if (!s) s = ss.insertSheet('Settings');
  if (s.getLastRow() === 0) {
    s.getRange('A:B').setNumberFormat('@');
    s.appendRow(['setting', 'value']);
    s.appendRow(['admin_pin', String(100000 + Math.floor(Math.random() * 900000))]);
    s.getRange(1, 1, 1, 2).setFontWeight('bold');
    s.autoResizeColumns(1, 2);
  }

  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

/** Website reads everything: GET (optionally ?pin=ADMINPIN for the admin view) */
function doGet(e) {
  const pin = e && e.parameter && e.parameter.pin;
  let admin = false;
  if (pin) {
    admin = isAdminPin_(pin);
    if (!admin) return json_({ ok: false, error: 'bad_pin' });
  }
  return json_(state_(admin, households_()));
}

/**
 * Website saves changes: POST {as: householdId, pin?: ADMINPIN, device: 'iPhone · Safari · #3f2a',
 *   changes: [{hid, date, value}],   // value '', 'out' or a number
 *   checks:  [{id, done}]}           // shared checklist ticks
 * Every change that actually changes something is appended to the History tab.
 */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hs = households_();
    const admin = body.pin ? isAdminPin_(body.pin) : false;
    if (body.pin && !admin) return json_({ ok: false, error: 'bad_pin' });
    const as = admin ? null : hs.find(x => x.id === body.as && !x.hidden && !x.admin);
    if (!admin && !as) return json_({ ok: false, error: 'choose_household' });
    const by = admin ? 'admin' : as.name;
    const device = String(body.device || '').replace(/[\r\n\t]/g, ' ').slice(0, 60);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = [];
    const now = new Date();

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Answers');
    const data = sh.getDataRange().getValues();
    const index = {}, cur = {};
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0]) + '|' + dateStr_(data[i][1]);
      index[k] = i + 1;
      cur[k] = describe_(data[i][2], data[i][3]);
    }

    const changes = Array.isArray(body.changes) ? body.changes.slice(0, 200) : [];
    let refused = 0;
    changes.forEach(c => {
      const target = hs.find(x => x.id === c.hid);
      if (!target || (!admin && target.id !== as.id) || DATES.indexOf(c.date) < 0) { refused++; return; }
      let answer = '', people = '';
      if (c.value === 'out') answer = 'out';
      else if (c.value !== '' && c.value != null) {
        const n = Math.round(Number(c.value));
        if (!(n >= 1 && n <= 30)) { refused++; return; }
        answer = 'in'; people = String(n);
      }
      const key = target.id + '|' + c.date;
      const before = cur[key] || '?', after = describe_(answer, people);
      if (before === after) return;
      const row = [target.id, c.date, answer, people, now, by];
      if (index[key]) sh.getRange(index[key], 1, 1, 6).setValues([row]);
      else { sh.appendRow(row); index[key] = sh.getLastRow(); }
      cur[key] = after;
      log.push([now, by, device, target.name + ' · ' + c.date + ': ' + before + ' → ' + after, target.id, c.date, before, after]);
    });

    const checks = Array.isArray(body.checks) ? body.checks.slice(0, 50) : [];
    if (checks.length) {
      const cs = ss.getSheetByName('Checklist');
      const rows = cs ? cs.getDataRange().getValues() : [];
      checks.forEach(c => {
        const i = rows.findIndex((r, k) => k > 0 && String(r[0]) === String(c.id));
        if (i < 1) { refused++; return; }
        const was = String(rows[i][2]).toLowerCase() === 'yes', now2 = !!c.done;
        if (was === now2) return;
        cs.getRange(i + 1, 3, 1, 3).setValues([[now2 ? 'yes' : 'no', now, by]]);
        rows[i][2] = now2 ? 'yes' : 'no';
        log.push([now, by, device, 'Checklist: ' + rows[i][1] + (now2 ? ' ✓ done' : ' ✗ not done'), '', '', was ? 'done' : 'not done', now2 ? 'done' : 'not done']);
      });
    }

    if (log.length) {
      const hist = ss.getSheetByName('History');
      if (hist) log.forEach(r => hist.appendRow(r));
    }

    const out = state_(admin, hs);
    out.refused = refused;
    return json_(out);
  } finally {
    lock.releaseLock();
  }
}

function households_() {
  const all = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Households').getDataRange().getValues();
  const head = all[0].map(x => String(x).trim().toLowerCase());
  const col = name => head.indexOf(name);
  const get = (r, name) => (col(name) >= 0 ? String(r[col(name)]).trim() : '');
  return all.slice(1)
    .filter(r => get(r, 'id'))
    .map(r => ({
      id: get(r, 'id'),
      name: get(r, 'name') || get(r, 'id'),
      size: Math.max(1, Number(get(r, 'usual_size')) || 1),
      admin: get(r, 'admin').toLowerCase() === 'yes',
      hidden: get(r, 'hidden').toLowerCase() === 'yes',
    }));
}

function state_(admin, hs) {
  // Hidden households are left out completely for everyone except the admin.
  const visible = hs.filter(x => admin || !x.hidden);
  const ids = {};
  visible.forEach(x => { ids[x.id] = true; });
  const resp = {};
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Answers').getDataRange().getValues().slice(1);
  rows.forEach(r => {
    const hid = String(r[0]), d = dateStr_(r[1]), ans = String(r[2]);
    if (!hid || !ids[hid] || DATES.indexOf(d) < 0) return;
    if (ans === 'in') (resp[hid] = resp[hid] || {})[d] = { in: true, n: Math.max(1, Number(r[3]) || 1) };
    else if (ans === 'out') (resp[hid] = resp[hid] || {})[d] = { in: false };
  });
  const cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Checklist');
  const checklist = cs ? cs.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim()).map(r => ({
    id: String(r[0]), item: String(r[1]), done: String(r[2]).toLowerCase() === 'yes',
    at: r[3] instanceof Date ? r[3].toISOString() : String(r[3] || ''), by: String(r[4] || ''),
  })) : null;
  const hs2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  let history = null;
  if (hs2) {
    const hrows = hs2.getDataRange().getValues().slice(1);
    history = [];
    for (let i = hrows.length - 1; i >= 0 && history.length < HISTORY_SHOWN; i--) {
      const r = hrows[i], hid = String(r[4] || '');
      if (hid && !ids[hid]) continue; // hidden household: only the admin sees its history
      history.push({ when: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), by: String(r[1]), device: String(r[2]), what: String(r[3]) });
    }
  }
  return {
    ok: true,
    admin: admin,
    checklist: checklist,
    history: history,
    households: visible.map(x => {
      const o = { id: x.id, name: x.name, size: x.size, admin: x.admin };
      if (admin && x.hidden) o.hidden = true;
      return o;
    }),
    resp: resp,
  };
}

/** Checks the admin PIN from the Settings tab. Locks out for an hour after 20 wrong tries. */
function isAdminPin_(pin) {
  pin = String(pin || '').trim();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfails') || 0);
  if (!pin || fails >= 20) return false;
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  let real = '';
  if (s) s.getDataRange().getValues().forEach(r => { if (String(r[0]).trim() === 'admin_pin') real = String(r[1]).trim(); });
  const ok = real.length >= 4 && pin === real;
  if (!ok) cache.put('pinfails', String(fails + 1), 3600);
  return ok;
}

function describe_(answer, people) {
  answer = String(answer || '');
  if (answer === 'in') return String(Math.max(1, Number(people) || 1));
  if (answer === 'out') return 'out';
  return '?';
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
