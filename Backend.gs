/**
 * Household App - Backend
 * Serves the home screen and reads/writes the Google Sheet.
 */

function doGet(e) {
  // Multi-household router: no household yet (or an invite link) -> the
  // sign-in-free create/join gate. Already set up -> straight into the app.
  var hh = null;
  try { hh = currentHousehold_(); } catch (err) { hh = null; }

  if (!hh) {
    var tmpl = HtmlService.createTemplateFromFile('Gate');
    tmpl.inviteSheetId = (e && e.parameter && e.parameter.hh) ? e.parameter.hh : '';
    tmpl.webAppUrl = getWebAppUrl();
    return tmpl.evaluate()
      .setTitle('PorterPilot')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PorterPilot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ---------- Helpers ---------- */

// Every data function in this project (Backend.gs and every other module)
// reads/writes through ss_(). That makes this the one place multi-tenancy
// lives: instead of the Sheet this script used to be bound to, each visitor
// gets their own household's Sheet, resolved from Household.gs.
function ss_() {
  var id = PropertiesService.getUserProperties().getProperty(HOUSEHOLD_PROP_KEY);
  if (!id) throw new Error('No household set up yet.');
  return SpreadsheetApp.openById(id);
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function dayStr_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }

function rows_(name) {
  const sh = ss_().getSheetByName(name);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}

function findRow_(sh, name) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(1, 1, n, 1).getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === name) return i + 1;
  }
  return 0;
}

function who_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const people = rows_('People');
  for (let i = 0; i < people.length; i++) {
    if (people[i][1] && String(people[i][1]).toLowerCase() === email) return people[i][0];
  }
  return email ? email.split('@')[0] : 'Someone';
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- Read everything the home screen needs ---------- */

function getHomeData() {
  const now = new Date();
  const today = dayStr_(now);
  const tomorrow = dayStr_(new Date(now.getTime() + 86400000));
  const out = {};

  out.me = who_();
  out.people = rows_('People').map(r => r[0]).filter(Boolean);

  // Stock
  out.stock = rows_('Stock')
    .filter(r => r[0])
    .map(r => ({
      item: r[0], category: r[1], unit: r[2],
      qty: Number(r[3]) || 0, low: Number(r[4]) || 0, buy: Number(r[5]) || 0
    }));

  // Baby: last time for each type
  const last = { Feed: 0, Nappy: 0, Sleep: 0 };
  out.pendingFeeds = []; // feeds logged by voice in the last 24 hours that still need details
  const dayAgo = now.getTime() - 86400000;
  rows_('BabyLog').forEach((r, i) => {
    const ts = r[0] instanceof Date ? r[0].getTime() : 0;
    if (last[r[1]] !== undefined && ts > last[r[1]]) last[r[1]] = ts;
    if (r[1] === 'Feed' && r[2] === 'Voice' && !r[4] && ts >= dayAgo) out.pendingFeeds.push({ row: i + 2, ts: ts });
  });
  out.pendingFeeds.sort((a, b) => a.ts - b.ts);
  out.last = { feed: last.Feed, nappy: last.Nappy, sleep: last.Sleep };

  // Chores
  const weekAgo = now.getTime() - 7 * 86400000;
  let weekDone = 0;
  const interval = { Daily: 1, Weekly: 7, Monthly: 30 };
  out.chores = rows_('Chores').filter(r => r[0]).map(r => {
    const ld = r[3] instanceof Date ? r[3] : null;
    if (ld && ld.getTime() >= weekAgo) weekDone++;
    const repeat = interval[r[2]] ? r[2] : 'Weekly';
    const doneToday = !!ld && dayStr_(ld) === today;
    let dueIn = 0; // never done = due now
    if (ld) {
      const days = Math.round((Date.parse(today) - Date.parse(dayStr_(ld))) / 86400000);
      dueIn = interval[repeat] - days;
    }
    const status = doneToday ? 'done' : (dueIn <= 0 ? 'due' : 'later');
    return { name: r[0], who: r[1], repeat: repeat, done: doneToday, status: status, dueIn: dueIn };
  });
  out.weekDone = weekDone;
  out.weekTotal = out.chores.length;

  // Bills
  const monthStr = Utilities.formatDate(now, tz_(), 'yyyy-MM');
  const dom = Number(Utilities.formatDate(now, tz_(), 'd'));
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  out.bills = rows_('Bills').filter(r => r[0]).map(r => {
    const dueRaw = Number(r[1]) || 1;
    const due = Math.min(dueRaw, daysInMonth);
    const lp = r[2] instanceof Date ? Utilities.formatDate(r[2], tz_(), 'yyyy-MM') : String(r[2]);
    // negative daysLeft = overdue (not paid and the due day has passed)
    return { name: r[0], due: dueRaw, paid: lp === monthStr, daysLeft: due - dom };
  });

  // Meals
  out.meals = { tonight: '', tomorrow: '' };
  rows_('Meals').forEach(r => {
    if (r[0] instanceof Date) {
      const d = dayStr_(r[0]);
      if (d === today) out.meals.tonight = r[1];
      if (d === tomorrow) out.meals.tomorrow = r[1];
    }
  });

  // Reminders
  out.reminders = rows_('Reminders')
    .filter(r => r[0] && r[1] instanceof Date && String(r[4]).toLowerCase() !== 'yes')
    .map(r => ({ name: r[0], category: r[2], daysLeft: Math.round((Date.parse(dayStr_(r[1])) - Date.parse(today)) / 86400000) }))
    .sort((a, b) => a.daysLeft - b.daysLeft);

  // Calendar events: reads EVERY calendar called "Family" and merges them
  out.events = [];
  out.calendarFound = false;
  try {
    const cals = familyCalendars_();
    if (cals.length) {
      out.calendarFound = true;
      const ppl = eventPeople_();
      const seen = {};
      cals.forEach(cal => {
        cal.getEventsForDay(now).forEach(e => {
          const id = e.getId();
          if (seen[id]) return;
          seen[id] = true;
          out.events.push({
            title: e.getTitle(),
            time: e.isAllDayEvent() ? 'All day' : Utilities.formatDate(e.getStartTime(), tz_(), 'HH:mm'),
            sub: e.getLocation() || '',
            who: parseWho_(e.getDescription(), ppl).label,
            start: e.getStartTime().getTime()
          });
        });
      });
      out.events.sort((a, b) => a.start - b.start);
    }
  } catch (err) { /* calendar not set up yet */ }

  // Next baby clinic visit (from the BabyClinic tab, once it exists)
  out.nextClinic = null;
  try {
    const cs = ss_().getSheetByName('BabyClinic');
    if (cs && cs.getLastRow() >= 2) {
      cs.getRange(2, 1, cs.getLastRow() - 1, 4).getValues().forEach(v => {
        if (!v[1] || !(v[0] instanceof Date) || String(v[3]).toLowerCase() === 'yes') return;
        const daysLeft = Math.round((Date.parse(dayStr_(v[0])) - Date.parse(today)) / 86400000);
        if (daysLeft >= 0 && (!out.nextClinic || daysLeft < out.nextClinic.daysLeft)) {
          out.nextClinic = { what: v[1], daysLeft: daysLeft };
        }
      });
    }
  } catch (err) { /* baby tabs not created yet */ }

  return out;
}

/* ---------- Actions the home screen can save ---------- */

function logBaby(type) {
  // Feeds, nappies and sleeps all need details now, so they are saved with logBabyAt instead
  throw new Error('Please refresh the app. Feeds, nappies and sleeps now need details.');
}

function useItem(item) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Stock');
    const r = findRow_(sh, item);
    if (!r) throw new Error('Item not found');
    const cur = Number(sh.getRange(r, 4).getValue()) || 0;
    if (cur <= 0) return cur;
    sh.getRange(r, 4).setValue(cur - 1);
    ss_().getSheetByName('StockLog').appendRow([new Date(), item, -1, who_()]);
    return cur - 1;
  });
}

function restockItem(item) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Stock');
    const r = findRow_(sh, item);
    if (!r) throw new Error('Item not found');
    const cur = Number(sh.getRange(r, 4).getValue()) || 0;
    const buy = Number(sh.getRange(r, 6).getValue()) || 0;
    sh.getRange(r, 4).setValue(cur + buy);
    ss_().getSheetByName('StockLog').appendRow([new Date(), item, buy, who_()]);
    return cur + buy;
  });
}

function toggleChore(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Chores');
    const r = findRow_(sh, name);
    if (!r) throw new Error('Chore not found');
    const cell = sh.getRange(r, 4);
    const v = cell.getValue();
    const doneToday = v instanceof Date && dayStr_(v) === dayStr_(new Date());
    if (doneToday) { cell.clearContent(); return false; }
    cell.setValue(new Date());
    return true;
  });
}

function markBillPaid(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Bills');
    const r = findRow_(sh, name);
    if (!r) throw new Error('Bill not found');
    const monthStr = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
    sh.getRange(r, 3).setNumberFormat('@').setValue(monthStr);
    return true;
  });
}

/* ---------- Stock screen: add, edit, adjust, delete ---------- */

// Last row that has something in the given column (ignores the auto-filled formula column)
function lastFilledRow_(sh, col) {
  const vals = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '' && vals[i][0] !== null) return i + 1;
  }
  return 1;
}

function cleanItem_(f) {
  const name = String(f.item || '').trim();
  if (!name) throw new Error('Please enter a name');
  return {
    item: name,
    category: String(f.category || '').trim() || 'General',
    unit: String(f.unit || '').trim() || 'pcs',
    qty: Math.max(0, Number(f.qty) || 0),
    low: Math.max(0, Number(f.low) || 0),
    buy: Math.max(0, Number(f.buy) || 0)
  };
}

function nameTaken_(sh, name, exceptRow) {
  const n = lastFilledRow_(sh, 1);
  if (n < 2) return false;
  const vals = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (i + 2 !== exceptRow && String(vals[i][0]).trim().toLowerCase() === name.toLowerCase()) return true;
  }
  return false;
}

function addStockItem(f) {
  return withLock_(() => {
    const it = cleanItem_(f);
    const sh = ss_().getSheetByName('Stock');
    if (nameTaken_(sh, it.item, 0)) throw new Error('That item already exists');
    const row = lastFilledRow_(sh, 1) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
    sh.getRange(row, 1, 1, 6).setValues([[it.item, it.category, it.unit, it.qty, it.low, it.buy]]);
    if (it.qty > 0) ss_().getSheetByName('StockLog').appendRow([new Date(), it.item, it.qty, who_()]);
    return true;
  });
}

function updateStockItem(oldName, f) {
  return withLock_(() => {
    const it = cleanItem_(f);
    const sh = ss_().getSheetByName('Stock');
    const r = findRow_(sh, oldName);
    if (!r) throw new Error('Item not found');
    if (nameTaken_(sh, it.item, r)) throw new Error('That name is already used');
    const before = Number(sh.getRange(r, 4).getValue()) || 0;
    sh.getRange(r, 1, 1, 6).setValues([[it.item, it.category, it.unit, it.qty, it.low, it.buy]]);
    if (it.qty !== before) ss_().getSheetByName('StockLog').appendRow([new Date(), it.item, it.qty - before, who_()]);
    return true;
  });
}

function adjustStock(item, delta) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Stock');
    const r = findRow_(sh, item);
    if (!r) throw new Error('Item not found');
    const cur = Number(sh.getRange(r, 4).getValue()) || 0;
    const next = Math.max(0, cur + Number(delta));
    if (next === cur) return cur;
    sh.getRange(r, 4).setValue(next);
    ss_().getSheetByName('StockLog').appendRow([new Date(), item, next - cur, who_()]);
    return next;
  });
}

// Removes an item without deleting the sheet row, so the auto-filled "OnShoppingList" formula stays intact
function deleteStockItem(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Stock');
    const n = lastFilledRow_(sh, 1);
    if (n < 2) return false;
    const rows = sh.getRange(2, 1, n - 1, 6).getValues();
    const keep = rows.filter(r => r[0] !== name);
    if (keep.length === rows.length) throw new Error('Item not found');
    sh.getRange(2, 1, rows.length, 6).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 6).setValues(keep);
    return true;
  });
}

/* ---------- Chores screen: add, edit, delete ---------- */

function cleanChore_(f) {
  const name = String(f.name || '').trim();
  if (!name) throw new Error('Please enter a name');
  const repeats = ['Daily', 'Weekly', 'Monthly'];
  return { name: name, who: String(f.who || '').trim(), repeat: repeats.indexOf(f.repeat) >= 0 ? f.repeat : 'Weekly' };
}

function addChore(f) {
  return withLock_(() => {
    const c = cleanChore_(f);
    const sh = ss_().getSheetByName('Chores');
    if (nameTaken_(sh, c.name, 0)) throw new Error('That chore already exists');
    const row = lastFilledRow_(sh, 1) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
    sh.getRange(row, 1, 1, 3).setValues([[c.name, c.who, c.repeat]]);
    return true;
  });
}

function updateChore(oldName, f) {
  return withLock_(() => {
    const c = cleanChore_(f);
    const sh = ss_().getSheetByName('Chores');
    const r = findRow_(sh, oldName);
    if (!r) throw new Error('Chore not found');
    if (nameTaken_(sh, c.name, r)) throw new Error('That name is already used');
    sh.getRange(r, 1, 1, 3).setValues([[c.name, c.who, c.repeat]]); // keeps the LastDone date
    return true;
  });
}

function deleteChore(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Chores');
    const r = findRow_(sh, name);
    if (!r) throw new Error('Chore not found');
    sh.deleteRow(r);
    return true;
  });
}

/* ---------- Bills screen: add, edit, delete, paid / not paid ---------- */

function cleanBill_(f) {
  const name = String(f.name || '').trim();
  if (!name) throw new Error('Please enter a name');
  const due = Math.min(31, Math.max(1, Math.round(Number(f.due) || 1)));
  return { name: name, due: due };
}

function addBill(f) {
  return withLock_(() => {
    const b = cleanBill_(f);
    const sh = ss_().getSheetByName('Bills');
    if (nameTaken_(sh, b.name, 0)) throw new Error('That bill already exists');
    const row = lastFilledRow_(sh, 1) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
    sh.getRange(row, 1, 1, 2).setValues([[b.name, b.due]]);
    return true;
  });
}

function updateBill(oldName, f) {
  return withLock_(() => {
    const b = cleanBill_(f);
    const sh = ss_().getSheetByName('Bills');
    const r = findRow_(sh, oldName);
    if (!r) throw new Error('Bill not found');
    if (nameTaken_(sh, b.name, r)) throw new Error('That name is already used');
    sh.getRange(r, 1, 1, 2).setValues([[b.name, b.due]]); // keeps the paid month
    return true;
  });
}

function deleteBill(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Bills');
    const r = findRow_(sh, name);
    if (!r) throw new Error('Bill not found');
    sh.deleteRow(r);
    return true;
  });
}

// Tap once = paid this month, tap again = back to not paid
function toggleBillPaid(name) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('Bills');
    const r = findRow_(sh, name);
    if (!r) throw new Error('Bill not found');
    const monthStr = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
    const cell = sh.getRange(r, 3);
    const v = cell.getValue();
    const cur = v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM') : String(v);
    if (cur === monthStr) { cell.clearContent(); return false; }
    cell.setNumberFormat('@').setValue(monthStr);
    return true;
  });
}

/* ---------- Baby screen: history, clinic visits, growth ---------- */

// Creates the tab the first time it is needed (no locking here, so it is safe to call inside withLock_)
function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('A2:A2000').setNumberFormat('yyyy-mm-dd');
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}
function clinicSheet_() { return ensureSheet_('BabyClinic', ['Date', 'What', 'Notes', 'Done']); }
function growthSheet_() { return ensureSheet_('BabyGrowth', ['Date', 'WeightKg', 'HeightCm', 'HeadCm', 'Note']); }

function dateCell_(v) { return v instanceof Date ? dayStr_(v) : ''; }

function getBabyData() {
  const now = new Date();
  const today = dayStr_(now);
  const out = { log: [], clinic: [], growth: [] };

  // Feeds, nappies and sleeps from the last 14 days
  const sh = babyLogSheet_();
  const n = sh.getLastRow();
  if (n >= 2) {
    const cutoff = now.getTime() - 14 * 86400000;
    sh.getRange(2, 1, n - 1, 7).getValues().forEach((r, i) => {
      const t = r[0] instanceof Date ? r[0].getTime() : 0;
      if (t >= cutoff && ['Feed', 'Nappy', 'Sleep'].indexOf(r[1]) >= 0) {
        out.log.push({
          row: i + 2, ts: t, type: r[1], by: r[2] || '', note: r[3] ? String(r[3]) : '',
          feedType: r[4] ? String(r[4]) : '', amount: r[5] === '' ? null : Number(r[5]), unit: r[6] ? String(r[6]) : ''
        });
      }
    });
    out.log.sort((a, b) => b.ts - a.ts);
  }

  const cs = clinicSheet_();
  if (cs.getLastRow() >= 2) {
    cs.getRange(2, 1, cs.getLastRow() - 1, 4).getValues().forEach((v, i) => {
      const ds = dateCell_(v[0]);
      if (!v[1] || !ds) return;
      out.clinic.push({
        row: i + 2, what: String(v[1]), notes: String(v[2] || ''), date: ds,
        done: String(v[3]).toLowerCase() === 'yes',
        daysLeft: Math.round((Date.parse(ds) - Date.parse(today)) / 86400000)
      });
    });
    out.clinic.sort((a, b) => a.date < b.date ? -1 : 1);
  }

  const gs = growthSheet_();
  if (gs.getLastRow() >= 2) {
    gs.getRange(2, 1, gs.getLastRow() - 1, 5).getValues().forEach((v, i) => {
      const ds = dateCell_(v[0]);
      if (!ds) return;
      const num = x => (x === '' || x === null || isNaN(Number(x))) ? null : Number(x);
      out.growth.push({ row: i + 2, date: ds, weight: num(v[1]), height: num(v[2]), head: num(v[3]), note: String(v[4] || '') });
    });
    out.growth.sort((a, b) => a.date < b.date ? -1 : 1);
  }
  return out;
}

// Adds the extra BabyLog columns the first time they are needed
function babyLogSheet_() {
  const sh = ss_().getSheetByName('BabyLog');
  if (!sh.getRange('E1').getValue()) {
    sh.getRange('E1:G1').setValues([['Detail', 'Amount', 'Unit']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
  } else if (String(sh.getRange('E1').getValue()) === 'FeedType') {
    sh.getRange('E1').setValue('Detail'); // the column now holds: feed type, nappy kind or sleep quality
  }
  return sh;
}

const FEED_UNITS = { 'Formula': 'ml', 'Growth milk': 'ml', 'Juice': 'ml', 'Water': 'ml', 'Tea': 'ml', 'Solids': 'spoons' };

// A feed must say what was fed and how much. Solids must also say what he ate.
function feedDetails_(feedType, amount, food) {
  const unit = FEED_UNITS[feedType];
  const a = Number(amount);
  if (!unit) throw new Error('Please choose what was fed');
  if (!(a > 0) || a > 1000) throw new Error('Please enter how much');
  let note = '';
  if (feedType === 'Solids') {
    note = String(food || '').trim();
    if (!note) throw new Error('Please type what he ate');
    if (note.length > 200) throw new Error('Please keep it under 200 characters');
  }
  return { feedType: feedType, amount: a, unit: unit, note: note };
}

const NAPPY_KINDS = ['Pee', 'Poop', 'Both'];
const SLEEP_QUALITY = ['Well', 'Restless', 'Woke often', 'Hard to settle'];

// A nappy must say what was in it (pee, poop or both). The note is optional (normal, runny, diarrhea...).
function nappyDetails_(kind, note) {
  if (NAPPY_KINDS.indexOf(kind) < 0) throw new Error('Please choose pee, poop or both');
  return { kind: kind, note: String(note || '').trim().slice(0, 200) };
}

// A sleep must say how long. How he slept (quality) and a note are optional.
function sleepDetails_(quality, minutes, note) {
  const m = Math.round(Number(minutes));
  if (!(m >= 1 && m <= 1440)) throw new Error('Please enter how long he slept');
  return {
    quality: SLEEP_QUALITY.indexOf(quality) >= 0 ? quality : '',
    minutes: m, note: String(note || '').trim().slice(0, 200)
  };
}

// Save a feed, nappy or sleep. Feeds MUST include what was fed and how much.
// Also used for things that happened earlier (the time is passed in).
function logBabyAt(type, ts, feedType, amount, food) {
  if (['Feed', 'Nappy', 'Sleep'].indexOf(type) < 0) throw new Error('Unknown type');
  const t = Number(ts);
  const now = Date.now();
  if (!t || t > now + 300000 || t < now - 30 * 86400000) throw new Error('Please pick a time in the last 30 days');
  let note = '', extra = ['', '', ''];
  if (type === 'Feed') {
    const d = feedDetails_(feedType, amount, food);
    note = d.note;
    extra = [d.feedType, d.amount, d.unit];
  } else if (type === 'Nappy') {
    const d = nappyDetails_(feedType, food);
    note = d.note;
    extra = [d.kind, '', ''];
  } else if (type === 'Sleep') {
    const d = sleepDetails_(feedType, amount, food);
    note = d.note;
    extra = [d.quality, d.minutes, 'min'];
  }
  const sh = babyLogSheet_();
  return withLock_(() => {
    sh.appendRow([new Date(t), type, who_(), note, extra[0], extra[1], extra[2]]);
    return true;
  });
}

// Fills in the details for a feed that was logged by voice
function setFeedDetails(row, ts, feedType, amount, food) {
  const d = feedDetails_(feedType, amount, food);
  const sh = babyLogSheet_();
  return withLock_(() => {
    if (row < 2 || row > sh.getLastRow()) throw new Error('That feed has changed. Please refresh.');
    const v = sh.getRange(row, 1, 1, 2).getValues()[0];
    const t = v[0] instanceof Date ? v[0].getTime() : 0;
    if (t !== Number(ts) || v[1] !== 'Feed') throw new Error('That feed has changed. Please refresh.');
    sh.getRange(row, 4).setValue(d.note);
    sh.getRange(row, 5, 1, 3).setValues([[d.feedType, d.amount, d.unit]]);
    return true;
  });
}

function deleteBabyLog(row, ts, type) {
  return withLock_(() => {
    const sh = ss_().getSheetByName('BabyLog');
    if (row < 2 || row > sh.getLastRow()) throw new Error('That entry has changed. Please refresh.');
    const v = sh.getRange(row, 1, 1, 2).getValues()[0];
    const t = v[0] instanceof Date ? v[0].getTime() : 0;
    if (t !== Number(ts) || v[1] !== type) throw new Error('That entry has changed. Please refresh.');
    sh.deleteRow(row);
    return true;
  });
}

/* --- Clinic visits --- */

function cleanClinic_(f) {
  const what = String(f.what || '').trim();
  if (!what) throw new Error('Please enter what the visit is for');
  const ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  return { what: what, notes: String(f.notes || '').trim(), date: Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd') };
}
function checkClinicRow_(sh, row, what) {
  if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 2).getValue()) !== what) {
    throw new Error('That visit has changed. Please refresh.');
  }
}

function addClinic(f) {
  const c = cleanClinic_(f);
  const sh = clinicSheet_();
  return withLock_(() => { sh.appendRow([c.date, c.what, c.notes, '']); return true; });
}
function updateClinic(row, oldWhat, f) {
  const c = cleanClinic_(f);
  const sh = clinicSheet_();
  return withLock_(() => {
    checkClinicRow_(sh, row, oldWhat);
    sh.getRange(row, 1, 1, 3).setValues([[c.date, c.what, c.notes]]); // keeps the Done mark
    return true;
  });
}
function toggleClinicDone(row, what) {
  const sh = clinicSheet_();
  return withLock_(() => {
    checkClinicRow_(sh, row, what);
    const cell = sh.getRange(row, 4);
    if (String(cell.getValue()).toLowerCase() === 'yes') { cell.clearContent(); return false; }
    cell.setValue('Yes');
    return true;
  });
}
function deleteClinic(row, what) {
  const sh = clinicSheet_();
  return withLock_(() => { checkClinicRow_(sh, row, what); sh.deleteRow(row); return true; });
}

/* --- Growth notes --- */

function cleanGrowth_(f) {
  const ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  const num = v => { const x = parseFloat(String(v).replace(',', '.')); return isFinite(x) && x > 0 ? x : ''; };
  const g = { weight: num(f.weight), height: num(f.height), head: num(f.head), note: String(f.note || '').trim() };
  if (g.weight === '' && g.height === '' && g.head === '' && !g.note) throw new Error('Enter a measurement or a note');
  g.date = Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd');
  return g;
}
function checkGrowthRow_(sh, row, dateStr) {
  if (row < 2 || row > sh.getLastRow() || dateCell_(sh.getRange(row, 1).getValue()) !== dateStr) {
    throw new Error('That entry has changed. Please refresh.');
  }
}

function addGrowth(f) {
  const g = cleanGrowth_(f);
  const sh = growthSheet_();
  return withLock_(() => { sh.appendRow([g.date, g.weight, g.height, g.head, g.note]); return true; });
}
function updateGrowth(row, oldDate, f) {
  const g = cleanGrowth_(f);
  const sh = growthSheet_();
  return withLock_(() => {
    checkGrowthRow_(sh, row, oldDate);
    sh.getRange(row, 1, 1, 5).setValues([[g.date, g.weight, g.height, g.head, g.note]]);
    return true;
  });
}
function deleteGrowth(row, dateStr) {
  const sh = growthSheet_();
  return withLock_(() => { checkGrowthRow_(sh, row, dateStr); sh.deleteRow(row); return true; });
}

/* ---------- Reminders screen ---------- */

const REPEAT_MONTHS = { 'Monthly': 1, '3 months': 3, '6 months': 6, 'Yearly': 12 };

// Adds the Repeat and Done columns the first time they are needed
function remindersSheet_() {
  let sh = ss_().getSheetByName('Reminders');
  if (!sh) {
    sh = ensureSheet_('Reminders', ['Reminder', 'DueDate', 'Category']);
    sh.getRange('B2:B2000').setNumberFormat('d mmm yyyy');
  }
  if (!sh.getRange('D1').getValue()) {
    sh.getRange('D1:E1').setValues([['Repeat', 'Done']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
  }
  return sh;
}

// Adds n months to a yyyy-MM-dd date, keeping the day where possible (31 Jan + 1 month = 28 Feb)
function addMonthsStr_(ds, n) {
  const p = ds.split('-').map(Number);
  const total = p[0] * 12 + (p[1] - 1) + n;
  const y = Math.floor(total / 12), m = total % 12;
  const dim = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  const pad = x => (x < 10 ? '0' : '') + x;
  return y + '-' + pad(m + 1) + '-' + pad(Math.min(p[2], dim));
}

function getRemindersData() {
  const sh = remindersSheet_();
  const today = dayStr_(new Date());
  const out = { items: [] };
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach((v, i) => {
      const ds = dateCell_(v[1]);
      if (!v[0] || !ds) return;
      out.items.push({
        row: i + 2, name: String(v[0]), category: String(v[2] || ''), date: ds,
        repeat: REPEAT_MONTHS[v[3]] ? String(v[3]) : '',
        done: String(v[4]).toLowerCase() === 'yes',
        daysLeft: Math.round((Date.parse(ds) - Date.parse(today)) / 86400000)
      });
    });
    out.items.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
  }
  return out;
}

function cleanReminder_(f) {
  const name = String(f.name || '').trim();
  if (!name) throw new Error('Please enter a name');
  const ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  return {
    name: name,
    category: String(f.category || '').trim() || 'Other',
    repeat: REPEAT_MONTHS[f.repeat] ? f.repeat : '',
    date: Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd')
  };
}
function checkReminderRow_(sh, row, name) {
  if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 1).getValue()) !== name) {
    throw new Error('That reminder has changed. Please refresh.');
  }
}

function addReminder(f) {
  const c = cleanReminder_(f);
  const sh = remindersSheet_();
  return withLock_(() => { sh.appendRow([c.name, c.date, c.category, c.repeat, '']); return true; });
}
function updateReminder(row, oldName, f) {
  const c = cleanReminder_(f);
  const sh = remindersSheet_();
  return withLock_(() => {
    checkReminderRow_(sh, row, oldName);
    sh.getRange(row, 1, 1, 4).setValues([[c.name, c.date, c.category, c.repeat]]); // keeps the Done mark
    return true;
  });
}

// Repeating reminders move to their next date. One-off reminders are marked done.
function completeReminder(row, name) {
  const sh = remindersSheet_();
  return withLock_(() => {
    checkReminderRow_(sh, row, name);
    const repeat = String(sh.getRange(row, 4).getValue());
    const months = REPEAT_MONTHS[repeat];
    if (!months) { sh.getRange(row, 5).setValue('Yes'); return { next: null }; }
    const today = dayStr_(new Date());
    let next = addMonthsStr_(dateCell_(sh.getRange(row, 2).getValue()) || today, months);
    for (let i = 0; next <= today && i < 240; i++) next = addMonthsStr_(next, months);
    sh.getRange(row, 2).setValue(Utilities.parseDate(next, tz_(), 'yyyy-MM-dd'));
    sh.getRange(row, 5).clearContent();
    return { next: next };
  });
}
function undoReminder(row, name) {
  const sh = remindersSheet_();
  return withLock_(() => { checkReminderRow_(sh, row, name); sh.getRange(row, 5).clearContent(); return true; });
}
function deleteReminder(row, name) {
  const sh = remindersSheet_();
  return withLock_(() => { checkReminderRow_(sh, row, name); sh.deleteRow(row); return true; });
}

/* ---------- Calendar screen: view, add, edit, delete family events ---------- */

// A small Settings tab remembers which calendar is the family one
function settingsSheet_() {
  let sh = ss_().getSheetByName('Settings');
  if (!sh) {
    sh = ss_().insertSheet('Settings');
    sh.getRange(1, 1, 1, 2).setValues([['Key', 'Value']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 380);
  }
  return sh;
}
function getSetting_(key) {
  const sh = settingsSheet_();
  const n = sh.getLastRow();
  if (n < 2) return '';
  const vals = sh.getRange(2, 1, n - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === key) return String(vals[i][1]);
  return '';
}
function setSetting_(key, value) {
  const sh = settingsSheet_();
  withLock_(() => {
    const n = sh.getLastRow();
    const vals = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getValues() : [];
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
    }
    sh.appendRow([key, value]);
  });
}

// Finds the family calendar. Google's own built-in "Family" calendar is ignored,
// and once the owner has opened the Calendar screen the exact calendar is remembered.
function familyCalendars_() {
  const savedId = getSetting_('FamilyCalendarId');
  if (savedId) {
    const c = CalendarApp.getCalendarById(savedId);
    if (c) return [c];
  }
  let cals = CalendarApp.getCalendarsByName('Family');
  const own = cals.filter(c => !/^family\d/.test(c.getId()));
  if (own.length) cals = own;
  const mine = cals.filter(c => c.isOwnedByMe());
  if (mine.length === 1) {
    setSetting_('FamilyCalendarId', mine[0].getId());
    return [mine[0]];
  }
  return cals;
}

// Who can take part in an event. Edit the EventPeople row in the Settings tab to change the names.
// Defaults to this household's own People tab the first time it's needed.
function eventPeople_() {
  let v = getSetting_('EventPeople');
  if (!v) {
    const names = rows_('People').map(r => r[0]).filter(Boolean);
    v = names.length ? names.join(', ') : 'Everyone';
    setSetting_('EventPeople', v);
  }
  return v.split(',').map(x => x.trim()).filter(Boolean);
}

// The people are stored on the first line of the event description, e.g. "Who: Luke, Marianka"
function parseWho_(desc, ppl) {
  const m = /^Who: ([^\n<]*)(?:\n|<br\s*\/?>)?([\s\S]*)$/i.exec(String(desc || ''));
  const who = m ? m[1].split(',').map(x => x.trim()).filter(Boolean) : [];
  const notes = m ? m[2] : String(desc || '');
  let label = '';
  if (who.length) label = (ppl.length && ppl.every(p => who.indexOf(p) >= 0)) ? 'Everyone' : who.join(', ');
  return { who: who, notes: notes, label: label };
}

function eventInfo_(e, calId, ppl) {
  const w = parseWho_(e.getDescription(), ppl);
  return {
    calId: calId, id: e.getId(), title: e.getTitle(),
    start: e.getStartTime().getTime(), end: e.getEndTime().getTime(),
    allDay: e.isAllDayEvent(), location: e.getLocation() || '', notes: w.notes,
    who: w.who, whoLabel: w.label,
    recurring: e.isRecurringEvent()
  };
}

// The next 31 days of family events
function getCalendarData() {
  const cals = familyCalendars_();
  const ppl = eventPeople_();
  const out = { calendarFound: cals.length > 0, events: [], people: ppl };
  if (!cals.length) return out;
  const start = Utilities.parseDate(dayStr_(new Date()), tz_(), 'yyyy-MM-dd');
  const end = new Date(start.getTime() + 31 * 86400000);
  const seen = {};
  cals.forEach(cal => {
    cal.getEvents(start, end).forEach(e => {
      const key = cal.getId() + '|' + e.getId() + '|' + e.getStartTime().getTime();
      if (seen[key]) return;
      seen[key] = true;
      out.events.push(eventInfo_(e, cal.getId(), ppl));
    });
  });
  out.events.sort((a, b) => a.start - b.start);
  return out;
}

function cleanEvent_(f) {
  const title = String(f.title || '').trim();
  if (!title) throw new Error('Please enter a title');
  const ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  const c = { title: title, allDay: !!f.allDay, location: String(f.location || '').trim(), notes: String(f.notes || '').trim() };
  const who = (Array.isArray(f.who) ? f.who : []).map(x => String(x).replace(/[,\n\r<]/g, ' ').trim().slice(0, 30)).filter(Boolean).slice(0, 10);
  c.description = (who.length ? 'Who: ' + who.join(', ') + (c.notes ? '\n' : '') : '') + c.notes;
  if (c.allDay) {
    c.day = Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd');
  } else {
    if (!/^\d{2}:\d{2}$/.test(String(f.start || ''))) throw new Error('Please pick a start time');
    c.start = Utilities.parseDate(ds + ' ' + f.start, tz_(), 'yyyy-MM-dd HH:mm');
    c.end = /^\d{2}:\d{2}$/.test(String(f.end || '')) ? Utilities.parseDate(ds + ' ' + f.end, tz_(), 'yyyy-MM-dd HH:mm') : new Date(c.start.getTime() + 3600000);
    if (c.end <= c.start) c.end = new Date(c.start.getTime() + 3600000);
  }
  return c;
}

function addCalendarEvent(f) {
  const c = cleanEvent_(f);
  const cals = familyCalendars_();
  if (!cals.length) throw new Error('The Family calendar was not found');
  const cal = cals.filter(x => x.isOwnedByMe())[0] || cals[0];
  const opts = { location: c.location, description: c.description };
  if (c.allDay) cal.createAllDayEvent(c.title, c.day, opts);
  else cal.createEvent(c.title, c.start, c.end, opts);
  return true;
}

function findEvent_(calId, eventId) {
  const cal = CalendarApp.getCalendarById(calId);
  const ev = cal ? cal.getEventById(eventId) : null;
  if (!ev) throw new Error('That event has changed. Please refresh.');
  if (ev.isRecurringEvent()) throw new Error('This event repeats. Please change it in Google Calendar.');
  return ev;
}

function updateCalendarEvent(calId, eventId, f) {
  const c = cleanEvent_(f);
  const ev = findEvent_(calId, eventId);
  ev.setTitle(c.title);
  if (c.allDay) ev.setAllDayDate(c.day); else ev.setTime(c.start, c.end);
  ev.setLocation(c.location);
  ev.setDescription(c.description);
  return true;
}

function deleteCalendarEvent(calId, eventId) {
  findEvent_(calId, eventId).deleteEvent();
  return true;
}

/* ---------- Meals page and AI recipe ideas (Google Gemini) ---------- */

function mealsSheet_() {
  const sh = ss_().getSheetByName('Meals');
  return sh || ensureSheet_('Meals', ['Date', 'Meal', 'Notes']);
}

function getMealsData() {
  const today = dayStr_(new Date());
  const sh = mealsSheet_();
  const out = { items: [] };
  const n = sh.getLastRow();
  if (n >= 2) {
    sh.getRange(2, 1, n - 1, 3).getValues().forEach((v, i) => {
      const ds = dateCell_(v[0]);
      if (!ds || !v[1] || ds < today) return;
      out.items.push({ row: i + 2, date: ds, meal: String(v[1]), notes: String(v[2] || '') });
    });
    out.items.sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
    out.items = out.items.slice(0, 30);
  }
  return out;
}

function addMeal(f) {
  const meal = String(f.meal || '').trim().slice(0, 100);
  if (!meal) throw new Error('Please enter the meal');
  const ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  const sh = mealsSheet_();
  return withLock_(() => {
    sh.appendRow([Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd'), meal, String(f.notes || '').trim().slice(0, 200)]);
    return true;
  });
}

function deleteMeal(row, meal) {
  const sh = mealsSheet_();
  return withLock_(() => {
    if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 2).getValue()) !== meal) {
      throw new Error('That meal has changed. Please refresh.');
    }
    sh.deleteRow(row);
    return true;
  });
}

// Puts missing ingredients on the shopping list (an item with 0 in stock shows up there automatically)
function addToShoppingList(names) {
  const list = (Array.isArray(names) ? names : []).map(x => String(x).trim().slice(0, 40)).filter(Boolean).slice(0, 10);
  const sh = ss_().getSheetByName('Stock');
  return withLock_(() => {
    let added = 0;
    list.forEach(name => {
      if (nameTaken_(sh, name, 0)) return;
      const row = lastFilledRow_(sh, 1) + 1;
      if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
      sh.getRange(row, 1, 1, 6).setValues([[name, 'Food', 'pcs', 0, 1, 1]]);
      added++;
    });
    return added;
  });
}

/* --- Gemini --- */

function geminiError_(code, model) {
  const which = model ? ' (' + model + ')' : '';
  if (code === 429) return 'The AI is busy or the free limit was reached' + which + '. Please try again in a minute.';
  if (code === 400 || code === 401 || code === 403) return 'The AI key was not accepted. Please check GEMINI_API_KEY in the project settings.';
  if (code === 500 || code === 503) return 'Google\'s AI is very busy right now' + which + ' (error ' + code + '). Please try again in a few minutes.';
  return 'The AI could not answer right now' + which + ' (error ' + code + '). Please try again.';
}

// Finds Gemini "flash" models this key can use, newest first, with a lite model as a backup.
// To pick one yourself, add a GeminiModel row in the Settings tab (for example gemini-3.5-flash).
function geminiModels_(key) {
  const override = getSetting_('GeminiModel');
  if (override) return [override];
  const cache = CacheService.getScriptCache();
  const cached = cache.get('gemini_models_v2');
  if (cached) return JSON.parse(cached);

  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error(geminiError_(res.getResponseCode()));
  const models = (JSON.parse(res.getContentText()).models || [])
    .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
    .map(m => String(m.name).replace(/^models\//, ''));
  const pick = re => models.map(n => {
    const m = re.exec(n);
    return m ? { n: n, v: Number(m[1]) * 1000 + Number(m[2] || 0) } : null;
  }).filter(Boolean).sort((a, b) => b.v - a.v).map(x => x.n);
  const flash = pick(/^gemini-(\d+)(?:\.(\d+))?-flash$/);
  const lite = pick(/^gemini-(\d+)(?:\.(\d+))?-flash-lite$/);
  const list = [];
  flash.slice(0, 2).concat(lite.slice(0, 1), flash.slice(2, 4), lite.slice(1, 2))
    .forEach(n => { if (list.indexOf(n) < 0) list.push(n); });
  if (!list.length) throw new Error('No Gemini model was found for this key.');
  cache.put('gemini_models_v2', JSON.stringify(list), 21600);
  return list;
}

const NON_FOOD_CATEGORIES = ['baby', 'cleaning', 'household'];

function cleanList_(v, maxItems, maxLen) {
  return (Array.isArray(v) ? v : []).map(x => String(x).trim().slice(0, maxLen)).filter(Boolean).slice(0, maxItems);
}

function getRecipeIdeas(opts) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('The AI key is not set up yet. Add GEMINI_API_KEY in the project settings.');

  const meal = ['Dinner', 'Lunch', 'Breakfast', 'Snack'].indexOf(opts && opts.meal) >= 0 ? opts.meal : 'Dinner';
  const wish = String((opts && opts.notes) || '').trim().slice(0, 200);
  const items = rows_('Stock')
    .filter(r => r[0] && Number(r[3]) > 0 && NON_FOOD_CATEGORIES.indexOf(String(r[1]).toLowerCase()) < 0)
    .map(r => r[0] + ' (' + r[3] + ' ' + r[2] + ')');
  if (!items.length) throw new Error('No food items in stock yet. Add some on the Stock screen first.');

  const system = 'You are a friendly home cook helping a family plan simple meals. ' +
    'Use mostly the food items listed. Assume salt, pepper, cooking oil and water are always available. ' +
    'Any other ingredient you need must be listed under "missing" (at most 4). Use metric units. ' +
    'Keep every recipe simple: at most 6 short steps, one sentence each. ' +
    'Reply with JSON only, in exactly this shape: ' +
    '{"recipes":[{"name":"","time":"","serves":4,"uses":[""],"missing":[""],"steps":[""]}]}';
  const prompt = 'Food we have at home: ' + items.join('; ') + '.\n' +
    'Meal: ' + meal + '.' + (wish ? '\nRequest from the family: ' + wish : '') + '\n' +
    'Suggest 3 different recipes.';

  const models = geminiModels_(key);
  const payload = JSON.stringify({
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json' }
  });
  let lastCode = 0, lastModel = '', body = '';
  for (let i = 0; i < models.length && i < 5 && !body; i++) {
    lastModel = models[i];
    for (let attempt = 0; attempt < 2; attempt++) {
      const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', {
        method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true, payload: payload
      });
      lastCode = res.getResponseCode();
      if (lastCode === 200) { body = res.getContentText(); break; }
      if (lastCode === 400 && /API key/i.test(res.getContentText())) throw new Error(geminiError_(400));
      if ((lastCode === 500 || lastCode === 503) && attempt === 0) { Utilities.sleep(1500); continue; } // busy: try once more
      break;
    }
  }
  if (!body) throw new Error(geminiError_(lastCode, lastModel));

  const data = JSON.parse(body);
  const parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
  const text = parts.filter(p => !p.thought).map(p => p.text || '').join('').replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
  let json;
  try { json = JSON.parse(text); } catch (err) { throw new Error('The AI reply could not be read. Please try again.'); }
  const recipes = (Array.isArray(json.recipes) ? json.recipes : []).slice(0, 3).map(r => ({
    name: String(r.name || '').trim().slice(0, 80),
    time: String(r.time || '').trim().slice(0, 30),
    serves: Number(r.serves) > 0 && Number(r.serves) <= 20 ? Number(r.serves) : 0,
    uses: cleanList_(r.uses, 12, 40),
    missing: cleanList_(r.missing, 6, 40),
    steps: cleanList_(r.steps, 8, 220)
  })).filter(r => r.name && r.steps.length);
  if (!recipes.length) throw new Error('The AI did not suggest any recipes. Please try again.');
  return { recipes: recipes };
}
