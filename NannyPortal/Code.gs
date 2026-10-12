/**
 * PorterPilot — standalone Nanny Portal.
 *
 * This is a SEPARATE, tiny Apps Script project — not part of the main
 * PorterPilot app — so your nanny never has to be added as a member of
 * your household (which would mean giving her edit access to your whole
 * household Google Sheet, including Money, Bills, Documents and Contacts).
 *
 * This script only ever touches a few things in your household Sheet: the
 * NannyLog tab (so she can log naps/meals/snacks/activities/notes, which
 * show up automatically in the main PorterPilot app's Nanny screen), the
 * BabyLog tab (so a feed she logs — what and how much — shows up exactly
 * like one logged from the main app's Baby screen, amount and all), and a
 * handful of rows in the Settings tab (her instructions and emergency
 * contact) — which she can only ever READ here. You still edit those from
 * the main app's Nanny > Info tab.
 *
 * Because this runs as its own Apps Script project deployed "Execute as:
 * Me", it can read/write your household Sheet using YOUR authorization,
 * even though your nanny herself is never given any access to that Sheet.
 * There is nothing else for her to find — this project's code is the only
 * thing she can ever reach, and it only exposes the functions below.
 *
 * SETUP (one-time, about 5 minutes):
 * 1. Go to script.google.com and click "New project".
 * 2. Delete the placeholder code in Code.gs, and paste in this entire file.
 * 3. Click the + next to "Files" → "HTML", name the file exactly "Index"
 *    (capital I, no .html — Apps Script adds that), and paste in the
 *    contents of Index.html from this same folder.
 * 4. Below, set SPREADSHEET_ID to your household Sheet's ID — it's the
 *    same ID that appears after "?hh=" in your PorterPilot invite link
 *    (open PorterPilot → More → Invite link, and copy the part after "hh=").
 * 5. Click Deploy → New deployment → type: Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    ("Anyone" is what lets your nanny use it without her own Google
 *    account, and without ever touching your household Sheet directly.)
 * 6. Click Deploy, authorize it with your own Google account when asked,
 *    then copy the Web app URL it gives you — that's the link to send
 *    your nanny. Bookmark it on her phone; nothing else is reachable
 *    from it.
 */

const SPREADSHEET_ID = 'PASTE_YOUR_HOUSEHOLD_SHEET_ID_HERE';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Nanny Log')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

const NANNY_LOG_TYPES = ['Nap', 'Meal', 'Snack', 'Activity', 'Note'];

// Matches the main app's NannyLog layout exactly, so entries logged here
// show up in PorterPilot's own Nanny > Log tab, and vice versa.
function nannyLogSheet_() {
  let sh = ss_().getSheetByName('NannyLog');
  if (!sh) {
    sh = ss_().insertSheet('NannyLog');
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Type', 'Note', 'By']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Read-only lookup into the same Settings tab the main app's Nanny > Info
// tab writes to (NannyInstructions / NannyEmergencyName / NannyEmergencyPhone).
function getSetting_(key) {
  const sh = ss_().getSheetByName('Settings');
  if (!sh) return '';
  const n = sh.getLastRow();
  if (n < 2) return '';
  const vals = sh.getRange(2, 1, n - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === key) return String(vals[i][1]);
  return '';
}

// Same BabyLog tab the main app's Baby screen uses — a feed logged here
// shows up there too, amount and all, and vice versa. Doesn't create the
// sheet or its extra columns; those are set up by the main app already.
function babyLogSheet_() {
  const sh = ss_().getSheetByName('BabyLog');
  if (!sh) throw new Error('BabyLog tab not found — open the main PorterPilot app at least once first.');
  return sh;
}

const FEED_UNITS = { 'Formula': 'ml', 'Growth milk': 'ml', 'Juice': 'ml', 'Water': 'ml', 'Tea': 'ml', 'Solids': 'spoons' };

// Same validation as the main app: a feed must say what was fed and how
// much; Solids must also say what he ate.
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

// Everything the portal shows her: instructions, emergency contact, and
// today's log entries (NannyLog + any feeds she's logged) so she can see
// what's already been done.
function getNannyPortalData() {
  const out = {
    instructions: getSetting_('NannyInstructions'),
    emName: getSetting_('NannyEmergencyName'),
    emPhone: getSetting_('NannyEmergencyPhone'),
    log: []
  };
  const startOfDay = new Date(); startOfDay.setHours(0, 0, 0, 0);
  const cutoff = startOfDay.getTime();

  const sh = nannyLogSheet_();
  const n = sh.getLastRow();
  if (n >= 2) {
    sh.getRange(2, 1, n - 1, 4).getValues().forEach(r => {
      const t = r[0] instanceof Date ? r[0].getTime() : 0;
      if (t >= cutoff && NANNY_LOG_TYPES.indexOf(r[1]) >= 0) {
        out.log.push({ ts: t, type: r[1], note: r[2] ? String(r[2]) : '' });
      }
    });
  }

  let bsh;
  try { bsh = babyLogSheet_(); } catch (e) { bsh = null; }
  if (bsh) {
    const bn = bsh.getLastRow();
    if (bn >= 2) {
      bsh.getRange(2, 1, bn - 1, 7).getValues().forEach(r => {
        const t = r[0] instanceof Date ? r[0].getTime() : 0;
        if (t >= cutoff && r[1] === 'Feed') {
          out.log.push({
            ts: t, type: r[4] ? String(r[4]) : 'Feed',
            note: r[3] ? String(r[3]) : '',
            amount: r[5] === '' ? null : Number(r[5]), unit: r[6] ? String(r[6]) : ''
          });
        }
      });
    }
  }

  out.log.sort((a, b) => b.ts - a.ts);
  return out;
}

// Log a nap/meal/snack/activity/note.
function addNannyPortalEntry(type, note) {
  if (NANNY_LOG_TYPES.indexOf(type) < 0) throw new Error('Please choose a type');
  const n = String(note || '').trim().slice(0, 200);
  const sh = nannyLogSheet_();
  return withLock_(() => {
    sh.appendRow([new Date(), type, n, 'Nanny']);
    return true;
  });
}

// Log a feed — same validation and same BabyLog row shape as the main
// app's Baby screen, so it appears there exactly as if she'd used it.
function addNannyPortalFeed(feedType, amount, food) {
  const d = feedDetails_(feedType, amount, food);
  const sh = babyLogSheet_();
  return withLock_(() => {
    sh.appendRow([new Date(), 'Feed', 'Nanny', d.note, d.feedType, d.amount, d.unit]);
    return true;
  });
}
