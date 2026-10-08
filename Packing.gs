// Packing.gs - packing lists for trips, with AI suggestions from Google Gemini.
// This is a separate file in the same Apps Script project. It uses helper functions from Backend.gs
// (including the Gemini model finder), so the GEMINI_API_KEY in Project Settings is used here too.
//
// Privacy: the AI only receives the trip type, the number of days, the kind of person going
// (adult, child or baby, never names) and your optional notes.

var PACK_CATS = ['Clothes', 'Toiletries', 'Health', 'Baby and kids', 'Documents', 'Electronics', 'Food and drink', 'Gear', 'Other'];
var PERSON_TYPES = ['Adult', 'Child', 'Baby'];

function tripsSheet_() {
  var sh = ss_().getSheetByName('Trips');
  if (!sh) {
    sh = ss_().insertSheet('Trips');
    sh.getRange(1, 1, 1, 8).setValues([['Id', 'Name', 'Trip', 'Days', 'People', 'Notes', 'Created', 'Archived']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 8);
  }
  return sh;
}

function packSheet_() {
  var sh = ss_().getSheetByName('PackItems');
  if (!sh) {
    sh = ss_().insertSheet('PackItems');
    sh.getRange(1, 1, 1, 6).setValues([['TripId', 'Person', 'Item', 'Qty', 'Category', 'Packed']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 6);
  }
  return sh;
}

function readTrips_() {
  var sh = tripsSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().map(function (v, i) {
    var people = [];
    try { people = JSON.parse(v[4] || '[]'); } catch (err) { people = []; }
    return {
      row: i + 2, id: String(v[0]), name: String(v[1]), trip: String(v[2] || ''), days: Number(v[3]) || 1,
      people: people, notes: String(v[5] || ''), created: v[6] instanceof Date ? v[6].getTime() : 0,
      archived: String(v[7]).toLowerCase() === 'yes'
    };
  }).filter(function (t) { return t.id; });
}

function readPack_() {
  var sh = packSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().map(function (v, i) {
    return {
      row: i + 2, tripId: String(v[0]), person: String(v[1] || 'Everyone'), item: String(v[2] || ''),
      qty: Number(v[3]) || 1, category: String(v[4] || 'Other'), packed: String(v[5]).toLowerCase() === 'yes'
    };
  }).filter(function (x) { return x.tripId && x.item; });
}

function personTypes_() {
  try { return JSON.parse(getSetting_('PersonTypes') || '{}'); } catch (err) { return {}; }
}

function getTripsData() {
  var counts = {};
  readPack_().forEach(function (x) {
    var c = counts[x.tripId] || (counts[x.tripId] = { total: 0, packed: 0 });
    c.total++;
    if (x.packed) c.packed++;
  });
  var trips = readTrips_().map(function (t) {
    var c = counts[t.id] || { total: 0, packed: 0 };
    t.total = c.total; t.packed = c.packed;
    return t;
  });
  trips.sort(function (a, b) { return (a.archived - b.archived) || (b.created - a.created); });
  return { trips: trips, people: eventPeople_(), types: personTypes_(), categories: PACK_CATS };
}

function getTripItems(tripId) {
  var trip = readTrips_().filter(function (t) { return t.id === tripId; })[0];
  if (!trip) throw new Error('That trip was not found. Please refresh.');
  var items = readPack_().filter(function (x) { return x.tripId === tripId; });
  return { trip: trip, items: items, categories: PACK_CATS };
}

function getPackingSummary() {
  var d = getTripsData();
  var active = d.trips.filter(function (t) { return !t.archived; });
  var top = active[0] || null;
  return { active: active.length, top: top ? { name: top.name, packed: top.packed, total: top.total } : null };
}

/* --- AI --- */

// Asks Gemini for JSON, trying the newest suitable models and retrying when Google is busy
function geminiJson_(system, prompt) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('The AI key is not set up yet. Add GEMINI_API_KEY in the project settings.');
  var models = geminiModels_(key);
  var payload = JSON.stringify({
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: 'application/json' }
  });
  var lastCode = 0, lastModel = '', body = '';
  for (var i = 0; i < models.length && i < 5 && !body; i++) {
    lastModel = models[i];
    for (var attempt = 0; attempt < 2; attempt++) {
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', {
        method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true, payload: payload
      });
      lastCode = res.getResponseCode();
      if (lastCode === 200) { body = res.getContentText(); break; }
      if (lastCode === 400 && /API key/i.test(res.getContentText())) throw new Error(geminiError_(400));
      if ((lastCode === 500 || lastCode === 503) && attempt === 0) { Utilities.sleep(1500); continue; }
      break;
    }
  }
  if (!body) throw new Error(geminiError_(lastCode, lastModel));
  var data = JSON.parse(body);
  var parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
  var text = parts.filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('')
    .replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
  try { return JSON.parse(text); } catch (err) { throw new Error('The AI reply could not be read. Please try again.'); }
}

// people = [{n: 'Luke', t: 'Adult'}, ...]. Names are replaced by P1, P2... before anything is sent.
function generatePacking_(days, people, trip, notes, existing, wish) {
  var ids = people.map(function (p, i) { return 'P' + (i + 1); });
  var who = people.map(function (p, i) { return ids[i] + ' = ' + p.t.toLowerCase(); }).join('; ');
  var system = 'You are a practical family travel planner. Create a packing list. ' +
    'Reply with JSON only, in exactly this shape: {"items":[{"who":"P1","item":"","qty":1,"category":""}]}. ' +
    'Rules: "who" is one of the given person ids, or "ALL" for shared family items (first aid kit, chargers, sunscreen). ' +
    'The category must be one of: ' + PACK_CATS.join(', ') + '. ' +
    'Give each person their own clothes and personal items, and scale clothing quantities to the number of days ' +
    '(assume washing is possible on trips longer than 7 days). Babies need nappies, wipes and feeding items scaled to the days. ' +
    'Keep item names short, do not repeat items, and give at most 70 items.';
  var prompt = 'Trip: ' + (trip || 'not specified') + '.\nLength: ' + days + ' days (' + Math.max(0, days - 1) + ' nights).\nPeople: ' + who + '.' +
    (notes ? '\nExtra notes: ' + notes : '');
  if (existing && existing.length) {
    prompt += '\nAlready on the list (do not repeat): ' + existing.join('; ') + '.\nSuggest up to 15 extra items only.' +
      (wish ? '\nExtra request: ' + wish : '');
  }
  var json = geminiJson_(system, prompt);
  var out = [], seen = {};
  (Array.isArray(json.items) ? json.items : []).slice(0, 90).forEach(function (r) {
    var w = String(r.who || '').trim();
    var idx = ids.indexOf(w);
    var person = idx < 0 ? 'Everyone' : people[idx].n;
    var item = String(r.item || '').trim().slice(0, 60);
    if (!item) return;
    var key = (person + '|' + item).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push({
      person: person, item: item,
      qty: Math.max(1, Math.min(99, Math.round(Number(r.qty)) || 1)),
      category: PACK_CATS.indexOf(r.category) >= 0 ? r.category : 'Other'
    });
  });
  if (!out.length) throw new Error('The AI did not suggest any items. Please try again.');
  return out;
}

/* --- Trips --- */

function cleanPeople_(list) {
  var out = [];
  (Array.isArray(list) ? list : []).slice(0, 10).forEach(function (p) {
    var n = String((p && p.n) || '').replace(/[<>]/g, '').trim().slice(0, 30);
    if (n) out.push({ n: n, t: PERSON_TYPES.indexOf(p.t) >= 0 ? p.t : 'Adult' });
  });
  return out;
}

function createTrip(f) {
  var name = String(f.name || '').trim().slice(0, 60);
  if (!name) throw new Error('Please enter a name for the trip');
  var days = Math.round(Number(f.days));
  if (!(days >= 1 && days <= 90)) throw new Error('Please enter the number of days (1 to 90)');
  var people = cleanPeople_(f.people);
  if (!people.length) throw new Error('Please choose who is going');
  var trip = String(f.trip || '').trim().slice(0, 100);
  var notes = String(f.notes || '').trim().slice(0, 200);

  var items = f.useAi === false ? [] : generatePacking_(days, people, trip, notes, null, '');

  var types = personTypes_();
  people.forEach(function (p) { types[p.n] = p.t; });
  setSetting_('PersonTypes', JSON.stringify(types));

  var id = 't' + Date.now().toString(36);
  var tsh = tripsSheet_(), psh = packSheet_();
  return withLock_(function () {
    tsh.appendRow([id, name, trip, days, JSON.stringify(people), notes, new Date(), '']);
    if (items.length) {
      var start = psh.getLastRow() + 1;
      psh.getRange(start, 1, items.length, 6).setValues(items.map(function (x) { return [id, x.person, x.item, x.qty, x.category, '']; }));
    }
    return { id: id, count: items.length };
  });
}

function suggestMoreItems(tripId, wish) {
  var t = readTrips_().filter(function (x) { return x.id === tripId; })[0];
  if (!t) throw new Error('That trip was not found. Please refresh.');
  var existing = readPack_().filter(function (x) { return x.tripId === tripId; }).map(function (x) { return x.item; });
  var items = generatePacking_(t.days, t.people, t.trip, t.notes, existing, String(wish || '').trim().slice(0, 200));
  var have = {};
  readPack_().filter(function (x) { return x.tripId === tripId; }).forEach(function (x) { have[(x.person + '|' + x.item).toLowerCase()] = true; });
  var fresh = items.filter(function (x) { return !have[(x.person + '|' + x.item).toLowerCase()]; });
  var psh = packSheet_();
  return withLock_(function () {
    if (fresh.length) {
      var start = psh.getLastRow() + 1;
      psh.getRange(start, 1, fresh.length, 6).setValues(fresh.map(function (x) { return [tripId, x.person, x.item, x.qty, x.category, '']; }));
    }
    return fresh.length;
  });
}

function setTripArchived(tripId, archived) {
  var sh = tripsSheet_();
  return withLock_(function () {
    var t = readTrips_().filter(function (x) { return x.id === tripId; })[0];
    if (!t) throw new Error('That trip was not found. Please refresh.');
    sh.getRange(t.row, 8).setValue(archived ? 'Yes' : '');
    return true;
  });
}

function deleteTrip(tripId) {
  var tsh = tripsSheet_(), psh = packSheet_();
  return withLock_(function () {
    var t = readTrips_().filter(function (x) { return x.id === tripId; })[0];
    if (!t) throw new Error('That trip was not found. Please refresh.');
    tsh.deleteRow(t.row);
    var n = psh.getLastRow();
    if (n >= 2) {
      var rows = psh.getRange(2, 1, n - 1, 6).getValues();
      var keep = rows.filter(function (r) { return String(r[0]) !== tripId; });
      psh.getRange(2, 1, rows.length, 6).clearContent();
      if (keep.length) psh.getRange(2, 1, keep.length, 6).setValues(keep);
    }
    return true;
  });
}

// Untick everything so a list can be reused
function resetPacked(tripId) {
  var sh = packSheet_();
  return withLock_(function () {
    var n = sh.getLastRow();
    if (n < 2) return true;
    var rows = sh.getRange(2, 1, n - 1, 6).getValues();
    rows.forEach(function (r) { if (String(r[0]) === tripId) r[5] = ''; });
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
    return true;
  });
}

/* --- Items --- */

function cleanPackItem_(f) {
  var item = String(f.item || '').trim().slice(0, 60);
  if (!item) throw new Error('Please enter the item');
  return {
    person: String(f.person || 'Everyone').replace(/[<>]/g, '').trim().slice(0, 30) || 'Everyone',
    item: item,
    qty: Math.max(1, Math.min(99, Math.round(Number(f.qty)) || 1)),
    category: PACK_CATS.indexOf(f.category) >= 0 ? f.category : 'Other'
  };
}

function checkPackRow_(sh, row, tripId, item) {
  if (row < 2 || row > sh.getLastRow()) throw new Error('That item has changed. Please refresh.');
  var v = sh.getRange(row, 1, 1, 3).getValues()[0];
  if (String(v[0]) !== tripId || String(v[2]) !== item) throw new Error('That item has changed. Please refresh.');
}

function addPackItem(tripId, f) {
  var c = cleanPackItem_(f);
  var sh = packSheet_();
  return withLock_(function () { sh.appendRow([tripId, c.person, c.item, c.qty, c.category, '']); return true; });
}

function updatePackItem(row, old, f) {
  var c = cleanPackItem_(f);
  var sh = packSheet_();
  return withLock_(function () {
    checkPackRow_(sh, row, old.tripId, old.item);
    sh.getRange(row, 2, 1, 4).setValues([[c.person, c.item, c.qty, c.category]]); // keeps the ticked state
    return true;
  });
}

function togglePacked(row, tripId, item) {
  var sh = packSheet_();
  return withLock_(function () {
    checkPackRow_(sh, row, tripId, item);
    var cell = sh.getRange(row, 6);
    if (String(cell.getValue()).toLowerCase() === 'yes') { cell.clearContent(); return false; }
    cell.setValue('Yes');
    return true;
  });
}

function deletePackItem(row, tripId, item) {
  var sh = packSheet_();
  return withLock_(function () { checkPackRow_(sh, row, tripId, item); sh.deleteRow(row); return true; });
}
