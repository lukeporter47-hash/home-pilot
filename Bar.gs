// Bar.gs - drinks and snacks for the bar, with its own shopping list and a Pour button.
// A separate file in the same Apps Script project. Uses helpers from Backend.gs (ss_, tz_, dayStr_,
// withLock_, who_, lastFilledRow_, nameTaken_, money-style restock logic).
//
// Liquids (Category Spirits, Wine, Beer, Mixers) are tracked in ml. Tapping Pour takes off one
// standard double (50 ml) unless the item has its own pour size. Snacks and Other are tracked by
// count, with + and - like the main Stock screen.

var BAR_CATS = ['Spirits', 'Wine', 'Beer', 'Mixers', 'Snacks', 'Other'];
var BAR_LIQUID_CATS = ['Spirits', 'Wine', 'Beer', 'Mixers'];
var BAR_DEFAULT_POUR_ML = 50; // a double

function barSheet_() {
  var sh = ss_().getSheetByName('BarStock');
  if (!sh) {
    sh = ss_().insertSheet('BarStock');
    sh.getRange(1, 1, 1, 8).setValues([['Item', 'Category', 'Unit', 'Qty', 'Standard', 'BottleMl', 'PourMl', 'OnShoppingList']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(E2:E>0,IF(D2:D<E2:E,"Yes","No"),"No")))');
    sh.autoResizeColumns(1, 8);
  }
  return sh;
}

function barLogSheet_() {
  var sh = ss_().getSheetByName('BarLog');
  if (!sh) {
    sh = ss_().insertSheet('BarLog');
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Item', 'Change', 'By']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 4);
  }
  return sh;
}

function isLiquidCat_(cat) { return BAR_LIQUID_CATS.indexOf(cat) >= 0; }

function readBarStock_() {
  var sh = barSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().map(function (v, i) {
    if (!v[0]) return null;
    return {
      row: i + 2, item: String(v[0]), category: String(v[1] || 'Other'), unit: String(v[2] || (isLiquidCat_(String(v[1])) ? 'ml' : 'pcs')),
      qty: Number(v[3]) || 0, std: Number(v[4]) || 0, bottleMl: Number(v[5]) || 0,
      pourMl: Number(v[6]) || (isLiquidCat_(String(v[1])) ? BAR_DEFAULT_POUR_ML : 0)
    };
  }).filter(Boolean);
}

function getBarData() {
  return { items: readBarStock_(), categories: BAR_CATS, defaultPour: BAR_DEFAULT_POUR_ML };
}

// For the home screen widget: how many items are below their standard
function getBarSummary() {
  var items = readBarStock_();
  var low = items.filter(function (i) { return i.std > 0 && i.qty < i.std; });
  return { total: items.length, toBuy: low.length };
}

function cleanBarItem_(f) {
  var item = String(f.item || '').trim().slice(0, 60);
  if (!item) throw new Error('Please enter a name');
  var category = BAR_CATS.indexOf(f.category) >= 0 ? f.category : 'Other';
  var liquid = isLiquidCat_(category);
  return {
    item: item, category: category, unit: liquid ? 'ml' : 'pcs',
    qty: Math.max(0, Math.round(Number(f.qty)) || 0),
    std: Math.max(0, Math.round(Number(f.std)) || 0),
    bottleMl: liquid ? Math.max(0, Math.round(Number(f.bottleMl)) || 0) : 0,
    pourMl: liquid ? Math.max(1, Math.round(Number(f.pourMl)) || BAR_DEFAULT_POUR_ML) : 0
  };
}

function addBarItem(f) {
  var it = cleanBarItem_(f);
  var sh = barSheet_();
  return withLock_(function () {
    if (nameTaken_(sh, it.item, 0)) throw new Error('That item already exists');
    var row = lastFilledRow_(sh, 1) + 1;
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
    sh.getRange(row, 1, 1, 7).setValues([[it.item, it.category, it.unit, it.qty, it.std, it.bottleMl, it.pourMl]]);
    return true;
  });
}

function updateBarItem(oldName, f) {
  var it = cleanBarItem_(f);
  var sh = barSheet_();
  return withLock_(function () {
    var row = findRow_(sh, oldName);
    if (!row) throw new Error('Item not found');
    if (nameTaken_(sh, it.item, row)) throw new Error('That name is already used');
    sh.getRange(row, 1, 1, 7).setValues([[it.item, it.category, it.unit, it.qty, it.std, it.bottleMl, it.pourMl]]);
    return true;
  });
}

function deleteBarItem(name) {
  var sh = barSheet_();
  return withLock_(function () {
    var n = lastFilledRow_(sh, 1);
    if (n < 2) return false;
    var rows = sh.getRange(2, 1, n - 1, 7).getValues();
    var keep = rows.filter(function (r) { return r[0] !== name; });
    if (keep.length === rows.length) throw new Error('Item not found');
    sh.getRange(2, 1, rows.length, 7).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 7).setValues(keep);
    return true;
  });
}

// Pour: takes one standard tot (or the item's own size) off a liquid. Never goes below zero.
function pourBarItem(item) {
  var sh = barSheet_();
  return withLock_(function () {
    var row = findRow_(sh, item);
    if (!row) throw new Error('Item not found');
    var v = sh.getRange(row, 2, 1, 6).getValues()[0]; // category, unit, qty, std, bottleMl, pourMl
    if (!isLiquidCat_(String(v[0]))) throw new Error('Pour is only for drinks measured in ml');
    var cur = Number(v[2]) || 0;
    var pour = Number(v[5]) || BAR_DEFAULT_POUR_ML;
    var next = Math.max(0, cur - pour);
    sh.getRange(row, 4).setValue(next);
    barLogSheet_().appendRow([new Date(), item, -(cur - next), who_()]);
    return next;
  });
}

// +1 / -1 for snacks and other count items
function adjustBarItem(item, delta) {
  var d = Number(delta);
  if (d !== 1 && d !== -1) throw new Error('Unknown change');
  var sh = barSheet_();
  return withLock_(function () {
    var row = findRow_(sh, item);
    if (!row) throw new Error('Item not found');
    var cur = Number(sh.getRange(row, 4).getValue()) || 0;
    var next = Math.max(0, cur + d);
    if (next === cur) return cur;
    sh.getRange(row, 4).setValue(next);
    barLogSheet_().appendRow([new Date(), item, next - cur, who_()]);
    return next;
  });
}

// Mark bought: liquids top up to a full bottle (or the standard, if set higher); counted items top up to the standard
// cost and category are optional, same as the household Bought box
function restockBarItem(item, cost, category) {
  var sh = barSheet_();
  var next = withLock_(function () {
    var row = findRow_(sh, item);
    if (!row) throw new Error('Item not found');
    var v = sh.getRange(row, 2, 1, 5).getValues()[0]; // category, unit, qty, std, bottleMl
    var cur = Number(v[2]) || 0, std = Number(v[3]) || 0, bottleMl = Number(v[4]) || 0;
    var target = isLiquidCat_(String(v[0])) ? Math.max(bottleMl, std) : std;
    var nxt = target > 0 ? target : cur + 1;
    sh.getRange(row, 4).setValue(nxt);
    barLogSheet_().appendRow([new Date(), item, nxt - cur, who_()]);
    return nxt;
  });
  var amount = parseFloat(String(cost == null ? '' : cost).replace(',', '.'));
  if (amount > 0) addExpense({ amount: amount, date: dayStr_(new Date()), category: category || 'Fun', note: item });
  return next;
}

/* ---------- Suggestions: a private box between whoever wrote it and you ---------- */
// Anyone can add a suggestion. Only you (the household's admin) and the person who wrote it
// can see it and its status. Set SuggestionsAdmin in the Settings tab if it should be someone
// other than the first name on the People tab.

var SUGGESTION_STATUSES = ['New', 'Noted', 'In progress', 'Done', "Won't do"];

function suggestionsSheet_() {
  var sh = ss_().getSheetByName('Suggestions');
  if (!sh) {
    sh = ss_().insertSheet('Suggestions');
    sh.getRange(1, 1, 1, 6).setValues([['Id', 'By', 'Text', 'Status', 'Reply', 'Created']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 6);
  }
  return sh;
}

function suggestionsAdmin_() {
  return getSetting_('SuggestionsAdmin') || (rows_('People')[0] || [''])[0];
}

function readSuggestions_() {
  var sh = suggestionsSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().map(function (v, i) {
    if (!v[0]) return null;
    return {
      row: i + 2, id: String(v[0]), by: String(v[1]), text: String(v[2] || ''),
      status: SUGGESTION_STATUSES.indexOf(v[3]) >= 0 ? String(v[3]) : 'New', reply: String(v[4] || ''),
      created: v[5] instanceof Date ? v[5].getTime() : 0
    };
  }).filter(Boolean);
}

// Everyone sees only their own suggestions and the admin's replies to them.
// The admin also sees everyone else's, so they can reply.
function getSuggestionsData() {
  var me = who_(), admin = suggestionsAdmin_();
  var isAdmin = me.toLowerCase() === admin.toLowerCase();
  var all = readSuggestions_().sort(function (a, b) { return b.created - a.created; });
  var mine = isAdmin ? all : all.filter(function (s) { return s.by.toLowerCase() === me.toLowerCase(); });
  return { me: me, isAdmin: isAdmin, items: mine, statuses: SUGGESTION_STATUSES };
}

function addSuggestion(text) {
  var t = String(text || '').trim().slice(0, 500);
  if (!t) throw new Error('Please type your suggestion');
  var sh = suggestionsSheet_();
  var id = 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  return withLock_(function () { sh.appendRow([id, who_(), t, 'New', '', new Date()]); return true; });
}

// Only the suggestion's own author can delete it, to keep things honest
function deleteSuggestion(id) {
  var sh = suggestionsSheet_();
  return withLock_(function () {
    var vals = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === id) {
        if (String(vals[i][1]).toLowerCase() !== who_().toLowerCase()) throw new Error('You can only remove your own suggestions.');
        sh.deleteRow(i + 2);
        return true;
      }
    }
    throw new Error('That suggestion was not found. Please refresh.');
  });
}

// Only the admin can set a status or write a reply
function updateSuggestionStatus(id, status, reply) {
  if (who_().toLowerCase() !== suggestionsAdmin_().toLowerCase()) throw new Error('Only the admin can update a suggestion.');
  var s = SUGGESTION_STATUSES.indexOf(status) >= 0 ? status : 'New';
  var sh = suggestionsSheet_();
  return withLock_(function () {
    var vals = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][0] === id) { sh.getRange(i + 2, 4, 1, 2).setValues([[s, String(reply || '').trim().slice(0, 500)]]); return true; }
    }
    throw new Error('That suggestion was not found. Please refresh.');
  });
}

/* ---------- AI bar snack ideas ---------- */
// Looks at the snack and mixer items in the bar stock and suggests dips and party snacks.
// Reuses geminiModels_() and geminiError_() from Backend.gs, and the family's Gemini key.

function getBarSnackIdeas(opts) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('The AI key is not set up yet. Add GEMINI_API_KEY in the project settings.');

  var wish = String((opts && opts.notes) || '').trim().slice(0, 200);
  var items = readBarStock_()
    .filter(function (i) { return i.category === 'Snacks' || i.category === 'Mixers'; })
    .map(function (i) { return i.item + ' (' + i.qty + ' ' + i.unit + ')'; });

  var system = 'You are a friendly party host helping plan dips and snacks for drinks with friends. ' +
    'Use the items listed where you can. Assume salt, pepper, oil, butter and basic pantry staples are available. ' +
    'Any other ingredient you need must be listed under "missing" (at most 4). ' +
    'Give a precise amount for every ingredient in "uses" and "missing" (for example "250 g cream cheese", "2 avocados", "1 tsp paprika"), using metric units. ' +
    'Keep every recipe simple and quick to prepare: at most 6 short steps, one sentence each. ' +
    'Reply with JSON only, in exactly this shape: ' +
    '{"recipes":[{"name":"","time":"","serves":6,"uses":[""],"missing":[""],"steps":[""]}]}';
  var prompt = (items.length ? 'Bar snacks and mixers we already have: ' + items.join('; ') + '.\n' : 'We don\u2019t have any snacks in yet.\n') +
    (wish ? 'Request: ' + wish + '.\n' : '') + 'Suggest 3 different dips or snacks for a party.';

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
  var json;
  try { json = JSON.parse(text); } catch (err) { throw new Error('The AI reply could not be read. Please try again.'); }
  var recipes = (Array.isArray(json.recipes) ? json.recipes : []).slice(0, 3).map(function (r) {
    return {
      name: String(r.name || '').trim().slice(0, 80), time: String(r.time || '').trim().slice(0, 30),
      serves: Number(r.serves) > 0 && Number(r.serves) <= 40 ? Number(r.serves) : 0,
      uses: (Array.isArray(r.uses) ? r.uses : []).map(function (x) { return String(x).trim().slice(0, 60); }).filter(Boolean).slice(0, 12),
      missing: (Array.isArray(r.missing) ? r.missing : []).map(function (x) { return String(x).trim().slice(0, 60); }).filter(Boolean).slice(0, 6),
      steps: (Array.isArray(r.steps) ? r.steps : []).map(function (x) { return String(x).trim().slice(0, 220); }).filter(Boolean).slice(0, 8)
    };
  }).filter(function (r) { return r.name && r.steps.length; });
  if (!recipes.length) throw new Error('The AI did not suggest anything. Please try again.');
  return { recipes: recipes };
}

// Adds a missing ingredient to the bar's own shopping list (a snack item with 0 in stock shows up there)
function addToBarShoppingList(names, category) {
  var list = (Array.isArray(names) ? names : []).map(function (x) { return String(x).trim().slice(0, 40); }).filter(Boolean).slice(0, 10);
  var cat = BAR_CATS.indexOf(category) >= 0 ? category : 'Snacks';
  var sh = barSheet_();
  return withLock_(function () {
    var added = 0;
    list.forEach(function (name) {
      if (nameTaken_(sh, name, 0)) return;
      var row = lastFilledRow_(sh, 1) + 1;
      if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
      sh.getRange(row, 1, 1, 5).setValues([[name, cat, 'pcs', 0, 1]]);
      added++;
    });
    return added;
  });
}

// Cocktail recipes, made up fresh each time rather than checked against the bar stock.
// Ingredients always come with an amount for one drink, so the person knows what to buy and how much.
function getCocktailIdeas(opts) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('The AI key is not set up yet. Add GEMINI_API_KEY in the project settings.');

  var wish = String((opts && opts.notes) || '').trim().slice(0, 200);
  var system = 'You are a skilled bartender suggesting cocktail recipes for a home bar. ' +
    'Do not assume anything is already in stock: list every ingredient the drink needs, each with a precise amount for one drink ' +
    '(for example "50 ml white rum", "2 dashes Angostura bitters", "1 sprig mint"), using metric measures. ' +
    'Keep steps short and simple, at most 6 steps, one sentence each. ' +
    'Reply with JSON only, in exactly this shape: ' +
    '{"recipes":[{"name":"","glass":"","time":"","ingredients":[""],"steps":[""]}]}';
  var prompt = (wish ? 'Request: ' + wish + '.\n' : 'No particular request, surprise us.\n') + 'Suggest 3 different cocktails.';

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
  var json;
  try { json = JSON.parse(text); } catch (err) { throw new Error('The AI reply could not be read. Please try again.'); }
  var recipes = (Array.isArray(json.recipes) ? json.recipes : []).slice(0, 3).map(function (r) {
    return {
      name: String(r.name || '').trim().slice(0, 80), glass: String(r.glass || '').trim().slice(0, 40), time: String(r.time || '').trim().slice(0, 30),
      ingredients: (Array.isArray(r.ingredients) ? r.ingredients : []).map(function (x) { return String(x).trim().slice(0, 60); }).filter(Boolean).slice(0, 12),
      steps: (Array.isArray(r.steps) ? r.steps : []).map(function (x) { return String(x).trim().slice(0, 220); }).filter(Boolean).slice(0, 8)
    };
  }).filter(function (r) { return r.name && r.ingredients.length && r.steps.length; });
  if (!recipes.length) throw new Error('The AI did not suggest anything. Please try again.');
  return { recipes: recipes };
}
