/**
 * PorterPilot — household account layer.
 *
 * This is what makes the rest of the project (Backend.gs and every other
 * module — Bar, Pets, Money, Contacts, Documents, Packing) multi-tenant
 * without having to touch any of their own logic: they all read/write
 * through Backend.gs's ss_(), and ss_() now resolves to *this* visitor's
 * household Spreadsheet instead of a single bound Sheet.
 *
 * The web app is deployed with:
 *   Execute as:        User accessing the web app
 *   Who has access:    Anyone with a Google account
 * so the visitor's own Google session is the login — there's no separate
 * auth step. Their household Spreadsheet ID is stored in their own
 * UserProperties (private to this script + that Google account).
 *
 * Joining an existing household works via an invite link of the form
 * <web app url>?hh=<spreadsheetId>. The admin shares the household
 * Spreadsheet (Drive "can edit") with the new member first, then sends
 * them that link.
 */

var HOUSEHOLD_MARKER = 'PORTERPILOT_HOUSEHOLD_V1';
var HOUSEHOLD_PROP_KEY = 'HOUSEHOLD_SHEET_ID';

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getMe_() {
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return { email: email, name: email ? email.split('@')[0] : 'You' };
}

// Cheap existence check used by doGet() to pick which page to serve.
function currentHousehold_() {
  var id = PropertiesService.getUserProperties().getProperty(HOUSEHOLD_PROP_KEY);
  if (!id) return null;
  return readHousehold_(id);
}

// ---------------------------------------------------------------------
// Called by Gate.html on load
// ---------------------------------------------------------------------
function getMyHousehold() {
  var me = getMe_();
  var hh = currentHousehold_();
  if (!hh) PropertiesService.getUserProperties().deleteProperty(HOUSEHOLD_PROP_KEY); // forget a stale pointer
  return { me: me, household: hh };
}

function createHousehold(name, peopleNames) {
  name = String(name || '').trim();
  if (!name) throw new Error('Please enter a household name.');
  peopleNames = (peopleNames || []).map(function (n) { return String(n || '').trim(); }).filter(Boolean);
  if (!peopleNames.length) throw new Error('Please enter at least one name.');

  var me = getMe_();
  var ss = SpreadsheetApp.create(name + ' — PorterPilot');
  var ssId = ss.getId();

  buildHouseholdSpreadsheet_(ss, name, peopleNames, me);

  PropertiesService.getUserProperties().setProperty(HOUSEHOLD_PROP_KEY, ssId);
  return { me: me, household: readHousehold_(ssId) };
}

function joinHousehold(sheetId) {
  sheetId = String(sheetId || '').trim();
  if (!sheetId) throw new Error('Missing invite link.');

  var ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    // Most common cause: the admin hasn't shared the Sheet with this
    // Google account yet (or shared it with a different one than the
    // person is currently signed in as).
    throw new Error("You don't have access to that household's Sheet yet. Ask the admin to share it with you (the Sheet's own File ▸ Share), then try the link again.");
  }

  var meta = ss.getSheetByName('Meta');
  if (!meta || meta.getRange('A1').getValue() !== HOUSEHOLD_MARKER) {
    throw new Error("That link doesn't point to a valid PorterPilot household.");
  }

  var hh = readHousehold_(sheetId);
  var me = getMe_();
  PropertiesService.getUserProperties().setProperty(HOUSEHOLD_PROP_KEY, sheetId);
  return { me: me, household: hh };
}

function leaveHousehold() {
  PropertiesService.getUserProperties().deleteProperty(HOUSEHOLD_PROP_KEY);
  return { me: getMe_(), household: null };
}

function getInviteLink() {
  var sheetId = PropertiesService.getUserProperties().getProperty(HOUSEHOLD_PROP_KEY);
  if (!sheetId) throw new Error('No household yet.');
  return getWebAppUrl() + '?hh=' + encodeURIComponent(sheetId);
}

// Called from the app's "More" screen — shares the household Sheet with
// someone's email (Drive "can edit") AND hands back the invite link in one
// step, so the admin doesn't have to go find the Sheet in Drive themselves.
function shareHouseholdWithEmail(email) {
  email = String(email || '').trim();
  if (!email) throw new Error('Please enter an email address.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('That doesn\'t look like a valid email address.');

  var sheetId = PropertiesService.getUserProperties().getProperty(HOUSEHOLD_PROP_KEY);
  if (!sheetId) throw new Error('No household yet.');

  try {
    DriveApp.getFileById(sheetId).addEditor(email);
  } catch (err) {
    throw new Error('Could not share the Sheet with ' + email + '. Double check the address and try again.');
  }

  return { link: getInviteLink(), email: email };
}

// ---------------------------------------------------------------------
// Internal
// ---------------------------------------------------------------------
function readHousehold_(sheetId) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (err) {
    return null; // doesn't exist, or we don't have access
  }
  var meta = ss.getSheetByName('Meta');
  if (!meta || meta.getRange('A1').getValue() !== HOUSEHOLD_MARKER) return null;

  var name = meta.getRange('B1').getValue();
  var adminEmail = meta.getRange('B2').getValue();

  var people = [];
  var pSheet = ss.getSheetByName('People');
  if (pSheet && pSheet.getLastRow() > 1) {
    var rows = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 2).getValues();
    rows.forEach(function (r) {
      if (r[0]) people.push({ name: r[0], email: r[1] || '' });
    });
  }

  return { id: sheetId, name: name, adminEmail: adminEmail, people: people };
}

// Creates every tab the rest of the app expects. Matches the original
// single-household Setup.gs layout, which Backend.gs (and Bar/Pets/Money/
// Contacts/Documents/Packing) already assume — those self-create their own
// extra tabs (BarStock, Pets, Expenses, ...) the first time they're used.
function buildHouseholdSpreadsheet_(ss, name, peopleNames, me) {
  var today = new Date();
  var day = 24 * 60 * 60 * 1000;

  var meta = ss.getSheets()[0];
  meta.setName('Meta');
  meta.getRange('A1:B3').setValues([
    [HOUSEHOLD_MARKER, name],
    ['AdminEmail', me.email],
    ['CreatedAt', today]
  ]);
  // Hidden only after the other tabs exist below — a Sheet can't have
  // every one of its sheets hidden, and right now Meta is the only one.

  var tabs = [
    {
      name: 'People',
      headers: ['Name', 'Email'],
      rows: peopleNames.map(function (n, i) { return [n, i === 0 ? me.email : '']; })
    },
    {
      name: 'Stock',
      headers: ['Item', 'Category', 'Unit', 'Qty', 'LowLevel', 'BuyQty', 'OnShoppingList'],
      rows: [
        ['Nappies', 'Baby', 'pcs', 14, 12, 48],
        ['Baby wipes', 'Baby', 'packs', 3, 2, 6],
        ['Milk', 'Food', 'L', 3, 2, 6],
        ['Toilet paper', 'Cleaning', 'rolls', 9, 4, 24],
        ['Dishwashing liquid', 'Cleaning', 'bottles', 1, 1, 3],
        ['Coffee', 'Food', 'bags', 2, 1, 3],
        ['Bin bags', 'Cleaning', 'rolls', 3, 1, 5]
      ],
      formula: { cell: 'G2', value: '=ARRAYFORMULA(IF(A2:A="","",IF(D2:D<=E2:E,"Yes","No")))' }
    },
    { name: 'StockLog', headers: ['Timestamp', 'Item', 'Change', 'By'], rows: [] },
    {
      name: 'BabyLog', headers: ['Timestamp', 'Type', 'By', 'Note'], rows: [],
      dropdown: { range: 'B2:B1000', values: ['Feed', 'Nappy', 'Sleep'] }
    },
    {
      name: 'Chores', headers: ['Chore', 'AssignedTo', 'Repeat', 'LastDone'],
      rows: [['Put the bins out', peopleNames[0], 'Weekly', ''], ['Run the laundry', peopleNames[peopleNames.length - 1], 'Weekly', '']]
    },
    {
      name: 'Bills', headers: ['Bill', 'DueDayOfMonth', 'LastPaidMonth'],
      rows: [['Electricity', 23, ''], ['Internet', 1, ''], ['Insurance', 1, '']]
    },
    {
      name: 'Meals', headers: ['Date', 'Meal', 'Notes'],
      rows: [[today, '', ''], [new Date(today.getTime() + day), '', '']]
    },
    {
      name: 'Reminders', headers: ['Reminder', 'DueDate', 'Category'],
      rows: [['Example: car licence renewal', new Date(today.getTime() + 18 * day), 'Car']]
    }
  ];

  tabs.forEach(function (t) {
    var sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers])
      .setFontWeight('bold').setBackground('#3B4D37').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    if (t.rows.length) sh.getRange(2, 1, t.rows.length, t.rows[0].length).setValues(t.rows);
    if (t.formula) sh.getRange(t.formula.cell).setFormula(t.formula.value);
    if (t.dropdown) {
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(t.dropdown.values, true).build();
      sh.getRange(t.dropdown.range).setDataValidation(rule);
    }
    if (t.name === 'Meals') sh.getRange('A2:A1000').setNumberFormat('ddd d mmm yyyy');
    if (t.name === 'Reminders') sh.getRange('B2:B1000').setNumberFormat('d mmm yyyy');
    if (t.name === 'StockLog' || t.name === 'BabyLog') sh.getRange('A2:A5000').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.autoResizeColumns(1, t.headers.length);
  });

  meta.hideSheet();
}
