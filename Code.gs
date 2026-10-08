/**
 * PorterPilot — multi-household backend.
 *
 * Runs as a Google Apps Script web app deployed with:
 *   Execute as:        User accessing the web app
 *   Who has access:    Anyone with a Google account
 *
 * Because it executes as the visiting user, there is no separate sign-in
 * step and no external auth provider (Firebase, etc.) — the person's own
 * Google session *is* the login. Each person's household Spreadsheet ID is
 * remembered in their own UserProperties, which is private to (this script,
 * that Google account) and persists across visits.
 *
 * Joining an existing household (rather than creating a new one) works via
 * an invite link of the form:  <web app url>?hh=<spreadsheetId>
 * The admin shares the household Spreadsheet (Drive "can edit") with the
 * new member first, then sends them that link.
 */

var MARKER = 'PORTERPILOT_HOUSEHOLD_V1';
var PROP_KEY = 'HOUSEHOLD_SHEET_ID';

// ---------------------------------------------------------------------
// Web app entry point
// ---------------------------------------------------------------------
function doGet(e) {
  var tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.inviteSheetId = (e && e.parameter && e.parameter.hh) ? e.parameter.hh : '';
  return tmpl.evaluate()
    .setTitle('PorterPilot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------
// Identity helpers
// ---------------------------------------------------------------------
function getMe_() {
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return { email: email, name: email ? email.split('@')[0] : 'You' };
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ---------------------------------------------------------------------
// Called by the client on load
// ---------------------------------------------------------------------
function getMyHousehold() {
  var me = getMe_();
  var props = PropertiesService.getUserProperties();
  var sheetId = props.getProperty(PROP_KEY);

  if (sheetId) {
    var hh = readHousehold_(sheetId);
    if (hh) return { me: me, household: hh };
    // Stale pointer (file deleted / access revoked) — forget it and fall through.
    props.deleteProperty(PROP_KEY);
  }
  return { me: me, household: null };
}

// ---------------------------------------------------------------------
// Create a brand-new household
// ---------------------------------------------------------------------
function createHousehold(name, peopleNames) {
  name = String(name || '').trim();
  if (!name) throw new Error('Please enter a household name.');
  peopleNames = (peopleNames || []).map(function (n) { return String(n || '').trim(); }).filter(Boolean);
  if (!peopleNames.length) throw new Error('Please enter at least one name.');

  var me = getMe_();
  var ss = SpreadsheetApp.create(name + ' — PorterPilot');
  var ssId = ss.getId();

  buildHouseholdSpreadsheet_(ss, name, peopleNames, me);

  PropertiesService.getUserProperties().setProperty(PROP_KEY, ssId);

  return { me: me, household: readHousehold_(ssId) };
}

// ---------------------------------------------------------------------
// Join a household you've been invited to (admin already shared the Sheet)
// ---------------------------------------------------------------------
function joinHousehold(sheetId) {
  sheetId = String(sheetId || '').trim();
  if (!sheetId) throw new Error('Missing invite link.');

  var hh;
  try {
    hh = readHousehold_(sheetId);
  } catch (err) {
    throw new Error("That invite link didn't work. Ask the admin to check it's been shared with you.");
  }
  if (!hh) throw new Error('That link is not a valid PorterPilot household.');

  var me = getMe_();
  PropertiesService.getUserProperties().setProperty(PROP_KEY, sheetId);
  return { me: me, household: hh };
}

function leaveHousehold() {
  PropertiesService.getUserProperties().deleteProperty(PROP_KEY);
  return { me: getMe_(), household: null };
}

function getInviteLink() {
  var sheetId = PropertiesService.getUserProperties().getProperty(PROP_KEY);
  if (!sheetId) throw new Error('No household yet.');
  return getWebAppUrl() + '?hh=' + encodeURIComponent(sheetId);
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
  if (!meta || meta.getRange('A1').getValue() !== MARKER) return null;

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

function buildHouseholdSpreadsheet_(ss, name, peopleNames, me) {
  var today = new Date();
  var day = 24 * 60 * 60 * 1000;

  // Meta — identifies this Sheet as a PorterPilot household and who admins it
  var meta = ss.getSheets()[0];
  meta.setName('Meta');
  meta.getRange('A1:B3').setValues([
    [MARKER, name],
    ['AdminEmail', me.email],
    ['CreatedAt', today]
  ]);
  meta.hideSheet();

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
}
