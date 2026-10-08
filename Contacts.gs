// Contacts.gs - plumber, doctor, school and other numbers. This is a separate file in the same Apps Script project.
// It uses helper functions that already live in Backend.gs.
// The WhatsApp button turns numbers that start with 0 into international numbers using the country code
// in the Settings tab (row "CountryCode", default 27).

var CONTACT_CATS = ['Home services', 'Medical', 'School', 'Family', 'Emergency', 'Other'];

function contactsSheet_() {
  var sh = ss_().getSheetByName('Contacts');
  if (!sh) {
    sh = ss_().insertSheet('Contacts');
    sh.getRange(1, 1, 1, 5).setValues([['Name', 'Category', 'Phone', 'Notes', 'Pinned']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('C2:C2000').setNumberFormat('@'); // keeps the leading 0 of phone numbers
    sh.autoResizeColumns(1, 5);
  }
  return sh;
}

function getContactsData() {
  var sh = contactsSheet_();
  var out = { countryCode: getSetting_('CountryCode') || '27', items: [] };
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (v, i) {
      if (!v[0]) return;
      out.items.push({
        row: i + 2, name: String(v[0]), category: String(v[1] || 'Other'), phone: String(v[2] || ''),
        notes: String(v[3] || ''), pinned: String(v[4]).toLowerCase() === 'yes'
      });
    });
  }
  return out;
}

function cleanContact_(f) {
  var name = String(f.name || '').trim().slice(0, 80);
  if (!name) throw new Error('Please enter a name');
  return {
    name: name,
    category: CONTACT_CATS.indexOf(f.category) >= 0 ? f.category : 'Other',
    phone: String(f.phone || '').replace(/[^\d+\-\s()]/g, '').trim().slice(0, 25),
    notes: String(f.notes || '').trim().slice(0, 200),
    pinned: f.pinned === 'Yes' || f.pinned === true ? 'Yes' : ''
  };
}

function checkContactRow_(sh, row, name) {
  if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 1).getValue()) !== name) {
    throw new Error('That contact has changed. Please refresh.');
  }
}

function addContact(f) {
  var c = cleanContact_(f);
  var sh = contactsSheet_();
  return withLock_(function () { sh.appendRow([c.name, c.category, c.phone, c.notes, c.pinned]); return true; });
}

function updateContact(row, oldName, f) {
  var c = cleanContact_(f);
  var sh = contactsSheet_();
  return withLock_(function () {
    checkContactRow_(sh, row, oldName);
    sh.getRange(row, 1, 1, 5).setValues([[c.name, c.category, c.phone, c.notes, c.pinned]]);
    return true;
  });
}

function deleteContact(row, name) {
  var sh = contactsSheet_();
  return withLock_(function () { checkContactRow_(sh, row, name); sh.deleteRow(row); return true; });
}
