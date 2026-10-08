// Documents.gs - warranties, policies and IDs. This is a separate file in the same Apps Script project.
// It uses helper functions that already live in Backend.gs.

/* ---------- Documents: warranties, policies, IDs ---------- */
// The Sheet keeps the title, category, who it belongs to, the expiry date and a link.
// The scan or photo itself is stored in a Google Drive folder called "Household Documents".
// Tip: do not type ID or policy numbers into the Notes; keep them in the scan.

var DOCS_COLUMNS = ['Title', 'Category', 'Person', 'Expires', 'FileUrl', 'Notes'];

function docsSheet_() {
  let sh = ss_().getSheetByName('Documents');
  if (!sh) {
    sh = ss_().insertSheet('Documents');
    sh.getRange(1, 1, 1, DOCS_COLUMNS.length).setValues([DOCS_COLUMNS])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('D2:D2000').setNumberFormat('d mmm yyyy');
    sh.autoResizeColumns(1, DOCS_COLUMNS.length);
  }
  return sh;
}

// Finds (or creates) the shared Drive folder. It is shared with the emails on the People tab.
function docsFolder_() {
  const id = getSetting_('DocumentsFolderId');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (err) { /* folder missing: make a new one */ }
  }
  const folder = DriveApp.createFolder('Household Documents');
  rows_('People').map(r => String(r[1] || '').trim()).filter(e => /@/.test(e)).forEach(email => {
    try { folder.addEditor(email); } catch (err) { /* could not share with this address */ }
  });
  setSetting_('DocumentsFolderId', folder.getId());
  return folder;
}

function saveDocFile_(name, mime, base64) {
  if (!base64) return '';
  const type = String(mime || '');
  if (type.indexOf('image/') !== 0 && type !== 'application/pdf') throw new Error('Please use a photo or a PDF');
  if (String(base64).length > 14000000) throw new Error('That file is too big. Please use one under about 10 MB.');
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), type, String(name || 'document').slice(0, 100));
  return docsFolder_().createFile(blob).getUrl();
}

var DOCS_CATEGORY_LIST = ['Warranty', 'Insurance', 'ID', 'Vehicle', 'Medical', 'Property', 'Other'];

function cleanDoc_(f) {
  const title = String(f.title || '').trim().slice(0, 100);
  if (!title) throw new Error('Please enter a title');
  const c = {
    title: title,
    category: DOCS_CATEGORY_LIST.indexOf(f.category) >= 0 ? f.category : 'Other',
    person: String(f.person || '').trim().slice(0, 40),
    notes: String(f.notes || '').trim().slice(0, 300),
    expires: ''
  };
  const ds = String(f.expires || '').trim();
  if (ds) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a valid expiry date');
    c.expires = Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd');
  }
  return c;
}

function getDocumentsData() {
  const sh = docsSheet_();
  const today = dayStr_(new Date());
  const out = { items: [] };
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach((v, i) => {
      if (!v[0]) return;
      const ds = v[3] instanceof Date ? dayStr_(v[3]) : '';
      out.items.push({
        row: i + 2, title: String(v[0]), category: String(v[1] || 'Other'), person: String(v[2] || ''),
        expires: ds, daysLeft: ds ? Math.round((Date.parse(ds) - Date.parse(today)) / 86400000) : null,
        url: String(v[4] || ''), notes: String(v[5] || '')
      });
    });
  }
  return out;
}

function addDocument(f) {
  const c = cleanDoc_(f);
  const sh = docsSheet_();
  const url = saveDocFile_(f.fileName, f.fileMime, f.fileData); // before the lock: it may create the folder
  return withLock_(() => { sh.appendRow([c.title, c.category, c.person, c.expires, url, c.notes]); return true; });
}

function updateDocument(row, oldTitle, f) {
  const c = cleanDoc_(f);
  const sh = docsSheet_();
  const newUrl = saveDocFile_(f.fileName, f.fileMime, f.fileData);
  return withLock_(() => {
    if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 1).getValue()) !== oldTitle) {
      throw new Error('That document has changed. Please refresh.');
    }
    const url = newUrl || String(sh.getRange(row, 5).getValue() || '');
    sh.getRange(row, 1, 1, 6).setValues([[c.title, c.category, c.person, c.expires, url, c.notes]]);
    return true;
  });
}

// Removes the entry from the list. The file itself stays in the Drive folder.
function deleteDocument(row, title) {
  const sh = docsSheet_();
  return withLock_(() => {
    if (row < 2 || row > sh.getLastRow() || String(sh.getRange(row, 1).getValue()) !== title) {
      throw new Error('That document has changed. Please refresh.');
    }
    sh.deleteRow(row);
    return true;
  });
}

// Used by the home screen: how many documents there are and which ones are expired or expiring in 60 days
function getDocsSummary() {
  const out = { total: 0, soon: [] };
  const sh = ss_().getSheetByName('Documents');
  if (!sh || sh.getLastRow() < 2) return out;
  const today = dayStr_(new Date());
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(v => {
    if (!v[0]) return;
    out.total++;
    if (v[3] instanceof Date) {
      const daysLeft = Math.round((Date.parse(dayStr_(v[3])) - Date.parse(today)) / 86400000);
      if (daysLeft <= 60) out.soon.push({ title: String(v[0]), daysLeft: daysLeft });
    }
  });
  out.soon.sort((a, b) => a.daysLeft - b.daysLeft);
  return out;
}
