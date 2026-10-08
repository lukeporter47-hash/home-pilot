// Money.gs - expenses and monthly budgets. This is a separate file in the same Apps Script project.
// It uses helper functions that already live in Backend.gs.
// Amounts are shown with the currency symbol in the Settings tab (row "Currency"). The default is R.

var MONEY_CATS = ['Groceries', 'Baby', 'Transport', 'Home', 'Eating out', 'Health', 'Fun', 'Other'];

function currency_() {
  return getSetting_('Currency') || 'R';
}

function expensesSheet_() {
  var sh = ss_().getSheetByName('Expenses');
  if (!sh) {
    sh = ss_().insertSheet('Expenses');
    sh.getRange(1, 1, 1, 5).setValues([['Date', 'Amount', 'Category', 'Note', 'By']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('A2:A5000').setNumberFormat('yyyy-mm-dd');
    sh.getRange('B2:B5000').setNumberFormat('#,##0.00');
    sh.autoResizeColumns(1, 5);
  }
  return sh;
}

function budgetSheet_() {
  var sh = ss_().getSheetByName('Budget');
  if (!sh) {
    sh = ss_().insertSheet('Budget');
    sh.getRange(1, 1, 1, 2).setValues([['Category', 'MonthlyBudget']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, MONEY_CATS.length, 2).setValues(MONEY_CATS.map(function (c) { return [c, 0]; }));
    sh.getRange('B2:B50').setNumberFormat('#,##0.00');
    sh.autoResizeColumns(1, 2);
  }
  return sh;
}

function readBudgets_() {
  var sh = budgetSheet_();
  var out = {};
  MONEY_CATS.forEach(function (c) { out[c] = 0; });
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (v) {
      var c = String(v[0]).trim();
      if (out[c] !== undefined) out[c] = Number(v[1]) || 0;
    });
  }
  return out;
}

function readExpenses_() {
  var sh = expensesSheet_();
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues().map(function (v, i) {
    return {
      row: i + 2, date: v[0] instanceof Date ? dayStr_(v[0]) : '', amount: Number(v[1]) || 0,
      category: String(v[2] || 'Other'), note: String(v[3] || ''), by: String(v[4] || '')
    };
  }).filter(function (x) { return x.date && x.amount; });
}

function monthOf_(m) {
  return /^\d{4}-\d{2}$/.test(String(m || '')) ? String(m) : dayStr_(new Date()).slice(0, 7);
}

// Everything the Money screen needs for one month (month looks like 2026-09)
function getMoneyData(month) {
  var m = monthOf_(month);
  var items = readExpenses_().filter(function (x) { return x.date.slice(0, 7) === m; });
  var spent = {};
  MONEY_CATS.forEach(function (c) { spent[c] = 0; });
  items.forEach(function (x) {
    var c = MONEY_CATS.indexOf(x.category) >= 0 ? x.category : 'Other';
    spent[c] = Math.round((spent[c] + x.amount) * 100) / 100;
  });
  items.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : b.row - a.row; });
  return { month: m, currency: currency_(), categories: MONEY_CATS, budgets: readBudgets_(), spent: spent, items: items };
}

// Used by the home screen: this month at a glance
function getMoneySummary() {
  var d = getMoneyData();
  var spent = 0, budget = 0, over = [], near = [];
  MONEY_CATS.forEach(function (c) {
    spent += d.spent[c];
    budget += d.budgets[c];
    if (d.budgets[c] > 0) {
      var pct = d.spent[c] / d.budgets[c];
      if (pct > 1) over.push(c); else if (pct >= 0.9) near.push(c);
    }
  });
  return { currency: d.currency, spent: Math.round(spent * 100) / 100, budget: Math.round(budget * 100) / 100, over: over, near: near };
}

function cleanExpense_(f) {
  var amount = Math.round(parseFloat(String(f.amount).replace(',', '.')) * 100) / 100;
  if (!(amount > 0) || amount > 10000000) throw new Error('Please enter an amount above zero');
  var ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  return {
    amount: amount,
    date: Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd'),
    category: MONEY_CATS.indexOf(f.category) >= 0 ? f.category : 'Other',
    note: String(f.note || '').trim().slice(0, 100)
  };
}

function addExpense(f) {
  var c = cleanExpense_(f);
  var sh = expensesSheet_();
  return withLock_(function () { sh.appendRow([c.date, c.amount, c.category, c.note, who_()]); return true; });
}

// "old" is what the row looked like when it was loaded, so we never change the wrong row
function checkExpenseRow_(sh, row, old) {
  if (row < 2 || row > sh.getLastRow()) throw new Error('That expense has changed. Please refresh.');
  var v = sh.getRange(row, 1, 1, 3).getValues()[0];
  var ds = v[0] instanceof Date ? dayStr_(v[0]) : '';
  if (ds !== old.date || Number(v[1]) !== Number(old.amount) || String(v[2]) !== old.category) {
    throw new Error('That expense has changed. Please refresh.');
  }
}

function updateExpense(row, old, f) {
  var c = cleanExpense_(f);
  var sh = expensesSheet_();
  return withLock_(function () {
    checkExpenseRow_(sh, row, old);
    sh.getRange(row, 1, 1, 4).setValues([[c.date, c.amount, c.category, c.note]]); // keeps who logged it
    return true;
  });
}

function deleteExpense(row, old) {
  var sh = expensesSheet_();
  return withLock_(function () { checkExpenseRow_(sh, row, old); sh.deleteRow(row); return true; });
}

// map looks like {"Groceries": 3000, "Baby": 1500, ...}
function setBudgets(map) {
  var sh = budgetSheet_();
  return withLock_(function () {
    var last = sh.getLastRow();
    var names = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
    MONEY_CATS.forEach(function (c) {
      var v = Math.max(0, Math.round((parseFloat(String((map || {})[c]).replace(',', '.')) || 0) * 100) / 100);
      var i = names.indexOf(c);
      if (i >= 0) sh.getRange(i + 2, 2).setValue(v);
      else sh.appendRow([c, v]);
    });
    return true;
  });
}
