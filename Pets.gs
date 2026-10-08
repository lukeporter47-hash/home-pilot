// Pets.gs - pet profiles (dog, cat or other), weight history and care schedule.
// This is a separate file in the same Apps Script project. It uses helper functions from Backend.gs.
// The care schedule (vaccinations, deworming, flea and tick...) is saved as repeating reminders on the
// Reminders tab, with the pet's name as the category, so it also shows on the Reminders screen.

var PET_SPECIES = ['Dog', 'Cat', 'Other'];
var PET_SEX = ['Male', 'Female', 'Unknown'];

function petsSheet_() {
  var sh = ss_().getSheetByName('Pets');
  if (!sh) {
    sh = ss_().insertSheet('Pets');
    sh.getRange(1, 1, 1, 11).setValues([['Id', 'Name', 'Species', 'Breed', 'Sex', 'Birthday', 'Chip', 'Vet', 'Food', 'Notes', 'Created']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('F2:F500').setNumberFormat('d mmm yyyy');
    sh.getRange('G2:G500').setNumberFormat('@'); // keeps leading zeros of the chip number
    sh.autoResizeColumns(1, 11);
  }
  return sh;
}

function petWeightsSheet_() {
  var sh = ss_().getSheetByName('PetWeights');
  if (!sh) {
    sh = ss_().insertSheet('PetWeights');
    sh.getRange(1, 1, 1, 3).setValues([['PetId', 'Date', 'Kg']])
      .setFontWeight('bold').setBackground('#1B7A6B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('B2:B2000').setNumberFormat('d mmm yyyy');
    sh.autoResizeColumns(1, 3);
  }
  return sh;
}

function readPets_() {
  var sh = petsSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues().map(function (v, i) {
    return {
      row: i + 2, id: String(v[0]), name: String(v[1]), species: String(v[2] || 'Dog'), breed: String(v[3] || ''),
      sex: String(v[4] || 'Unknown'), birthday: v[5] instanceof Date ? dayStr_(v[5]) : '', chip: String(v[6] || ''),
      vet: String(v[7] || ''), food: String(v[8] || ''), notes: String(v[9] || '')
    };
  }).filter(function (p) { return p.id && p.name; });
}

function readPetWeights_() {
  var sh = petWeightsSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().map(function (v, i) {
    return { row: i + 2, petId: String(v[0]), date: v[1] instanceof Date ? dayStr_(v[1]) : '', kg: Number(v[2]) || 0 };
  }).filter(function (w) { return w.petId && w.date && w.kg; });
}

function getPetsData() {
  var weights = readPetWeights_();
  weights.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return { pets: readPets_(), weights: weights };
}

// For the home screen: each pet and its next care item
function getPetsSummary() {
  var pets = readPets_();
  var rem = getRemindersData().items.filter(function (r) { return !r.done; });
  return {
    pets: pets.map(function (p) {
      var mine = rem.filter(function (r) { return r.category === p.name; }).sort(function (a, b) { return a.daysLeft - b.daysLeft; });
      return {
        name: p.name, species: p.species,
        next: mine[0] ? { name: mine[0].name, daysLeft: mine[0].daysLeft } : null
      };
    })
  };
}

function cleanPet_(f) {
  var name = String(f.name || '').replace(/[<>]/g, '').trim().slice(0, 40);
  if (!name) throw new Error('Please enter a name');
  var bd = String(f.birthday || '').trim(), birthday = '';
  if (bd) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(bd)) throw new Error('Please pick a valid birthday');
    birthday = Utilities.parseDate(bd, tz_(), 'yyyy-MM-dd');
  }
  return {
    name: name,
    species: PET_SPECIES.indexOf(f.species) >= 0 ? f.species : 'Dog',
    breed: String(f.breed || '').trim().slice(0, 60),
    sex: PET_SEX.indexOf(f.sex) >= 0 ? f.sex : 'Unknown',
    birthday: birthday,
    chip: String(f.chip || '').replace(/[^\dA-Za-z\s-]/g, '').trim().slice(0, 30),
    vet: String(f.vet || '').trim().slice(0, 100),
    food: String(f.food || '').trim().slice(0, 100),
    notes: String(f.notes || '').trim().slice(0, 300)
  };
}

function addPet(f) {
  var c = cleanPet_(f);
  if (readPets_().some(function (p) { return p.name.toLowerCase() === c.name.toLowerCase(); })) throw new Error('You already have a pet with that name');
  var sh = petsSheet_();
  var id = 'p' + Date.now().toString(36);
  return withLock_(function () {
    sh.appendRow([id, c.name, c.species, c.breed, c.sex, c.birthday, c.chip, c.vet, c.food, c.notes, new Date()]);
    return { id: id };
  });
}

function updatePet(id, f) {
  var c = cleanPet_(f);
  var pets = readPets_();
  var p = pets.filter(function (x) { return x.id === id; })[0];
  if (!p) throw new Error('That pet was not found. Please refresh.');
  if (pets.some(function (x) { return x.id !== id && x.name.toLowerCase() === c.name.toLowerCase(); })) throw new Error('You already have a pet with that name');
  var sh = petsSheet_();
  return withLock_(function () {
    sh.getRange(p.row, 2, 1, 9).setValues([[c.name, c.species, c.breed, c.sex, c.birthday, c.chip, c.vet, c.food, c.notes]]);
    if (c.name !== p.name) renamePetReminders_(p.name, c.name);
    return true;
  });
}

// Keeps the care schedule attached to the pet when it is renamed
function renamePetReminders_(oldName, newName) {
  var sh = remindersSheet_();
  var n = sh.getLastRow();
  if (n < 2) return;
  var rng = sh.getRange(2, 1, n - 1, 3);
  var vals = rng.getValues();
  vals.forEach(function (r) {
    if (String(r[2]) === oldName) {
      r[2] = newName;
      if (String(r[0]).indexOf(oldName + ' ') === 0) r[0] = newName + String(r[0]).slice(oldName.length);
    }
  });
  rng.setValues(vals);
}

function deletePet(id) {
  var p = readPets_().filter(function (x) { return x.id === id; })[0];
  if (!p) throw new Error('That pet was not found. Please refresh.');
  var sh = petsSheet_(), wsh = petWeightsSheet_(), rsh = remindersSheet_();
  return withLock_(function () {
    sh.deleteRow(p.row);
    var wn = wsh.getLastRow();
    if (wn >= 2) {
      var rows = wsh.getRange(2, 1, wn - 1, 3).getValues();
      var keep = rows.filter(function (r) { return String(r[0]) !== id; });
      wsh.getRange(2, 1, rows.length, 3).clearContent();
      if (keep.length) wsh.getRange(2, 1, keep.length, 3).setValues(keep);
    }
    var rn = rsh.getLastRow();
    for (var r = rn; r >= 2; r--) {
      if (String(rsh.getRange(r, 3).getValue()) === p.name) rsh.deleteRow(r); // the pet's care schedule
    }
    return true;
  });
}

/* --- Weight --- */

function cleanPetWeight_(f) {
  var ds = String(f.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) throw new Error('Please pick a date');
  var kg = Math.round(parseFloat(String(f.kg).replace(',', '.')) * 100) / 100;
  if (!(kg > 0 && kg <= 200)) throw new Error('Please enter the weight in kg');
  return { date: Utilities.parseDate(ds, tz_(), 'yyyy-MM-dd'), kg: kg };
}

function checkPetWeightRow_(sh, row, old) {
  if (row < 2 || row > sh.getLastRow()) throw new Error('That entry has changed. Please refresh.');
  var v = sh.getRange(row, 1, 1, 3).getValues()[0];
  if (String(v[0]) !== old.petId || (v[1] instanceof Date ? dayStr_(v[1]) : '') !== old.date || Number(v[2]) !== Number(old.kg)) {
    throw new Error('That entry has changed. Please refresh.');
  }
}

function addPetWeight(petId, f) {
  var c = cleanPetWeight_(f);
  var sh = petWeightsSheet_();
  return withLock_(function () { sh.appendRow([petId, c.date, c.kg]); return true; });
}

function updatePetWeight(row, old, f) {
  var c = cleanPetWeight_(f);
  var sh = petWeightsSheet_();
  return withLock_(function () {
    checkPetWeightRow_(sh, row, old);
    sh.getRange(row, 2, 1, 2).setValues([[c.date, c.kg]]);
    return true;
  });
}

function deletePetWeight(row, old) {
  var sh = petWeightsSheet_();
  return withLock_(function () { checkPetWeightRow_(sh, row, old); sh.deleteRow(row); return true; });
}
