/**
 * Habits – JSON-API för fristående frontend (GitHub Pages).
 *
 * SETUP:
 * 1. Skapa ett nytt Google Sheet (eller återanvänd ett), kopiera dess ID hit.
 * 2. Tillägg -> Apps Script, klistra in denna kod.
 * 3. Byt TOKEN till en egen hemlig sträng (samma som i index.html).
 * 4. Deploy -> New deployment -> Web app
 *      Execute as: Me
 *      Who has access: Anyone      <-- viktigt!
 * 5. Kopiera /exec-URL:en till API_URL i index.html.
 *
 * Fliken "Habits" (loggen): Datum | Habit | Värde
 *   Värde = 1 (klar) / 0 (ej klar) för ja/nej, eller siffran för mätbara.
 * Habitlistan sparas i Script Properties, inte i Sheetet.
 */

var SHEET_ID = '13GlzRS4Aavvx5b4r4bFQYk07fZR5lVlBS5B00TwHPTY';
var SHEET_NAME = 'Habits';
var TOKEN = 'habits';
var DAYS_BACK = 90; // hur långt bak getCheckins hämtar

function doGet(e) {
  if (e.parameter.token !== TOKEN) return jsonOut_({ error: 'unauthorized' });
  var a = e.parameter.action;
  if (a === 'getHabits') return jsonOut_(getHabits());
  if (a === 'getCheckins') return jsonOut_(getCheckins());
  return jsonOut_({ error: 'unknown action' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut_({ error: 'bad request' }); }
  if (body.token !== TOKEN) return jsonOut_({ error: 'unauthorized' });

  var a = body.action;
  if (a === 'saveCheckins') return jsonOut_(saveCheckins(body.checkins));
  if (a === 'addHabit') return jsonOut_(addHabit(body.habit));
  if (a === 'removeHabit') return jsonOut_(removeHabit(body.name));
  return jsonOut_({ error: 'unknown action' });
}

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Datum', 'Habit', 'Värde']);
  // Tvinga Datum-kolumnen till text så Sheets inte auto-konverterar "2026-08-06"
  // till ett Date-objekt (vilket kan skifta en dag beroende på tidszon).
  sheet.getRange('A:A').setNumberFormat('@');
  return sheet;
}

function dateKey_(v) {
  if (v instanceof Date) {
    var tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v);
}

/**
 * Upsert: en rad per habit+datum. Uppdaterar befintlig rad om den finns,
 * annars läggs en ny till. Tomt värde tar bort raden (markeras som ej loggad).
 */
function saveCheckins(list) {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];

  // index: "datum|habit" -> radnummer i arket
  var index = {};
  for (var i = 0; i < data.length; i++) {
    index[dateKey_(data[i][0]) + '|' + data[i][1]] = i + 2;
  }

  var toAppend = [];
  var toClear = [];

  for (var j = 0; j < list.length; j++) {
    var c = list[j];
    var key = c.date + '|' + c.habit;
    var isEmpty = (c.value === '' || c.value === null || c.value === undefined);

    if (index[key]) {
      if (isEmpty) toClear.push(index[key]);
      else sheet.getRange(index[key], 3).setValue(c.value);
    } else if (!isEmpty) {
      toAppend.push([c.date, c.habit, c.value]);
    }
  }

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 3).setValues(toAppend);
  }
  // radera bakifrån så radnummer inte förskjuts
  toClear.sort(function (a, b) { return b - a; });
  toClear.forEach(function (r) { sheet.deleteRow(r); });

  return { ok: true };
}

// Returnerar { habitnamn: { "2026-08-02": 1, ... } }
function getCheckins() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - DAYS_BACK);
  var cutoffKey = Utilities.formatDate(cutoff, 'Europe/Stockholm', 'yyyy-MM-dd');

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var out = {};
  for (var i = 0; i < data.length; i++) {
    var dk = dateKey_(data[i][0]);
    var habit = data[i][1];
    if (!habit || dk < cutoffKey) continue;
    if (!out[habit]) out[habit] = {};
    out[habit][dk] = data[i][2];
  }
  return out;
}

// ---- Habitlista ----
function getHabits() {
  var props = PropertiesService.getUserProperties();
  var s = props.getProperty('HABITS');
  if (!s) return [];
  return JSON.parse(s);
}

function addHabit(habit) {
  var props = PropertiesService.getUserProperties();
  var list = JSON.parse(props.getProperty('HABITS') || '[]');
  var exists = false;
  for (var i = 0; i < list.length; i++) if (list[i].name === habit.name) exists = true;
  if (!exists) {
    list.push(habit);
    props.setProperty('HABITS', JSON.stringify(list));
  }
  return list;
}

function removeHabit(name) {
  var props = PropertiesService.getUserProperties();
  var list = JSON.parse(props.getProperty('HABITS') || '[]');
  var out = [];
  for (var i = 0; i < list.length; i++) if (list[i].name !== name) out.push(list[i]);
  props.setProperty('HABITS', JSON.stringify(out));
  return out;
}
