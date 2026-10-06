/**
 * Archivierungs-Funktion für Pferde (V2)
 * Archiviert alle Daten zu einem Pferd in EINER Zeile im Archive-Sheet,
 * statt wie bisher pro Quell-Sheet eine eigene Zeile anzulegen.
 *
 * Ablauf: 1) Daten aus allen Sheets sammeln  2) ins Archiv schreiben
 *         3) erst danach aus den Quell-Sheets löschen
 * -> Wenn Schritt 2 fehlschlägt, wurde noch nichts gelöscht.
 */

function archiveHorses() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) { SpreadsheetApp.getUi().alert('Archiving is already running. Please try again shortly.'); return; }
  try { archiveHorsesLocked_(); } finally { lock.releaseLock(); }
}
function archiveAllHorses() {
  const ui = SpreadsheetApp.getUi();
  const herd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Herd');
  if (!herd) { ui.alert('Herd sheet not found.'); return; }
  const ids = [...new Set(herd.getDataRange().getValues().slice(1).map(row => String(row[2] || '').trim()).filter(Boolean))].sort();
  if (!ids.length) { ui.alert('There are no horses in Herd to archive.'); return; }
  const answer = ui.prompt('Archive All Horses',
    'Archive all ' + ids.length + ' horses in Herd, including active horses?\n\n' +
    'Uses the regular archive workflow: Herd, Stats, ICE_Stats, KATH_Stats, Colour Genetics and the two Results tabs. Other tabs, settings, branding, themes and logs are not cleared. Existing archive entries are retained.\n\n' +
    'Type ARCHIVE ALL to confirm.', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK || answer.getResponseText().trim() !== 'ARCHIVE ALL') return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) { ui.alert('Archiving is already running. Please try again shortly.'); return; }
  try { archiveHorsesLocked_(true, ids); } finally { lock.releaseLock(); }
}
function archiveHorsesLocked_(allHorses, expectedIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const herdTrackerSheet = ss.getSheetByName('Herd');

  if (!herdTrackerSheet) {
    ui.alert('Fehler: Sheet "Herd" nicht gefunden!');
    return;
  }

  const archiveStatuses = ['retired', 'sold', 'passed'];
  const statusColumn = 1;     // Spalte B
  const idColumnHerd = 2;     // Spalte C
  const nameColumnHerd = 3;   // Spalte D -- bitte prüfen/anpassen!

  const herdData = herdTrackerSheet.getDataRange().getValues();
  if (allHorses === true && expectedIds) {
    const current = [...new Set(herdData.slice(1).map(row => String(row[idColumnHerd] || '').trim()).filter(Boolean))].sort();
    if (JSON.stringify(current) !== JSON.stringify(expectedIds)) {
      ui.alert('The herd changed while the confirmation was open. Nothing was archived. Please start again.');
      return;
    }
  }
  let horsesToArchive = []; // [{id, name}]

  for (let i = 1; i < herdData.length; i++) {
    const status = herdData[i][statusColumn];
    const id = herdData[i][idColumnHerd];
    const name = herdData[i][nameColumnHerd];
    if ((allHorses === true || archiveStatuses.includes(status)) && id) {
      horsesToArchive.push({ id: String(id).trim(), name: name ? String(name).trim() : '' });
    }
  }

  // Recover interrupted clean-up only for archived horses no longer present in Herd.
  // Horses still active in Herd must never be removed on the basis of an old archive entry.
  const existingArchive = ss.getSheetByName('Archive');
  const archivedIds = archiveIds_(existingArchive);
  const herdIds = new Set(herdData.slice(1).map(row => String(row[idColumnHerd] || '').trim()));
  const resumeIds = archivedIds.filter(id => !herdIds.has(id));
  if (horsesToArchive.length === 0 && resumeIds.length === 0) {
    ui.alert('Keine Pferde zum Archivieren gefunden.');
    return;
  }

  const idsToArchive = [...new Set(horsesToArchive.map(h => h.id).concat(resumeIds))];

  // horseData: { "id123": { "Herd Tracker - Status": "retired", "Horse Stats - Height": 165, ... } }
  let horseData = {};
  horsesToArchive.forEach(h => { horseData[h.id] = {}; });

  const normalSheets = [
    { name: 'Herd', idColumn: 2 },
    { name: 'Stats', idColumn: 1 },
    { name: 'ICE_Stats', idColumn: 1 },
    { name: 'KATH_Stats', idColumn: 1 },
    { name: 'Colour Genetics', idColumn: 1 }
  ];

  const resultsSheets = ['Conf. Results', 'Comp. Results'];

  let masterFields = []; // feste Spaltenreihenfolge fürs Archiv

  // ===== Schritt 1: Sammeln (noch nichts verändern) =====
  normalSheets.forEach(cfg => {
    const fields = collectFromNormalSheet(ss, cfg.name, idsToArchive, cfg.idColumn, horseData);
    fields.forEach(f => { if (masterFields.indexOf(f) === -1) masterFields.push(f); });
  });

  resultsSheets.forEach(name => {
    const fields = collectFromResultsSheet(ss, name, idsToArchive, horseData);
    fields.forEach(f => { if (masterFields.indexOf(f) === -1) masterFields.push(f); });
  });

  // ===== Schritt 2: Ins Archiv schreiben (1 Zeile pro Pferd) =====
  let archiveSheet;
  if (horsesToArchive.length) {
  try {
    archiveSheet = getOrCreateArchiveSheet(ss, masterFields);
  } catch (e) {
    ui.alert('Abgebrochen, es wurde nichts gelöscht:\n\n' + e.message);
    return;
  }
  writeHorsesToArchive(archiveSheet, horsesToArchive, horseData, masterFields);
  SpreadsheetApp.flush();
  }

  // ===== Schritt 3: Aus den Quell-Sheets löschen =====
  let archiveReport = [];
  normalSheets.forEach(cfg => {
    const removed = deleteFromNormalSheet(ss, cfg.name, idsToArchive, cfg.idColumn);
    archiveReport.push((removed > 0 ? '✓ ' : '○ ') + cfg.name + ': ' + removed + ' Zeile(n) entfernt');
  });
  resultsSheets.forEach(name => {
    const removed = deleteFromResultsSheet(ss, name, idsToArchive);
    archiveReport.push((removed > 0 ? '✓ ' : '○ ') + name + ': ' + removed + ' Spalte(n) entfernt');
  });

  let message = horsesToArchive.length + ' Pferd(e) archiviert (je 1 Zeile im Archive-Sheet):\n\n';
  message += horsesToArchive.map(h => '  • ' + h.name + ' (ID: ' + h.id + ')').join('\n');
  if (resumeIds.length) message += '\nPreviously archived horses checked for remaining source data: ' + resumeIds.length;
  message += '\n\n' + archiveReport.join('\n');
  ui.alert(message);
}


/**
 * Liest ein "normales" Sheet (zeilenbasiert) und trägt passende Zeilen
 * in horseData ein. Gibt die verwendeten Feldnamen zurück.
 */
function collectFromNormalSheet(ss, sheetName, idsToArchive, idColumn, horseData) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) return [];

  const header = data[0];
  const fieldNames = header.map((h, colIdx) => {
    const label = h ? String(h).trim() : ('Spalte ' + (colIdx + 1));
    return sheetName + ' - ' + label;
  });

  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][idColumn];
    const rowIdStr = rowId ? String(rowId).trim() : '';
    if (rowIdStr && idsToArchive.indexOf(rowIdStr) !== -1) {
      if (!horseData[rowIdStr]) horseData[rowIdStr] = {};
      for (let col = 0; col < data[i].length; col++) {
        horseData[rowIdStr][fieldNames[col]] = data[i][col];
      }
    }
  }

  return fieldNames;
}


/**
 * Liest ein Results-Sheet (Pferde als Spalten, inkl. Spalte A).
 * Struktur: Zeile 1 = ID, Zeile 2 = Name, ab Zeile 3 = einzelne Einträge
 * (unterschiedlich viele pro Pferd, z.B. Turnierergebnisse über die Zeit).
 */
function collectFromResultsSheet(ss, sheetName, idsToArchive, horseData) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastColumn < 1 || lastRow < 1) return [];

  const idRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]; // Zeile 1 = ID
  const nameRow = lastRow >= 2 ? sheet.getRange(2, 1, 1, lastColumn).getValues()[0] : [];
  const entryCount = Math.max(0, lastRow - 2); // Einträge ab Zeile 3

  let fieldNames = [sheetName + ' - Name'];
  for (let e = 1; e <= entryCount; e++) {
    fieldNames.push(sheetName + ' - Eintrag ' + e);
  }

  for (let col = 0; col < lastColumn; col++) {
    const colId = idRow[col];
    const colIdStr = colId ? String(colId).trim() : '';
    if (colIdStr && idsToArchive.indexOf(colIdStr) !== -1) {
      if (!horseData[colIdStr]) horseData[colIdStr] = {};
      horseData[colIdStr][fieldNames[0]] = nameRow[col] || '';

      if (entryCount > 0) {
        const entryValues = sheet.getRange(3, col + 1, entryCount, 1).getValues();
        for (let e = 0; e < entryCount; e++) {
          horseData[colIdStr][fieldNames[e + 1]] = entryValues[e][0];
        }
      }
    }
  }

  return fieldNames;
}


/**
 * Holt das Archive-Sheet oder legt es neu an.
 * Wirft einen Fehler, wenn ein vorhandenes Archive-Sheet noch das
 * alte Format hat (Spalte A ≠ "Horse ID") -- dann lieber nicht überschreiben.
 */
function getOrCreateArchiveSheet(ss, masterFields) {
  const baseHeader = ['Horse ID', 'Name', 'Archive Date'];
  let sheet = ss.getSheetByName('Archive');

  if (!sheet) {
    sheet = ss.insertSheet('Archive');
    ensureArchiveGrid_(sheet, 1, baseHeader.length + masterFields.length);
    sheet.getRange(1, 1, 1, baseHeader.length + masterFields.length)
      .setValues([baseHeader.concat(masterFields)]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, baseHeader.length + masterFields.length)
      .setValues([baseHeader.concat(masterFields)]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const existingHeader = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (existingHeader[0] !== 'Horse ID') {
    throw new Error(
      'Das vorhandene "Archive"-Sheet hat noch das alte Format ' +
      '(Spalte A ist "' + existingHeader[0] + '", nicht "Horse ID").\n' +
      'Bitte erst umbenennen (z.B. "Archive alt") oder leeren, ' +
      'damit hier nichts überschrieben wird.'
    );
  }

  const existingFields = existingHeader.slice(baseHeader.length);
  let changed = false;
  masterFields.forEach(f => {
    if (existingFields.indexOf(f) === -1) {
      existingFields.push(f);
      changed = true;
    }
  });

  if (changed) {
    ensureArchiveGrid_(sheet, 1, baseHeader.length + existingFields.length);
    sheet.getRange(1, 1, 1, baseHeader.length + existingFields.length)
      .setValues([baseHeader.concat(existingFields)]);
  }

  return sheet;
}


/**
 * Schreibt eine Zeile pro Pferd ins Archive-Sheet, in der Spaltenreihenfolge
 * des aktuellen Headers.
 */
function writeHorsesToArchive(sheet, horsesToArchive, horseData, masterFields) {
  const archiveDate = new Date();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fieldOrder = header.slice(3); // nach Horse ID, Name, Archive Date

  const existingIds = new Set(archiveIds_(sheet));
  const rows = horsesToArchive.filter(h => !existingIds.has(h.id)).map(h => {
    const data = horseData[h.id] || {};
    const fieldValues = fieldOrder.map(f => (f in data ? data[f] : ''));
    return [h.id, h.name, archiveDate].concat(fieldValues);
  });

  const startRow = sheet.getLastRow() + 1;
  if (rows.length) {
    ensureArchiveGrid_(sheet, startRow + rows.length - 1, header.length);
    sheet.getRange(startRow, 1, rows.length, header.length).setValues(rows);
  }
  try { applyArchiveLayout_(sheet, getActiveTheme()); }
  catch (error) { Logger.log("Archive styling: " + error.message); }
}


function deleteFromNormalSheet(ss, sheetName, idsToArchive, idColumn) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let removed = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowId = data[i][idColumn];
    const rowIdStr = rowId ? String(rowId).trim() : '';
    if (rowIdStr && idsToArchive.indexOf(rowIdStr) !== -1) {
      if (i + 1 <= sheet.getFrozenRows()) throw new Error(sheetName + ': matching horse is in a frozen header row; no deletion performed.');
      if (sheet.getMaxRows() <= sheet.getFrozenRows() + 1) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  return removed;
}


function deleteFromResultsSheet(ss, sheetName, idsToArchive) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return 0;

  const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]; // Zeile 1 = ID
  let removed = 0;

  for (let col = lastColumn - 1; col >= 0; col--) { // jede Spalte inkl. A ist ein Pferd
    const colId = headerRow[col];
    const colIdStr = colId ? String(colId).trim() : '';
    if (colIdStr && idsToArchive.indexOf(colIdStr) !== -1) {
      if (col + 1 <= sheet.getFrozenColumns()) throw new Error(sheetName + ': matching horse is in a frozen header column; no deletion performed.');
      if (sheet.getMaxColumns() <= sheet.getFrozenColumns() + 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
      sheet.deleteColumn(col + 1);
      removed++;
    }
  }
  return removed;
}


/**
 * Richtet einen täglichen Trigger für automatische Archivierung ein
 */
function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'archiveHorses') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('archiveHorses')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Automatische Archivierung wurde aktiviert!\n\nDas Script läuft jetzt täglich um 2 Uhr nachts.');
}


/**
 * Hilfsfunktion: Preview (unverändert)
 */
function previewArchiveCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const herdTrackerSheet = ss.getSheetByName('Herd Tracker');

  if (!herdTrackerSheet) {
    SpreadsheetApp.getUi().alert('Fehler: Sheet "Herd Tracker" nicht gefunden!');
    return;
  }

  const archiveStatuses = ['Retired', 'Sold', 'Passed'];
  const herdData = herdTrackerSheet.getDataRange().getValues();

  let candidates = [];

  for (let i = 1; i < herdData.length; i++) {
    const status = herdData[i][1];
    const id = herdData[i][2];
    const name = herdData[i][3];

    if (archiveStatuses.indexOf(status) !== -1) {
      candidates.push('  • ' + name + ' (' + status + ') - ID: ' + id);
    }
  }

  if (candidates.length === 0) {
    SpreadsheetApp.getUi().alert('Keine Pferde zum Archivieren gefunden.');
    return;
  }

  let message = 'Folgende ' + candidates.length + ' Pferd(e) würden archiviert:\n\n';
  message += candidates.join('\n');

  SpreadsheetApp.getUi().alert(message);
}

function cleanUpOrphanedHorses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Die Master-IDs aus "Herd Tracker" holen
  const trackerSheet = ss.getSheetByName("Herd");
  const trackerData = trackerSheet.getRange("B2:B" + trackerSheet.getLastRow()).getValues();
  // Wir erstellen ein Set für blitzschnellen Abgleich (IDs als Strings speichern)
  const activeIDs = new Set(trackerData.map(row => row[0].toString().trim()));

  // 2. Die Blätter definieren, die bereinigt werden sollen
  // HINWEIS: Prüfe, ob "Comp result" exakt so geschrieben wird (Groß-/Kleinschreibung)
  const sheetsToClean = ["Conf. Results", "Comp result"]; 

  sheetsToClean.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.warn("Blatt nicht gefunden: " + sheetName);
      return;
    }

    // Wir gehen davon aus, dass die IDs in Zeile 1 stehen (laut deinem Screenshot)
    // Wir arbeiten uns von rechts nach links vor, damit sich die Indizes beim Löschen nicht verschieben
    const lastCol = sheet.getLastRow() > 0 ? sheet.getLastColumn() : 0;
    if (lastCol === 0) return;

    const idRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Von hinten nach vorne loopen
    for (let col = idRow.length - 1; col >= 0; col--) {
      let currentID = idRow[col].toString().trim();
      
      // Wenn die Zelle nicht leer ist UND die ID nicht im Herd Tracker existiert
      if (currentID !== "" && !activeIDs.has(currentID)) {
        sheet.deleteColumn(col + 1); // +1 weil Sheet-Indizes bei 1 beginnen
        console.log(`Gelöscht aus ${sheetName}: ID ${currentID} (Spalte ${col + 1})`);
      }
    }
  });
  
  SpreadsheetApp.getUi().alert("Bereinigung abgeschlossen!");
}
function archiveIds_(sheet) {
  if (!sheet || sheet.getLastRow()<2) return [];
  if (sheet.getRange(1,1).getValue() !== 'Horse ID') throw new Error('Archive header is not recognized; clean-up stopped.');
  return [...new Set(sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().map(([id])=>String(id || '').trim()).filter(Boolean))];
}
function ensureArchiveGrid_(sheet, rows, cols) {
  if (sheet.getMaxRows()<rows) sheet.insertRowsAfter(sheet.getMaxRows(),rows-sheet.getMaxRows());
  if (sheet.getMaxColumns()<cols) sheet.insertColumnsAfter(sheet.getMaxColumns(),cols-sheet.getMaxColumns());
}
