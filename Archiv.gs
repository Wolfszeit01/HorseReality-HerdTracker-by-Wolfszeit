/***********************
 * ARCHIVE SCRIPT — COMMUNITY VERSION
 * Archives horses with status: retired / sold / passed
 ***********************/

function archiveHorses() {
  const ss               = SpreadsheetApp.getActiveSpreadsheet();
  const herdTrackerSheet = ss.getSheetByName('Herd Tracker');
  if (!herdTrackerSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Herd Tracker" not found!');
    return;
  }

  const archiveStatuses = ['retired', 'sold', 'passed'];
  const herdData        = herdTrackerSheet.getDataRange().getValues();
  const statusColumn    = 1; // Column B
  const idColumnHerd    = 2; // Column C

  const idsToArchive = [];
  for (let i = 1; i < herdData.length; i++) {
    const status = herdData[i][statusColumn];
    const id     = herdData[i][idColumnHerd];
    if (archiveStatuses.includes(String(status).toLowerCase()) && id) {
      idsToArchive.push(String(id).trim());
    }
  }

  if (idsToArchive.length === 0) {
    SpreadsheetApp.getUi().alert('No horses found to archive.');
    return;
  }

  // Get or create Archive sheet
  let archiveSheet = ss.getSheetByName('Archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Archive');
    archiveSheet.getRange('A1').setValue('Source Sheet');
    archiveSheet.getRange('B1').setValue('Archive Date');
    archiveSheet.getRange('C1').setValue('Data Type');
  }

  let totalArchived = 0;
  const archiveReport = [];
  const archiveDate   = new Date();

  // Row-based sheets
  const normalSheets = [
    { name: 'Herd Tracker',    idColumn: 2 },
    { name: 'Horse Stats',     idColumn: 1 },
    { name: 'ICE_Horse Stats', idColumn: 1 },
    { name: 'KATH_Horse Stats',idColumn: 1 },
    { name: 'Colour Genetics', idColumn: 1 }
  ];
  normalSheets.forEach(config => {
    const archived = archiveFromNormalSheet(ss, config.name, idsToArchive, config.idColumn, archiveSheet, archiveDate);
    archiveReport.push((archived > 0 ? '✓' : '○') + ' ' + config.name + ': ' + archived + ' row(s)');
    totalArchived += archived;
  });

  // Column-based sheets
  ['Conf. Results', 'Comp. Results'].forEach(sheetName => {
    const archived = archiveFromResultsSheet(ss, sheetName, idsToArchive, archiveSheet, archiveDate);
    archiveReport.push((archived > 0 ? '✓' : '○') + ' ' + sheetName + ': ' + archived + ' column(s)');
    totalArchived += archived;
  });

  SpreadsheetApp.getUi().alert(
    'Archive complete: ' + totalArchived + ' entries\n' +
    idsToArchive.length + ' horse(s) with status Retired/Sold/Passed\n\n' +
    archiveReport.join('\n')
  );
}

function archiveFromNormalSheet(ss, sheetName, idsToArchive, idColumn, archiveSheet, archiveDate) {
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) return 0;

  const data = sourceSheet.getDataRange().getValues();
  let archivedCount = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowId = data[i][idColumn];
    if (!rowId || !idsToArchive.includes(String(rowId).trim())) continue;

    const sourceRange  = sourceSheet.getRange(i + 1, 1, 1, sourceSheet.getLastColumn());
    const rowData      = sourceRange.getValues()[0];
    const rowFormats   = sourceRange.getBackgrounds()[0];
    const rowFonts     = sourceRange.getFontColors()[0];
    const targetRow    = archiveSheet.getLastRow() + 1;

    archiveSheet.getRange(targetRow, 1).setValue(sheetName);
    archiveSheet.getRange(targetRow, 2).setValue(archiveDate);
    archiveSheet.getRange(targetRow, 3).setValue('Row Data');
    archiveSheet.getRange(targetRow, 4, 1, rowData.length)
      .setValues([rowData]).setBackgrounds([rowFormats]).setFontColors([rowFonts]);

    sourceSheet.deleteRow(i + 1);
    archivedCount++;
  }
  return archivedCount;
}

function archiveFromResultsSheet(ss, sheetName, idsToArchive, archiveSheet, archiveDate) {
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) return 0;

  const lastColumn  = sourceSheet.getLastColumn();
  const lastRow     = sourceSheet.getLastRow();
  const headerRow   = sourceSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let archivedCount = 0;

  for (let col = lastColumn - 1; col >= 0; col--) {
    const colId = headerRow[col];
    if (!colId || !idsToArchive.includes(String(colId).trim())) continue;

    const colData   = sourceSheet.getRange(1, col + 1, lastRow, 1).getValues();
    const colFmts   = sourceSheet.getRange(1, col + 1, lastRow, 1).getBackgrounds();
    const colFonts  = sourceSheet.getRange(1, col + 1, lastRow, 1).getFontColors();
    const rowData   = colData.map(r => r[0]);
    const rowFmts   = colFmts.map(r => r[0]);
    const rowFonts2 = colFonts.map(r => r[0]);
    const targetRow = archiveSheet.getLastRow() + 1;

    archiveSheet.getRange(targetRow, 1).setValue(sheetName);
    archiveSheet.getRange(targetRow, 2).setValue(archiveDate);
    archiveSheet.getRange(targetRow, 3).setValue('Column Data');
    archiveSheet.getRange(targetRow, 4, 1, rowData.length)
      .setValues([rowData]).setBackgrounds([rowFmts]).setFontColors([rowFonts2]);

    sourceSheet.deleteColumn(col + 1);
    archivedCount++;
  }
  return archivedCount;
}

function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'archiveHorses')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('archiveHorses').timeBased().atHour(2).everyDays(1).create();
  SpreadsheetApp.getUi().alert('Auto-archive enabled. Runs daily at 2:00 AM.');
}

function removeTrigger() {
  const removed = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'archiveHorses')
    .map(t => { ScriptApp.deleteTrigger(t); return true; }).length;

  SpreadsheetApp.getUi().alert(
    removed > 0 ? '✓ Auto-archive disabled.' : 'No active auto-archive trigger found.'
  );
}

function previewArchiveCandidates() {
  const ss               = SpreadsheetApp.getActiveSpreadsheet();
  const herdTrackerSheet = ss.getSheetByName('Herd Tracker');
  if (!herdTrackerSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Herd Tracker" not found!');
    return;
  }

  const archiveStatuses = ['Retired', 'Sold', 'Passed'];
  const herdData        = herdTrackerSheet.getDataRange().getValues();
  const candidates      = [];

  for (let i = 1; i < herdData.length; i++) {
    const status = herdData[i][1];
    const id     = herdData[i][2];
    const name   = herdData[i][3];
    if (archiveStatuses.includes(status)) candidates.push('  • ' + name + ' (' + status + ') — ID: ' + id);
  }

  if (candidates.length === 0) {
    SpreadsheetApp.getUi().alert('No horses found to archive.');
    return;
  }

  SpreadsheetApp.getUi().alert(candidates.length + ' horse(s) would be archived:\n\n' + candidates.join('\n'));
}
