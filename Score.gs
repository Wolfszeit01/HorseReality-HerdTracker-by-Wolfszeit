/***********************
 * SCORE ENTRY BACKEND — COMMUNITY VERSION
 ***********************/

function openAddScoresModal() {
  try {
    const html = HtmlService.createTemplateFromFile('addScoresModalHtml')
      .evaluate()
      .setWidth(450)
      .setHeight(550)
      .setTitle('Score Entry');
    SpreadsheetApp.getUi().showModalDialog(html, ' ');
  } catch (error) {
    SpreadsheetApp.getActive().toast('❌ Error opening modal: ' + error.message, 'Error', 5);
  }
}

function getHorseListForSearch() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const herdSheet = ss.getSheetByName('Herd Tracker');
    if (!herdSheet) { console.error("Sheet 'Herd Tracker' not found!"); return []; }

    const lastRow = herdSheet.getLastRow();
    if (lastRow < 2) return [];

    return herdSheet.getRange(2, 3, lastRow - 1, 2).getValues()
      .filter(row => row[0] && row[0].toString().trim() !== "" &&
                     row[1] && row[1].toString().trim() !== "")
      .map(row => ({ id: row[0].toString().trim(), name: row[1].toString().trim() }))
      .sort((a, b) => a.name.localeCompare(b.name));

  } catch (error) {
    console.error("Error in getHorseListForSearch:", error);
    return [];
  }
}

function saveBulkScores(data) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = (data.target === 'conf') ? 'Conf. Results' : 'Comp. Results';
    const sheet     = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

    const horseId = data.id.toString().trim();
    const col     = findExistingInResultsSheet(sheet, horseId);
    if (!col) throw new Error('Horse with ID ' + horseId + ' not found in ' + sheetName + '.');

    const startRow   = findFirstEmptyRowInColumn(sheet, col, 6);
    const scoresArray = data.scores
      .map(score => parseFloat(score.toString().replace(',', '.')))
      .filter(score => !isNaN(score));

    if (scoresArray.length === 0) throw new Error('No valid scores found.');

    sheet.getRange(startRow, col, scoresArray.length, 1).setValues(scoresArray.map(s => [s]));

    return {
      success: true,
      message: scoresArray.length + ' scores saved to ' + sheetName + ' (rows ' + startRow + '–' + (startRow + scoresArray.length - 1) + ')'
    };

  } catch (error) {
    console.error("Error in saveBulkScores:", error);
    throw new Error("Save error: " + error.message);
  }
}

function findExistingInResultsSheet(sheet, horseId) {
  if (!sheet) return null;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let col = 0; col < lastCol; col++) {
    if ((headerRow[col] || '').toString().trim() === horseId) return col + 1;
  }
  return null;
}

function findFirstEmptyRowInColumn(sheet, col, startRow) {
  const lastRow = sheet.getMaxRows();
  const values  = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0].toString().trim() === '') return startRow + i;
  }
  return lastRow + 1;
}
