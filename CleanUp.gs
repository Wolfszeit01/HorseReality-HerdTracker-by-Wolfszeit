/***********************
 * CLEAN UP ORPHANED HORSE RESULTS
 * Removes columns from result sheets for horses no longer in Herd Tracker
 ***********************/
function cleanUpOrphanedHorses() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName("Herd Tracker");
  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Herd Tracker" not found!');
    return;
  }

  const trackerData = trackerSheet.getRange("C2:C" + trackerSheet.getLastRow()).getValues();
  const activeIDs   = new Set(trackerData.map(row => row[0].toString().trim()));

  const sheetsToClean = ["Conf. Results", "Comp. Results"];

  sheetsToClean.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { console.warn("Sheet not found: " + sheetName); return; }

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const idRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let col = idRow.length - 1; col >= 0; col--) {
      const currentID = idRow[col].toString().trim();
      if (currentID !== "" && !activeIDs.has(currentID)) {
        sheet.deleteColumn(col + 1);
      }
    }
  });

  SpreadsheetApp.getUi().alert('Cleanup complete!');
}
