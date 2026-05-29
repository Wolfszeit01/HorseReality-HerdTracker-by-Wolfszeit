/***********************
 * SYNC COMPETITION TEAM
 * Copies horses with TRUE in column Z to Comp. Team columns B & C
 ***********************/
function syncCompTeam() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet     = ss.getSheetByName("Herd Tracker");
  const compTeamSheet = ss.getSheetByName("Comp. Team");

  if (!mainSheet || !compTeamSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheets not found!\n\nRequired:\n1. "Herd Tracker"\n2. "Comp. Team"');
    return;
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found in Herd Tracker.');
    return;
  }

  // Read IDs (col C), Names (col D), Comp status (col Z)
  const ids      = mainSheet.getRange("C2:C" + lastRow).getValues();
  const names    = mainSheet.getRange("D2:D" + lastRow).getValues();
  const statuses = mainSheet.getRange("Z2:Z" + lastRow).getValues();

  const compTeamData = [];
  for (let i = 0; i < ids.length; i++) {
    const id     = ids[i][0];
    const name   = names[i][0];
    const status = statuses[i][0];
    if ((status === true || status === "TRUE" || status === "true" || status === 1)
        && id && id.toString().trim() !== ""
        && name && name.toString().trim() !== "") {
      compTeamData.push({ id: id.toString().trim(), name: name.toString().trim() });
    }
  }

  if (compTeamData.length === 0) {
    SpreadsheetApp.getUi().alert('No competition horses found',
      'No horses with TRUE in column Z.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Set headers if missing
  if (!compTeamSheet.getRange("B1").getValue()) compTeamSheet.getRange("B1").setValue("ID");
  if (!compTeamSheet.getRange("C1").getValue()) compTeamSheet.getRange("C1").setValue("Name");

  // Clear old data
  const lastCompRow = compTeamSheet.getLastRow();
  if (lastCompRow > 1) compTeamSheet.getRange(2, 2, lastCompRow - 1, 2).clearContent();

  // Write sorted data
  const outputData = compTeamData
    .map(h => [h.id, h.name])
    .sort((a, b) => (a[0] || "").localeCompare(b[0] || ""));

  compTeamSheet.getRange(2, 2, outputData.length, 2).setValues(outputData);

  SpreadsheetApp.getUi().alert('Done',
    `Competition Team synced!\n\n${compTeamData.length} horses written to Comp. Team.`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}
