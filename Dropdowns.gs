/***********************
 * UPDATE STALLION DROPDOWN
 * Updates dropdown in Broodmares column M with all stallions from:
 * - Herd Tracker (column D, filtered by gender)
 * - Outside Studs (column C, all entries)
 ***********************/
function updateStallionDropdown() {
  const ss               = SpreadsheetApp.getActiveSpreadsheet();
  const herdTrackerSheet = ss.getSheetByName("Herd Tracker");
  const broodmareSheet   = ss.getSheetByName("Broodmares");
  const outsideStudSheet = ss.getSheetByName("Outside Studs");

  if (!herdTrackerSheet || !broodmareSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found!');
    return;
  }

  const stallions = [];

  // 1. Stallions from Herd Tracker
  const lastRowHerd = herdTrackerSheet.getLastRow();
  if (lastRowHerd >= 2) {
    const names   = herdTrackerSheet.getRange("D2:D" + lastRowHerd).getValues();
    const genders = herdTrackerSheet.getRange("F2:F" + lastRowHerd).getValues();

    for (let i = 0; i < names.length; i++) {
      const name      = names[i][0];
      const genderStr = genders[i][0] ? genders[i][0].toString().toLowerCase().trim() : "";
      if (name && name.toString().trim() !== "") {
        if (genderStr === "stallion" || genderStr === "male" || genderStr === "colt") {
          const horseName = name.toString().trim();
          if (!stallions.includes(horseName)) stallions.push(horseName);
        }
      }
    }
  }

  // 2. Stallions from Outside Studs (column C)
  if (outsideStudSheet) {
    const lastRowStud  = outsideStudSheet.getLastRow();
    if (lastRowStud >= 2) {
      const outsideNames = outsideStudSheet.getRange("C2:C" + lastRowStud).getValues();
      for (let i = 0; i < outsideNames.length; i++) {
        const name = outsideNames[i][0];
        if (name && name.toString().trim() !== "") {
          const horseName = name.toString().trim();
          if (!stallions.includes(horseName)) stallions.push(horseName);
        }
      }
    }
  }

  if (stallions.length === 0) {
    SpreadsheetApp.getUi().alert('No stallions found',
      'No stallions found in Herd Tracker or Outside Studs.\n\n' +
      'Valid gender values in Herd Tracker column F: Stallion, Male, Colt\n' +
      'Outside Studs: check names in column C.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  stallions.sort();

  // 3. Apply dropdown validation to Broodmares column M
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(stallions, true)
    .setAllowInvalid(false)
    .setHelpText('Select a stallion from the list')
    .build();

  broodmareSheet.getRange("M2:M").setDataValidation(rule);

  if (!broodmareSheet.getRange("M1").getValue()) {
    broodmareSheet.getRange("M1").setValue("Stallion");
  }

  SpreadsheetApp.getUi().alert('Done',
    `✅ Stallion dropdown updated!\n\nTotal stallions: ${stallions.length}`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}
