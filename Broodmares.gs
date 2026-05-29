function updateBroodmares() {
  const ss             = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet      = ss.getSheetByName("Herd Tracker");
  const broodmareSheet = ss.getSheetByName("Broodmares");

  if (!mainSheet || !broodmareSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheets not found!');
    return;
  }

  // 1. Read Herd Tracker (columns C–P)
  const mainData           = mainSheet.getRange("C2:P" + mainSheet.getLastRow()).getValues();
  const currentBroodmareIDs = [];
  const mainMap            = new Map();

  for (let i = 0; i < mainData.length; i++) {
    const id     = mainData[i][0] ? mainData[i][0].toString().trim() : "";
    const name   = mainData[i][1];
    const breed  = mainData[i][2];
    const status = mainData[i][13]; // Column P

    if (id !== "" && status && status.toString().toLowerCase().replace(/\s+/g, " ").trim().includes("broodmare")) {
      mainMap.set(id, { name, breed });
      currentBroodmareIDs.push(id);
    }
  }

  // 2. Check existing Broodmares sheet
  let lastBMRow = Math.max(broodmareSheet.getLastRow(), 2);
  let bmIDs     = broodmareSheet.getRange(2, 3, Math.max(1, lastBMRow - 1), 1)
                    .getValues().map(r => r[0].toString().trim());

  let removedCount = 0;
  let addedCount   = 0;

  // Remove rows no longer broodmares (iterate backwards to keep indices valid)
  for (let j = bmIDs.length - 1; j >= 0; j--) {
    const bmID = bmIDs[j];
    if (bmID === "") continue;
    if (!mainMap.has(bmID)) {
      broodmareSheet.deleteRow(j + 2);
      removedCount++;
    }
  }

  // 3. Add new or update existing
  const updatedLastBMRow = broodmareSheet.getLastRow();
  const updatedBMIDs     = updatedLastBMRow >= 2
    ? broodmareSheet.getRange(2, 3, updatedLastBMRow - 1, 1).getValues().map(r => r[0].toString().trim())
    : [];

  mainMap.forEach((data, id) => {
    const existingIndex = updatedBMIDs.indexOf(id);
    if (existingIndex === -1) {
      const nextRow = broodmareSheet.getLastRow() + 1;
      broodmareSheet.getRange(nextRow, 3, 1, 3).setValues([[id, data.name, data.breed]]);
      addedCount++;
    } else {
      broodmareSheet.getRange(existingIndex + 2, 4, 1, 2).setValues([[data.name, data.breed]]);
    }
  });

  // 4. Result
  SpreadsheetApp.getUi().alert('Done',
    `Sync complete:\n\n➕ Added: ${addedCount}\n🗑️ Removed: ${removedCount}`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}
