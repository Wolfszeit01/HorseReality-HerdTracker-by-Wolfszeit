/**
 * ═══════════════════════════════════════════════════════════════
 * ZENTRALER onEdit-DISPATCHER
 * Ersetzt: onEdit aus CompTeam.gs, BreedingTab.gs (Broodmare-Teil)
 * StylingEngine.gs onEdit bleibt SEPARAT (umbenennen s.u.)
 * ═══════════════════════════════════════════════════════════════
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // ─── Styling (separate Datei, umbenannt) ───
  if (typeof onEdit_Styling === "function") onEdit_Styling(e);

  // ─── HERD: Spalte AA (27) → Comp. Team Sync ───
  if (sheetName === "Herd" && col === 27 && row >= 2 && e.value) {
    onEdit_CompTeam(sheet);
    return;
  }

  // ─── HERD: Spalte Q (17) → Broodmare Sync (Hinzufügen UND Entfernen) ───
  if (sheetName === "Herd" && col === 17 && row >= 2) {
    onEdit_Broodmare(sheet);
    return;
  }

  // ─── HERD: Spalte F (6) → Hengst-Dropdown ───
  if (sheetName === "Herd" && col === 6 && row >= 2) {
    updateStallionFromCache("herd");
    return;
  }

  // ─── OUTSIDE STUDS: Spalte C (3) → Hengst-Dropdown ───
  if (sheetName === "Outside Studs" && col === 3 && row >= 2) {
    updateStallionFromCache("outside");
    return;
  }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * COMP. TEAM SYNC
 * ═══════════════════════════════════════════════════════════════
 */
function onEdit_CompTeam(mainSheet) {
  const triggerLock = PropertiesService.getScriptProperties().getProperty('syncInProgress');
  if (triggerLock === 'true') {
    console.log("CompTeam-Sync bereits aktiv, überspringe...");
    return;
  }

  try {
    PropertiesService.getScriptProperties().setProperty('syncInProgress', 'true');
    syncCompTeamFromTrigger(mainSheet);
  } finally {
    PropertiesService.getScriptProperties().setProperty('syncInProgress', 'false');
  }
}

/**
 * HILFSFUNKTION: Comp. Team Synchronisation
 * Zeile 1: Last sync + Zeitstempel
 * Zeile 2: Spaltenüberschriften
 * Zeile 3+: Daten
 */
function syncCompTeamFromTrigger(mainSheet) {
  const ss = mainSheet.getParent();
  const compTeamSheet = ss.getSheetByName("Comp. Team");

  if (!compTeamSheet) {
    SpreadsheetApp.getUi().alert('Fehler: Tabelle "Comp. Team" nicht gefunden!');
    return;
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return;

  const allData = mainSheet.getRange(2, 3, lastRow - 1, 2).getValues(); // C & D (ID, Name)
  const statuses = mainSheet.getRange("AA2:AA" + lastRow).getValues(); // AA

  const compTeamData = [];

  for (let i = 0; i < statuses.length; i++) {
    const status = statuses[i][0];
    const rowData = allData[i];
    const id = rowData[0];
    const name = rowData[1];

    if ((status === true || status === "TRUE" || status === 1)) {
      if (id && id.toString().trim() !== "" && name && name.toString().trim() !== "") {
        compTeamData.push([id.toString().trim(), name.toString().trim()]);
      }
    }
  }

  const lastSyncTime = new Date();
  const horseCount = compTeamData.length;

  compTeamSheet.getRange("B1").setValue("Last sync:");
  const formattedTime = Utilities.formatDate(lastSyncTime, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  compTeamSheet.getRange("C1").setValue(formattedTime + " (" + horseCount + ")");
  compTeamSheet.getRange("B2").setValue("ID");
  const lastCompRow = compTeamSheet.getLastRow();
  if (lastCompRow >= 3) compTeamSheet.getRange(3, 2, lastCompRow - 2, 2).clearContent();

  if (horseCount === 0) {
    console.log("Sync: Keine Competition-Pferde gefunden, Liste geleert.");
    return;
  }

  compTeamData.sort((a, b) => (a[0] || "").localeCompare(b[0] || ""));
  compTeamSheet.getRange(3, 2, horseCount, 2).setValues(compTeamData);

  console.log(`CompTeam-Sync abgeschlossen: ${horseCount} Pferde @ ${formattedTime}`);
}

/**
 * ═══════════════════════════════════════════════════════════════
 * BROODMARE SYNC
 * ═══════════════════════════════════════════════════════════════
 */
function onEdit_Broodmare(mainSheet) {
  const lockKey = 'broodmareSyncInProgress';
  const lockValue = PropertiesService.getUserProperties().getProperty(lockKey);
  if (lockValue === 'true') {
    console.log("⚠️ Lock hängend, wird zurückgesetzt...");
    PropertiesService.getUserProperties().deleteProperty(lockKey);
    Utilities.sleep(200);
  }

  try {
    PropertiesService.getUserProperties().setProperty(lockKey, 'true');
    syncBroodmaresFromTrigger(mainSheet);
  } finally {
    PropertiesService.getUserProperties().deleteProperty(lockKey);
    console.log("✓ Lock freigegeben nach Sync");
  }
}

/**
 * HILFSFUNKTION: Broodmares Synchronisation
 * Datenbereich: C2:Q = 15 Spalten (C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q)
 * Läuft bei JEDER Änderung in Spalte Q — Hinzufügen UND Entfernen.
 */
function syncBroodmaresFromTrigger(mainSheet) {
  const ss = mainSheet.getParent();
  const broodmareSheet = ss.getSheetByName("Broodmares");

  if (!broodmareSheet) {
    console.error('Fehler: Tabelle "Broodmares" nicht gefunden!');
    return;
  }

  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) {
    console.log("Keine Daten in Herd Sheet");
    return;
  }

  // C2:Q = 15 Spalten (C=Spalte 3, Q=Spalte 17)
  const mainData = mainSheet.getRange(2, 3, lastRow - 1, 15).getValues();
  const mainMap = new Map();

  let addedCount = 0;
  let removedCount = 0;
  let updatedCount = 0;

  for (let i = 0; i < mainData.length; i++) {
    const id = mainData[i][0] ? String(mainData[i][0]).trim() : "";
    const name = mainData[i][1];
    const breed = mainData[i][2];
    const status = mainData[i][14]; // Q = Index 14

    if (id !== "" && status && String(status).toLowerCase().includes("broodmare")) {
      mainMap.set(id, { name: name, breed: breed });
    }
  }

  // Bestehende IDs in "Broodmares" lesen
  let lastBMRow = broodmareSheet.getLastRow();
  if (lastBMRow < 2) lastBMRow = 2;

  let bmIDs = [];
  if (lastBMRow >= 2) {
    bmIDs = broodmareSheet.getRange(2, 3, lastBMRow - 1, 1).getValues().map(r => String(r[0]).trim());
  }

  // 1. Entfernungen: IDs in Broodmares, aber nicht mehr in Herd mit Broodmare-Status
  for (let j = bmIDs.length - 1; j >= 0; j--) {
    const bmID = bmIDs[j];
    if (bmID === "") continue;

    if (!mainMap.has(bmID)) {
      broodmareSheet.deleteRow(j + 2);
      removedCount++;
      console.log(`🗑️ Entfernt: ${bmID}`);
    }
  }

  // 2. Neue Einlesen nach Löschung
  let updatedLastBMRow = broodmareSheet.getLastRow();
  let updatedBMIDs = [];
  if (updatedLastBMRow >= 2) {
    updatedBMIDs = broodmareSheet.getRange(2, 3, updatedLastBMRow - 1, 1).getValues().map(r => String(r[0]).trim());
  }

  // 3. Hinzufügen/Update
mainMap.forEach((data, id) => {
  const existingIndex = updatedBMIDs.indexOf(id);
  if (existingIndex === -1) {
    const nextRow = _findFreeBroodmareRow(broodmareSheet);
    broodmareSheet.getRange(nextRow, 3, 1, 3).setValues([[id, data.name, data.breed]]);
    updatedBMIDs[nextRow - 2] = id;
    addedCount++;
    console.log(`➕ Hinzugefügt: ${id}`);
  } else {
    broodmareSheet.getRange(existingIndex + 2, 4, 1, 2).setValues([[data.name, data.breed]]);
    updatedCount++;
  }
});

  const lastSyncTime = new Date();
  const formattedTime = Utilities.formatDate(lastSyncTime, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
  broodmareSheet.getRange("B1").setValue("Confirmed");
  const syncHeader = broodmareSheet.getRange("B1");
  const previousNote = (syncHeader.getNote() || "").replace(/\n?Last sync:[^\n]*/gi, "");
  syncHeader.setNote((previousNote ? previousNote + "\n" : "") + `Last sync: ${formattedTime} (${mainMap.size} mares)`);

  console.log(`Broodmare Sync abgeschlossen: ➕${addedCount} ➖${removedCount} 🔄${updatedCount}`);
}

/**
 * ═══════════════════════════════════════════════════════════════
 * STALLION DROPDOWN
 * ═══════════════════════════════════════════════════════════════
 */
function updateStallionFromCache(source) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const broodmareSheet = ss.getSheetByName("Broodmares");

  if (!broodmareSheet) return;

  let herdStallions = [];
  let outsideStallions = [];

  if (source === "herd") {
    const herdSheet = ss.getSheetByName("Herd");
    if (herdSheet) {
      herdStallions = scanHerdTrackerStallions(herdSheet);
      PropertiesService.getScriptProperties().setProperty('cachedHerdStallions', JSON.stringify(herdStallions));
      PropertiesService.getScriptProperties().setProperty('lastHerdUpdate', new Date().toISOString());
    }
    const cachedOutside = PropertiesService.getScriptProperties().getProperty('cachedOutsideStallions');
    if (cachedOutside) outsideStallions = JSON.parse(cachedOutside);

  } else if (source === "outside") {
    const outsideSheet = ss.getSheetByName("Outside Studs");
    if (outsideSheet) {
      outsideStallions = scanOutsideStuds(outsideSheet);
      PropertiesService.getScriptProperties().setProperty('cachedOutsideStallions', JSON.stringify(outsideStallions));
      PropertiesService.getScriptProperties().setProperty('lastOutsideUpdate', new Date().toISOString());
    }
    const cachedHerd = PropertiesService.getScriptProperties().getProperty('cachedHerdStallions');
    if (cachedHerd) herdStallions = JSON.parse(cachedHerd);
  }

  const allStallions = [...new Set([...herdStallions, ...outsideStallions])];
  allStallions.sort();

  updateDropdown(allStallions, false);
}

function scanHerdTrackerStallions(herdSheet) {
  const stallions = [];
  const lastRow = herdSheet.getLastRow();

  if (lastRow < 2) return stallions;

  const names = herdSheet.getRange("D2:D" + lastRow).getValues();
  const genders = herdSheet.getRange("F2:F" + lastRow).getValues();

  for (let i = 0; i < names.length; i++) {
    const name = names[i][0];
    const gender = genders[i][0]?.toString().toLowerCase().trim() || "";
    const validGenders = ["stallion", "hengst", "male", "männlich", "colt", "wallach"];

    if (name && validGenders.includes(gender)) {
      const horseName = String(name).trim();
      if (!stallions.includes(horseName)) {
        stallions.push(horseName);
      }
    }
  }

  return stallions;
}

function scanOutsideStuds(outsideSheet) {
  const stallions = [];
  const lastRow = outsideSheet.getLastRow();

  if (lastRow < 2) return stallions;

  const names = outsideSheet.getRange("C2:C" + lastRow).getValues();

  for (let i = 0; i < names.length; i++) {
    const name = names[i][0];
    if (name && String(name).trim() !== "") {
      const horseName = String(name).trim();
      if (!stallions.includes(horseName)) {
        stallions.push(horseName);
      }
    }
  }

  return stallions;
}

function updateDropdown(stallions, showAlert) {
  const broodmareSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Broodmares");
  if (!broodmareSheet) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(stallions, true)
    .setAllowInvalid(false)
    .setHelpText('Wähle einen Hengst aus der Liste')
    .build();

  broodmareSheet.getRange("M2:M").setDataValidation(rule);

  if (broodmareSheet.getRange("M1").getValue() === "") {
    broodmareSheet.getRange("M1").setValue("Hengst");
  }

 if (showAlert) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Hengstdropdowns aktualisiert! Anzahl Hengste: ${stallions.length}`,
    'Fertig', 5
  );
}

  console.log(`Stallion Update: ${stallions.length} Hengste`);
}

function updateStallionDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName("Herd");
  const outsideSheet = ss.getSheetByName("Outside Studs");

  let herdStallions = herdSheet ? scanHerdTrackerStallions(herdSheet) : [];
  let outsideStallions = outsideSheet ? scanOutsideStuds(outsideSheet) : [];

  PropertiesService.getScriptProperties().setProperty('cachedHerdStallions', JSON.stringify(herdStallions));
  PropertiesService.getScriptProperties().setProperty('cachedOutsideStallions', JSON.stringify(outsideStallions));

  const allStallions = [...new Set([...herdStallions, ...outsideStallions])];
  allStallions.sort();

  updateDropdown(allStallions, true);
}

function _findFreeBroodmareRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const ids = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === "" || ids[i][0] === null) return i + 2;
  }
  return lastRow + 1;
}

/**
 * ═══════════════════════════════════════════════════════════════
 * MANUELLE TEST-/FALLBACK-FUNKTIONEN
 * ═══════════════════════════════════════════════════════════════
 */
function testBroodmareSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName("Herd");
  if (herdSheet) {
    syncBroodmaresFromTrigger(herdSheet);
    SpreadsheetApp.getUi().alert("Broodmare-Sync abgeschlossen!");
  }
}

function testCompTeamSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName("Herd");
  if (herdSheet) {
    syncCompTeamFromTrigger(herdSheet);
    SpreadsheetApp.getUi().alert("CompTeam-Sync abgeschlossen!");
  }
}

function resetBroodmareLock() {
  PropertiesService.getUserProperties().deleteProperty('broodmareSyncInProgress');
  console.log("✓ Lock zurückgesetzt");
}

function resetCompTeamLock() {
  PropertiesService.getScriptProperties().setProperty('syncInProgress', 'false');
  console.log("✓ Lock zurückgesetzt");
}


