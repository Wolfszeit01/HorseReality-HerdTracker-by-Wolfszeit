/***********************
  SCORES NACHTRAGEN
 ***********************/

function openAddScoresModal() {
  try {
    const html = HtmlService.createTemplateFromFile('addScores')
      .evaluate()
       .setWidth(800)
    .setHeight(798)
      .setTitle('Add Scores');
    
    SpreadsheetApp.getUi().showModalDialog(html, ' '); 
  } catch (error) {
    SpreadsheetApp.getActive().toast('Error opening dialog: ' + error.message, 'Error', 5);
  }
}

function getHorseListForSearch() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const herdSheet = ss.getSheetByName('Herd');
    
    if (!herdSheet) {
      console.error("Sheet 'Herd' not found!");
      return [];
    }

    const lastRow = herdSheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = herdSheet.getRange(2, 3, lastRow - 1, 2).getValues();
    
    const horseList = data
      .filter(row => row[0] && row[0].toString().trim() !== "" && 
                     row[1] && row[1].toString().trim() !== "")
      .map(row => ({
        id: row[0].toString().trim(),
        name: row[1].toString().trim()
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    
    console.log("Horses found: " + horseList.length);
    return horseList;
    
  } catch (error) {
    console.error("Error in getHorseListForSearch:", error);
    return [];
  }
}

function saveBulkScores(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = (data.target === 'conf') ? 'Conf. Results' : 'Comp. Results';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

    const horseId = data.id.toString().trim();
    const col = findExistingInResultsSheet(sheet, horseId);

    if (!col) {
      throw new Error('Horse with ID ' + horseId + ' was not found in ' + sheetName + '.');
    }

    const dataStartRow = (data.target === 'conf') ? 8 : 7;
    const lastRow = sheet.getLastRow();

    // 1. Hole alle bereits existierenden Scores in dieser Spalte
    let existingScores = [];
    if (lastRow >= dataStartRow) {
      const existingRangeValues = sheet.getRange(dataStartRow, col, lastRow - dataStartRow + 1, 1).getValues();
      existingScores = existingRangeValues
        .flat()
        .filter(v => v !== "" && !isNaN(v))
        .map(v => Number(v));
    }

    // Ein Set für den schnellen Abgleich (existierende Werte + bereits verarbeitete neue Werte)
    const seenValues = new Set(existingScores);

    // 2. Rohdaten parsen
    const rawLines = typeof data.scores === 'string' 
      ? data.scores.split('\n') 
      : (Array.isArray(data.scores) ? data.scores : [data.scores]);

    const newUniqueScores = [];
    let duplicatesIgnored = 0;

    rawLines.forEach(line => {
      const trimmed = line.toString().trim();
      
      // Überspringe leere Zeilen, Datumsangaben, Ränge oder Text
      if (!trimmed || trimmed.includes('-') || trimmed.includes('#') || /[a-zA-Z]/.test(trimmed)) {
        return;
      }

      const num = parseFloat(trimmed.replace(',', '.'));
      if (!isNaN(num)) {
        // Prüfen, ob der Score schon existiert (weder in der Tabelle noch im aktuellen Batch)
        if (!seenValues.has(num)) {
          seenValues.add(num);
          newUniqueScores.push(num);
        } else {
          duplicatesIgnored++;
        }
      }
    });
    
    // 3. Nur die neuen, wirklich eindeutigen Scores unten anhängen
    let savedCount = 0;
    if (newUniqueScores.length > 0) {
      const startRow = findFirstEmptyRowInColumn(sheet, col, dataStartRow);
      const valuesToWrite = newUniqueScores.map(score => [score]);
      sheet.getRange(startRow, col, valuesToWrite.length, 1).setValues(valuesToWrite);
      savedCount = newUniqueScores.length;
    }
    
    // 4. Sicherheitshalber die Spalte noch mal komplett bereinigen (falls alter Müll/Duplikate drin waren)
    let cleanResult = cleanSingleHorse(horseId);
    refreshResultColumnStyle_(sheet, col);
    
    // 5. Manuelles Stats speichern
    let statsInfo = { updated: false, sheets: [], breeds: [] };
    try {
      statsInfo = updateHorseStats(horseId, data.target, data.manStats);
    } catch (statErr) {
      console.error("Warning: stats update failed: " + statErr.message);
    }
    
    return {
      success: true,
      message: `${savedCount} new scores saved (${duplicatesIgnored} duplicates ignored)`,
      duplicatesRemoved: cleanResult.removed || 0,
      statsUpdated: statsInfo.updated,
      statsSheets: statsInfo.sheets,
      horseBreed: statsInfo.breeds[0] || ''
    };
    
  } catch (error) {
    console.error("Error in saveBulkScores:", error);
    throw new Error("Error while saving: " + error.message);
  }
}

function findExistingInResultsSheet(sheet, horseId) {
  if (!sheet) return null;
  
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  for (let col = 0; col < lastCol; col++) {
    const cellValue = headerRow[col] ? headerRow[col].toString().trim() : '';
    if (cellValue === horseId.toString().trim()) {
      return col + 1;
    }
  }
  return null;
}

function findFirstEmptyRowInColumn(sheet, col, startRow) {
  const lastRow = sheet.getMaxRows();
  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0].toString().trim() === '') {
      return startRow + i;
    }
  }
  return lastRow + 1;
}

// ====== STATS UPDATE SYSTEM (MANUELLE EINGABE) ======

function updateHorseStats(horseId, targetType, manStats) {
  const statsInfo = { updated: false, sheets: [], breeds: [] };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const herdSheet = ss.getSheetByName('Herd');
    if (!herdSheet) return statsInfo;

    const lastRow = herdSheet.getLastRow();
    if (lastRow < 2) return statsInfo;

    const data = herdSheet.getRange(2, 3, lastRow - 1, 3).getValues();
    
    let horseBreed = "";
    let horseName = "";
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === horseId.toString().trim()) {
        horseBreed = data[i][2] ? data[i][2].toString().trim() : "";
        horseName = data[i][1] ? data[i][1].toString().trim() : "";
        break;
      }
    }

    if (!horseBreed) return statsInfo;

    statsInfo.breeds.push(horseBreed);

    const isIcelandic = horseBreed.toLowerCase().includes('isländer') || horseBreed.toLowerCase().includes('icelandic');
    const isKathiwari = horseBreed.toLowerCase().includes('kathiwari');

    const sheetsToUpdate = ["Stats"];
    
    if (isIcelandic) {
      sheetsToUpdate.push("ICE_Stats");
    } else if (isKathiwari) {
      sheetsToUpdate.push("KATH_Horse-Stats");
    }

    for (let i = 0; i < sheetsToUpdate.length; i++) {
      const success = updateSingleStatsSheet(ss, horseId, horseName, targetType, sheetsToUpdate[i], manStats);
      if (success) {
        statsInfo.updated = true;
        statsInfo.sheets.push(sheetsToUpdate[i]);
      }
    }

    console.log(`Stats for ${horseName} (${horseBreed}) updated: ${statsInfo.sheets.join(', ')}`);

  } catch (error) {
    console.error("Error updating stats:", error);
  }
  
  return statsInfo;
}

function updateSingleStatsSheet(ss, horseId, horseName, targetType, sheetName, manStats) {
  try {
    const statsSheet = ss.getSheetByName(sheetName);
    if (!statsSheet) return false;

    // Spalten-Mapping
    let mapping = null;
    if (sheetName === "Stats") {
      mapping = { dc: 50, "1st": 51, "2nd": 52, "3rd": 53, first: 59, second: 60, third: 61 };
    } else if (sheetName === "ICE_Stats") {
      mapping = { dc: 2, "1st": 3, "2nd": 4, "3rd": 5, first: 58, second: 60, third: 61 };
    } else if (sheetName === "KATH_Horse-Stats") {
      mapping = { dc: 2, "1st": 3, "2nd": 4, "3rd": 5, first: 58, second: 60, third: 61 };
    }

    if (!mapping) return false;

    // Finde die Zeile für das Pferd (ID in Spalte B)
    const statsLastRow = statsSheet.getLastRow();
    let statsRow = -1;
    
    for (let r = 2; r <= statsLastRow; r++) {
      const cellValue = statsSheet.getRange(r, 2).getValue();
      if (cellValue && cellValue.toString().trim() === horseId.toString().trim()) {
        statsRow = r;
        break;
      }
    }

    if (statsRow === -1) {
      statsRow = statsLastRow + 1;
      statsSheet.getRange(statsRow, 2).setValue(horseId);
      statsSheet.getRange(statsRow, 3).setValue(horseName);
    }

    // Manuelles Stats verwenden (wenn übergeben)
    if (manStats && Object.keys(manStats).length > 0) {
      if (manStats.dc !== undefined && manStats.dc !== "") statsSheet.getRange(statsRow, mapping.dc).setValue(parseFloat(manStats.dc));
      if (manStats["1st"] !== undefined && manStats["1st"] !== "") statsSheet.getRange(statsRow, mapping["1st"]).setValue(parseFloat(manStats["1st"]));
      if (manStats["2nd"] !== undefined && manStats["2nd"] !== "") statsSheet.getRange(statsRow, mapping["2nd"]).setValue(parseFloat(manStats["2nd"]));
      if (manStats["3rd"] !== undefined && manStats["3rd"] !== "") statsSheet.getRange(statsRow, mapping["3rd"]).setValue(parseFloat(manStats["3rd"]));
      if (manStats.first !== undefined && manStats.first !== "") statsSheet.getRange(statsRow, mapping.first).setValue(parseFloat(manStats.first));
      if (manStats.second !== undefined && manStats.second !== "") statsSheet.getRange(statsRow, mapping.second).setValue(parseFloat(manStats.second));
      if (manStats.third !== undefined && manStats.third !== "") statsSheet.getRange(statsRow, mapping.third).setValue(parseFloat(manStats.third));
    }

    return true;

  } catch (error) {
    console.error("Error in updateSingleStatsSheet:", error);
    return false;
  }
}

// ====== DUPLIKAT-CHECK & CLEAN ======

function checkSingleHorseDuplicates(horseId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ["Conf. Results", "Comp. Results"];
    let totalDuplicates = 0;
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = ss.getSheetByName(sheets[i]);
      if (!sheet) continue;
      
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      let targetCol = -1;
      for (let j = 0; j < headers.length; j++) {
        if (headers[j] && headers[j].toString().trim() === horseId.toString().trim()) {
          targetCol = j + 1;
          break;
        }
      }
      
      if (targetCol === -1) continue;
      
      const startRow = (sheets[i] === 'Conf. Results') ? 8 : 7;
      const lastRow = sheet.getLastRow();
      if (lastRow < startRow) continue;
      
      const values = sheet.getRange(startRow, targetCol, lastRow - startRow + 1, 1)
                         .getValues()
                         .flat()
                         .filter(v => v !== "" && !isNaN(v));
      
      if (values.length > 0) {
        const uniqueCount = [...new Set(values)].length;
        const duplicates = values.length - uniqueCount;
        totalDuplicates += duplicates;
      }
    }
    
    return {
      success: true,
      hasDuplicates: totalDuplicates > 0,
      duplicateCount: totalDuplicates
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function cleanSingleHorse(horseId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ["Conf. Results", "Comp. Results"];
    let totalRemoved = 0;
    
    for (let i = 0; i < sheets.length; i++) {
      const sheetName = sheets[i];
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      let targetCol = -1;
      for (let j = 0; j < headers.length; j++) {
        if (headers[j] && headers[j].toString().trim() === horseId.toString().trim()) {
          targetCol = j + 1;
          break;
        }
      }
      
      if (targetCol === -1) continue;
      
      const startRow = (sheetName === 'Conf. Results') ? 8 : 7;
      const lastRow = sheet.getLastRow();
      if (lastRow < startRow) continue;
      
      const numRows = lastRow - startRow + 1;
      const range = sheet.getRange(startRow, targetCol, numRows, 1);
      const allValues = range.getValues().flat();
      
      const numericValues = allValues
        .filter(v => v !== "" && !isNaN(v))
        .map(v => Number(v));
      
      if (numericValues.length === 0) continue;
      
      const uniqueValues = [...new Set(numericValues)].sort((a, b) => b - a);
      const removed = numericValues.length - uniqueValues.length;
      totalRemoved += removed;
      
      if (removed > 0) {
        range.clearContent();
        if (uniqueValues.length > 0) {
          const outputValues = uniqueValues.map(v => [v]);
          sheet.getRange(startRow, targetCol, outputValues.length, 1)
               .setValues(outputValues);
        }
      }
    }
    
    return {
      success: true,
      message: `${totalRemoved} duplicates removed`,
      removed: totalRemoved
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function cleanAllDuplicatesInSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');

    const lastCol = sheet.getLastColumn();
    if (lastCol < 2) return { success: false, message: 'No data found' };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let totalRemoved = 0;
    let horsesCleaned = 0;

    const startRow = (sheetName === 'Conf. Results') ? 8 : 7;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < startRow) {
      return { success: true, message: 'No scores to clean up', totalRemoved: 0 };
    }

    // Über jede Spalte iterieren (jedes Pferd)
    for (let col = 1; col <= lastCol; col++) {
      const horseId = headers[col - 1];
      if (!horseId || horseId.toString().trim() === '') continue;
      
      const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1)
                         .getValues()
                         .flat()
                         .filter(v => v !== "" && !isNaN(v));
      
      if (values.length === 0) continue;
      
      const uniqueCount = [...new Set(values)].length;
      const duplicates = values.length - uniqueCount;
      
      if (duplicates > 0) {
        const uniqueValues = [...new Set(values)].sort((a, b) => b - a);
        sheet.getRange(startRow, col, lastRow - startRow + 1, 1).clearContent();
        sheet.getRange(startRow, col, uniqueValues.length, 1)
             .setValues(uniqueValues.map(v => [v]));
        
        totalRemoved += duplicates;
        horsesCleaned++;
        console.log(`${horseId}: ${duplicates} duplicates removed`);
      }
    }

    return {
      success: true,
      message: `${totalRemoved} duplicates removed across ${horsesCleaned} horse columns`,
      totalRemoved: totalRemoved,
      horsesCleaned: horsesCleaned
    };

  } catch (error) {
    console.error("Error in cleanAllDuplicatesInSheet:", error);
    return { success: false, error: error.message };
  }
}