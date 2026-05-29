/***********************
 * DUPLICATE CLEANER BACKEND — COMMUNITY VERSION
 ***********************/

function openDuplicateCleanerModal() {
  const html = HtmlService.createHtmlOutputFromFile('duplicateCleanerModalHtml')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Duplicate Cleaner');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getHorsesForDuplicateCleaner() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    let herdSheet   = ss.getSheetByName("Herd Tracker");

    if (!herdSheet) {
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const name = sheet.getName().toLowerCase();
        if (name.includes("herd") || name.includes("tracker")) { herdSheet = sheet; break; }
      }
    }

    if (!herdSheet) return { success: false, horses: [], error: "Sheet 'Herd Tracker' not found." };

    const lastRow = herdSheet.getLastRow();
    if (lastRow < 2) return { success: true, horses: [] };

    const horses = herdSheet.getRange(2, 3, lastRow - 1, 2).getValues()
      .filter(row => row[0] && row[0].toString().trim() !== "")
      .map(row => ({ id: row[0].toString().trim(), name: (row[1] || "Unknown").toString().trim() }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, horses, count: horses.length };

  } catch (error) {
    console.error("Error in getHorsesForDuplicateCleaner:", error.message);
    return { success: false, horses: [], error: "Error loading horses: " + error.message };
  }
}

function checkSingleHorseDuplicates(horseId) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ["Conf. Results", "Comp. Results"];
    let report   = `Duplicate check for horse: ${horseId}\n${'─'.repeat(40)}\n\n`;
    let totalDuplicates = 0;
    let totalValues     = 0;

    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) { report += `${sheetName}: Sheet not found\n`; return; }

      const lastCol   = sheet.getLastColumn();
      const headers   = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let targetCol   = -1;
      for (let i = 0; i < headers.length; i++) {
        if ((headers[i] || '').toString().trim() === horseId.toString().trim()) { targetCol = i + 1; break; }
      }

      if (targetCol === -1) { report += `${sheetName}: Horse not found\n`; return; }

      const startRow = 6;
      const lastRow  = sheet.getLastRow();
      if (lastRow < startRow) { report += `${sheetName}: No data\n`; return; }

      const values = sheet.getRange(startRow, targetCol, lastRow - startRow + 1, 1)
        .getValues().flat().filter(v => v !== "" && !isNaN(v));

      totalValues += values.length;
      const uniqueCount  = [...new Set(values)].length;
      const duplicates   = values.length - uniqueCount;
      totalDuplicates   += duplicates;

      report += `${sheetName} (col ${targetCol}):\n`;
      report += `  ${values.length} values · ${uniqueCount} unique · ${duplicates} duplicates\n`;

      if (duplicates > 0) {
        const valueMap = {};
        values.forEach(v => { valueMap[v] = (valueMap[v] || 0) + 1; });
        const dups = Object.entries(valueMap).filter(([,c]) => c > 1).map(([v,c]) => `${v} (${c}x)`);
        if (dups.length) report += `  Duplicates: ${dups.join(', ')}\n`;
      }
      report += '\n';
    });

    report += `${'─'.repeat(40)}\nTotal: ${totalValues} values · ${totalDuplicates} duplicates`;

    return { success: true, report, hasDuplicates: totalDuplicates > 0, duplicateCount: totalDuplicates };

  } catch (error) {
    return { success: false, error: error.message };
  }
}

function cleanSingleHorse(horseId) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheets  = ["Conf. Results", "Comp. Results"];
    let results   = [];
    let totalRemoved = 0;

    for (const sheetName of sheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) { results.push(`${sheetName}: Sheet not found`); continue; }

      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      let targetCol = -1;
      for (let j = 0; j < headers.length; j++) {
        if ((headers[j] || '').toString().trim() === horseId.toString().trim()) { targetCol = j + 1; break; }
      }

      if (targetCol === -1) { results.push(`${sheetName}: Horse not found`); continue; }

      const startRow    = 6;
      const lastRow     = sheet.getLastRow();
      if (lastRow < startRow) { results.push(`${sheetName}: No data`); continue; }

      const range       = sheet.getRange(startRow, targetCol, lastRow - startRow + 1, 1);
      const allValues   = range.getValues().flat().filter(v => v !== "" && !isNaN(v)).map(Number);
      if (allValues.length === 0) { results.push(`${sheetName}: No numeric values`); continue; }

      const uniqueValues  = [...new Set(allValues)].sort((a, b) => b - a);
      const removed       = allValues.length - uniqueValues.length;
      totalRemoved       += removed;

      if (removed > 0) {
        range.clearContent();
        sheet.getRange(startRow, targetCol, uniqueValues.length, 1).setValues(uniqueValues.map(v => [v]));
        results.push(`${sheetName}: ${removed} duplicates removed, ${uniqueValues.length} unique values kept`);
      } else {
        results.push(`${sheetName}: No duplicates found`);
      }
    }

    return {
      success: true,
      message: "Cleaning complete!\n\n" + results.join('\n') + `\n\nTotal removed: ${totalRemoved} duplicates`,
      removed: totalRemoved,
      details: results
    };

  } catch (error) {
    console.error("Error in cleanSingleHorse:", error);
    return { success: false, error: error.message };
  }
}

function cleanAllHorsesFast() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheets  = ["Conf. Results", "Comp. Results"];
    let totalDuplicates    = 0;
    let processedColumns   = 0;
    let results            = [];

    for (const sheetName of sheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) { results.push(`${sheetName}: Sheet not found`); continue; }

      const lastCol  = sheet.getLastColumn();
      const lastRow  = sheet.getLastRow();
      const startRow = 6;
      if (lastRow < startRow) { results.push(`${sheetName}: No data`); continue; }

      let sheetDuplicates = 0;
      let sheetColumns    = 0;

      for (let col = 1; col <= lastCol; col++) {
        if (!sheet.getRange(1, col).getValue()) continue;

        const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1)
          .getValues().flat().filter(v => v !== "" && !isNaN(v));
        if (values.length === 0) continue;

        const uniqueValues = [...new Set(values)].sort((a, b) => b - a);
        const removed      = values.length - uniqueValues.length;
        if (removed > 0) {
          sheetDuplicates += removed;
          sheetColumns++;
          sheet.getRange(startRow, col, lastRow - startRow + 1, 1).clearContent();
          sheet.getRange(startRow, col, uniqueValues.length, 1).setValues(uniqueValues.map(v => [v]));
        }
      }

      totalDuplicates  += sheetDuplicates;
      processedColumns += sheetColumns;
      results.push(sheetDuplicates > 0
        ? `${sheetName}: ${sheetDuplicates} duplicates removed in ${sheetColumns} columns`
        : `${sheetName}: No duplicates found`);
    }

    return {
      success: true,
      message: "Cleaning complete!\n\n" + results.join('\n') + `\n\nTotal: ${totalDuplicates} duplicates removed in ${processedColumns} columns`,
      totalDuplicates,
      processedColumns,
      details: results
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}
