/* ══════════════════════════════════════════════════════════════
   SHEET STYLING ENGINE (Mit Conditional Formatting & Rahmen)
   ══════════════════════════════════════════════════════════════ */

function applyThemeToAllSheets() {
  const theme = getActiveTheme();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const FONT_FAMILY = "Inter";
  const failedSheets = [];

  sheets.forEach(sheet => {
    try {
    const name = sheet.getName();
    const config = resolveSheetStyleConfig_(name);

    // ----------------------------------------------------
    // FALL 1: SPEZIFISCH DEFINIERTER TAB (aus SHEET_CONFIG)
    // ----------------------------------------------------
    if (config && config.archiveLayout) {
      applyArchiveLayout_(sheet, theme);
      return;
    }
    if (config && config.settingsLayout) {
      applySettingsLayout_(sheet, theme);
      return;
    }
    if (config && config.resultMatrix) {
      applyResultMatrixStyling_(sheet, config, theme);
      return;
    }
    if (config) {
      if (config.compTeamLayout) prepareCompTeamLayout_(sheet);
      if (config.marketLayout) prepareHorseMarketLayout_(sheet, config.marketOutputColumns);
      if (config.broodmaresLayout) prepareBroodmaresLayout_(sheet);
      const lastRow = config.endRow || Math.max(sheet.getLastRow(), config.startDataRow || 2);
      const lastCol = config.dynamicEndCol ? sheet.getLastColumn() : config.endCol;

      if (lastRow === 0 || lastCol === 0) return;

      // Schriftart & Gridlines setzen
      try {
        const fullRange = sheet.getRange(1, config.startCol, lastRow, (lastCol - config.startCol + 1));
        fullRange.setFontFamily(FONT_FAMILY);
      } catch (e) {
      }
      sheet.setHiddenGridlines(false);
      if (config.frozenRows !== undefined) sheet.setFrozenRows(config.frozenRows);
      if (config.frozenColumns !== undefined) sheet.setFrozenColumns(config.frozenColumns);
      Object.entries(config.columnWidths || {}).forEach(([col, width]) => sheet.setColumnWidth(Number(col), width));

        // ----------------------------------------------------
        // GLOBAL: Alle Zellen dieses Tabs zentrieren
        // ----------------------------------------------------
        const allRange = sheet.getRange(1, config.startCol, lastRow, (lastCol - config.startCol + 1));
        allRange.setHorizontalAlignment("center");
        allRange.setVerticalAlignment("middle");

      // A) MATRIX-TABS
    if (config.isMatrix) {
      const headerRows = config.headerEndRow - config.headerStartRow + 1;

      // Spalte B: volle Header-Höhe dunkel hinterlegen
      const colBRange = sheet.getRange(config.headerStartRow, config.startCol, headerRows, 1);
      colBRange.setBackground(theme.primary.deep);
      colBRange.setFontColor(theme.neutrals.white);
      colBRange.setFontWeight("bold");
      colBRange.setHorizontalAlignment("center");
      colBRange.setVerticalAlignment("middle");

      // Ab Spalte C: nur die ersten N Zeilen dunkel hinterlegen
      const restStartCol = config.startCol + 1;
      if (lastCol >= restStartCol) {
        const restRows = config.narrowHeaderRows || headerRows;
        const restRange = sheet.getRange(config.headerStartRow, restStartCol, restRows, (lastCol - restStartCol + 1));
        restRange.setBackground(theme.primary.deep);
        restRange.setFontColor(theme.neutrals.white);
        restRange.setFontWeight("bold");
        restRange.setHorizontalAlignment("center");
        restRange.setVerticalAlignment("middle");
      }


      // ----------------------------------------------------
      //HORIZONTALE TRENNLINIE UNTER ZEILE X IN MATRIX-TABS
      // ----------------------------------------------------
      const targetDividerRow = config.dividerAfterHeaderRow || 7;
      if (lastRow >= targetDividerRow) {
        // --- Berechne die letzte Spalte basierend auf dem HEADER (nicht auf gesamten Daten) ---
        let headerEndCol = lastCol; // Fallback: bisherige letzte Spalte
        
        if (config.dynamicEndCol) {
          // Durchsuche die Header-Zeilen (headerStartRow bis headerEndRow) nach der letzten gefüllten Spalte
        const headerStart = config.headerStartRow;
        const headerEnd = config.narrowHeaderRows ? (config.headerStartRow + config.narrowHeaderRows - 1) : config.headerEndRow;
          let maxCol = config.startCol - 1; // Start bei Spalte vor startCol

      for (let r = headerStart; r <= headerEnd; r++) {
      const rowValues = sheet.getRange(r, config.startCol, 1, sheet.getLastColumn() - config.startCol + 1).getValues()[0];
      for (let c = rowValues.length - 1; c >= 0; c--) {
        if (rowValues[c] !== "" && rowValues[c] !== null) {
          const currentCol = config.startCol + c;
          if (currentCol > maxCol) maxCol = currentCol;
          break; // diese Zeile hat ihr Maximum gefunden, nächste Zeile
        }
      }
    }
    if (maxCol >= config.startCol) {
      headerEndCol = maxCol;
    }
  }

  // Jetzt verwenden wir headerEndCol für die Linienbreite
  const numCols = headerEndCol - config.startCol + 1;
  if (numCols > 0) {
    const dividerRange = sheet.getRange(targetDividerRow, config.startCol, 1, numCols);

    // --- Stil & Farbe aus Config (wie zuvor) ---
    const borderStyleMap = {
      "SOLID_THIN": SpreadsheetApp.BorderStyle.SOLID_THIN,
      "SOLID_MEDIUM": SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      "SOLID_THICK": SpreadsheetApp.BorderStyle.SOLID_THICK,
      "DASHED": SpreadsheetApp.BorderStyle.DASHED,
      "DOTTED": SpreadsheetApp.BorderStyle.DOTTED,
      "DASHED_DOTTED": SpreadsheetApp.BorderStyle.DASHED_DOTTED
    };
    const styleKey = config.dividerBorderStyle || "SOLID_MEDIUM";
    const borderStyle = borderStyleMap[styleKey] || SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const borderColor = config.dividerColor || theme.primary.deep;

    dividerRange.setBorder(
      null, null, true, null, null, null,
      borderColor,
      borderStyle
    );
  }
}

        // SPEZIAL-FALL: Wenn Tab "Conf. Results" ODER "Comp. Results" ist
        if (name === "Conf. Results" || name === "Comp. Results") {
          const startColForClear = 3; // Spalte C

          if (lastCol >= startColForClear) {
            const numColsToClear = lastCol - startColForClear + 1;
            const rowsToClear = lastRow - 2;

            if (rowsToClear > 0) {
              const clearRange = sheet.getRange(3, startColForClear, rowsToClear, numColsToClear);
              clearRange.setBackground(null);
            }
          }
        }
      }
      // B) LISTEN-TABS (z. B. Herd, Stats)
      else {
      // Ermittelt die dynamische Endspalte für den Header (entweder aus Config oder echten Daten)
      let headerEndCol = config.endCol;

      if (config.dynamicEndCol || config.endCol === "DYNAMIC") {
        const maxSheetCol = sheet.getLastColumn();
        if (maxSheetCol >= config.startCol) {
          const headerValues = sheet.getRange(config.headerRow, config.startCol, 1, maxSheetCol - config.startCol + 1).getValues()[0];
          
          let lastFilledIndex = -1;
          for (let i = headerValues.length - 1; i >= 0; i--) {
            if (headerValues[i] !== "" && headerValues[i] !== null) {
              lastFilledIndex = i;
              break;
            }
          }
          if (lastFilledIndex !== -1) {
            headerEndCol = config.startCol + lastFilledIndex;
          }
        }
      }
        const numHeaderCols = headerEndCol - config.startCol + 1;
        if (numHeaderCols > 0) {
          const topHeader = sheet.getRange(config.headerRow, config.startCol, 1, numHeaderCols);
          topHeader.setBackground(theme.primary.deep);
          topHeader.setFontColor(theme.neutrals.white);
          topHeader.setFontWeight("bold");
          topHeader.setHorizontalAlignment("center");
          topHeader.setVerticalAlignment("middle");
        }

        // Zebra-Muster Zeile für Zeile anwenden
        if (config.useZebra && lastRow >= config.startDataRow) {
          for (let r = config.startDataRow; r <= lastRow; r++) {
            const rowRange = sheet.getRange(r, config.startCol, 1, (lastCol - config.startCol + 1));
            const color = (r % 2 === 0) ? theme.primary.ultramuted : theme.primary.lightmuted;
            rowRange.setBackground(color);
          }
        }

        // Alle Datenzeilen zentrieren
        if (config.centerAllData && lastRow >= config.startDataRow) {
          const dataRange = sheet.getRange(config.startDataRow, config.startCol, (lastRow - config.startDataRow + 1), (lastCol - config.startCol + 1));
          dataRange.setHorizontalAlignment("center");
        }

        applyConfiguredAlignment_(sheet, config, config.startDataRow, lastRow);
        if (config.compTeamLayout) {
          sheet.getRange(1,2,1,8).setBackground(theme.primary.ultramuted).setFontColor(theme.primary.deep).setHorizontalAlignment("left");
          sheet.getRange("C1").setWrap(false);
        }

        // Spezifische Spalten fett machen
        if (config.boldColumns && config.boldColumns.length > 0) {
          config.boldColumns.forEach(colIndex => {
            if (colIndex >= config.startCol && colIndex <= lastCol && lastRow >= config.startDataRow) {
              const boldRange = sheet.getRange(config.startDataRow, colIndex, (lastRow - config.startDataRow + 1), 1);
              boldRange.setFontWeight("bold");
            }
          });
        }

        // ------------------------------------------------
        // DROPDOWN- & DATUMS-FORMATIERUNG ANWENDEN
        // ------------------------------------------------
        const startRow = config.startDataRow || 2;

        // 1) STATUS DROPDOWN
        if (config.statusDropdownCol) {
          const statusMapping = [
            { text: "active",      bg: theme.tertiary.light,  fg: theme.neutrals.white },
            { text: "observation", bg: theme.accents.accent5,   fg: theme.neutrals.white },
            { text: "companion",   bg: theme.accents.accent1, fg: theme.neutrals.white },
            { text: "for sale",    bg: theme.accents.accent2, fg: theme.neutrals.white },
            { text: "sold",        bg: theme.neutrals.grey,   fg: theme.neutrals.white },
            { text: "retired",     bg: theme.neutrals.grey,   fg: theme.neutrals.white },
            { text: "passed",      bg: theme.neutrals.grey,   fg: theme.neutrals.white }
          ];
          applyDropdownFormatting(sheet, config.statusDropdownCol, startRow, config.softThemeColors ? softenThemeMapping_(statusMapping, theme) : statusMapping);
        }

        // 2) GENDER DROPDOWN
        if (config.genderDropdownCol !== false && (!config.gradeColorRange || config.genderDropdownCol)) {
        const genderCol = config.genderDropdownCol || 6;
        const genderMapping = [
          { text: "Stallion", bg: theme.accents.accent3, fg: theme.neutrals.white },
          { text: "Mare",     bg: theme.accents.accent1,   fg: theme.neutrals.white },
          { text: "Gelding",  bg: theme.neutrals.grey,   fg: theme.neutrals.white }
        ];
        applyDropdownFormatting(sheet, genderCol, startRow, config.softThemeColors ? softenThemeMapping_(genderMapping, theme) : genderMapping);

        }

        // 3) DATUMS-SPALTE
        if (config.dateColumns && config.dateColumns.length > 0) {
          config.dateColumns.forEach(colIndex => {
            applyRegionalDateFormat(sheet, colIndex, startRow);
          });
        }

          // 3b) UHRZEIT-SPALTEN (neu)
        if (config.timeColumns && config.timeColumns.length > 0) {
          config.timeColumns.forEach(colIndex => {
            applyRegionalTimeFormat(sheet, colIndex, startRow);
          });
        }

                // 3c) DATUM+UHRZEIT-SPALTEN
        if (config.datetimeColumns && config.datetimeColumns.length > 0) {
          config.datetimeColumns.forEach(colIndex => {
            applyRegionalDateTimeFormat(sheet, colIndex, startRow);
          });
        }

        // 4) RASSEDURCHSCHNITT-HERVORHEBUNG
        if (config.breedAverageHighlights && config.breedAverageHighlights.length > 0) {
        config.breedAverageHighlights.forEach(bah => {
          applyBreedAverageHighlight(
            sheet, bah.targetCol, bah.compareCol, bah.breedCol,
            startRow, config.softThemeColors ? mixThemeColor_(theme.secondary.base, theme.neutrals.white, 0.28) : theme.secondary.base, bah.validBreedsRange, lastRow
          );
        });
        } else if (config.breedAverageHighlight) {
          const bah = config.breedAverageHighlight;
          applyBreedAverageHighlight(
            sheet, bah.targetCol, bah.compareCol, bah.breedCol,
            startRow, config.softThemeColors ? mixThemeColor_(theme.secondary.base, theme.neutrals.white, 0.28) : theme.secondary.base, bah.validBreedsRange, lastRow
          );
        }

        // 5) RATING-DROPDOWNS
        if (config.ratingDropdownCols && config.ratingDropdownCols.length > 0) {
          const ratingMapping = [
            { text: "Excellent", bg: theme.tertiary.mid,   fg: theme.neutrals.white },
            { text: "Good",      bg: theme.tertiary.light,    fg: theme.neutrals.white },
            { text: "Average",   bg: theme.accents.accent5,   fg: theme.neutrals.white },
            { text: "Fair",      bg: theme.accents.accent1, fg: theme.neutrals.white }
          ];
          config.ratingDropdownCols.forEach(colIndex => {
            applyDropdownFormatting(sheet, colIndex, startRow, config.softThemeColors ? softenThemeMapping_(ratingMapping, theme) : ratingMapping);
          });
        }

        // ------------------------------------------------
        // TRENNLINIEN & SPALTEN-UNTERLEGUNG
        // ------------------------------------------------
        if (config.dividerBeforeCols && config.dividerBeforeCols.length > 0) {
          const lineColor = config.subtleDividers ? mixThemeColor_(theme.primary.deep, theme.neutrals.white, 0.3) : theme.primary.deep;
          config.dividerBeforeCols.forEach(targetColIndex => {
            const borderColIndex = targetColIndex - 1;
            if (borderColIndex >= config.startCol && borderColIndex <= lastCol && lastRow >= startRow) {
              const dividerRange = sheet.getRange(startRow, borderColIndex, (lastRow - startRow + 1), 1);
              dividerRange.setBorder(null, null, null, true, null, null, lineColor, config.subtleDividers ? SpreadsheetApp.BorderStyle.SOLID : SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
            }
          });
        }

        if (config.dividerBeforeColsDashed && config.dividerBeforeColsDashed.length > 0) {
          const lineColor = config.subtleDividers ? mixThemeColor_(theme.primary.deep, theme.neutrals.white, 0.3) : theme.primary.deep;
          config.dividerBeforeColsDashed.forEach(targetColIndex => {
            const borderColIndex = targetColIndex - 1;
            if (borderColIndex >= config.startCol && borderColIndex <= lastCol && lastRow >= startRow) {
              const dividerRange = sheet.getRange(startRow, borderColIndex, (lastRow - startRow + 1), 1);
              dividerRange.setBorder(null, null, null, true, null, null, lineColor, SpreadsheetApp.BorderStyle.DASHED);
            }
          });
        }

        if (config.subtleShadeCols && config.subtleShadeCols.length > 0) {
          config.subtleShadeCols.forEach(colIndex => {
            if (colIndex >= config.startCol && colIndex <= lastCol && lastRow >= startRow) {
              const shadeRange = sheet.getRange(startRow, colIndex, (lastRow - startRow + 1), 1);
              shadeRange.setBackground((config.softStatsColors || config.shadeWeight !== undefined) ? mixThemeColor_(theme.neutrals.neutralcolor, theme.neutrals.white, config.shadeWeight !== undefined ? config.shadeWeight : 0.4) : theme.neutrals.neutralcolor);
            }
          });
        }

        if (config.gradeColorRange && lastRow >= startRow) {
          const gcr = config.gradeColorRange;
          applyGradeColorCoding(sheet, gcr.startCol, gcr.endCol, startRow, lastRow, theme, config.softStatsColors);
        }

        if (config.heatmapRange && lastRow >= startRow) {
          const hr = config.heatmapRange;
          applyHeatmapGradient(sheet, hr.startCol, hr.endCol, startRow, lastRow, theme, config.softStatsColors);
        }
      }
    }
    // ----------------------------------------------------
    // FALL 2: AUTOMATISCHER FALLBACK
    // ----------------------------------------------------
    else {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow === 0 || lastCol === 0) return;

      try {
        const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
        fullRange.setFontFamily(FONT_FAMILY);
      } catch (e) {}
      sheet.setHiddenGridlines(false);

      const topHeader = sheet.getRange(1, 1, 1, lastCol);
      topHeader.setBackground(theme.primary.deep);
      topHeader.setFontColor(theme.neutrals.white);
      topHeader.setFontWeight("bold");
      topHeader.setHorizontalAlignment("center");
      topHeader.setVerticalAlignment("middle");

      if (lastRow > 1) {
        for (let r = 2; r <= lastRow; r++) {
          const rowRange = sheet.getRange(r, 1, 1, lastCol);
          rowRange.setBackground(r % 2 === 0 ? theme.primary.ultra : theme.primary.light);
        }
      }
    }
    } catch (sheetError) {
      failedSheets.push(sheet.getName() + ": " + sheetError.message);
      Logger.log("FEHLER beim Styling von Tab '" + sheet.getName() + "': " + sheetError.message + " (Stack: " + sheetError.stack + ")");
    }
  });
  if (failedSheets.length) {
    const message = "Formatting could not be completed for these sheets:\n\n" +
      failedSheets.join("\n\n") +
      "\n\nPlease copy this message or take a screenshot before closing this window.";
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Styling incomplete", message, ui.ButtonSet.OK);
    } catch (uiError) {
      // Background executions may not have access to a dialog.
      ss.toast(message, "Styling incomplete", -1);
    }
  }
}

/* ══════════════════════════════════════════════════════════════
   HILFSFUNKTIONEN (DÜRFEN NICHT FEHLEN)
   ══════════════════════════════════════════════════════════════ */

function applyDropdownFormatting(sheet, colIndex, startRow, colorMapping) {
  const lastRow = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow(), startRow + 50));
  if (lastRow < startRow) return;
  const range = sheet.getRange(startRow, colIndex, (lastRow - startRow + 1), 1);

  let allRules = sheet.getConditionalFormatRules();
  let cleanRules = allRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === colIndex);
  });

  colorMapping.forEach(item => {
    if (item.bg) {
      const newRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(item.text)
        .setBackground(item.bg)
        .setFontColor(item.fg)
        .setBold(true)
        .setRanges([range])
        .build();
      cleanRules.push(newRule);
    }
  });

  sheet.setConditionalFormatRules(cleanRules);
}

function applyBreedAverageHighlight(sheet, targetCol, compareCol, breedCol, startRow, bgColor, validBreedsRange, lastRow) {
  lastRow = lastRow || sheet.getLastRow();
  const range = sheet.getRange(startRow, targetCol, (lastRow - startRow + 1), 1);

  const breedColLetter   = columnToLetter(breedCol);
  const compareColLetter = columnToLetter(compareCol);

  const breedCellRel   = "$" + breedColLetter + startRow;
  const compareCellRel = "$" + compareColLetter + startRow;
  const breedColFull    = "$" + breedColLetter + "$" + startRow + ":$" + breedColLetter + "$" + lastRow;
  const compareColFull  = "$" + compareColLetter + "$" + startRow + ":$" + compareColLetter + "$" + lastRow;

  let conditions = [
    compareCellRel + "<>\"\"",
    compareCellRel + ">AVERAGE(IFERROR(FILTER(" + compareColFull + "," + breedColFull + "=" + breedCellRel + "),0))"
  ];

  if (validBreedsRange) {
    conditions.splice(1, 0,
      "ISNUMBER(MATCH(" + breedCellRel + ",INDIRECT(\"" + validBreedsRange + "\"),0))"
    );
  }

  const formula = "=AND(" + conditions.join(",") + ")";
  let allRules = sheet.getConditionalFormatRules();

  let cleanRules = allRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === targetCol);
  });

  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(bgColor)
    .setRanges([range])
    .build();

  cleanRules.push(newRule);
  sheet.setConditionalFormatRules(cleanRules);
}
function applyGradeColorCoding(sheet, startCol, endCol, startRow, lastRow, theme, softColors) {
  const range = sheet.getRange(startRow, startCol, (lastRow - startRow + 1), (endCol - startCol + 1));

  const gradeMapping = [
    { text: "VG", bg: theme.tertiary.deep,   fg: theme.neutrals.white },
    { text: "G+", bg: theme.tertiary.dark,    fg: theme.neutrals.white },
    { text: "G",  bg: theme.tertiary.mid,   fg: theme.neutrals.white },
    { text: "G-", bg: theme.tertiary.light, fg: theme.neutrals.white },
    { text: "A",  bg: theme.accents.accent5,  fg: theme.neutrals.white },
    { text: "BA", bg: theme.accents.accent6, fg: theme.neutrals.white },
    { text: "P",  bg: theme.accents.accent2,    fg: theme.neutrals.white }
  ];

  let allRules = sheet.getConditionalFormatRules();
  let cleanRules = allRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() >= startCol && r.getColumn() <= endCol);
  });

  (softColors ? statsGradeMapping_(theme) : gradeMapping).forEach(item => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(item.text)
      .setBackground(item.bg)
      .setFontColor(item.fg)
      .setBold(true)
      .setRanges([range])
      .build();
    cleanRules.push(rule);
  });

  sheet.setConditionalFormatRules(cleanRules);
}

function applyHeatmapGradient(sheet, startCol, endCol, startRow, lastRow, theme, softColors) {
  const range = sheet.getRange(startRow, startCol, (lastRow - startRow + 1), (endCol - startCol + 1));
  range.setFontColor(softColors ? theme.primary.deep : theme.neutrals.white);
  const heatColor = color => softColors ? mixThemeColor_(color, theme.neutrals.white, 0.65) : color;

  let allRules = sheet.getConditionalFormatRules();
  let cleanRules = allRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() >= startCol && r.getColumn() <= endCol);
  });

  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(heatColor(theme.accents.accent2), SpreadsheetApp.InterpolationType.MIN, "")
    .setGradientMidpointWithValue(heatColor(softColors ? theme.secondary.base : theme.tertiary.light), SpreadsheetApp.InterpolationType.PERCENTILE, "50")
    .setGradientMaxpointWithValue(heatColor(softColors ? theme.tertiary.light : theme.tertiary.mid), SpreadsheetApp.InterpolationType.MAX, "")
    .setRanges([range])
    .build();

  cleanRules.push(newRule);
  sheet.setConditionalFormatRules(cleanRules);
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function applyRegionalDateFormat(sheet, colIndex, startRow) {
  const userTimeZone = Session.getScriptTimeZone();
  let dateFormat = "yyyy-mm-dd";

  if (userTimeZone.includes("Europe") || userTimeZone.includes("Berlin") || userTimeZone.includes("Vienna") || userTimeZone.includes("Zurich")) {
    dateFormat = "dd.mm.yyyy";
  } else if (userTimeZone.includes("America") || userTimeZone.includes("US")) {
    dateFormat = "mm/dd/yyyy";
  } else if (userTimeZone.includes("London") || userTimeZone.includes("GB")) {
    dateFormat = "dd/mm/yyyy";
  }

  const lastRow = Math.max(sheet.getLastRow(), startRow + 50);
  const range = sheet.getRange(startRow, colIndex, (lastRow - startRow + 1), 1);
  try {
    range.setNumberFormat(dateFormat);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("Datumsformat: Spalte " + colIndex + " ist typisiert, übersprungen: " + e.message);
  }
}

function applyRegionalTimeFormat(sheet, colIndex, startRow) {
  const userTimeZone = Session.getScriptTimeZone();
  let timeFormat = "HH:mm"; // Standard: 24h

  // Regionale Anpassung: USA → 12h mit AM/PM
  if (userTimeZone.includes("America") || userTimeZone.includes("US")) {
    timeFormat = "h:mm AM/PM";
  } else if (userTimeZone.includes("Europe") || userTimeZone.includes("Berlin") || 
             userTimeZone.includes("Vienna") || userTimeZone.includes("Zurich") ||
             userTimeZone.includes("London") || userTimeZone.includes("GB")) {
    timeFormat = "HH:mm"; // 24h
  }

  const lastRow = Math.max(sheet.getLastRow(), startRow + 50);
  const range = sheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1);
  try {
    range.setNumberFormat(timeFormat);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("Uhrzeitformat: Spalte " + colIndex + " ist typisiert, übersprungen: " + e.message);
  }
}

function applyRegionalDateTimeFormat(sheet, colIndex, startRow) {
  const userTimeZone = Session.getScriptTimeZone();
  let format = "dd.mm.yyyy HH:mm"; // Standard Europa

  if (userTimeZone.includes("America") || userTimeZone.includes("US")) {
    format = "mm/dd/yyyy h:mm AM/PM";
  } else if (userTimeZone.includes("Europe") || userTimeZone.includes("Berlin") || 
             userTimeZone.includes("Vienna") || userTimeZone.includes("Zurich")) {
    format = "dd.mm.yyyy HH:mm";
  } else if (userTimeZone.includes("London") || userTimeZone.includes("GB")) {
    format = "dd/mm/yyyy HH:mm";
  }

  const lastRow = Math.max(sheet.getLastRow(), startRow + 50);
  const range = sheet.getRange(startRow, colIndex, lastRow - startRow + 1, 1);
  try {
    range.setNumberFormat(format);
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("Datum+Zeit-Format: Spalte " + colIndex + " ist typisiert, übersprungen: " + e.message);
  }
}

function onEdit_Styling(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const name = sheet.getName();
  const col = range.getColumn();
  const row = range.getRow();
  const config = resolveSheetStyleConfig_(name);

  if (config && config.resultMatrix) {
    applyResultMatrixStyling_(sheet, config, getActiveTheme(), col, col + range.getNumColumns() - 1);
    return;
  }

  if (config && config.settingsLayout) {
    applySettingsLayout_(sheet, getActiveTheme());
    return;
  }

  // LIVERENDER für Header & Trennlinie (unverändert)
  if (config && (row === 1 || row === 2)) {
    const theme = getActiveTheme();
    if (e.value) {
      const headerCell = sheet.getRange(1, col, 2, 1);
      headerCell.setBackground(theme.primary.deep);
      headerCell.setFontColor(theme.neutrals.white);
      headerCell.setFontWeight("bold");
      headerCell.setHorizontalAlignment("center");
      headerCell.setVerticalAlignment("middle");
      headerCell.setFontFamily("Inter");

      if (config.isMatrix && config.dividerAfterHeaderRow) {
        const lineRow = config.dividerAfterHeaderRow;
        const lineCell = sheet.getRange(lineRow, col);
        const styleKey = config.dividerBorderStyle || "DASHED";
        const borderStyle = (styleKey === "DASHED")
          ? SpreadsheetApp.BorderStyle.DASHED
          : SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
        lineCell.setBorder(null, null, true, null, null, null, theme.primary.deep, borderStyle);
      }
    }
  }

  // NEU: Live-Styling für neue/bearbeitete Zeilen in Listen-Tabs
  if (config && !config.isMatrix) {
    const startDataRow = config.startDataRow || 2;
    const lastEditedRow = row + range.getNumRows() - 1;
    if (lastEditedRow >= startDataRow) {
      const theme = getActiveTheme();
      const from = Math.max(row, startDataRow);
      for (let r = from; r <= lastEditedRow; r++) {
        applyLiveRowStyling(sheet, config, theme, r);
      }
    }
  }

  // Multiselect (Spalte 18 / R) — unverändert
  if (col === 18 && row >= 2) {
    const newValue = e.value;
    const oldValue = e.oldValue;
    if (!newValue) return;
    let currentValues = [];
    if (oldValue) currentValues = oldValue.split(", ").map(v => v.trim());
    if (currentValues.includes(newValue)) {
      currentValues = currentValues.filter(v => v !== newValue);
    } else {
      currentValues.push(newValue);
    }
    range.setValue(currentValues.join(", "));
  }
}

/**
 * Wendet Zebra, Zentrierung, fette Spalten, Trennlinien & Subtle-Shading
 * live auf eine einzelne Datenzeile an — läuft bei jedem Edit statt erst
 * beim nächsten manuellen applyThemeToAllSheets()-Lauf.
 */
function applyLiveRowStyling(sheet, config, theme, row) {
  const lastCol = config.dynamicEndCol ? sheet.getLastColumn() : config.endCol;
  if (!lastCol || lastCol < config.startCol) return;

  const rowRange = sheet.getRange(row, config.startCol, 1, (lastCol - config.startCol + 1));
  rowRange.setFontFamily("Inter");

  if (config.useZebra) {
    const color = (row % 2 === 0) ? theme.primary.ultramuted : theme.primary.lightmuted;
    rowRange.setBackground(color);
  }

  if (config.centerAllData) {
    rowRange.setHorizontalAlignment("center");
    rowRange.setVerticalAlignment("middle");
  }

  applyConfiguredAlignment_(sheet, config, row, row);

  if (config.boldColumns && config.boldColumns.length > 0) {
    config.boldColumns.forEach(colIndex => {
      if (colIndex >= config.startCol && colIndex <= lastCol) {
        sheet.getRange(row, colIndex).setFontWeight("bold");
      }
    });
  }

  if (config.dividerBeforeCols && config.dividerBeforeCols.length > 0) {
    config.dividerBeforeCols.forEach(targetColIndex => {
      const borderColIndex = targetColIndex - 1;
      if (borderColIndex >= config.startCol && borderColIndex <= lastCol) {
        sheet.getRange(row, borderColIndex).setBorder(
          null, null, null, true, null, null,
          theme.primary.deep, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
      }
    });
  }

  if (config.dividerBeforeColsDashed && config.dividerBeforeColsDashed.length > 0) {
    config.dividerBeforeColsDashed.forEach(targetColIndex => {
      const borderColIndex = targetColIndex - 1;
      if (borderColIndex >= config.startCol && borderColIndex <= lastCol) {
        sheet.getRange(row, borderColIndex).setBorder(
          null, null, null, true, null, null,
          theme.primary.deep, SpreadsheetApp.BorderStyle.DASHED
        );
      }
    });
  }

  if (config.subtleShadeCols && config.subtleShadeCols.length > 0) {
    config.subtleShadeCols.forEach(colIndex => {
      if (colIndex >= config.startCol && colIndex <= lastCol) {
        sheet.getRange(row, colIndex).setBackground((config.softStatsColors || config.shadeWeight !== undefined) ? mixThemeColor_(theme.neutrals.neutralcolor, theme.neutrals.white, config.shadeWeight !== undefined ? config.shadeWeight : 0.4) : theme.neutrals.neutralcolor);
      }
    });
  }
}

/**
 * Hebt den Pferdenamen oder den GAP hervor, wenn GAP > 0 (Potenzial vorhanden)
 */
function applyGapHighlight(sheet, startCol, lastCol, theme) {
  if (lastCol < startCol) return;

  const numCols = lastCol - startCol + 1;
  
  // Wir wenden die Regel z. B. direkt auf die Namen in Zeile 2 an
  const nameRange = sheet.getRange(2, startCol, 1, numCols); // Zeile 2 = Name

  // Formel prüft, ob die GAP-Zelle in Zeile 7 größer als 0 ist (und nicht leer)
  // Wichtig: Relative Spaltenreferenz (C$7), damit es für jede Spalte einzeln wandert
  const firstColLetter = columnToLetter(startCol);
  const formula = `=AND(${firstColLetter}$7<>"", ${firstColLetter}$7>0)`;

  let allRules = sheet.getConditionalFormatRules();
  
  // Alte GAP-Regeln auf dem Bereich filtern
  let cleanRules = allRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getRow() === 2 && r.getColumn() >= startCol);
  });

  const newRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(theme.secondary.base) // Oder z. B. theme.accents.accent5
    .setFontColor(theme.primary.deep)
    .setBold(true)
    .setRanges([nameRange])
    .build();

  cleanRules.push(newRule);
  sheet.setConditionalFormatRules(cleanRules);
}
// Blends only active-theme colours; no fixed palette is introduced.
function mixThemeColor_(foreground, background, weight) {
  const a = foreground.replace('#',''), b = background.replace('#','');
  return '#' + [0,2,4].map(i => Math.round(parseInt(a.slice(i,i+2),16)*weight + parseInt(b.slice(i,i+2),16)*(1-weight)).toString(16).padStart(2,'0')).join('');
}
function softenThemeMapping_(mapping, theme) {
  return mapping.map(item => ({...item, bg:mixThemeColor_(item.bg,theme.neutrals.white,0.40), fg:mixThemeColor_(item.bg,theme.primary.deep,0.3)}));
}
function applyConfiguredAlignment_(sheet, config, firstRow, lastRow) {
  if (lastRow < firstRow) return;
  Object.entries(config.numberFormatColumns || {}).forEach(([format,columns]) => columns.forEach(col => {
    try {
      sheet.getRange(firstRow,col,lastRow-firstRow+1,1).setNumberFormat(format);
    } catch (error) {
      Logger.log("Number format skipped for " + sheet.getName() + " column " + col + ": " + error.message);
    }
  }));
  if (!config.alignColumns) return;
  Object.entries(config.alignColumns).forEach(([alignment,columns]) => columns.forEach(col => {
    sheet.getRange(firstRow,col,lastRow-firstRow+1,1).setHorizontalAlignment(alignment);
  }));
}

// Accept cosmetic differences in tab names without guessing unrelated sheets.
function resolveSheetStyleConfig_(name) {
  if (SHEET_CONFIG[name]) return SHEET_CONFIG[name];
  const normalize = value => value.replace(/[ _]/g, '').toLowerCase();
  const matches = Object.keys(SHEET_CONFIG).filter(key => normalize(key) === normalize(name));
  return matches.length === 1 ? SHEET_CONFIG[matches[0]] : undefined;
}
function statsGradeMapping_(theme) {
  return [
    ["VG", theme.tertiary.light, 0.85],
    ["G+", theme.tertiary.ultra, 0.72],
    ["G", theme.tertiary.ultra, 0.50],
    ["G-", theme.tertiary.ultra, 0.28],
    ["A", theme.accents.accent5, 0.65],
    ["BA", theme.accents.accent6, 0.60],
    ["P", theme.accents.accent2, 0.60]
  ].map(([text, color, weight]) => ({text, bg:mixThemeColor_(color, theme.neutrals.white, weight), fg:theme.primary.deep}));
}

// Optional column bounds keep score imports and edits proportional to their changes.
function applyResultMatrixStyling_(sheet, config, theme, fromCol, toCol) {
  const lastCol = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), config.headerEndRow);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
  sheet.setColumnWidth(2, 110);
  sheet.setRowHeight(2, 48);
  sheet.getRange(1, 2, config.headerEndRow, 1)
    .setBackground(theme.primary.deep).setFontColor(theme.neutrals.white)
    .setFontFamily('Inter').setFontWeight('bold').setHorizontalAlignment('left');
  const first = Math.max(3, fromCol || 3);
  const end = Math.min(lastCol, toCol || lastCol);
  if (end < first) return;
  const count = end - first + 1;
  sheet.setColumnWidths(first, count, 155);
  sheet.getRange(1, first, lastRow, count).setFontFamily('Inter').setVerticalAlignment('middle');
  sheet.getRange(1, first, 2, count).setBackground(theme.primary.deep)
    .setFontColor(theme.neutrals.white).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, first, 1, count).setWrap(true);
  const bodyRows = lastRow - 2;
  const backgrounds = Array.from({length:bodyRows}, (_,i) => {
    const row = i + 3;
    const color = row <= config.headerEndRow
      ? mixThemeColor_(theme.neutrals.neutralcolor, theme.neutrals.white, 0.55)
      : ((row - config.headerEndRow) % 2 ? theme.primary.lightmuted : theme.primary.ultramuted);
    return Array(count).fill(color);
  });
  sheet.getRange(3, first, bodyRows, count).setBackgrounds(backgrounds)
    .setFontColor(theme.primary.deep).setFontWeight('normal').setHorizontalAlignment('right');
  sheet.getRange(config.headerEndRow, first, 1, count).setBorder(null, null, true, null, null, null,
    mixThemeColor_(theme.primary.deep, theme.neutrals.white, 0.35), SpreadsheetApp.BorderStyle.SOLID);
}
function refreshResultColumnStyle_(sheet, col) {
  const config = resolveSheetStyleConfig_(sheet.getName());
  if (!config || !config.resultMatrix) return;
  try { applyResultMatrixStyling_(sheet, config, getActiveTheme(), col, col); }
  catch (error) { Logger.log('Result column styling: ' + error.message); }
}

function prepareCompTeamLayout_(sheet) {
  const oldTimestamp = sheet.getRange('B2').getValue();
  if (oldTimestamp && oldTimestamp !== 'ID' && !sheet.getRange('C1').getValue()) {
    sheet.getRange('C1').setValue(oldTimestamp);
  }
  sheet.getRange('B1').setValue('Last sync:');
  sheet.getRange('B2').setValue('ID');
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const formulas = sheet.getRange(3,4,lastRow-2,6).getFormulas();
  formulas.forEach((row,i) => row.forEach((formula,j) => {
    const guarded = guardCompTeamFormula_(formula, i+3);
    if (guarded !== formula) sheet.getRange(i+3,j+4).setFormula(guarded);
  }));
}
function guardCompTeamFormula_(formula, row) {
  if (!formula || /ARRAYFORMULA/i.test(formula) || !/VLOOKUP/i.test(formula)) return formula;
  if (formula.startsWith('=IF($B' + row + '="","",')) return formula;
  // Only guard formulas that explicitly use this row's horse ID.
  const refs = formula.match(/\$?B\$?\d+/gi) || [];
  if (!refs.some(ref => ref.replace(/\$/g,'').toUpperCase() === 'B'+row)) return formula;
  return '=IF($B' + row + '="","",' + formula.slice(1) + ')';
}

function guardBroodmareDueFormula_(formula, row) {
  if (!formula || /ARRAYFORMULA/i.test(formula)) return formula;
  if (formula.startsWith('=IF($F' + row + '="","",')) return formula;
  const refs = formula.match(/\$?F\$?\d+/gi) || [];
  if (!refs.some(ref => ref.replace(/\$/g,'').toUpperCase() === 'F'+row)) return formula;
  return '=IF($F' + row + '="","",' + formula.slice(1) + ')';
}
function prepareBroodmaresLayout_(sheet) {
  const header = sheet.getRange('B1');
  const previous = header.getValue();
  if (previous && previous !== 'Confirmed') {
    const existing = header.getNote();
    header.setNote((existing ? existing + '\n' : '') + String(previous));
  }
  header.setValue('Confirmed');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const formulas = sheet.getRange(2,11,lastRow-1,1).getFormulas();
  formulas.forEach((values,i) => {
    const next = guardBroodmareDueFormula_(values[0], i+2);
    if (next !== values[0]) sheet.getRange(i+2,11).setFormula(next);
  });
  const values = sheet.getRange(2,14,lastRow-1,1).getValues();
  let dates=0, texts=0, other=0;
  values.forEach(([value]) => {
    if (value === '' || value == null) return;
    if (value instanceof Date) dates++;
    else if (typeof value === 'string') texts++;
    else other++;
  });
  const dateHeader = sheet.getRange('N1');
  const note = (dateHeader.getNote() || '').replace(/\n?\[Date audit\][^\n]*/g,'');
  dateHeader.setNote((note ? note+'\n' : '') + '[Date audit] Date values: '+dates+'; text values: '+texts+'; other values: '+other+'. Text dates have been preserved to avoid ambiguous day/month conversion.');
}

function guardHorseMarketFormula_(formula, row) {
  if (!formula || /ARRAYFORMULA|\b(SEQUENCE|FILTER|QUERY|IMPORTRANGE)\s*\(/i.test(formula)) return formula;
  const prefix = '=IF($B' + row + '="","",';
  if (formula.startsWith(prefix)) return formula;
  return prefix + formula.slice(1) + ')';
}
function prepareHorseMarketLayout_(sheet, outputColumns) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // Only the visible output columns; hidden pricing inputs remain untouched.
  (outputColumns || [3,13,14]).forEach(col => {
    const formulas = sheet.getRange(2,col,lastRow-1,1).getFormulas();
    formulas.forEach(([formula],i) => {
      const guarded = guardHorseMarketFormula_(formula,i+2);
      if (guarded !== formula) sheet.getRange(i+2,col).setFormula(guarded);
    });
  });
}

function applySettingsLayout_(sheet, theme) {
  const lastRow = sheet.getLastRow();
  const widths = {2:55,3:260,4:135,5:110,6:110,7:90,8:125,9:205,10:90,11:100,12:28,13:170,14:110,15:260,16:85,17:125};
  Object.entries(widths).forEach(([col,width]) => sheet.setColumnWidth(Number(col),width));
  const border = mixThemeColor_(theme.primary.deep,theme.neutrals.white,0.25);
  function block(top,left,bottom,right,header) {
    if(bottom<top) return;
    const height=bottom-top+1, width=right-left+1;
    const range=sheet.getRange(top,left,height,width);
    range.setFontFamily('Inter').setFontSize(10).setFontColor(theme.primary.deep)
      .setVerticalAlignment('middle').setHorizontalAlignment('left')
      .setBackgrounds(Array.from({length:height},(_,i)=>Array(width).fill(i%2?theme.primary.lightmuted:theme.primary.ultramuted)));
    range.setBorder(true,true,true,true,false,false,border,SpreadsheetApp.BorderStyle.SOLID);
    if(header) sheet.getRange(top,left,1,width).setBackground(theme.primary.deep)
      .setFontColor(theme.neutrals.white).setFontWeight('bold');
  }
  // Find the end of the breed list independently of the blocks beside it.
  let breedEnd=2;
  if(lastRow>=3) sheet.getRange(3,3,lastRow-2,1).getValues().forEach(([value],i)=>{
    if(value!=='' && value!=null) breedEnd=i+3;
  });
  block(2,2,breedEnd,11,true);
  if(breedEnd>=3) {
    [2,5,6,7,8,10].forEach(col=>sheet.getRange(3,col,breedEnd-2,1).setHorizontalAlignment('right'));
    sheet.getRange(3,4,breedEnd-2,1).setHorizontalAlignment('center');
  }
  block(2,13,5,15,true);
  sheet.getRange(3,14,3,1).setHorizontalAlignment('right');
  sheet.getRange(3,15,3,1).setHorizontalAlignment('center');
  block(7,13,17,15,true);
  sheet.getRange(8,14,10,1).setHorizontalAlignment('right');
  sheet.getRange(8,15,10,1).setWrap(true);
  block(2,16,5,17,false);
  sheet.getRange(2,16,4,1).setFontWeight('bold');
  sheet.getRange(2,17,4,1).setHorizontalAlignment('right');
}

function archiveBlocks_(headers) {
  const blocks=[];
  headers.slice(3).forEach((header,i)=>{
    const text=String(header || '');
    const split=text.indexOf(' - ');
    const source=split<0 ? 'Other' : text.slice(0,split);
    const previous=blocks[blocks.length-1];
    if(previous && previous.source===source) previous.end=i+4;
    else blocks.push({source,start:i+4,end:i+4});
  });
  return blocks;
}
function applyArchiveLayout_(sheet, theme) {
  const cols=sheet.getLastColumn(), rows=sheet.getLastRow();
  if(cols<3 || rows<1) return;
  const headers=sheet.getRange(1,1,1,cols).getValues()[0];
  sheet.setFrozenRows(1); sheet.setFrozenColumns(3);
  sheet.setColumnWidths(1,cols,140);
  sheet.setColumnWidth(1,110); sheet.setColumnWidth(2,230); sheet.setColumnWidth(3,145);
  sheet.setRowHeight(1,64);
  sheet.getRange(1,1,rows,cols).setFontFamily('Inter').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(1,1,1,cols).setBackground(theme.primary.deep).setFontColor(theme.neutrals.white)
    .setFontWeight('bold').setWrap(true).setHorizontalAlignment('left');
  if(rows>1) {
    sheet.getRange(2,1,rows-1,cols).setFontColor(theme.primary.deep);
    for(let row=2;row<=rows;row++) sheet.getRange(row,1,1,cols)
      .setBackground(row%2===0 ? theme.primary.ultramuted : theme.primary.lightmuted);
  }
  headers.forEach((header,i)=>{
    if(i<3) return;
    const label=String(header).split(' - ').slice(1).join(' - ');
    if(/name|sire|dam|coat|genetic|notes|link/i.test(label)) sheet.setColumnWidth(i+1,220);
  });
  const border=mixThemeColor_(theme.primary.deep,theme.neutrals.white,0.35);
  archiveBlocks_(headers).forEach(block=>{
    sheet.getRange(1,block.start,rows,1).setBorder(null,true,null,null,null,null,border,SpreadsheetApp.BorderStyle.SOLID);
    // Keep the first column visible as a section marker and separate adjacent groups.
    const start=block.start+1, count=block.end-block.start;
    if(count<1) return;
    let group=null;
    if(sheet.getColumnGroupDepth(start)>0) group=sheet.getColumnGroup(start,1);
    if(group) {
      const range=group.getRange();
      if(range.getColumn()!==start) return; // Preserve a user-created overlapping group.
      if(range.getNumColumns()===count) return;
      const collapsed=group.isCollapsed();
      group.remove();
      sheet.getRange(1,start,1,count).shiftColumnGroupDepth(1);
      if(collapsed) sheet.getColumnGroup(start,1).collapse();
    } else {
      // Do not nest over existing custom groups.
      for(let col=start;col<=block.end;col++) if(sheet.getColumnGroupDepth(col)>0) return;
      sheet.getRange(1,start,1,count).shiftColumnGroupDepth(1);
    }
  });
}
