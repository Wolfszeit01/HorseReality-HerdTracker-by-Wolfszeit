/**
 * Zentrale Konfiguration aller Import-Profile.
 * Neuer CSV-Tab? Einfach hier einen neuen Eintrag hinzufügen -
 * der Rest des Scripts bleibt unverändert.
 *
 * headerRowIndex: 0-basierter Index der Zeile im CSV, die die echten
 *                 Spaltennamen enthält (0 = erste Zeile, 1 = zweite Zeile usw.)
 * idField:        Name der Spalte im CSV, die als eindeutige ID dient
 *                 (Zeilen ohne Wert hier werden beim Import übersprungen)
 * matchColumn:    Spalte im Zielsheet, die auf "leer" geprüft wird,
 *                 um die nächste freie Zeile zu finden
 * mapping:        CSV-Spaltenname -> Zielspalte (Buchstabe) im Zielsheet
 */


function showCsvImportDialog() {
  var tpl = HtmlService.createTemplateFromFile("CsvImportDialog");
  var html = tpl.evaluate()
    .setWidth(600)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, " ");
}
var IMPORT_PROFILES = {
  herd: {
    label: "Herd (Status, ID, Name, Breed, ...)",
    targetSheet: "Herd",
    headerRowIndex: 0,
    idField: "ID",
    matchColumn: "C",
    mapping: {
      "Status": "B",
      "ID": "C",
      "Name": "D",
      "Breed": "E",
      "Gender": "F",
      "DOB": "G",
      "Times Aged": "H",
      "Project": "K",
      "Role": "Q",
      "Predicates": "R",
      "Training": "Z",
      "Colic Res.": "AD",
      "Hoof Quality": "AE",
      "Back Problems": "AF",
      "Resp. Disease": "AG",
      "Lameness Res.": "AH",
      "Fertility": "AI"
    }
  },

  stats: {
    label: "Stats (Conformation, Abilities, Placements)",
    targetSheet: "Stats",
    headerRowIndex: 1, // Zeile 2 im CSV enthält die echten Header
    idField: "ID",
    matchColumn: "B",
    mapping: {
      "ID": "B",
      "Name": "C",
      "WLK": "D",
      "TRT": "E",
      "CANT": "F",
      "GLP": "G",
      "POST": "H",
      "HD": "I",
      "NCK": "J",
      "BCK": "K",
      "SHLD": "L",
      "FRTL": "M",
      "HQU": "N",
      "SCK": "O",
      "ACC": "AI",
      "AGI": "AJ",
      "BALA": "AK",
      "BASC": "AL",
      "PULL": "AM",
      "SPEED": "AN",
      "SPR": "AO",
      "STA": "AP",
      "STRE": "AQ",
      "SURE": "AR",
      "Day Champ": "AX",
      "1st": "AY",
      "2nd": "AZ",
      "3rd": "BA",
      "First": "BS",
      "Second": "BT",
      "Third": "BU"
    }
  },

  ice_stats: {
    label: "ICE Stats (Islandpferd: Tölt/Pace)",
    targetSheet: "ICE_Stats",
    headerRowIndex: 1, // Zeile 2 im CSV enthält die echten Header
    idField: "ID",
    matchColumn: "B",
    mapping: {
      "ID": "B",
      "Name": "C",
      "WLK": "D",
      "TRT": "E",
      "CANT": "F",
      "GLP": "G",
      "Tölt": "H",
      "Pace": "I",
      "POST": "J",
      "HD": "K",
      "NCK": "L",
      "BCK": "M",
      "SHLD": "N",
      "FRTL": "O",
      "HQU": "P",
      "SCK": "Q",
      "ACC": "AM",
      "AGI": "AN",
      "BALA": "AO",
      "BASC": "AP",
      "PULL": "AQ",
      "SPEED": "AR",
      "SPR": "AS",
      "STA": "AT",
      "STRE": "AU",
      "SURE": "AV",
      "Day Champ": "BB",
      "1st": "BC",
      "2nd": "BD",
      "3rd": "BE",
      "First": "BW",
      "Second": "BX",
      "Third": "BY"
    }
  },

  kath_stats: {
    label: "KATH Stats (Kathiawari: REVAAL)",
    targetSheet: "KATH_Stats",
    headerRowIndex: 1, // Zeile 2 im CSV enthält die echten Header
    idField: "ID",
    matchColumn: "B",
    mapping: {
      "ID": "B",
      "Name": "C",
      "WLK": "D",
      "TRT": "E",
      "CANT": "F",
      "GLP": "G",
      "REVAAL": "H",
      "POST": "I",
      "HD": "J",
      "NCK": "K",
      "BCK": "L",
      "SHLD": "M",
      "FRTL": "N",
      "HQU": "O",
      "SCK": "P",
      "ACC": "AK",
      "AGI": "AL",
      "BALA": "AM",
      "BASC": "AN",
      "PULL": "AO",
      "SPEED": "AP",
      "SPR": "AQ",
      "STA": "AR",
      "STRE": "AS",
      "SURE": "AT",
      "Day Champ": "AZ",
      "1st": "BA",
      "2nd": "BB",
      "3rd": "BC",
      "First": "BU",
      "Second": "BV",
      "Third": "BW"
    }
  },

  pedigree: {
    label: "Pedigree (Sireline, Damline, Sire, Dam, ...)",
    targetSheet: "Pedigree",
    headerRowIndex: 0,
    idField: "ID",
    matchColumn: "B",
    mapping: {
      "ID": "B",
      "Name": "C",
      "Sireline": "E",
      "Damline": "F",
      "Sire": "G",
      "Dam": "H",
      "GS (P)": "I",
      "GD (P)": "J",
      "GS (M)": "K",
      "GD (M)": "L",
      "GGS (PP)": "M",
      "GGD (PP)": "N",
      "GGS (PM)": "O",
      "GGD (PM)": "P",
      "GGS (MP)": "Q",
      "GGD (MP)": "R",
      "GGS (MM)": "S",
      "GGD (MM)": "T"
    }
  },

  colour_genetics: {
    label: "Colour Genetics",
    targetSheet: "Colour Genetics",
    headerRowIndex: 0,
    idField: "ID",
    matchColumn: "B",
    mapping: {
      "ID": "B",
      "Name": "C",
      "Ext.": "D",
      "Ag.": "E",
      "Grey": "F",
      "Cr/Prl": "G",
      "Dun/Ps.Dun": "H",
      "Chmpgn": "I",
      "Slvr": "J",
      "Mush.": "K",
      "Overo": "L",
      "Leop.": "M",
      "PATN1": "N",
      "PATN2": "O",
      "SW1/SW3": "P",
      "SW 2": "Q",
      "SW4": "R",
      "SW5": "S",
      "SW6": "T",
      "SW7": "U",
      "Tobiano": "V",
      "Roan": "W",
      "Sab1": "X",
      "W1": "Y",
      "W2": "Z",
      "W3": "AA",
      "W4": "AB",
      "W5": "AC",
      "W6": "AD",
      "W7": "AE",
      "W8": "AF",
      "9": "AG",
      "W10": "AH",
      "11": "AI",
      "12": "AJ",
      "13": "AK",
      "14": "AL",
      "15": "AM",
      "W16": "AN",
      "17": "AO",
      "18": "AP",
      "W19": "AQ",
      "W20": "AR",
      "W21": "AS",
      "W22": "AT",
      "W23": "AU",
      "Flaxen": "AV",
      "Sooty": "AW",
      "Pangaré": "AX",
      "Hidden Sab / White Spotting": "AY",
      "Rab": "AZ",
      "White Markings": "BA"
    }
  },

  conf_results: {
    // Andere Struktur als die übrigen Profile: Pferde stehen in SPALTEN,
    // nicht in Zeilen. Wird von processCsvText erkannt (type: "columnar")
    // und an importColumnarProfile() weitergereicht.
    label: "Conf. Results (Score-Historie, Pferde in Spalten)",
    type: "columnar",
    targetSheet: "Conf. Results",
    idRow: 1,               // Zeile im Zielsheet, die die ID enthält
    nameRow: 2,              // Zeile im Zielsheet, die den Namen enthält
    dataStartCol: 3,         // Spalte C = erste Pferde-Spalte (A leer, B = Labels)
    csvIdRowIndex: 0,        // 0-basierter Index der ID-Zeile im CSV
    csvNameRowIndex: 1,      // 0-basierter Index der Name-Zeile im CSV
    csvHistoryStartRow1based: 6, // ab hier (1-basiert im CSV) beginnt die echte Score-Historie
    historyRowOffset: 2,     // Zielzeile = CSV-Zeile(1-basiert) + historyRowOffset
    protectedRows: [3, 4, 5, 6, 7] // MIN/AVG/MAX/RANGE/GAP - niemals überschreiben
  },

  comp_results: {
    label: "Comp. Results (Wettkampf-Historie, Pferde in Spalten)",
    type: "columnar",
    targetSheet: "Comp. Results",
    idRow: 1,
    nameRow: 2,
    dataStartCol: 3,
    csvStartCol: 1,          // Spalte A im CSV ("Highlight") wird übersprungen
    csvIdRowIndex: 0,
    csvNameRowIndex: 1,
    csvHistoryStartRow1based: 3, // Werte-Historie beginnt direkt nach ID/Name
    historyRowOffset: 4,     // Zielzeile = CSV-Zeile(1-basiert) + 4  ->  CSV-Zeile 3 = Zielzeile 7
    protectedRows: [3, 4, 5, 6] // MIN/AVG/MAX/RANGE - kein GAP hier
  }
};


function getImportProfiles() {
  var list = [];
  for (var key in IMPORT_PROFILES) {
    list.push({ key: key, label: IMPORT_PROFILES[key].label });
  }
  return list;
}


function processCsvText(csvText, profileKey) {
  if (!csvText || csvText.trim() === "") {
    throw new Error("Es wurde kein CSV-Text übergeben.");
  }

  var profile = IMPORT_PROFILES[profileKey];
  if (!profile) {
    throw new Error('Unbekanntes Profil: "' + profileKey + '"');
  }

  if (profile.type === "columnar") {
    return importColumnarProfile(csvText, profile);
  }

  var parsed = Utilities.parseCsv(csvText);
  if (parsed.length <= profile.headerRowIndex + 1) {
    throw new Error("CSV enthält keine Datenzeilen.");
  }

  var headers = parsed[profile.headerRowIndex];
  var srcRows = parsed.slice(profile.headerRowIndex + 1);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = ss.getSheetByName(profile.targetSheet);
  if (!dstSheet) throw new Error('Sheet "' + profile.targetSheet + '" wurde nicht gefunden.');

  function colIndex(name) {
    var idx = headers.indexOf(name);
    if (idx === -1) throw new Error('Spalte "' + name + '" wurde im CSV-Header nicht gefunden.');
    return idx;
  }

  var fieldIndices = {};
  for (var field in profile.mapping) {
    fieldIndices[field] = colIndex(field);
  }
  var idColIndex = colIndex(profile.idField);

  // matchColumn-Buchstabe -> Spaltennummer für getRange
  var matchColNum = columnLetterToNumber(profile.matchColumn);

  // Leere Zeilen im Zielsheet finden (matchColumn leer), ab Zeile 2
  var dstLastRow = Math.max(dstSheet.getLastRow(), 2);
  var matchValues = dstSheet.getRange(2, matchColNum, dstLastRow - 1, 1).getValues();

  var emptyRows = [];
  for (var i = 0; i < matchValues.length; i++) {
    var v = matchValues[i][0];
    if (v === "" || v === null) {
      emptyRows.push(2 + i);
    }
  }

  var nextAppendRow = dstLastRow + 1;
  var writtenCount = 0;
  var appendedCount = 0;
  var skippedCount = 0;

  for (var r = 0; r < srcRows.length; r++) {
    var row = srcRows[r];

    var idValue = row[idColIndex];
    if (idValue === "" || idValue === null || idValue === undefined) {
      skippedCount++;
      continue;
    }

    var targetRow;
    if (writtenCount < emptyRows.length) {
      targetRow = emptyRows[writtenCount];
      writtenCount++;
    } else {
      targetRow = nextAppendRow;
      nextAppendRow++;
      appendedCount++;
    }

    for (var f in profile.mapping) {
      var colLetter = profile.mapping[f];
      var value = row[fieldIndices[f]];
      dstSheet.getRange(colLetter + targetRow).setValue(value);
    }
  }

  return {
    written: writtenCount,
    appended: appendedCount,
    skipped: skippedCount
  };
}

/**
 * Import-Logik für "columnar" Profile, bei denen Pferde in SPALTEN stehen
 * (z.B. Conf. Results). Matching erfolgt über die ID-Zeile im Zielsheet.
 * Rows in profile.protectedRows werden NIE beschrieben, auch nicht bei
 * neu angelegten Spalten.
 */
function importColumnarProfile(csvText, profile) {
  var parsed = Utilities.parseCsv(csvText);
  if (parsed.length <= profile.csvIdRowIndex) {
    throw new Error("CSV enthält keine ID-Zeile.");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dstSheet = ss.getSheetByName(profile.targetSheet);
  if (!dstSheet) throw new Error('Sheet "' + profile.targetSheet + '" wurde nicht gefunden.');

  var protectedRows = {};
  (profile.protectedRows || []).forEach(function (r) { protectedRows[r] = true; });

  // Bestehende ID-Zeile im Zielsheet einlesen -> Spalten-Zuordnung
  var dstLastCol = Math.max(dstSheet.getLastColumn(), profile.dataStartCol - 1);
  var idRowValues = dstSheet.getRange(profile.idRow, 1, 1, dstLastCol).getValues()[0];

  var idToCol = {};
  var emptyCols = []; // Spalten im bestehenden Bereich, die noch keine ID haben (Lücken)
  for (var c = profile.dataStartCol - 1; c < idRowValues.length; c++) {
    var v = idRowValues[c];
    if (v !== "" && v !== null && v !== undefined) {
      idToCol[String(v).trim()] = c + 1; // 1-basierte Spaltennummer
    } else {
      emptyCols.push(c + 1);
    }
  }

  var nextAppendCol = dstLastCol + 1;
  var emptyColPointer = 0;

  var csvCols = parsed[profile.csvIdRowIndex].length;
  var updatedCols = 0;
  var addedCols = 0;
  var historyValuesWritten = 0;
  var skippedCols = 0;

  var historyStartIdx = profile.csvHistoryStartRow1based - 1; // 0-basiert
  var csvStartCol = profile.csvStartCol || 0; // manche CSVs haben eine nicht-relevante erste Spalte

  for (var col = csvStartCol; col < csvCols; col++) {
    var idVal = parsed[profile.csvIdRowIndex][col];
    if (idVal === "" || idVal === null || idVal === undefined) {
      skippedCols++;
      continue;
    }
    idVal = String(idVal).trim();
    var nameVal = parsed[profile.csvNameRowIndex] ? parsed[profile.csvNameRowIndex][col] : "";

    var targetCol;
    if (idToCol.hasOwnProperty(idVal)) {
      targetCol = idToCol[idVal];
      updatedCols++;
    } else if (emptyColPointer < emptyCols.length) {
      targetCol = emptyCols[emptyColPointer];
      emptyColPointer++;
      idToCol[idVal] = targetCol;
      addedCols++;

      if (!protectedRows[profile.idRow]) {
        dstSheet.getRange(profile.idRow, targetCol).setValue(idVal);
      }
      if (!protectedRows[profile.nameRow]) {
        dstSheet.getRange(profile.nameRow, targetCol).setValue(nameVal);
      }
    } else {
      targetCol = nextAppendCol;
      nextAppendCol++;
      idToCol[idVal] = targetCol;
      addedCols++;

      // Neue Spalte: ID und Name eintragen (Zeilen 1/2 sind nicht geschützt)
      if (!protectedRows[profile.idRow]) {
        dstSheet.getRange(profile.idRow, targetCol).setValue(idVal);
      }
      if (!protectedRows[profile.nameRow]) {
        dstSheet.getRange(profile.nameRow, targetCol).setValue(nameVal);
      }
    }

    // Score-Historie ab csvHistoryStartRow1based übertragen
    for (var r = historyStartIdx; r < parsed.length; r++) {
      var rawVal = parsed[r][col];
      if (rawVal === "" || rawVal === undefined) continue; // leere CSV-Zellen nicht schreiben

      var targetRow = (r + 1) + profile.historyRowOffset; // r ist 0-basiert -> +1 = 1-basierte CSV-Zeile
      if (protectedRows[targetRow]) continue; // harte Sperre für Formel-Zeilen

      dstSheet.getRange(targetRow, targetCol).setValue(rawVal);
      historyValuesWritten++;
    }
  }

  return {
    written: updatedCols,
    appended: addedCols,
    skipped: skippedCols,
    historyValuesWritten: historyValuesWritten
  };
}

/**
 * Hilfsfunktion: Spaltenbuchstabe (z.B. "AI") -> Spaltennummer (z.B. 35)
 */
function columnLetterToNumber(letters) {
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col;
}