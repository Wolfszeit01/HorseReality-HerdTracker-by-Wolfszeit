/* ══════════════════════════════════════════════════════════════
  Pedidree Import
   ══════════════════════════════════════════════════════════════ */

function openPedigreeModal() {
  const tpl = HtmlService.createTemplateFromFile('pedigreeModal');
  const html = tpl.evaluate()
   .setWidth(1000)
    .setHeight(1200)
  SpreadsheetApp.getUi().showModalDialog(html,' ');
}

function openPedigreeModalWithData(pedigreeData) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('pendingPedigreeData', JSON.stringify(pedigreeData));
  const tpl = HtmlService.createTemplateFromFile('pedigreeModal');
  const html = tpl.evaluate()
    .setWidth(1000)
    .setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getPedigreeModalData() {
  try {
    const props       = PropertiesService.getScriptProperties();
    const pendingData = props.getProperty('pendingPedigreeData');
    if (!pendingData) return { success: false, error: "No pending data" };

    let data = JSON.parse(pendingData);

    if (data.damName && data.damName !== "Unknown") {
      const d = getExistingPedigreeInfo(data.damName);
      if (d) {
        data.damline = data.damline || d.damline;
        data.gs_m    = d.sire;
        data.gd_m    = d.dam;
        data.ggs_mp  = d.gs_p;
        data.ggd_mp  = d.gd_p;
        data.ggs_mm  = d.gs_m;
        data.ggd_mm  = d.gd_m;
      }
    }

    if (data.sireName && data.sireName !== "Unknown") {
      const s = getExistingPedigreeInfo(data.sireName);
      if (s) {
        data.stallionline = data.stallionline || s.stallionline;
        data.gs_p         = s.sire;
        data.gd_p         = s.dam;
        data.ggs_pp       = s.gs_p;
        data.ggd_pp       = s.gd_p;
        data.ggs_pm       = s.gs_m;
        data.ggd_pm       = s.gd_m;
      }
    }

    data.success = true;
    return data;
  } catch (e) {
    console.error("Error in getPedigreeModalData: " + e.message);
    return { success: false, error: e.message };
  }
}

function getKnownHorseNames() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const names = new Set();

  const herdSheet = ss.getSheetByName('Herd');
  if (herdSheet && herdSheet.getLastRow() >= 2) {
    herdSheet.getRange('D2:D' + herdSheet.getLastRow()).getValues().forEach(row => {
      if (row[0] && row[0].toString().trim())
        names.add(row[0].toString().trim().toLowerCase());
    });
  }

  const pedSheet = ss.getSheetByName('Pedigree');
  if (pedSheet && pedSheet.getLastRow() >= 2) {
    const pedData = pedSheet.getDataRange().getValues();
    for (let i = 1; i < pedData.length; i++) {
      for (let j = 2; j <= 18; j++) {
        if (pedData[i][j] && pedData[i][j].toString().trim())
          names.add(pedData[i][j].toString().trim().toLowerCase());
      }
    }
  }
  return names;
}

// ─── SHARED HELPERS ───────────────────────────────────────────────────────────

function _isPlaceholder(l) {
  return /^(Foundation Breeder|Unknown|n\/a)$/i.test(l.trim());
}

function _isStableName(l) {
  if (!l) return false;
  return /(Estate|Stables|Gardens|Meadows|Stuteri|Stable|Farm|Ranch|Stud\b|Park|Acres|Solitude|National|Breeding\s+Stud|hevostalli|hevostila|Academy|Centre|Center|Ridge|Valley|Grove)/i.test(l);
}

function _isTagline(l) {
  if (!l) return false;
  if (/^\d+\s*[\|\/]/.test(l))                                               return true; // 835 | ...
  if (/\bGP\d{3,}\b/i.test(l))                                               return true; // GP817
  if (/\{GP\d+\}/i.test(l))                                                  return true; // {GP828}
  if (/\d+VG\b/i.test(l))                                                    return true; // 12VG
  if (/\d+:\d+/.test(l))                                                     return true; // 12:0
  // Score-Zahl nur als Tagline wenn Zeile NICHT mit Buchstabe beginnt
  if (/\d{2,}\.\d{1,}/.test(l) && !/^[A-Za-zÀ-öø-ÿ\!$]/.test(l.trim()))   return true; // 89.469 solo
  if (/\[\d+[\.\,]\d+\]/.test(l))                                            return true; // [96.461]
  if (/\d+,\d+\/\d+/.test(l))                                                return true; // 0,3/6
  // Reine Zahl: NUR Tagline wenn > 1 Ziffer (einstellige = Pferdename wie "1")
  if (/^\d{2,}$/.test(l))                                                    return true; // 831
  if (/\d{3,}\s+\d+\s+\d/.test(l))                                          return true; // 756 12 93.5
  if (/^[A-Z]{2,3}\d{3,}/.test(l))                                          return true; // EN627
  if (/\d+[A-Z]{1,3}[I|]\d+[A-Z]{1,3}/i.test(l))                           return true; // 651GPI12VGI
  if (/︱|︲/.test(l))                                                         return true;
  if (/FULLYTRAINED|BETA/i.test(l))                                          return true;
  if (/^(Driving|Endurance|Training|Dressur|Dressage|Jumping|Western|Eventing)\b/i.test(l)) return true;
  if (/^\s*\|/.test(l))                                                      return true;
  if (/GP\d+\s*\|/.test(l))                                                  return true; // GP756 | ...
  // Reine Genetics-Zeile: "11VG 660 A Z SW1"
  if (/^\d+\s*(VG|G\+?|A|BA|P)\b/i.test(l))                                return true;
  return false;
}

function _normalizeDecoName(l) {
  return l
    .replace(/[ℬℌℛℐℑℒℓ]/g, m => ({'ℬ':'B','ℌ':'H','ℛ':'R','ℐ':'I','ℑ':'I','ℒ':'L','ℓ':'l'}[m] || m))
    .replace(/[ᴀʙᴄᴅᴇꜰɢʜɪᴊᴋʟᴍɴᴏᴘǫʀꜱᴛᴜᴠᴡxʏᴢᴸᴿ]/g, c => c.normalize('NFKD')[0] || c)
    .replace(/[σ]/g, 's');
}

function _cleanName(name) {
  if (!name) return '';
  let n = name;

  // "!94.453 ℘ Name" oder "$92.7 Name" → Score-Präfix strippen
  n = n.replace(/^[!$]?\d+[\.\,]\d+\s*[-–]?\s*\d*[\.\,]?\d*\s*[℘\s]*/,'').trim();

  // "95.029︱Name" → "Name"
  const scorePrefix = n.match(/^\d+[\.\,]\d+\s*[|︱]\s*(.+)$/);
  if (scorePrefix) n = scorePrefix[1];

  // Deceased
  n = n.replace(/Deceased/gi, '');

  // Führende Deko-Zeichen
  n = n.replace(/^[★☆✶✧˚ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃〔〕♟˚₊✧⋆°·•﹚ℙ♕⭐️♡♥*\-͟͞➳❥༅\s]+/, '').trim();

  // "Name || SU" → "Name"
  n = n.replace(/\s*\|\|.*$/, '').trim();

  // "Name (3.2)" → "Name"
  n = n.replace(/\s*\(\d+[\.\,]\d+\)\s*$/, '').trim();

  // "Name 92.365" → "Name"  (trailing score ≥ 3 Dezimalstellen)
  n = n.replace(/\s+\d+[\.\,]\d{3,}\s*$/, '').trim();

  // "Name | 92.3" oder "Name︱..."
  n = n.replace(/\s*[|︱︲\/]\s*[\d\s].*$/, '').trim();

  // Tagline-Anhängsel nach Pipe: "Name|12VG"
  n = n.split(/Tagline/i)[0].trim();

  // Trailing Deko
  n = n.replace(/[★☆✶✧˚ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃〔〕♟˚₊✧⋆°·•﹚⭐️♡♥\s]+$/, '').trim();

  return n; // "" erlaubt
}

/**
 * Gemeinsame Slot-Extraktion für getPreviewData + parsePedigreeString.
 *
 * Game slot mapping (14 slots):
 * 0=Sire, 1=GS(P), 2=GGS(PP), 3=GGD(PP), 4=GD(P), 5=GGS(PM), 6=GGD(PM)
 * 7=Dam,  8=GS(M), 9=GGS(MP), 10=GGD(MP), 11=GD(M), 12=GGS(MM), 13=GGD(MM)
 */
function _extractPedigreeSlots(pedLines) {
  const slots = [];
  let i = 0;

  while (i < pedLines.length && slots.length < 14) {
    const line = pedLines[i];

    // Placeholder → leerer Slot, 1 Zeile vorwärts
    if (_isPlaceholder(line)) {
      slots.push('');
      i += 1;
      continue;
    }

    // Stable ohne vorangehenden Namen → überspringen
    if (_isStableName(line)) {
      i += 1;
      continue;
    }

    // Aktuelle Zeile = Name (NIE per _isTagline filtern!)
    const cleaned = _cleanName(_normalizeDecoName(line));
    slots.push(cleaned); // auch "" oder "1" erlaubt

    // Nächste Zeile bestimmt Schrittweite
    const nextLine = pedLines[i + 1] || '';
    i += _isTagline(nextLine) ? 3 : 2;
  }

  return slots;
}

function _slotsToResult(slots) {
  const g = (idx) => slots[idx] || '';
  return {
    sire:   g(0),  dam:    g(7),
    gs_p:   g(1),  gd_p:   g(4),
    gs_m:   g(8),  gd_m:   g(11),
    ggs_pp: g(2),  ggd_pp: g(3),
    ggs_pm: g(5),  ggd_pm: g(6),
    ggs_mp: g(9),  ggd_mp: g(10),
    ggs_mm: g(12), ggd_mm: g(13),
  };
}

// ─── parsePedigreeString ──────────────────────────────────────────────────────

function parsePedigreeString(rawText) {
  try {
    if (!rawText || rawText.trim() === '') {
      return { success: false, error: "No text provided" };
    }

    const pedSplit = rawText.split(/Pedigree/i);
    const pedRaw   = pedSplit.length > 1 ? pedSplit[1] : rawText;
    const pedText  = pedRaw.split(/Pregnancy|This page was/i)[0];

    const pedLines = pedText
      .split(/\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0 && !l.match(/^COI:/i));

    const result = _slotsToResult(_extractPedigreeSlots(pedLines));
    result.success = true;
    return result;
  } catch (e) {
    console.error('parsePedigreeString: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ─── getExistingPedigreeInfo ──────────────────────────────────────────────────

function getExistingPedigreeInfo(nameOrId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd');
  const pedSheet  = ss.getSheetByName('Pedigree');

  if (!nameOrId) return null;

  const clean = (s) => s ? s.toString().toLowerCase()
    .replace(/[★☆✶✧˚ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃〔〕♟˚₊✧⋆°·•﹚⭐️♡♥ℙ♕*#\s\-\|➳❥]+/g, '').trim() : '';

  const searchStr = clean(nameOrId);
  const searchId  = nameOrId.toString().trim();

  let result = {
    id: '', name: '',
    sire: '', dam: '', stallionline: '', damline: '',
    gs_p: '', gd_p: '', gs_m: '', gd_m: ''
  };

  if (herdSheet) {
    const herdData = herdSheet.getDataRange().getValues();
    for (let i = 1; i < herdData.length; i++) {
      const hId   = herdData[i][2] ? herdData[i][2].toString().trim() : '';
      const hName = herdData[i][3] ? herdData[i][3].toString() : '';
      if (hId === searchId || clean(hName) === searchStr) {
        result.id   = hId;
        result.name = hName;
        break;
      }
    }
  }

  if (pedSheet) {
    const pedData = pedSheet.getDataRange().getValues();
    for (let i = 1; i < pedData.length; i++) {
      const pId   = pedData[i][1] ? pedData[i][1].toString().trim() : '';
      const pName = pedData[i][2] ? pedData[i][2].toString() : '';
      if ((result.id && pId === result.id) || clean(pName) === searchStr) {
        if (!result.id)   result.id   = pId;
        if (!result.name) result.name = pName;
        // B(1)=ID, C(2)=Name, D(3)=Sireline, E(4)=Damline
        // F(5)=Sire, G(6)=Dam
        // H(7)=GS(P), I(8)=GD(P), J(9)=GS(M), K(10)=GD(M)
        result.stallionline = pedData[i][4]  || '';
        result.damline      = pedData[i][5]  || '';
        result.sire         = pedData[i][6]  || '';
        result.dam          = pedData[i][7]  || '';
        result.gs_p         = pedData[i][8]  || '';
        result.gd_p         = pedData[i][9]  || '';
        result.gs_m         = pedData[i][10]  || '';
        result.gd_m         = pedData[i][12] || '';
        break;
      }
    }
  }

  return (result.id || result.name) ? result : null;
}

// ─── getPedigreeAutocompleteData ──────────────────────────────────────────────

function getPedigreeAutocompleteData() {
  try {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    let horseNames = [];
    let stallionlines = [];
    let damlines   = [];

    const herdSheet = ss.getSheetByName('Herd');
    if (herdSheet && herdSheet.getLastRow() >= 2) {
      herdSheet.getRange('D2:D' + herdSheet.getLastRow()).getValues().forEach(row => {
        if (row[0] && row[0].toString().trim() !== '')
          horseNames.push(row[0].toString().trim());
      });
    }

    const pedigreeSheet = ss.getSheetByName('Pedigree');
    if (pedigreeSheet && pedigreeSheet.getLastRow() >= 2) {
      pedigreeSheet.getRange('C2:F' + pedigreeSheet.getLastRow()).getValues().forEach(row => {
        // C = Name, E = Stallionline, F = Damline
        if (row[0] && row[0].toString().trim() !== '') {
          const name = row[0].toString().trim();
          if (!horseNames.includes(name)) horseNames.push(name);
        }
        if (row[2] && row[2].toString().trim() !== '') {
          stallionlines.push(row[2].toString().trim());
        }
        if (row[3] && row[3].toString().trim() !== '') {
          damlines.push(row[3].toString().trim());
        }
      });
    }

    return {
      success: true,
      horseNames: [...new Set(horseNames)].sort(),
      stallionlines: [...new Set(stallionlines)].sort(),
      damlines: [...new Set(damlines)].sort()
    };
  } catch (error) {
    return { success: false, error: error.message, horseNames: [], stallionlines: [], damlines: [] };
  }
}

// ─── saveManualPedigree ───────────────────────────────────────────────────────

// ─── saveManualPedigree ───────────────────────────────────────────────────────
function saveManualPedigree(pedigreeData) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pedigree');
    if (!sheet) throw new Error('Sheet "Pedigree" not found!');

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;

    const searchId   = pedigreeData.id   ? pedigreeData.id.toString().trim()                : 'NO_ID';
    const searchName = pedigreeData.name ? pedigreeData.name.toString().toLowerCase().trim() : 'NO_NAME';

    for (let i = 1; i < data.length; i++) {
      const rowId   = data[i][1] ? data[i][1].toString().trim()               : '';
      const rowName = data[i][2] ? data[i][2].toString().toLowerCase().trim() : '';
      if ((searchId !== 'NO_ID' && rowId === searchId) || rowName === searchName) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

    const rowValues = [
      pedigreeData.id           || '', // B (Spalte 2)
      pedigreeData.name         || '', // C (Spalte 3)
      pedigreeData.coi          || '', // D (Spalte 4)
      pedigreeData.stallionline || '', // E (Spalte 5)
      pedigreeData.damline      || '', // F (Spalte 6)
      pedigreeData.sire         || '', // G (Spalte 7)
      pedigreeData.dam          || '', // H (Spalte 8)
      pedigreeData.gs_p         || '', // I (Spalte 9)
      pedigreeData.gd_p         || '', // J (Spalte 10)
      pedigreeData.gs_m         || '', // K (Spalte 11)
      pedigreeData.gd_m         || '', // L (Spalte 12)
      pedigreeData.ggs_pp       || '', // M (Spalte 13)
      pedigreeData.ggd_pp       || '', // N (Spalte 14)
      pedigreeData.ggs_pm       || '', // O (Spalte 15)
      pedigreeData.ggd_pm       || '', // P (Spalte 16)
      pedigreeData.ggs_mp       || '', // Q (Spalte 17)
      pedigreeData.predicates   || '', // R (Spalte 18)
      pedigreeData.ggs_mm       || '', // S (Spalte 19)
      pedigreeData.ggd_mm       || ''  // T (Spalte 20)
    ];

    sheet.getRange(targetRow, PEDIGREE_COLS.ID).setValue(pedigreeData.id || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.NAME).setValue(pedigreeData.name || '');
    // COI NICHT schreiben - behalten
    sheet.getRange(targetRow, PEDIGREE_COLS.STALLIONLINE).setValue(pedigreeData.stallionline || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.DAMLINE).setValue(pedigreeData.damline || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.SIRE).setValue(pedigreeData.sire || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.DAM).setValue(pedigreeData.dam || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GS_P).setValue(pedigreeData.gs_p || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GD_P).setValue(pedigreeData.gd_p || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GS_M).setValue(pedigreeData.gs_m || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GD_M).setValue(pedigreeData.gd_m || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_PP).setValue(pedigreeData.ggs_pp || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_PP).setValue(pedigreeData.ggd_pp || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_PM).setValue(pedigreeData.ggs_pm || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_PM).setValue(pedigreeData.ggd_pm || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_MP).setValue(pedigreeData.ggs_mp || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_MP).setValue(pedigreeData.ggd_mp || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_MM).setValue(pedigreeData.ggs_mm || '');
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_MM).setValue(pedigreeData.ggd_mm || '');
    PropertiesService.getScriptProperties().deleteProperty('pendingPedigreeData');
    return { success: true, row: targetRow };

  } catch (error) {
    throw new Error('Error: ' + error.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// NEU: Speichert Pedigree UND führt vollständigen Import durch
// ═══════════════════════════════════════════════════════════════
function saveAndImportPedigree(fullData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pedigreeData = fullData.pedigreeData || fullData;
    const opts = fullData.opts || { herd: true, stats: true, colour: true, scores: true };
    const manualData = fullData.manualData || {};
    
    // 1. Pedigree im Sheet speichern
    const pedResult = saveManualPedigree(pedigreeData);
    
    // 2. Vollständigen Import durchführen mit allen Daten
    const data = {
      id: pedigreeData.id,
      name: pedigreeData.name,
      breed: pedigreeData.breed,
      gender: pedigreeData.gender,
      dob: pedigreeData.dob,
      predicates: pedigreeData.predicates,
      agedWithDP: pedigreeData.agedWithDP,
      sire: pedigreeData.sire,
      dam: pedigreeData.dam,
      sireIsInside: true,  // Nach Save sind sie Inside
      damIsInside: true,
      gs_p: pedigreeData.gs_p || "",
      gd_p: pedigreeData.gd_p || "",
      gs_m: pedigreeData.gs_m || "",
      gd_m: pedigreeData.gd_m || "",
      coi: pedigreeData.coi || "",
      gp: pedigreeData.gp || {},
      confo: pedigreeData.confo || {},
      genetics: pedigreeData.genetics || {},
      health: pedigreeData.health || {},
      achieve: pedigreeData.achieve || {},
      showScores: pedigreeData.showScores || [],
      compScores: pedigreeData.compScores || [],
      pregnancy: pedigreeData.pregnancy || null
    };
    
    const results = [];
    const convertedDOB = _convertDOB(data.dob);
    
    if (opts.herd) {
      const r = _writeHerdTracker(data, convertedDOB);
      if (r) results.push(r);
    }

    if (opts.stats) {
      const r = _writeStats(ss, data);
      if (r) results.push(r);
    }

    if (opts.colour) {
      const r = _writeColourGenetics(ss, data, manualData);
      if (r) results.push(r);
    }

    if (opts.scores) {
      results = results.concat(_writeScores(ss, data));
    }

    // Pedigree ist bereits gespeichert oben, aber hier auch referenzieren
    const pedRef = "✓ Pedigree saved";
    results.unshift(pedRef);

    return {
      success: true,
      message: "✓ Complete Import!\n" + results.join("\n")
    };
    
  } catch (error) {
    throw new Error("Import error: " + error.message);
  }
}

// ═══════════════════════════════════════════════════════════════
// Speichert nur Pedigree (Rest ist schon beim Import gespeichert)
// Wird vom Pedigree-Modal aufgerufen nach outside-Eltern-Input
// ═══════════════════════════════════════════════════════════════
function savePedigreeOnly(params) {
  try {
    const pedigreeData = params.pedigreeData || params;
    
    // Pedigree-Zeile speichern
    const pedSaveResult = saveManualPedigree(pedigreeData);
    
    // Properties clearen
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('pendingPedigreeData');
    
    return {
      success: true,
      message: "✓ Pedigree saved!"
    };
  } catch (error) {
    throw new Error("Save error: " + error.message);
  }
}