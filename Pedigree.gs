/***********************
 * PEDIGREE MODAL BACKEND — FINAL
 ***********************/

function openPedigreeModal() {
  const html = HtmlService.createHtmlOutputFromFile('pedigreeModal')
    .setWidth(620)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pedigree Editor');
}

function openPedigreeModalWithData(pedigreeData) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('pendingPedigreeData', JSON.stringify(pedigreeData));
  const html = HtmlService.createHtmlOutputFromFile('pedigreeModal')
    .setWidth(620)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Complete Pedigree');
}

function getPedigreeModalData() {
  try {
    const props       = PropertiesService.getScriptProperties();
    const pendingData = props.getProperty('pendingPedigreeData');
    if (!pendingData) return { success: false, error: "No pending data" };

    let data = JSON.parse(pendingData);

    // --- DEEP SEARCH: DAM ---
    if (data.damName && data.damName !== "Unknown") {
      const d = getExistingPedigreeInfo(data.damName);
      if (d) {
        data.damline = data.damline || d.damline;
        data.gs_m    = d.sire;   // Dam's Sire
        data.gd_m    = d.dam;    // Dam's Dam
        // Great-grandparents via Dam's Sire (MP)
        data.ggs_mp  = d.gs_p;   // Dam's Sire's Sire
        data.ggd_mp  = d.gd_p;   // Dam's Sire's Dam
        // Great-grandparents via Dam's Dam (MM)
        data.ggs_mm  = d.gs_m;   // Dam's Dam's Sire
        data.ggd_mm  = d.gd_m;   // Dam's Dam's Dam
      }
    }

    // --- DEEP SEARCH: SIRE ---
    if (data.sireName && data.sireName !== "Unknown") {
      const s = getExistingPedigreeInfo(data.sireName);
      if (s) {
        data.stallionline = data.stallionline || s.stallionline;
        data.gs_p         = s.sire;   // Sire's Sire
        data.gd_p         = s.dam;    // Sire's Dam
        // Great-grandparents via Sire's Sire (PP)
        data.ggs_pp       = s.gs_p;   // Sire's Sire's Sire
        data.ggd_pp       = s.gd_p;   // Sire's Sire's Dam
        // Great-grandparents via Sire's Dam (PM)
        data.ggs_pm       = s.gs_m;   // Sire's Dam's Sire
        data.ggd_pm       = s.gd_m;   // Sire's Dam's Dam
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

  const herdSheet = ss.getSheetByName('Herd Tracker');
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

/**
 * Parses a pedigree string from the game.
 *
 * Primary:  3-line blocks (Name / Tagline / Stable)
 * Fallback: 2-line blocks (Name / Stable) when no tagline detected
 * Special:  1-line for placeholders (Foundation Breeder / Unknown)
 *
 * Game slot structure (14 slots):
 * 0=Sire, 1=GS(P), 2=GGS(PP), 3=GGD(PP), 4=GD(P), 5=GGS(MP), 6=GGD(MP)
 * 7=Dam,  8=GS(M), 9=GGS(PM), 10=GGD(PM), 11=GD(M), 12=GGS(MM), 13=GGD(MM)
 */
function parsePedigreeString(rawText) {
  try {
    const result = {
      sire: '', dam: '',
      gs_p: '', gd_p: '', gs_m: '', gd_m: '',
      ggs_pp: '', ggd_pp: '', ggs_mp: '', ggd_mp: '',
      ggs_pm: '', ggd_pm: '', ggs_mm: '', ggd_mm: ''
    };

    if (!rawText || rawText.trim() === '') return result;

    const pedSplit = rawText.split(/Pedigree/i);
    const pedRaw   = pedSplit.length > 1 ? pedSplit[1] : rawText;
    const pedText  = pedRaw.split(/Pregnancy|This page was/i)[0];

    const pedLines = pedText
      .split(/\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0 && !l.match(/^COI:/i));

    const isPlaceholder = (l) =>
      /^(Foundation Breeder|Unknown|n\/a)$/i.test(l.trim());

    const isTagline = (l) => {
      if (!l) return false;
      if (/^\d+\s*[\|\/]/.test(l))                    return true; // "835 | ..."
      if (/\bGP\d{3,}\b/i.test(l))                    return true; // GP817
      if (/\{GP\d+\}/i.test(l))                       return true; // {GP828}
      if (/\d+VG\b/.test(l))                          return true; // "12VG"
      if (/\d+:\d+/.test(l))                          return true; // "12:0" or "680:11"
      if (/\d{2,}\.\d{1,}/.test(l))                   return true; // "78.3" or "96.352"
      if (/\[\d+[\.\,]\d+\]/.test(l))                 return true; // "[96.461]"
      if (/\d+,\d+\/\d+/.test(l))                     return true; // "0,3/6"
      if (/^\d+$/.test(l))                             return true; // "831"
      if (/\d{3,}\s+\d+\s+\d/.test(l))                return true; // "756 12 93.5 G-AGGGA"
      if (/^[A-Z]{2,3}\d{3,}/.test(l))                return true; // EN627, RA606
      if (/︱|︲/.test(l))                              return true;
      if (/FULLYTRAINED|BETA/i.test(l))                return true;
      if (/^(Driving|Endurance|Training)\b/i.test(l)) return true;
      if (/^\s*\|/.test(l))                            return true;
      if (/GP\d+\s*\|/.test(l))                       return true; // "GP756 | 92.383"
      return false;
    };

    const normalizeDecoName = (l) => l
      .replace(/[ℬℌℛℐℑℒℓ]/g, m => ({'ℬ':'B','ℌ':'H','ℛ':'R','ℐ':'I','ℑ':'I','ℒ':'L','ℓ':'l'}[m] || m))
      .replace(/[ᴀʙᴄᴅᴇꜰɢʜɪᴊᴋʟᴍɴᴏᴘǫʀꜱᴛᴜᴠᴡxʏᴢᴸᴿ]/g, c => c.normalize('NFKD')[0] || c)
      .replace(/[σ]/g, 's');

    const cleanName = (name) => {
      if (!name) return '';
      let n = name;
      // Score-prefix: "95.029︱Söpö" → "Söpö"
      const scorePrefix = n.match(/^\d+[\.\,]\d+\s*[|︱]\s*(.+)$/);
      if (scorePrefix) n = scorePrefix[1];
      // Strip leading deco chars
      n = n.replace(/Deceased/gi, '')
           .replace(/^[★☆✶✧˚ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃〔〕♟˚₊✧⋆°·•﹚ℙ♕⭐️♡♥*\-͟͞➳❥\s]+/, '')
           .trim();
      // Strip "|| SU" or "|| XX" suffixes (double-pipe qualifiers)
      n = n.replace(/\s*\|\|.*$/, '').trim();
      // Strip trailing score: "KT Sukulaku 92.365" → "KT Sukulaku"
      n = n.replace(/\s+\d+[\.\,]\d{3,}\s*$/, '').trim();
      // Strip trailing score after single pipe followed by digits
      n = n.replace(/\s*[|︱︲\/]\s*[\d\s].*$/, '').trim();
      // Strip trailing deco chars
      n = n.replace(/\s*[⪐+⭐️♡♥★☆✶✧˚➳❥\s]+$/, '').trim();
      n = n.replace(/[★☆✶✧˚ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃〔〕♟˚₊✧⋆°·•﹚⭐️♡♥\s]+$/, '').trim();
      return n || '';
    };

    const slots = [];
    let i = 0;

    while (i < pedLines.length && slots.length < 14) {
      const line = pedLines[i];

      if (isPlaceholder(line)) {
        slots.push('');
        i += 1;
        continue;
      }

      const cleaned = cleanName(normalizeDecoName(line));
      slots.push(cleaned.length > 1 ? cleaned : '');

      const nextLine = pedLines[i + 1] || '';
      i += isTagline(nextLine) ? 3 : 2;
    }

    const g = (idx) => slots[idx] || '';

    // Game slot mapping:
    // 0=Sire, 1=GS(P), 2=GGS(PP), 3=GGD(PP), 4=GD(P), 5=GGS(PM), 6=GGD(PM)
    // 7=Dam,  8=GS(M), 9=GGS(MP), 10=GGD(MP), 11=GD(M), 12=GGS(MM), 13=GGD(MM)
    result.sire    = g(0);
    result.dam     = g(7);
    result.gs_p    = g(1);   // Sire's Sire
    result.gd_p    = g(4);   // Sire's Dam
    result.gs_m    = g(8);   // Dam's Sire
    result.gd_m    = g(11);  // Dam's Dam
    result.ggs_pp  = g(2);   // via Sire's Sire (PP)
    result.ggd_pp  = g(3);
    result.ggs_pm  = g(5);   // via Sire's Dam (PM)
    result.ggd_pm  = g(6);
    result.ggs_mp  = g(9);   // via Dam's Sire (MP)
    result.ggd_mp  = g(10);
    result.ggs_mm  = g(12);  // via Dam's Dam (MM)
    result.ggd_mm  = g(13);

    return result;
  } catch (e) {
    console.error('Error in parsePedigreeString: ' + e.message);
    return {
      sire:'', dam:'', gs_p:'', gd_p:'', gs_m:'', gd_m:'',
      ggs_pp:'', ggd_pp:'', ggs_mp:'', ggd_mp:'',
      ggs_pm:'', ggd_pm:'', ggs_mm:'', ggd_mm:''
    };
  }
}

function getExistingPedigreeInfo(nameOrId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd Tracker');
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
        // Pedigree sheet columns:
        // B(1)=ID, C(2)=Name, D(3)=Sireline, E(4)=Damline
        // F(5)=Sire, G(6)=Dam
        // H(7)=GS(P), I(8)=GD(P), J(9)=GS(M), K(10)=GD(M)
        result.stallionline = pedData[i][3]  || '';
        result.damline      = pedData[i][4]  || '';
        result.sire         = pedData[i][5]  || '';
        result.dam          = pedData[i][6]  || '';
        result.gs_p         = pedData[i][7]  || '';  // GS(P) = Sire's Sire
        result.gd_p         = pedData[i][8]  || '';  // GD(P) = Sire's Dam
        result.gs_m         = pedData[i][9]  || '';  // GS(M) = Dam's Sire
        result.gd_m         = pedData[i][10] || '';  // GD(M) = Dam's Dam
        break;
      }
    }
  }

  return (result.id || result.name) ? result : null;
}

function getPedigreeAutocompleteData() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    let horseNames   = [];

    const herdSheet = ss.getSheetByName('Herd Tracker');
    if (herdSheet && herdSheet.getLastRow() >= 2) {
      herdSheet.getRange('D2:D' + herdSheet.getLastRow()).getValues().forEach(row => {
        if (row[0] && row[0].toString().trim() !== '')
          horseNames.push(row[0].toString().trim());
      });
    }

    const pedigreeSheet = ss.getSheetByName('Pedigree');
    if (pedigreeSheet && pedigreeSheet.getLastRow() >= 2) {
      pedigreeSheet.getRange('C2:C' + pedigreeSheet.getLastRow()).getValues().forEach(row => {
        if (row[0] && row[0].toString().trim() !== '') {
          const name = row[0].toString().trim();
          if (!horseNames.includes(name)) horseNames.push(name);
        }
      });
    }

    return { success: true, horseNames: [...new Set(horseNames)].sort() };
  } catch (error) {
    return { success: false, error: error.message, horseNames: [] };
  }
}

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
      const rowId   = data[i][1] ? data[i][1].toString().trim()                : '';
      const rowName = data[i][2] ? data[i][2].toString().toLowerCase().trim()  : '';
      if ((searchId !== 'NO_ID' && rowId === searchId) || rowName === searchName) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

    // Pedigree sheet column mapping:
    // B(2)=ID, C(3)=Name, D(4)=Sireline, E(5)=Damline
    // F(6)=Sire, G(7)=Dam
    // H(8)=GS(P), I(9)=GD(P), J(10)=GS(M), K(11)=GD(M)
    // L(12)=GGS(PP), M(13)=GGD(PP), N(14)=GGS(PM), O(15)=GGD(PM)
    // P(16)=GGS(MP), Q(17)=GGD(MP), R(18)=GGS(MM), S(19)=GGD(MM)
    const rowValues = [
      pedigreeData.id           || '',  // B
      pedigreeData.name         || '',  // C
      pedigreeData.stallionline || '',  // D
      pedigreeData.damline      || '',  // E
      pedigreeData.sire         || '',  // F
      pedigreeData.dam          || '',  // G
      pedigreeData.gs_p         || '',  // H — GS(P)  = Sire's Sire
      pedigreeData.gd_p         || '',  // I — GD(P)  = Sire's Dam
      pedigreeData.gs_m         || '',  // J — GS(M)  = Dam's Sire
      pedigreeData.gd_m         || '',  // K — GD(M)  = Dam's Dam
      pedigreeData.ggs_pp       || '',  // L — GGS(PP) = via Sire's Sire
      pedigreeData.ggd_pp       || '',  // M — GGD(PP)
      pedigreeData.ggs_pm       || '',  // N — GGS(PM) = via Sire's Dam
      pedigreeData.ggd_pm       || '',  // O — GGD(PM)
      pedigreeData.ggs_mp       || '',  // P — GGS(MP) = via Dam's Sire
      pedigreeData.ggd_mp       || '',  // Q — GGD(MP)
      pedigreeData.ggs_mm       || '',  // R — GGS(MM) = via Dam's Dam
      pedigreeData.ggd_mm       || '',  // S — GGD(MM)
    ];

    sheet.getRange(targetRow, 2, 1, rowValues.length).setValues([rowValues]);
    PropertiesService.getScriptProperties().deleteProperty('pendingPedigreeData');
    return '✓ Pedigree saved in row ' + targetRow;

  } catch (error) {
    throw new Error('Error: ' + error.message);
  }
}

function getPedigreeModalUi(horseId) {
  const tpl = HtmlService.createTemplateFromFile('pedigreeModal');
  tpl.preselectedHorseId = horseId || '';
  return tpl.evaluate().setWidth(620).setHeight(800);
}
