/***********************
 * TAGLINE PRINTER BACKEND — FINAL
 ***********************/

function createTaglinePrinter() {
  const tpl = HtmlService.createTemplateFromFile('TaglineModal');
  const html = tpl.evaluate()
    .setWidth(550)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/** Quick list — only ID + Name for fast search */
function getQuickHorseList() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd');
  if (!herdSheet) return [];

  return herdSheet.getDataRange().getValues()
    .slice(1)
    .map(row => ({ id: row[2].toString().trim(), name: row[3] ? row[3].toString().trim() : '' }))
    .filter(h => h.id !== '');
}

/** Full horse data for a single horse — called when selected */
function getHorseDataForTagline(horseId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd');
  if (!herdSheet) return null;

  const data = herdSheet.getDataRange().getValues();
  let horse  = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() === horseId.toString()) {
      horse = {
        id:         data[i][2].toString().trim(),   // C
        name:       data[i][3] ? data[i][3].toString().trim() : '', // D
        breed:      data[i][4] ? data[i][4].toString().trim() : '', // E
        gender:     data[i][5] ? data[i][5].toString().trim() : '', // F
        coatColor:  data[i][14] ? data[i][14].toString().trim() : '', // O
        color:      data[i][15] ? data[i][15].toString().trim() : '', // P
        predicates: data[i][17] ? data[i][17].toString() : '',      // R
        gp:         data[i][19] || null,                            // T
        maxConfo:   data[i][21] || null,                            // V
        discipline: data[i][24] ? data[i][24].toString().trim() : '', // Y
        vg: null, g: null, a: null, confo: null,
        sireline: '', damline: '', sire: '', dam: ''
      };
      // Da maxConfo im Herd Tracker liegt, setzen wir confo direkt gleich (oder lassen es parallel laufen)
      horse.confo = horse.maxConfo; 
      break;
    }
  }

  if (!horse) return null;

  // Enrich with pedigree lines & parents
  const pedSheet = ss.getSheetByName('Pedigree');
  if (pedSheet) {
    const pedData = pedSheet.getDataRange().getValues();
    for (let i = 1; i < pedData.length; i++) {
      if (pedData[i][1].toString().trim() === horseId.toString()) { // B
        horse.sireline = pedData[i][4] ? pedData[i][4].toString().trim() : ''; // E
        horse.damline  = pedData[i][5] ? pedData[i][5].toString().trim() : ''; // F
        horse.sire     = pedData[i][6] ? pedData[i][6].toString().trim() : ''; // G
        horse.dam      = pedData[i][7] ? pedData[i][7].toString().trim() : ''; // H
        break;
      }
    }
  }

  const sheetName = horse.breed === 'Icelandic Horse' ? 'ICE_Stats'
                  : horse.breed === 'Kathiawari'      ? 'KATH_Stats'
                  : 'Stats';
  const offset    = horse.breed === 'Icelandic Horse' ? 2
                  : horse.breed === 'Kathiawari'      ? 1 : 0;

  const statsSheet = ss.getSheetByName(sheetName);
  if (statsSheet) {
    const statsData = statsSheet.getDataRange().getValues();
    for (let i = 1; i < statsData.length; i++) {
      if (statsData[i][1].toString().trim() === horseId.toString()) { // B
        const row    = statsData[i];
        horse.vg     = row[27 + offset] ?? null; // AB
        const gPlus  = row[28 + offset] || 0; // AC
        const g      = row[29 + offset] || 0; // AD
        const gMinus = row[30 + offset] || 0; // AE
        horse.g      = gPlus + g + gMinus;
        horse.a      = row[31 + offset] ?? null; // AF
        break;
      }
    }
  }

  return horse;
}

/** Save tagline settings to script properties */
function saveTaglineSettings(jsonString) {
  PropertiesService.getUserProperties().setProperty('publicTaglineSettings', jsonString);
}

/** Load tagline settings from script properties */
function getTaglineSettings() {
  return PropertiesService.getUserProperties().getProperty('publicTaglineSettings') || null;
}