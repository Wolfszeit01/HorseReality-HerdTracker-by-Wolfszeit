/***********************
 * TAGLINE PRINTER BACKEND — FINAL
 ***********************/

function createTaglinePrinter() {
  const html = HtmlService.createHtmlOutputFromFile('TaglineModal')
    .setWidth(550)
    .setHeight(720)
    .setTitle('Tagline Printer');
  SpreadsheetApp.getUi().showModalDialog(html, 'Tagline Printer');
}

/** Quick list — only ID + Name for fast search */
function getQuickHorseList() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd Tracker');
  if (!herdSheet) return [];

  return herdSheet.getDataRange().getValues()
    .slice(1)
    .map(row => ({ id: row[2].toString().trim(), name: row[3] ? row[3].toString().trim() : '' }))
    .filter(h => h.id !== '');
}

/** Full horse data for a single horse — called when selected */
function getHorseDataForTagline(horseId) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const herdSheet = ss.getSheetByName('Herd Tracker');
  if (!herdSheet) return null;

  const data = herdSheet.getDataRange().getValues();
  let horse  = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][2].toString().trim() === horseId.toString()) {
      horse = {
        id:         data[i][2].toString().trim(),
        name:       data[i][3] ? data[i][3].toString().trim() : '',
        breed:      data[i][4] ? data[i][4].toString().trim() : '',
        gender:     data[i][5] ? data[i][5].toString().trim() : '',
        coatColor:  data[i][13] ? data[i][13].toString().trim() : '',
        color:      data[i][14] ? data[i][14].toString().trim() : '',
        predicates: data[i][16] ? data[i][16].toString() : '',
        gp:         data[i][18] || null,
        maxConfo:   data[i][20] || null,
        discipline: data[i][23] ? data[i][23].toString().trim() : '',
        vg: null, g: null, a: null, confo: null
      };
      break;
    }
  }

  if (!horse) return null;

  // Enrich with pedigree lines
  const pedSheet = ss.getSheetByName('Pedigree');
  if (pedSheet) {
    const pedData = pedSheet.getDataRange().getValues();
    for (let i = 1; i < pedData.length; i++) {
      if (pedData[i][1].toString().trim() === horseId.toString()) {
        horse.sireline = pedData[i][3] ? pedData[i][3].toString().trim() : '';
        horse.damline  = pedData[i][4] ? pedData[i][4].toString().trim() : '';
        break;
      }
    }
  }
  const sheetName = horse.breed === 'Icelandic Horse' ? 'ICE_Horse Stats'
                  : horse.breed === 'Kathiawari'      ? 'KATH_Horse Stats'
                  : 'Horse Stats';
  const offset    = horse.breed === 'Icelandic Horse' ? 2
                  : horse.breed === 'Kathiawari'      ? 1 : 0;

  const statsSheet = ss.getSheetByName(sheetName);
  if (statsSheet) {
    const statsData = statsSheet.getDataRange().getValues();
    for (let i = 1; i < statsData.length; i++) {
      if (statsData[i][1].toString().trim() === horseId.toString()) {
        const row    = statsData[i];
        horse.vg     = row[27 + offset] ?? null;
        const gPlus  = row[28 + offset] || 0;
        const g      = row[29 + offset] || 0;
        const gMinus = row[30 + offset] || 0;
        horse.g      = gPlus + g + gMinus;
        horse.a      = row[31 + offset] ?? null;
        horse.confo  = row[63 + offset] ?? null;
        break;
      }
    }
  }

  return horse;
}

/** Save tagline settings to script properties */
function saveTaglineSettings(jsonString) {
  PropertiesService.getUserProperties().setProperty('taglineSettings', jsonString);
}

/** Load tagline settings from script properties */
function getTaglineSettings() {
  return PropertiesService.getUserProperties().getProperty('taglineSettings') || null;
}
