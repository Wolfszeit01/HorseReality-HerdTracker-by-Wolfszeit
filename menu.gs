/***********************
 * UI - MAIN MENU — COMMUNITY VERSION
 ***********************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🐴 Horse Management')

    .addSubMenu(ui.createMenu('📥 Import & Pedigree')
      .addItem('Universal Import',       'openUniversalImportModal')
      .addItem('Outside Stud Import',    'openOutsideStudsModal')
      .addItem('Manual Pedigree Editor', 'openPedigreeModal'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🏆 Scores')
      .addItem('Add Competition & Confo Scores', 'openAddScoresModal')
      .addSeparator()
      .addItem('Remove Duplicate Scores',        'openDuplicateCleanerModal')
      .addItem('Remove Orphaned Score Columns',  'cleanUpOrphanedHorseResults'))

    .addSeparator()

    .addSubMenu(ui.createMenu('🌿 Breeding')
      .addItem('Foal Calculator',           'openBreedingOptimizer')
      .addItem('Sync Broodmares',           'updateBroodmares')
      .addItem('Update Stallion Dropdowns', 'updateStallionDropdown'))

    .addSeparator()

    .addItem('🖨 Tagline Printer', 'createTaglinePrinter')

    .addSeparator()

    .addSubMenu(ui.createMenu('⚙️ System & Team')
      .addItem('Competition Team Update', 'syncCompTeam')
      .addSeparator()
      .addItem('Archive Horses',             'archiveHorses')
      .addItem('Preview Archive Candidates', 'previewArchiveCandidates')
      .addSeparator()
      .addItem('Enable Auto-Archive',  'setupTrigger')
      .addItem('Disable Auto-Archive', 'removeTrigger'))

    .addSeparator()

    .addItem('📊 Dashboard', 'showDashboard')

    .addSeparator()

    .addItem('↺ Reload Menu', 'onOpen')

    .addToUi();
}

/***********************
 * DATA MAINTENANCE
 ***********************/
function cleanUpOrphanedHorseResults() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName('Herd Tracker');

  if (!trackerSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Herd Tracker" not found!');
    return;
  }

  const trackerData = trackerSheet.getRange('C2:C' + trackerSheet.getLastRow()).getValues();
  const activeIDs   = new Set(
    trackerData.map(row => row[0].toString().trim()).filter(id => id !== '')
  );

  const sheetsToClean = ['Conf. Results', 'Comp. Results'];
  let totalDeleted = 0;

  sheetsToClean.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const idRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    for (let col = idRow.length - 1; col >= 0; col--) {
      const currentID = idRow[col].toString().trim();
      if (currentID !== '' && !activeIDs.has(currentID) && currentID.toLowerCase() !== 'id') {
        sheet.deleteColumn(col + 1);
        totalDeleted++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    '✓ Cleanup complete!\n' +
    totalDeleted + ' orphaned column' + (totalDeleted !== 1 ? 's' : '') + ' removed.'
  );
}
