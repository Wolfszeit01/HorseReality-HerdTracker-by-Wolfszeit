function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Herd Management');
  menu.addItem('Dashboard', 'showDashboard')
    .addItem('Daily Snapshot', 'showXpBalanceModal')
    .addItem('Tagline Printer', 'createTaglinePrinter')
    .addSeparator();

  menu.addSubMenu(ui.createMenu('Import & Update')
    .addItem('Horse Import', 'openUniversalImportModal')
    .addItem('Outside Stud Import', 'openOutsideStudsModal')
    .addItem('Pedigree Import', 'openPedigreeModal')
    .addItem('Add Scores', 'openAddScoresModal')
    .addItem('Paste Notifications', 'el_showNotificationParserModal'));

  menu.addSubMenu(ui.createMenu('Breeding')
    .addItem('Foal Calculator', 'openBreedingOptimizer')
    .addItem('Breeding Evaluation', 'openZielabgleichModal'));
  menu.addSeparator();

  menu.addSubMenu(ui.createMenu('Appearance')
    .addItem('Branding', 'showBrandingSettings')
    .addItem('Theme Editor', 'showThemeEditor')
    .addSeparator()
    .addItem('Apply Current Theme', 'applyThemeToAllSheets')
    .addSeparator()
    .addItem('Ocean Blue', 'switchToOceanBlue')
    .addItem('Forest & Gold', 'switchToForestGold')
    .addItem('Midnight & Silver', 'switchToMidnightSilver'));

  menu.addSubMenu(ui.createMenu('Tools & Maintenance')
    .addItem('Update Stallion Dropdowns', 'updateStallionDropdown')
    .addSeparator()
    .addItem('Show Log Tabs', 'fd_showLogSheets')
    .addItem('Hide Log Tabs', 'fd_hideLogSheets')
    .addSeparator()
    .addItem('Archive Horses', 'archiveHorses')
    .addItem('Archive All Horses…', 'archiveAllHorses')
    .addItem('Import Old Sheet', 'showCsvImportDialog')
    .addSeparator()
    .addItem('Reload Menu', 'onOpen'));

  menu.addSeparator().addItem('Help & Guide', 'showHelpGuide');
  menu.addToUi();
  checkDailyXpReminder();
}


function switchToOceanBlue() {
  switchTheme('OCEAN_BLUE');
}
function switchToForestGold() { switchTheme('FOREST_GOLD'); }
function switchToMidnightSilver() { switchTheme('MIDNIGHT_SILVER'); }
