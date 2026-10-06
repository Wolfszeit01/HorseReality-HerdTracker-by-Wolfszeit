/* ══════════════════════════════════════════════════════════════
   TABELLEN-BEREICHE & TAB-KONFIGURATION
   Reines Tab-Layout (Spalten, Zeilen, Trennlinien etc.). 
   Keine Farbdefinitionen – die kommen aus Themes.gs / ThemeBridge.gs.
   ══════════════════════════════════════════════════════════════ */

const SHEETS = {
  TRACKER:    'Herd',
  STATS:      'Stats',
  STATS_ICE:  'ICE_Stats',
  STATS_KATH: 'KATH_Stats',
  PEDIGREE:   'Pedigree',
  BROODMARES: 'Broodmares',
  ACTIVITY:   'Activity Log'
};

const DASH_TRACKER_COLS = {
  STATUS: 2, ID: 3, NAME: 4, BREED: 5, GENDER: 6, BIRTH_DATE: 7,
  AGE: 9, COAT: 15, BREEDING_STATUS: 17, GP: 20, MAX_CONFO: 22,
  RECC_DISCIPLINE: 25, TRAINED_DISCIPLINE: 26
};

const DASH_STATS_COLS      = { ID: 2, NAME: 3, VG_COUNT: 28, TOTAL_GP: 44, CONFO_MAX: 63 };
const DASH_STATS_COLS_ICE  = { ID: 2, NAME: 3, VG_COUNT: 32, TOTAL_GP: 46, CONFO_MAX: 65 };
const DASH_STATS_COLS_KATH = { ID: 2, NAME: 3, VG_COUNT: 29, TOTAL_GP: 45, CONFO_MAX: 64 };

const DASH_PEDIGREE_COLS = { ID: 2, NAME: 3, SIRELINE: 5, DAMLINE: 6 };

const BROODMARES_COLS = {
  CONFIRMED: 2, ID: 3, NAME: 4, BREED: 5, COVER_DATE: 6,
  COVER_TIME: 7, DUE_DATE: 8, GENDER: 12, STUD: 13,
  FOALED_ON: 14, COVERINGS_FAILED: 15
};

const AL_COLS = {
  DATE: 1, TYPE: 2, HORSE_ID: 3, HORSE_NAME: 4, DETAIL: 5,
  DETAIL2: 6, DETAIL3: 7, CURRENCY: 8, AMOUNT: 9, FEE: 10,
  NET: 11, RAW: 12
};

const INCOME_TYPES = ['conf', 'comp', 'stud_in', 'sale', 'interest', 'chores', 'exchange', 'refund'];
const EXPENSE_TYPES = ['stud_out', 'show_fee', 'comp_fee', 'colour_test', 'ultrasound', 'foal_pastures', 'purchase'];
const NEUTRAL_TYPES = ['savings_in', 'savings_out'];
const EVENT_TYPES = ['birth', 'covering_failed'];

const TRAIT_LIST = ['Walk','Trot','Canter','Gallop','Posture','Head','Neck','Back','Shoulders','Frontlegs','Hindquarters','Socks'];
const TRAIT_COLS_STD  = { Walk:4, Trot:5, Canter:6, Gallop:7, Posture:8,  Head:9,  Neck:10, Back:11, Shoulders:12, Frontlegs:13, Hindquarters:14, Socks:15 };
const TRAIT_COLS_ICE  = { Walk:4, Trot:5, Canter:6, Gallop:7, Posture:10, Head:11, Neck:12, Back:13, Shoulders:14, Frontlegs:15, Hindquarters:16, Socks:17 };
const TRAIT_COLS_KATH = { Walk:4, Trot:5, Canter:6, Gallop:7, Posture:9,  Head:10, Neck:11, Back:12, Shoulders:13, Frontlegs:14, Hindquarters:15, Socks:16 };

const INTEREST_TIERS = [
  { max: 100000,   rate: 1.50 },
  { max: 250000,   rate: 1.00 },
  { max: 500000,   rate: 0.80 },
  { max: 1000000,  rate: 0.50 },
  { max: 2000000,  rate: 0.30 },
  { max: 5000000,  rate: 0.15 },
  { max: 10000000, rate: 0.10 },
  { max: 25000000, rate: 0.03 },
  { max: Infinity, rate: 0 }
];
const PROFITABLE_CAPS = [100000, 250000, 500000, 1000000, 2000000, 5000000, 10000000, 25000000];

const SHEET_CONFIG = {
  "Archive": { archiveLayout: true },
  "Settings": { settingsLayout: true },
  "Public Studs": {
    headerRow: 1, startCol: 2, endCol: 17, startDataRow: 2,
    marketLayout: true, marketOutputColumns: [3,13,15,16,17],
    useZebra: true, boldColumns: [3],
    genderDropdownCol: false, frozenRows: 1,
    alignColumns: {left:[3],right:[2,13,15,16,17]},
    columnWidths: {2:110,3:240,13:140,15:85,16:85,17:85},
    numberFormatColumns: {"#,##0":[13],"0":[15,16,17]},
    dividerBeforeCols: [13], subtleDividers: true
  },
  "Horse Market": {
    headerRow: 1, startCol: 2, endCol: 14, startDataRow: 2,
    marketLayout: true, useZebra: true, boldColumns: [3],
    genderDropdownCol: false, frozenRows: 1,
    alignColumns: {left:[3],right:[2,13,14]},
    columnWidths: {2:110,3:240,13:140,14:140},
    numberFormatColumns: {"#,##0":[13,14]},
    dividerBeforeCols: [13], subtleDividers: true
  },
  "Herd": {
    headerRow: 1,
    startCol: 2,         // Spalte B
    endCol: 38,          // Spalte AL
    startDataRow: 2,
    useZebra: true,
    boldColumns: [4],
    centerAllData: true,
    softThemeColors: true,
    alignColumns: { left: [4, 12, 13, 15, 16], right: [14, 19, 20, 22, 23, 28, 38] },
    statusDropdownCol: 2,
    dateColumns: [7],

    // Trennlinien (hrPurple), jeweils LINKS von der genannten Spalte:
    // Q (bereits über Spalte P/16 vorhanden), U, Y, AD
    dividerBeforeCols: [17, 21, 25, 30], // Q, U, Y, AD

    // Rassedurchschnitt-Hervorhebung:
    breedAverageHighlights: [
      { targetCol: 20, compareCol: 20, breedCol: 5, validBreedsRange: "Settings!I3:I100" }, // T
      { targetCol: 22, compareCol: 22, breedCol: 5, validBreedsRange: "Settings!I3:I100" }  // V
    ],
    ratingDropdownCols: [30, 31, 32, 33, 34, 35]
  },

  "Stats": {
    headerRow: 1,
    startCol: 2,
    endCol: 73,

    startDataRow: 2,
    softStatsColors: true,
    alignColumns: { left: [3, 48, 68], center: Array.from({length:12}, (_,i)=>i+4), right: [2].concat(Array.from({length:58}, (_,i)=>i+16).filter(c=>c!==48 && c!==68)) },
    useZebra: true,
    boldColumns: [3],
    timeColumns: [],
    datetimeColumns: [],

    dividerBeforeCols: [4, 16, 35, 45, 54, 56, 58, 60, 62, 64, 66, 68],
    dividerBeforeColsDashed: [47, 49, 50, 69, 71],
    subtleShadeCols: [55, 57, 59, 61, 63, 65, 67],
    gradeColorRange: { startCol: 4, endCol: 15 },
    heatmapRange: { startCol: 35, endCol: 44 }
  },

  "ICE_Stats": {
    headerRow: 1,
    startCol: 2,
    endCol: 77,

    startDataRow: 2,
    softStatsColors: true,
    alignColumns: {"left":[3,52,72],"center":[4,5,6,7,8,9,10,11,12,13,14,15,16,17],"right":[2,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,73,74,75,76,77]},
    // Typed table columns own their number formats; styling must preserve them.
    useZebra: true,
    boldColumns: [3],

    dividerBeforeCols: [4, 18, 39, 49, 58, 60, 62, 64, 66, 68, 70, 72],
    dividerBeforeColsDashed: [51, 54, 73, 74],
    subtleShadeCols: [59, 61, 63, 65, 67, 69, 71],
    gradeColorRange: { startCol: 4, endCol: 17 },
    heatmapRange: { startCol: 39, endCol: 48 }
  },

  "KATH_Stats": {
    headerRow: 1,
    startCol: 2,
    endCol: 75,

    startDataRow: 2,
    softStatsColors: true,
    alignColumns: {"left":[3,50,70],"center":[4,5,6,7,8,9,10,11,12,13,14,15,16],"right":[2,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,71,72,73,74,75]},
    // Typed table columns own their number formats; styling must preserve them.
    useZebra: true,
    boldColumns: [3],

    dividerBeforeCols: [4, 17, 37, 47, 56, 58, 60, 62, 64, 66, 68, 70],
    dividerBeforeColsDashed: [49, 51, 52, 71, 73],
    subtleShadeCols: [57, 59, 61, 63, 65, 67, 69],
    gradeColorRange: { startCol: 4, endCol: 16 },
    heatmapRange: { startCol: 37, endCol: 46 }
  },

  "Pedigree": {
    headerRow: 1,
    startCol: 2,
    endCol: 20,

    startDataRow: 2,
    useZebra: true,
    boldColumns: [3],
    genderDropdownCol: false,
    frozenColumns: 3,
    alignColumns: { left: [3].concat(Array.from({length:16}, (_,i)=>i+5)), right: [2, 4] },
    columnWidths: Object.assign({2: 100, 3: 220, 4: 85, 5: 145, 6: 145},
      Object.fromEntries(Array.from({length:14}, (_,i)=>[i+7, 210]))),
    subtleDividers: true,
    shadeWeight: 0.55,

    dividerBeforeCols: [5, 7, 9, 13, 17],
    subtleShadeCols: [6, 8, 10, 12, 14, 16, 18, 20],
    dateColumns: [],
    timeColumns: [],
    datetimeColumns: []
  },

  "Colour Genetics": {
    headerRow: 1,
    startCol: 2,
    endCol: 80,

    startDataRow: 2,
    useZebra: true,
    genderDropdownCol: false,
    alignColumns: { left: [3, 79, 80], right: [2], center: Array.from({length:50}, (_,i)=>i+4) },
    columnWidths: { 3: 210, 8: 135, 14: 140, 15: 140, 16: 140, 17: 130, 53: 155, 79: 300, 80: 380 },
    dividerBeforeCols: [4, 6, 12, 25, 48, 79],
    subtleDividers: true,
    boldColumns: [3],
    dateColumns: [],
    timeColumns: [],
    datetimeColumns: []
  },

  "Conf. Results": {
    headerStartRow: 1,
    headerEndRow: 7,
    startCol: 2,
    dynamicEndCol: true,
    narrowHeaderRows: 2,
    resultMatrix: true,
    endCol: 2,
    isMatrix: true,
    dividerAfterHeaderRow: 7,
    dividerBorderStyle: "DASHED"
  },

  "Comp. Results": {
    headerStartRow: 1,
    headerEndRow: 6,
    startCol: 2,
    dynamicEndCol: true,
    narrowHeaderRows: 2,
    resultMatrix: true,
    endCol: 2,
    isMatrix: true,
    dividerAfterHeaderRow: 6,
    dividerBorderStyle: "DASHED"
  },

  "Comp. Team": {
    headerRow: 2,
    compTeamLayout: true,
    frozenRows: 2,
    genderDropdownCol: false,
    alignColumns: {left:[3,4,5], right:[2,7,8], center:[6,9]},
    columnWidths: {2:110,3:230,4:155,5:135,6:105,7:115,8:105,9:140},
    dividerBeforeCols: [7],
    subtleDividers: true,
    startCol: 2,
    endCol: 9,

    startDataRow: 3,
    useZebra: true,
    boldColumns: [3],
    dateColumns: [],
    timeColumns: [],
    datetimeColumns: []
  },

  "Broodmares": {
    headerRow: 1,
    startCol: 2,
    endCol: 17,

    startDataRow: 2,
    useZebra: true,
    broodmaresLayout: true,
    frozenRows: 1,
    frozenColumns: 4,
    genderDropdownCol: false,
    alignColumns: {left:[4,5,13,15,17],right:[3],center:[2,6,7,11,12,14]},
    columnWidths: {2:100,3:110,4:230,5:150,6:120,7:105,11:205,12:110,13:230,14:125,15:185,17:280},
    dividerBeforeCols: [6,12,17],
    subtleDividers: true,
    boldColumns: [4],
    dateColumns: [],
    timeColumns: [],
    datetimeColumns: []
  },

  "Outside Studs": {
    headerRow: 1,
    startCol: 2,
    endCol: 76,

    startDataRow: 2,
    useZebra: true,
    boldColumns: [3],
    softStatsColors: true,
    genderDropdownCol: false,
    frozenRows: 1,
    frozenColumns: 3,
    alignColumns: {
      left: [3,4,10,73,76],
      center: Array.from({length:14}, (_,i)=>i+11),
      right: [2,5,6,7,8,9].concat(Array.from({length:52}, (_,i)=>i+25).filter(c=>c!==73 && c!==76))
    },
    columnWidths: {2:105,3:230,4:160,5:145,6:80,7:80,8:80,9:90,10:210,73:150,74:120,76:380},
    subtleDividers: true,

    dividerBeforeCols: [5,10,11,25,46,56,59,61,63,65,67,69,71,73,76],
    dividerBeforeColsDashed: [],
    subtleShadeCols: [60,62,64,66,68,70,72],
    gradeColorRange: { startCol: 11, endCol: 24 },
    heatmapRange: { startCol: 46, endCol: 55 }
  },
};