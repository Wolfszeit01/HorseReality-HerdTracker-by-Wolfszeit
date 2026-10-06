/***************************************************************************
 * SHOW DASHBOARD
 ***************************************************************************/

function showDashboard() {
  const html = HtmlService.createTemplateFromFile('HerdDashboard')
    .evaluate()
    .setWidth(1200)
    .setHeight(1300)
    .setTitle('Herd Tracker Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/***************************************************************************
 * TEMPLATE INCLUDE HELPER
 ***************************************************************************/
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/***************************************************************************
 * MAIN DATA FUNCTION — one call, all tabs
 ***************************************************************************/

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const trackerData   = readSheet_(ss, SHEETS.TRACKER);
    const statsData     = readSheet_(ss, SHEETS.STATS);
    const statsIceData  = readSheet_(ss, SHEETS.STATS_ICE);
    const statsKathData = readSheet_(ss, SHEETS.STATS_KATH);
    const pedigreeData  = readSheet_(ss, SHEETS.PEDIGREE);
    const breedingData  = readSheet_(ss, SHEETS.BROODMARES);
    const activityData  = readSheet_(ss, SHEETS.ACTIVITY);

    // Conf/Comp Results lesen
    const confResults = readScoreSheet_(ss, 'Conf. Results');
    const compResults = readScoreSheet_(ss, 'Comp. Results');

    const statsMap    = parseAllStatsData(statsData, statsIceData, statsKathData);
    const pedigreeMap = parsePedigreeData(pedigreeData);
    const pregnancies = parseBreedingData(breedingData);
    const horses      = parseTrackerData(trackerData, statsMap, pedigreeMap);
    const activities  = parseActivityLog(activityData);

    const breedStats = calculateBreedStats(horses, pregnancies);
    const breeds     = getUniqueBreeds(horses);

    const xpEntry = getLastXpEntry();

    // Analytics wird für healthCheckEvents benötigt
    const analytics = calculateAnalytics_(ss, horses, trackerData, compResults, [statsData, statsIceData, statsKathData]);

    // Neue Berechnungen
    const agenda = calculateAgenda(activities, pregnancies, horses, analytics.health);
    const breedingSnapshot = calculateBreedingSnapshot(pregnancies, horses);
    const geneticOverview = calculateGeneticOverview(horses, confResults, compResults);
    const bloodlinePerformance = calculateBloodlinePerformance(horses, confResults, compResults);
    const records = calculateRecords(horses, activities, pregnancies, confResults, compResults);
    const healthCheckEvents = calculateHealthCheckEvents(horses, analytics.health);

    return {
      // Tab 1: Overview
      totalHorses: horses.length,
      stallions: horses.filter(h => h.gender === 'Stallion').length,
      mares: horses.filter(h => h.gender === 'Mare').length,
      foals: horses.filter(h => h.isFoal).length,
      pregnant: pregnancies.filter(p => p.confirmed).length,
      genderRatio: calculateGenderRatio(horses),
      breeds: breeds,
      breedStats: breedStats,
      roleStats: calculateRoleStats(horses),

      // Tab 2: Genetics (ehemals Breeding Progress)
      vgStats: calculateVGStats(horses),
      generationComparison: calculateGenerationComparison(horses, compScoreMap_(compResults)),
      generationByBreed: calculateGenerationComparisonByBreed(horses, compScoreMap_(compResults)),
      disciplineFocus: calculateDisciplineFocus(horses),
      colorStats: calculateColorStats(horses),

      // Tab 3: Bloodlines
      sirelines: calculateLineageStats(horses, 'sireline', compResults),
      damlines: calculateLineageStats(horses, 'damline', compResults),

      // Tab 4: Calendar (ehemals Dates)
      growthEvents: calculateGrowthEvents(horses),
      foalPastureEvents: calculateFoalPastureEvents(activities),
      breedingCheckEvents: calculateBreedingCheckEvents(pregnancies),

      // Tab 5: Breeding
      upcomingBirths: calculateUpcomingBirths(pregnancies),
      breedingSuccessRate: calculateBreedingSuccessRate(activities, pregnancies),
      studUtilization: calculateStudUtilization(activities),
      foalGenders: calculateFoalGenders(pregnancies),

      // Tab 6: Performance
      topPerformers: calculateTopPerformers(activities),
      eventEfficiency: calculateEventEfficiency(activities),
      showPerformanceIndex: calculateShowPerformanceIndex(activities),
      disciplineBreakdown: calculateDisciplineBreakdown(activities),
      earningsOverview: calculateEarningsOverview(activities, breeds),

      // Tab 7: Finances
      balance: getBalance_(ss),
      earningsByType: calculateEarningsByType(activities),
      breedingCosts: calculateBreedingCosts(activities),
      roiByLine: calculateROIByLine(activities, horses, pedigreeMap),
      dailyPassiveIncome: calculateDailyPassiveIncome(activities),

      // Tab 8: Records (ehemals Elite, erweitert)
      topGP: getTopHorses(horses, 'gp', 30),
      topConfo: getTopHorses(horses, 'confo', 30),
      marketValueEstimate: calculateMarketValueEstimate(activities, horses),

      // XP + Net Worth Trend + Savings Interest
      xp: xpEntry,
      netWorthTrend: getXpHistory(),
      interest: calculateInterestInfo_(xpEntry),

      // Analytics
      analytics: analytics,

      // Shared
      recentActivity: getRecentActivity(activities, 300),

      // Erweiterte Tab-Properties
      agenda: agenda,
      breedingSnapshot: breedingSnapshot,
      geneticOverview: geneticOverview,
      bloodlinePerformance: bloodlinePerformance,
      records: records,
      healthCheckEvents: healthCheckEvents
    };
  } catch (e) {
    return { error: e.toString() };
  }
}


/***************************************************************************
 * SHEET READERS
 ***************************************************************************/

function readSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}

function lastDataRowByColumnA_(sheet) {
  const numRows = Math.max(1, sheet.getLastRow());
  const colA = sheet.getRange(1, 1, numRows, 1).getValues();
  for (let r = colA.length - 1; r >= 0; r--) {
    if (colA[r][0] !== '' && colA[r][0] !== null) return r + 1;
  }
  return 1;
}

// 8. HELPER: Score Sheets lesen (Conf. Results / Comp. Results)
function readScoreSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  if (sheet.getLastRow() < 5 || !sheet.getLastColumn()) return [];
  const data = sheet.getRange(1, 1, 5, sheet.getLastColumn()).getValues();

  // Zeile 1 = ID, Zeile 2 = Name, Zeile 3 = Max Score, Zeile 4 = Avg Score, Zeile 5 = Min Score
  const results = [];
  for (let i = 0; i < data[0].length; i++) {
    const id = String(data[0][i] || '').trim();
    if (!id) continue;
    results.push({
      id: id,
      name: String(data[1][i] || '').trim(),
      maxScore: parseFloat(data[2][i]) || 0,
      avgScore: parseFloat(data[3][i]) || 0,
      minScore: parseFloat(data[4][i]) || 0
    });
  }
  return results;
}

/***************************************************************************
 * HERD PARSERS
 ***************************************************************************/

function parseTrackerData(data, statsMap, pedigreeMap) {
  const horses = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[DASH_TRACKER_COLS.ID - 1] || '').trim();
    if (!id) continue;
    const horseStats = statsMap[id] || { vgCount: null, totalGP: 0, confoMax: 0 };
    horses.push({
      id: id,
      name: row[DASH_TRACKER_COLS.NAME - 1] || 'Unnamed',
      breed: row[DASH_TRACKER_COLS.BREED - 1] || 'Unknown',
      gender: normalizeGender(row[DASH_TRACKER_COLS.GENDER - 1]),
      birthDate: row[DASH_TRACKER_COLS.BIRTH_DATE - 1],
      age: row[DASH_TRACKER_COLS.AGE - 1],
      isFoal: Number.isFinite(parseAge(row[DASH_TRACKER_COLS.AGE - 1])) && parseAge(row[DASH_TRACKER_COLS.AGE - 1]) < 3,
      breedingStatus: row[DASH_TRACKER_COLS.BREEDING_STATUS - 1] || '',
      coat: row[DASH_TRACKER_COLS.COAT - 1] || 'Unknown',
      gp: parseFloat(row[DASH_TRACKER_COLS.GP - 1]) || horseStats.totalGP || 0,
      maxConfo: parseFloat(row[DASH_TRACKER_COLS.MAX_CONFO - 1]) || horseStats.confoMax || 0,
      reccDiscipline: row[DASH_TRACKER_COLS.RECC_DISCIPLINE - 1] || '',
      trainedDiscipline: row[DASH_TRACKER_COLS.TRAINED_DISCIPLINE - 1] || '',
      stats: horseStats,
      pedigree: pedigreeMap[id] || { sireline: 'Unknown', damline: 'Unknown' }
    });
  }
  return horses;
}

function parseActivityLog(data) {
  const activities = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const type = String(row[AL_COLS.TYPE - 1] || '').trim();
    if (!type) continue;
    activities.push({
      date: row[AL_COLS.DATE - 1], type: type,
      horseId: String(row[AL_COLS.HORSE_ID - 1] || '').trim(),
      horseName: String(row[AL_COLS.HORSE_NAME - 1] || '').trim(),
      detail: String(row[AL_COLS.DETAIL - 1] || '').trim(),
      detail2: String(row[AL_COLS.DETAIL2 - 1] || '').trim(),
      detail3: String(row[AL_COLS.DETAIL3 - 1] || '').trim(),
      currency: String(row[AL_COLS.CURRENCY - 1] || 'HRC').trim(),
      amount: Number(row[AL_COLS.AMOUNT - 1]) || 0,
      fee: Number(row[AL_COLS.FEE - 1]) || 0,
      net: Number(row[AL_COLS.NET - 1]) || 0
    });
  }
  return activities;
}

function parsePedigreeData(data) {
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][DASH_PEDIGREE_COLS.ID - 1] || '').trim();
    if (id) map[id] = { damline: data[i][DASH_PEDIGREE_COLS.DAMLINE - 1] || 'Unknown', sireline: data[i][DASH_PEDIGREE_COLS.SIRELINE - 1] || 'Unknown' };
  }
  return map;
}

function parseBreedingData(data) {
  const pregnancies = [];
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][BROODMARES_COLS.ID - 1] || '').trim();
    if (!id) continue;
    const rawConfirmed = data[i][BROODMARES_COLS.CONFIRMED - 1];
    const isConfirmed = rawConfirmed === true || String(rawConfirmed).toLowerCase() === 'true';
    const foalGender = data[i][BROODMARES_COLS.GENDER - 1] || '';
    const stud = data[i][BROODMARES_COLS.STUD - 1] || '';
    let coverDate = data[i][BROODMARES_COLS.COVER_DATE - 1];
    if (typeof coverDate === 'string' && coverDate.includes('.')) { const parts = coverDate.split('.'); coverDate = new Date(parts[2], parts[1] - 1, parts[0]); }
    let coverDateTime = null;
    if (coverDate instanceof Date && !isNaN(coverDate)) {
      coverDateTime = new Date(coverDate);
      const coverTime = data[i][BROODMARES_COLS.COVER_TIME - 1];
      if (coverTime instanceof Date) coverDateTime.setHours(coverTime.getHours(), coverTime.getMinutes());
    }
    let rawDueDate = data[i][BROODMARES_COLS.DUE_DATE - 1];
    let dueDate = null;
    if (rawDueDate instanceof Date && !isNaN(rawDueDate.getTime())) { dueDate = rawDueDate; }
    else if (rawDueDate) { let dateString = String(rawDueDate).replace("EFD:", "").trim().split(" ")[0]; let parts = dateString.includes('.') ? dateString.split('.') : dateString.split('/'); if (parts.length === 3) dueDate = dateString.includes('.') ? new Date(parts[2], parts[1] - 1, parts[0]) : new Date(parts[2], parts[0] - 1, parts[1]); }
    let phase = "Waiting";
    if (coverDateTime) {
      const hoursPassed = (now - coverDateTime) / (1000 * 60 * 60);
      if (hoursPassed < 48) phase = "Check: 48h (in " + Math.round(48 - hoursPassed) + "h)";
      else if (hoursPassed < 96) phase = "Check: 96h (in " + Math.round(96 - hoursPassed) + "h)";
      else if (hoursPassed < 120) phase = "Ultrasound (in " + Math.round(120 - hoursPassed) + "h)";
      else phase = "Pregnant";
    }
    pregnancies.push({ id: id, confirmed: isConfirmed, name: data[i][BROODMARES_COLS.NAME - 1], breed: data[i][BROODMARES_COLS.BREED - 1], dueDate: dueDate, coverDateTime: coverDateTime, phase: phase, foalGender: foalGender, stud: stud });
  }
  return pregnancies;
}

function dashGrade_(value) {
  const grade = String(value || '').trim().toUpperCase();
  return ({'VERY GOOD':'VG', 'GOOD':'G', 'AVERAGE':'A', 'BELOW AVERAGE':'BA', 'POOR':'P'})[grade] || grade;
}

function parseAllStatsData(standardData, iceData, kathData) {
  const map = {};
  [[kathData,1,TRAIT_COLS_KATH],[iceData,2,TRAIT_COLS_ICE],[standardData,0,TRAIT_COLS_STD]].forEach(([rows,extra,fallback]) => {
    if (!rows || rows.length < 2) return;
    const traitCols = resolveTraitCols_(rows[1], fallback);
    const maxCol = fcColumn_(rows, ['conf max','confo max','conformation max','max confo score'],46+2*extra);
    rows.slice(1).forEach(row => {
      const id = String(row[1] || '').trim();
      if (!id || /^(id|horse id)$/i.test(id) || map[id]) return;
      const grades = row.slice(3,15+extra).map(dashGrade_);
      const knownGrades = grades.filter(g => ['VG','VG-','G+','G','G-','A','BA','P'].includes(g));
      const gp = Array.from({length:10},(_,i)=>fcNumber_(row[34+2*extra+i]));
      map[id] = {
        vgCount: knownGrades.length ? grades.filter(g=>g==='VG'||g==='VG-').length : null,
        totalGP: gp.every(n=>n!==null) ? gp.reduce((sum,n)=>sum+n,0) : 0,
        confoMax: fcNumber_(row[maxCol]) || 0,
        traits: readTraitGradesRow_(row,traitCols)
      };
    });
  });
  return map;
}

/***************************************************************************
 * TRAIT-SPALTEN-AUFLÖSUNG (aus Config verschoben)
 ***************************************************************************/

// Liest die 12 Basis-Trait-Grades einer Zeile anhand der übergebenen Spalten-Map
function readTraitGradesRow_(row, colMap) {
  const traits = {};
  TRAIT_LIST.forEach(t => {
    const col = colMap[t];
    traits[t] = col ? String(row[col - 1] || '').trim() : '';
  });
  return traits;
}

// Baut die Trait-Spalten-Map dynamisch aus der Header-Zeile (Zeile 2 des Sheets = data[1]),
// da Icelandic (Tolt/Pace) und Kathiawari (Revaal) zusätzliche Gait-Spalten einschieben und
// sich die Spaltenposition der gemeinsamen Traits dadurch verschiebt. Fällt auf die statische
// Fallback-Map zurück, falls ein Header nicht gefunden wird.
function resolveTraitCols_(headerRow, fallback) {
  const map = {};
  TRAIT_LIST.forEach(t => {
    let found = null;
    if (headerRow) {
      for (let c = 3; c < headerRow.length; c++) { // Spalte D (Index 3) aufwärts
        if (String(headerRow[c] || '').trim().toLowerCase() === t.toLowerCase()) { found = c + 1; break; }
      }
    }
    map[t] = found || fallback[t];
  });
  return map;
}

/***************************************************************************
 * TAB 1: OVERVIEW
 ***************************************************************************/

function parseAge(s) {
  if (s === null || s === undefined || String(s).trim() === '') return NaN;
  if (!/^(?:\d+(?:[.,]\d+)?|.*\d+\s*[ym].*)$/i.test(String(s).trim())) return NaN;
  return fcParseAge_(String(s).replace(',','.'));
}
function getUniqueBreeds(h) { return [...new Set(h.map(x => x.breed))].sort(); }
function calculateGenderRatio(h) { const s = h.filter(x => x.gender === 'Stallion').length; const m = h.filter(x => x.gender === 'Mare').length; return m === 0 ? s + ':0' : (s / m).toFixed(2) + ':1'; }

function calculateBreedStats(horses, pregnancies) {
  const s = {};
  horses.forEach(h => {
    if (!s[h.breed]) s[h.breed] = { total: 0, stallions: 0, mares: 0, foals: 0, pregnant: 0, gpSum: 0, gpCount: 0, vgSum: 0, vgCount: 0, confoSum: 0, confoCount: 0, roles: { stud: 0, broodmare: 0, public: 0, none: 0 } };
    const b = s[h.breed]; b.total++;
    if (h.gender === 'Stallion') b.stallions++;
    if (h.gender === 'Mare') b.mares++;
    if (h.isFoal) b.foals++;
    if (pregnancies.some(p => p.id === h.id && p.confirmed)) b.pregnant++;
    const bs = h.breedingStatus.toLowerCase();
    if (bs.includes('stud') && !bs.includes('public')) b.roles.stud++;
    else if (bs.includes('broodmare')) b.roles.broodmare++;
    else if (bs.includes('public')) b.roles.public++;
    else b.roles.none++;
    if (h.gp > 0) { b.gpSum += h.gp; b.gpCount++; }
    if (h.stats.vgCount !== null) { b.vgSum += h.stats.vgCount; b.vgCount++; }
    if (h.maxConfo > 0) { b.confoSum += h.maxConfo; b.confoCount++; }
  });
  for (let b in s) { const i = s[b]; i.avgGP = i.gpCount > 0 ? Math.round(i.gpSum / i.gpCount) : 0; i.avgVG = i.vgCount > 0 ? (i.vgSum / i.vgCount).toFixed(2) : '—'; i.avgConfo = i.confoCount > 0 ? (i.confoSum / i.confoCount).toFixed(3) : "0.000"; }
  return s;
}

function calculateRoleStats(h) {
  const r = { stud: 0, broodmare: 0, public: 0, none: 0 };
  h.forEach(x => {
    const s = x.breedingStatus.toLowerCase();
    if (s.includes('stud') && !s.includes('public')) r.stud++;
    else if (s.includes('broodmare')) r.broodmare++;
    else if (s.includes('public')) r.public++;
    else r.none++;
  });
  return r;
}

// TODAY'S AGENDA
function calculateAgenda(activities, pregnancies, horses, healthData) {
  const now = new Date();
  const todayStr = now.toLocaleDateString('en-US');

  // Ultrasounds: Heute fällige Ultraschall-Termine (120h nach Deckung)
  const ultrasounds = pregnancies.filter(p => {
    if (!p.coverDateTime || p.confirmed) return false;
    const targetDate = new Date(p.coverDateTime.getTime() + 120 * 60 * 60 * 1000);
    return targetDate.toLocaleDateString('en-US') === todayStr;
  }).length;

  // Births: Heute fällige Geburten
  const births = pregnancies.filter(p => {
    if (!p.dueDate) return false;
    return p.dueDate.toLocaleDateString('en-US') === todayStr;
  }).length;

  // Rideable Foals: Heute 3 Jahre alt werdende Fohlen
  const rideableFoals = horses.filter(h => {
    if (!h.birthDate || !h.isFoal) return false;
    const adultDate = new Date(h.birthDate.getTime() + 48 * 86400000);
    return adultDate.toLocaleDateString('en-US') === todayStr;
  }).length;

  // Pasture Returns: Heute zurückkehrende Fohlen von der Weide
  const pastureReturns = activities.filter(a => {
    if (a.type !== 'foal_pastures' || a.detail !== 'Returned') return false;
    if (!(a.date instanceof Date)) return false;
    return a.date.toLocaleDateString('en-US') === todayStr;
  }).length;

  // Health Tests Needed: erwachsene Pferde mit fehlenden Gesundheitstests
  const healthTestsNeeded = (healthData && healthData.actionNeeded) ? healthData.actionNeeded.length : 0;

  return { ultrasounds, births, rideableFoals, pastureReturns, healthTestsNeeded };
}

// HEALTH CHECK EVENTS für Calendar
function calculateHealthCheckEvents(horses, healthData) {
  const events = [];
  const now = new Date();

  healthData.actionNeeded.forEach(h => {
    // Fälligkeitsdatum: Heute + 30 Tage als Erinnerung
    const dueDate = new Date(now.getTime() + 30 * 86400000);
    events.push({
      name: h.name,
      breed: h.breed,
      missingTests: h.missing.join(', '),
      dueDate: dueDate.toLocaleDateString('en-US'),
      daysLeft: '30'
    });
  });

  return events;
}

// RECORDS / HALL OF FAME
function calculateRecords(horses, activities, pregnancies, confResults, compResults) {
  const breedByName_ = {}, breedById_ = {};
  horses.forEach(h => { breedByName_[h.name] = h.breed; breedById_[h.id] = h.breed; });

  // Highest GP
  const sortedGP = [...horses].sort((a, b) => b.gp - a.gp);
  const highestGP = sortedGP.length ? { value: Math.round(sortedGP[0].gp), name: sortedGP[0].name, breed: sortedGP[0].breed } : null;

  // Most VG
  const sortedVG = [...horses].filter(h => h.stats?.vgCount > 0).sort((a, b) => b.stats.vgCount - a.stats.vgCount);
  const highestVG = sortedVG.length ? { value: sortedVG[0].stats.vgCount, name: sortedVG[0].name, breed: sortedVG[0].breed } : null;

  // Best Confo
  const sortedConfo = [...horses].filter(h => h.maxConfo > 0).sort((a, b) => b.maxConfo - a.maxConfo);
  const bestConfo = sortedConfo.length ? { value: sortedConfo[0].maxConfo.toFixed(3), name: sortedConfo[0].name, breed: sortedConfo[0].breed } : null;

  // Highest Earnings
  const earnings = {};
  activities.filter(a => a.type === 'comp' || a.type === 'conf').forEach(a => {
    if (!a.horseId) return;
    earnings[a.horseId] = (earnings[a.horseId] || 0) + a.net;
  });
  const sortedEarnings = Object.entries(earnings).sort((a, b) => b[1] - a[1]);
  const highestEarningsHorse = sortedEarnings.length ? horses.find(h => h.id === sortedEarnings[0][0]) : null;
  const highestEarnings = sortedEarnings.length ? { value: sortedEarnings[0][1], name: highestEarningsHorse?.name || 'Unknown', breed: highestEarningsHorse?.breed || '' } : null;

  // Most Successful Stud (basierend auf stud_in Aktivitäten)
  const studEarnings = {};
  activities.filter(a => a.type === 'stud_in').forEach(a => {
    if (!a.horseName) return;
    studEarnings[a.horseName] = (studEarnings[a.horseName] || 0) + a.net;
  });
  const sortedStuds = Object.entries(studEarnings).sort((a, b) => b[1] - a[1]);
  const mostSuccessfulStud = sortedStuds.length
    ? { value: Math.round(sortedStuds[0][1]).toLocaleString('en-US'), name: sortedStuds[0][0], breed: breedByName_[sortedStuds[0][0]] || '' }
    : null;

  // Best Mare (basierend auf Nachkommen und Erfolgen)
  const mareSuccess = {};
  activities.filter(a => a.type === 'birth').forEach(a => {
    const key = a.horseId || a.horseName;
    if (!key) return;
    if (!mareSuccess[key]) mareSuccess[key] = { births: 0, name: a.horseName };
    mareSuccess[key].births++;
  });
  const sortedMares = Object.values(mareSuccess).sort((a, b) => b.births - a.births);
  const bestMare = sortedMares.length ? { value: sortedMares[0].births + ' foals', name: sortedMares[0].name, breed: breedByName_[sortedMares[0].name] || '' } : null;

  // Most Titles (basierend auf Platzierungen)
  const titles = {};
  activities.filter(a => a.type === 'comp' || a.type === 'conf').forEach(a => {
    if (!a.horseName) return;
    const place = (a.detail3 || a.detail2 || '').toLowerCase();
    if (place.includes('1st') || place.includes('champion')) {
      titles[a.horseName] = (titles[a.horseName] || 0) + 1;
    }
  });
  const sortedTitles = Object.entries(titles).sort((a, b) => b[1] - a[1]);
  const mostTitles = sortedTitles.length ? { value: sortedTitles[0][1], name: sortedTitles[0][0], breed: breedByName_[sortedTitles[0][0]] || '' } : null;

  // Oldest Horse
  const withAge = horses.filter(h => h.birthDate instanceof Date);
  const oldest = withAge.length ? withAge.sort((a, b) => a.birthDate - b.birthDate)[0] : null;
  const oldestHorse = oldest ? { value: Math.floor((new Date() - oldest.birthDate) / (16 * 86400000)) + ' years', name: oldest.name, breed: oldest.breed } : null;

  // Highest Produced GP (Durchschnitt der Nachkommen-GP pro Zuchtlinie)
  const damGP = {};
  horses.forEach(h => {
    const dam = h.pedigree?.damline || 'Unknown';
    if (!damGP[dam]) damGP[dam] = { sum: 0, count: 0, name: dam };
    if (h.gp > 0) { damGP[dam].sum += h.gp; damGP[dam].count++; }
  });
  const sortedDamGP = Object.values(damGP).filter(d => d.count > 0).sort((a, b) => (b.sum / b.count) - (a.sum / a.count));
  const highestProducedGP = sortedDamGP.length ? { value: Math.round(sortedDamGP[0].sum / sortedDamGP[0].count), name: sortedDamGP[0].name, breed: '' } : null;

  return {
    highestGP, highestVG, bestConfo, highestEarnings,
    mostSuccessfulStud, bestMare, mostTitles, oldestHorse, highestProducedGP
  };
}

/***************************************************************************
 * TAB 2: GENETICS / BREEDING PROGRESS
 ***************************************************************************/

function calculateVGStats(horses) {
  const res = { all: { vg12: 0, vg11: 0, vg10: 0, totalWithStats: 0 } };
  horses.filter(h => h.stats.vgCount !== null).forEach(h => { if (!res[h.breed]) res[h.breed] = { vg12: 0, vg11: 0, vg10: 0, totalWithStats: 0 }; [res.all, res[h.breed]].forEach(t => { if (h.stats.vgCount >= 12) t.vg12++; else if (h.stats.vgCount === 11) t.vg11++; else if (h.stats.vgCount === 10) t.vg10++; t.totalWithStats++; }); });
  return res;
}

function compScoreMap_(compResults) {
  const m = {};
  (compResults || []).forEach(c => { m[c.id] = c.maxScore; });
  return m;
}

function dashMean_(values) { return values.length ? values.reduce((s,v)=>s+v,0)/values.length : 0; }
function calculateGenerationComparison(arr, compMap) {
  compMap = compMap || {};
  const calc = (list) => ({
    avgGP: dashMean_(list.map(h=>h.gp).filter(v=>v>0)),
    avgConfo: dashMean_(list.map(h=>h.maxConfo).filter(v=>v>0)),
    avgVG: dashMean_(list.map(h=>h.stats.vgCount).filter(v=>v!==null)),
    avgComp: (() => { const withComp = list.filter(h => (compMap[h.id] || 0) > 0); return withComp.length ? withComp.reduce((a, b) => a + compMap[b.id], 0) / withComp.length : 0; })()
  });
  const adults = arr.filter(h => parseAge(h.age) >= 3), foals = arr.filter(h => h.isFoal);
  const counts = list => ({avgGP:list.filter(h=>h.gp>0).length, avgConfo:list.filter(h=>h.maxConfo>0).length, avgVG:list.filter(h=>h.stats.vgCount!==null).length, avgComp:list.filter(h=>(compMap[h.id]||0)>0).length});
  return { adults: calc(adults), foals: calc(foals), counts:{adults:counts(adults),foals:counts(foals)} };
}

function calculateGenerationComparisonByBreed(horses, compMap) { const res = {}; getUniqueBreeds(horses).forEach(b => { res[b] = calculateGenerationComparison(horses.filter(h => h.breed === b), compMap); }); return res; }

function calculateDisciplineFocus(horses) {
  const res = { all: {} };
  horses.forEach(h => { const d = h.reccDiscipline || h.trainedDiscipline || 'None'; if (!res[h.breed]) res[h.breed] = {}; [res.all, res[h.breed]].forEach(t => { if (!t[d]) t[d] = { count: 0, gpSum: 0, confoSum: 0, valid: 0 }; t[d].count++; if (h.gp > 0) { t[d].gpSum += h.gp; t[d].confoSum += h.maxConfo; t[d].valid++; } }); });
  for (let b in res) { for (let d in res[b]) { const i = res[b][d]; i.avgGP = i.valid > 0 ? i.gpSum / i.valid : 0; i.avgConfo = i.valid > 0 ? i.confoSum / i.valid : 0; } }
  return res;
}

function calculateColorStats(horses) {
  const patterns = ["tobiano", "overo", "sabino", "rabicano", "splash", "white", "leopard", "blanket", "varnish", "spotted", "appaloosa", "snowflake"];
  const res = { all: { solid: {}, patterned: {} } };
  horses.forEach(h => { const cat = patterns.some(p => h.coat.toLowerCase().includes(p)) ? 'patterned' : 'solid'; if (!res[h.breed]) res[h.breed] = { solid: {}, patterned: {} }; [res.all, res[h.breed]].forEach(t => { t[cat][h.coat] = (t[cat][h.coat] || 0) + 1; }); });
  for (let b in res) { const total = Object.values(res[b].solid).reduce((a, c) => a + c, 0) + Object.values(res[b].patterned).reduce((a, c) => a + c, 0); const toArr = (source) => Object.keys(source).map(c => ({ color: c, count: source[c], percentage: ((source[c] / (total || 1)) * 100).toFixed(1) })).sort((a, b) => b.count - a.count); res[b] = { solid: toArr(res[b].solid), patterned: toArr(res[b].patterned) }; }
  return res;
}

function calculateGeneticOverview(horses, confResults, compResults) {
  const withStats = horses.filter(h => h.stats && h.stats.vgCount > 0);
  const avgVG = withStats.length ? (withStats.reduce((s, h) => s + h.stats.vgCount, 0) / withStats.length).toFixed(1) : '—';
  const avgGP = withStats.length ? Math.round(withStats.reduce((s, h) => s + (h.gp || 0), 0) / withStats.length) : '—';
  const avgConfo = withStats.length ? (withStats.reduce((s, h) => s + (h.maxConfo || 0), 0) / withStats.length).toFixed(2) : '—';

  // Conf/Comp Scores aus den Ergebnissheets
  const confScores = confResults || [];
  const compScores = compResults || [];
  const avgConf = confScores.length ? (confScores.reduce((s, c) => s + c.maxScore, 0) / confScores.length).toFixed(2) : '—';
  const avgComp = compScores.length ? (compScores.reduce((s, c) => s + c.maxScore, 0) / compScores.length).toFixed(2) : '—';

  return { avgVG, avgGP, avgConfo, avgConf, avgComp };
}

/***************************************************************************
 * TAB 5: BREEDING
 ***************************************************************************/

function calculateUpcomingBirths(pregnancies) {
  const now = new Date(); now.setHours(0, 0, 0, 0); const cutoffDate = new Date(now); cutoffDate.setDate(now.getDate() - 5);
  return pregnancies.filter(x => x.dueDate && x.dueDate >= cutoffDate).map(x => { const diffDays = Math.ceil((x.dueDate - now) / (1000 * 60 * 60 * 24)); let statusText = x.phase; if (diffDays < 0) statusText = "Overdue (" + Math.abs(diffDays) + " days)"; else if (diffDays === 0) statusText = "DUE TODAY"; return { mareName: x.name, breed: x.breed, dueDate: x.dueDate.toLocaleDateString('en-US'), daysUntil: diffDays, phase: statusText, foalGender: x.foalGender || '—', stud: x.stud || '—' }; }).sort((a, b) => a.daysUntil - b.daysUntil);
}

function calculateBreedingSuccessRate(activities, pregnancies) {
  const births = activities.filter(a => a.type === 'birth'); const failures = activities.filter(a => a.type === 'covering_failed'); const totalAttempts = births.length + failures.length;
  const mareStats = {};
  births.forEach(a => { const name = a.horseName; if (!name) return; if (!mareStats[name]) mareStats[name] = { births: 0, failures: 0 }; mareStats[name].births++; });
  failures.forEach(a => { const name = a.horseName; if (!mareStats[name]) mareStats[name] = { births: 0, failures: 0 }; mareStats[name].failures++; });
  return { totalBirths: births.length, totalFailures: failures.length, totalAttempts: totalAttempts, successRate: totalAttempts > 0 ? ((births.length / totalAttempts) * 100).toFixed(1) : '—', mareStats: Object.keys(mareStats).map(name => ({ name: name, births: mareStats[name].births, failures: mareStats[name].failures, rate: ((mareStats[name].births / (mareStats[name].births + mareStats[name].failures)) * 100).toFixed(0) })).sort((a, b) => b.births + b.failures - a.births - a.failures).slice(0, 20) };
}

function calculateStudUtilization(activities) {
  const now = new Date(); const sevenDaysAgo = new Date(now.getTime() - 7 * 86400000); const thirtyDaysAgo = new Date(now.getTime() - 30 * 86400000);
  const studIns = activities.filter(a => a.type === 'stud_in'); const studOuts = activities.filter(a => a.type === 'stud_out');
  const studs = {};
  studIns.forEach(a => { const name = a.horseName; if (!studs[name]) studs[name] = { name: name, bookings7d: 0, bookings30d: 0, totalEarnings: 0, totalBookings: 0 }; studs[name].totalBookings++; studs[name].totalEarnings += a.net; const d = a.date instanceof Date ? a.date : new Date(a.date); if (d >= sevenDaysAgo) studs[name].bookings7d++; if (d >= thirtyDaysAgo) studs[name].bookings30d++; });
  return { totalIncoming: studIns.length, totalOutgoing: studOuts.length, studs: Object.values(studs).sort((a, b) => b.totalBookings - b.totalBookings) };
}

function calculateFoalGenders(pregnancies) {
  const covered = pregnancies.filter(p => p.coverDateTime); const withGender = covered.filter(p => p.foalGender && p.foalGender.trim() !== '' && p.foalGender !== '—'); const confirmed = pregnancies.filter(p => p.confirmed);
  return { totalCovered: covered.length, totalConfirmed: confirmed.length, colt: withGender.filter(p => p.foalGender.toLowerCase() === 'colt' || p.foalGender.toLowerCase() === 'stallion').length, filly: withGender.filter(p => ['filly','mare'].includes(p.foalGender.toLowerCase())).length, unknown: covered.length - withGender.length };
}

function calculateBreedingSnapshot(pregnancies, horses) {
  const covered = pregnancies.filter(p => p.coverDateTime).length;
  const pregnant = pregnancies.filter(p => p.confirmed).length;

  const availableMares = horses.filter(h =>
    h.gender === 'Mare' &&
    parseAge(h.age) >= 3 &&
    !pregnancies.some(p => p.id === h.id && p.confirmed)
  ).length;

  const availableStuds = horses.filter(h =>
    h.gender === 'Stallion' &&
    parseAge(h.age) >= 3 &&
    h.breedingStatus &&
    h.breedingStatus.toLowerCase().includes('stud')
  ).length;

  return { covered, pregnant, availableMares, availableStuds };
}

/***************************************************************************
 * TAB 3: BLOODLINES
 ***************************************************************************/

function calculateLineageStats(horses, lineType, compResults) {
  const compMap = {};
  (compResults || []).forEach(c => { compMap[c.id] = c.maxScore; });
  const res = { all: {} };
  horses.forEach(horse => { if (!horse.pedigree) return; const breed = horse.breed || 'Unknown'; const lineName = lineType === 'sireline' ? horse.pedigree.sireline : horse.pedigree.damline; if (!lineName) return; if (!res[breed]) res[breed] = {}; [res.all, res[breed]].forEach(target => { if (!target[lineName]) target[lineName] = { name: lineName, total: 0, stallions: 0, mares: 0, gpSum: 0, validCount: 0, vgSum: 0, validVG: 0, confoSum: 0, validConfo: 0, compSum: 0, validComp: 0 }; const l = target[lineName]; l.total++; if (horse.gender === 'Stallion') l.stallions++; if (horse.gender === 'Mare') l.mares++; const gp = horse.gp || (horse.stats ? horse.stats.totalGP : 0); if (gp > 0) { l.gpSum += gp; l.validCount++; } if (horse.stats && horse.stats.vgCount !== null) { l.vgSum += horse.stats.vgCount; l.validVG++; } if (horse.maxConfo > 0) { l.confoSum += horse.maxConfo; l.validConfo++; } const compScore = compMap[horse.id] || 0; if (compScore > 0) { l.compSum += compScore; l.validComp++; } }); });
  const final = {};
  const theme = getActiveTheme();
  Object.keys(res).forEach(b => { const bTotal = Object.values(res[b]).reduce((acc, curr) => acc + curr.total, 0) || 1; final[b] = Object.values(res[b]).map(line => { const share = (line.total / bTotal) * 100; const stallionRatio = (line.stallions / line.total) * 100; let domColor = theme.tertiary.deep, domLabel = 'STABLE'; if (line.total <= 2) { domColor = theme.accents.accent6; domLabel = 'AT RISK'; } else if (share > 40) { domColor = theme.accents.accent2; domLabel = 'DOMINANT'; } return { ...line, share: share.toFixed(1), avgGP: line.validCount > 0 ? Math.round(line.gpSum / line.validCount) : 0, avgVG: line.validVG > 0 ? (line.vgSum / line.validVG).toFixed(1) : '—', avgConfo: line.validConfo > 0 ? (line.confoSum / line.validConfo).toFixed(2) : '—', avgComp: line.validComp > 0 ? (line.compSum / line.validComp).toFixed(2) : '—', domColor, domLabel, stallionRatio, balanceWarning: stallionRatio > 60 && line.total > 2 }; }).sort((a, b) => b.total - a.total); });
  return final;
}

/***************************************************************************
 * ⚠️ VERMUTLICH UNGENUTZT — wird berechnet, aber im Frontend (renderBloodlines)
 * nirgends gelesen. Bitte prüfen, ob Feature noch gebraucht wird — sonst löschen.
 ***************************************************************************/

function calculateBloodlinePerformance(horses, confResults, compResults) {
  // Conf/Comp Scores nach Pferd-ID mappen
  const confMap = {};
  (confResults || []).forEach(c => { confMap[c.id] = c.maxScore; });
  const compMap = {};
  (compResults || []).forEach(c => { compMap[c.id] = c.maxScore; });

  const lines = {};
  horses.forEach(h => {
    const sire = h.pedigree?.sireline || 'Unknown';
    const dam = h.pedigree?.damline || 'Unknown';
    [sire, dam].forEach((line, isSire) => {
      if (!lines[line]) lines[line] = { name: line, total: 0, stallions: 0, mares: 0, gpSum: 0, vgSum: 0, confoSum: 0, compSum: 0, validGP: 0, validVG: 0, validConfo: 0, validComp: 0 };
      const l = lines[line];
      l.total++;
      if (h.gender === 'Stallion') l.stallions++;
      if (h.gender === 'Mare') l.mares++;
      if (h.gp > 0) { l.gpSum += h.gp; l.validGP++; }
      if (h.stats?.vgCount > 0) { l.vgSum += h.stats.vgCount; l.validVG++; }
      if (h.maxConfo > 0) { l.confoSum += h.maxConfo; l.validConfo++; }
      const compScore = compMap[h.id] || 0;
      if (compScore > 0) { l.compSum += compScore; l.validComp++; }
    });
  });

  return Object.values(lines).map(l => ({
    ...l,
    avgGP: l.validGP ? Math.round(l.gpSum / l.validGP) : 0,
    avgVG: l.validVG ? (l.vgSum / l.validVG).toFixed(1) : '—',
    avgConfo: l.validConfo ? (l.confoSum / l.validConfo).toFixed(2) : '—',
    avgComp: l.validComp ? (l.compSum / l.validComp).toFixed(2) : '—'
  }));
}

/***************************************************************************
 * TAB 6/PERFORMANCE-TAB: SCORE-ANALYTICS (Confo/Comp Verteilung, Health, Traits)
 ***************************************************************************/

function buildScoreAnalytics_(horses, scoreFn) {
  const list = horses.filter(h => scoreFn(h) > 0).map(h => ({
    name: h.name, breed: h.breed, gender: h.gender, isFoal: h.isFoal,
    score: scoreFn(h), gp: h.gp
  })).sort((a, b) => b.score - a.score);

  const scores = list.map(h => h.score);
  const median = scores.length ? scores[Math.floor(scores.length / 2)] : 0;
  const best = scores.length ? scores[0] : 0;
  const bestHorse = list.length ? list[0].name : '—';

  const genderCats = {};
  ['Stallion', 'Mare', 'Colt', 'Filly'].forEach(g => {
    const filtered = g === 'Colt' ? list.filter(h => h.isFoal && h.gender === 'Stallion') :
                     g === 'Filly' ? list.filter(h => h.isFoal && h.gender === 'Mare') :
                     g === 'Stallion' ? list.filter(h => !h.isFoal && h.gender === 'Stallion') :
                     list.filter(h => !h.isFoal && h.gender === 'Mare');
    if (filtered.length) {
      const s = filtered.map(h => h.score);
      genderCats[g] = { n: filtered.length, lowest: Math.min(...s), median: s[Math.floor(s.length / 2)], highest: Math.max(...s) };
    }
  });

  const top5 = list.slice(0, 5).map(h => ({ ...h, vsMedian: h.score - median }));
  const bottom5 = list.slice(-5).reverse().map(h => ({ ...h, vsMedian: h.score - median }));

  return { median, best, bestHorse, totalMeasured: list.length, totalHorses: horses.length, genderCats, top5, bottom5, allScores: list };
}

function calculateAnalytics_(ss, horses, trackerData, compResults, statsTables) {
  // ── Conformation Score Distribution ──
  const confoAnalytics = buildScoreAnalytics_(horses, h => h.maxConfo || 0);

  // ── Competition Score Distribution ──
  const compMap = {};
  (compResults || []).forEach(c => { compMap[c.id] = c.maxScore; });
  const compAnalytics = buildScoreAnalytics_(horses, h => compMap[h.id] || 0);

  // ── Trait Grade Analysis (aggregiert) ──
  const traitData = readTraitGrades_(ss, statsTables, new Set(horses.map(h => h.id)));

  // ── Trait Grades pro Pferd (für Heatmap) ──
  const perHorseTraits = buildPerHorseTraits_(horses);

  // ── Health Analysis ──
  const healthData = readHealthData_(trackerData, horses);

  // ── Generation Progress (Confo) ──
  const genProgress = {};
  const allBreeds = [...new Set(horses.map(h => h.breed))];
  allBreeds.forEach(b => {
    const adults = horses.filter(h => h.breed === b && !h.isFoal && h.maxConfo > 0);
    const foals = horses.filter(h => h.breed === b && h.isFoal && h.maxConfo > 0);
    genProgress[b] = {
      adultsAvg: adults.length ? Math.round(adults.reduce((s, h) => s + h.maxConfo, 0) / adults.length) : 0,
      foalsAvg: foals.length ? Math.round(foals.reduce((s, h) => s + h.maxConfo, 0) / foals.length) : 0,
      adultsCount: adults.length, foalsCount: foals.length,
      improvement: 0
    };
    if (genProgress[b].adultsAvg > 0 && genProgress[b].foalsAvg > 0) {
      genProgress[b].improvement = genProgress[b].foalsAvg - genProgress[b].adultsAvg;
    }
  });

  return {
    confo: confoAnalytics,
    comp: compAnalytics,
    traits: traitData,
    traitList: TRAIT_LIST,
    perHorseTraits: perHorseTraits,
    health: healthData,
    genProgress: genProgress
  };
}

// Baut die Liste für die Trait-Grades-Heatmap: ein Eintrag pro Pferd mit Score + Trait-Grades
function buildPerHorseTraits_(horses) {
  return horses
    .filter(h => h.maxConfo > 0 && h.stats && h.stats.traits)
    .map(h => ({
      name: h.name, breed: h.breed, gender: h.gender, isFoal: h.isFoal,
      score: h.maxConfo, traits: h.stats.traits
    }))
    .sort((a, b) => b.score - a.score);
}

function readTraitGrades_(ss, statsTables, horseIds) {
  const results = {};
  const sheetConfigs = [
    { name: 'Stats', traitStart: 4, traitEnd: 15 },      // D-O (1-indexed)
    { name: 'ICE_Stats', traitStart: 4, traitEnd: 17 },   // D-Q
    { name: 'KATH_Stats', traitStart: 4, traitEnd: 16 }   // D-P
  ];

  const traitTotals = {};

  sheetConfigs.forEach((cfg, index) => {
    const table = statsTables ? statsTables[index] : readSheet_(ss, cfg.name);
    if (!table || table.length < 2) return;
    const fallback = index === 1 ? TRAIT_COLS_ICE : index === 2 ? TRAIT_COLS_KATH : TRAIT_COLS_STD;
    const headers = Array.from({length: cfg.traitEnd - cfg.traitStart + 1}, (_, i) => {
      const col = cfg.traitStart + i;
      return Object.keys(fallback).find(t => fallback[t] === col) || (index === 1 ? (col === 8 ? 'Tolt' : 'Pace') : 'Revaal');
    });
    const data = table.slice(1).filter(r => r[1] && (!horseIds || horseIds.has(String(r[1]).trim())))
      .map(r => r.slice(cfg.traitStart - 1, cfg.traitEnd));

    headers.forEach((traitName, colIdx) => {
      const tn = String(traitName || '').trim();
      if (!tn || tn === '#VG' || tn === '#G+' || tn === '#G' || tn.startsWith('#')) return;
      if (!traitTotals[tn]) traitTotals[tn] = { VG: 0, 'G+': 0, G: 0, 'G-': 0, A: 0, BA: 0, P: 0, total: 0 };

      data.forEach(row => {
        const raw = String(row[colIdx] || '').trim();
        if (!raw) return;
        const grade = dashGrade_(raw);
        let mapped = null;
        if (grade === 'VG' || grade === 'VG-') mapped = 'VG';
        else if (grade === 'G+') mapped = 'G+';
        else if (grade === 'G') mapped = 'G';
        else if (grade === 'G-') mapped = 'G-';
        else if (grade === 'A') mapped = 'A';
        else if (grade === 'BA') mapped = 'BA';
        else if (grade === 'P') mapped = 'P';
        if (mapped) { traitTotals[tn][mapped]++; traitTotals[tn].total++; }
      });
    });
  });

  const sorted = Object.keys(traitTotals)
    .filter(name => traitTotals[name].total > 0)
    .map(name => {
      const t = traitTotals[name];
      return { name, ...t, vgPercent: t.total > 0 ? Math.round((t.VG / t.total) * 100) : 0 };
    }).sort((a, b) => a.vgPercent - b.vgPercent);

  return { traits: sorted, strongest: sorted.length ? sorted[sorted.length - 1] : null, weakest: sorted.length ? sorted[0] : null };
}

function readHealthData_(trackerData, horses) {
  // Health columns in HERD TRACKER: AD=30, AE=31, AF=32, AG=33, AH=34, AI=35 (1-indexed)
  const HEALTH_COLS = { 'Colic Res.': 30, 'Hoof Quality': 31, 'Back Problems': 32, 'Resp. Disease': 33, 'Lameness Res.': 34, 'Fertility': 35 };
  const testNames = Object.keys(HEALTH_COLS);
  const gradeValues = { 'Excellent': 5, 'Good': 4, 'Average': 3, 'Fair': 2, 'Poor': 1 };

  const horseHealth = [];
  let fullyTested = 0, untested = 0;
  const testCoverage = {};
  const testGradeCounts = {};
  testNames.forEach(t => { testCoverage[t] = { tested: 0, total: 0 }; testGradeCounts[t] = { Excellent: 0, Good: 0, Average: 0, Fair: 0, Poor: 0 }; });

  // Weakest test tracking
  const testAvgScores = {};
  testNames.forEach(t => { testAvgScores[t] = { sum: 0, count: 0 }; });

  for (let i = 1; i < trackerData.length; i++) {
    const row = trackerData[i];
    const id = String(row[DASH_TRACKER_COLS.ID - 1] || '').trim();
    if (!id) continue;
    const horse = horses.find(h => h.id === id);
    if (!horse) continue;

    const hEntry = { name: horse.name, breed: horse.breed, isFoal: horse.isFoal, tests: {} };
    let testedCount = 0;

    testNames.forEach(testName => {
      const val = String(row[HEALTH_COLS[testName] - 1] || '').trim();
      testCoverage[testName].total++;
      if (val && val !== '' && val !== '-') {
        hEntry.tests[testName] = val;
        testedCount++;
        testCoverage[testName].tested++;
        // Normalize grade name
        const normalized = val.charAt(0).toUpperCase() + val.slice(1).toLowerCase();
        if (testGradeCounts[testName].hasOwnProperty(normalized)) testGradeCounts[testName][normalized]++;
        if (gradeValues[normalized]) { testAvgScores[testName].sum += gradeValues[normalized]; testAvgScores[testName].count++; }
      } else {
        hEntry.tests[testName] = null;
      }
    });

    if (testedCount === testNames.length) fullyTested++;
    if (testedCount === 0) untested++;
    horseHealth.push(hEntry);
  }

  // Weakest test
  let weakestTest = '—', weakestAvg = 99;
  testNames.forEach(t => {
    if (testAvgScores[t].count > 0) {
      const avg = testAvgScores[t].sum / testAvgScores[t].count;
      if (avg < weakestAvg) { weakestAvg = avg; weakestTest = t; }
    }
  });

  // Action needed: untested adults
  const actionNeeded = horseHealth
    .filter(h => !h.isFoal && Object.values(h.tests).some(v => v === null))
    .map(h => ({ name: h.name, breed: h.breed, missing: testNames.filter(t => !h.tests[t]) }))
    .slice(0, 20);

  return {
    fullyTested, untested, totalHorses: horseHealth.length, weakestTest,
    testNames, testCoverage, testGradeCounts,
    horses: horseHealth.sort((a, b) => {
      const aCount = Object.values(a.tests).filter(v => v).length;
      const bCount = Object.values(b.tests).filter(v => v).length;
      return bCount - aCount;
    }),
    actionNeeded
  };
}

/***************************************************************************
 * TAB 7: FINANCE — Balance, Costs, Passive Income
 ***************************************************************************/

function getBalance_(ss) {
  const sheet = ss.getSheetByName('Settings');
  const bal = { hrc: 0, dp: 0, ft: 0, wt: 0, netWorth: 0 };
  if (!sheet) return bal;
  const data = sheet.getRange('P1:Q20').getValues();
  data.forEach(r => { const key = String(r[0]).trim().toUpperCase(); const val = Number(r[1]) || 0; if (key === 'HRC') bal.hrc = val; else if (key === 'DP') bal.dp = val; else if (key === 'FT') bal.ft = val; else if (key === 'WT') bal.wt = val; else if (key === 'NET WORTH' || key === 'NETWORTH') bal.netWorth = val; });
  if (!bal.netWorth && (bal.hrc || bal.dp || bal.ft || bal.wt)) { const rates = loadCurrencyRates_(ss); bal.netWorth = bal.hrc + (bal.dp * rates.DP) + (bal.ft * rates.FT) + (bal.wt * rates.WT); }
  return bal;
}

function loadCurrencyRates_(ss) {
  const sheet = ss.getSheetByName('Settings'); const rates = { DP: 0, FT: 0, WT: 0 };
  if (!sheet) return rates;
  const data = sheet.getRange('M1:N20').getValues();
  data.forEach(r => { const key = String(r[0]).trim().toUpperCase(); if (rates.hasOwnProperty(key)) rates[key] = Number(r[1]) || 0; });
  return rates;
}

function calculateEarningsByType(activities) {
  const byType = {}; activities.forEach(a => { if (!byType[a.type]) byType[a.type] = { type: a.type, total: 0, count: 0 }; byType[a.type].total += a.net; byType[a.type].count++; });
  const income = INCOME_TYPES.filter(t => byType[t]).map(t => byType[t]); const expenses = EXPENSE_TYPES.filter(t => byType[t]).map(t => byType[t]);
  return { income: income.sort((a, b) => b.total - a.total), expenses: expenses.sort((a, b) => a.total - b.total), totalIncome: income.reduce((s, e) => s + e.total, 0), totalExpenses: expenses.reduce((s, e) => s + e.total, 0) };
}

function calculateBreedingCosts(activities) {
  const types = ['ultrasound', 'colour_test', 'foal_pastures', 'stud_out']; const breakdown = {};
  types.forEach(t => { const items = activities.filter(a => a.type === t); breakdown[t] = { type: t, total: Math.abs(items.reduce((s, a) => s + a.net, 0)), count: items.length }; });
  return { breakdown, totalCost: Object.values(breakdown).reduce((s, b) => s + b.total, 0) };
}

function calculateDailyPassiveIncome(activities) {
  const now = new Date(); const thirtyDaysAgo = new Date(now.getTime() - 30 * 86400000);
  const chores = activities.filter(a => a.type === 'chores' && toDate_(a.date) >= thirtyDaysAgo); const interest = activities.filter(a => a.type === 'interest' && toDate_(a.date) >= thirtyDaysAgo);
  const choresTotal = chores.reduce((s, a) => s + a.net, 0); const interestTotal = interest.reduce((s, a) => s + a.net, 0); const days = Math.max(1, Math.ceil((now - thirtyDaysAgo) / 86400000));
  return { choresTotal, interestTotal, choresDaily: Math.round(choresTotal / days), interestDaily: Math.round(interestTotal / days), totalDaily: Math.round((choresTotal + interestTotal) / days) };
}

function calculateROIByLine(activities, horses, pedigreeMap) {
  const earnings = {}; activities.filter(a => a.type === 'comp' || a.type === 'conf').forEach(a => { if (!a.horseId) return; if (!earnings[a.horseId]) earnings[a.horseId] = 0; earnings[a.horseId] += a.net; });
  const lineEarnings = {};
  horses.forEach(h => { const sireline = h.pedigree ? h.pedigree.sireline : 'Unknown'; if (!lineEarnings[sireline]) lineEarnings[sireline] = { line: sireline, totalEarnings: 0, horseCount: 0, horses: [] }; const hEarnings = earnings[h.id] || 0; lineEarnings[sireline].totalEarnings += hEarnings; lineEarnings[sireline].horseCount++; if (hEarnings > 0) lineEarnings[sireline].horses.push({ name: h.name, earnings: hEarnings }); });
  return Object.values(lineEarnings).map(l => ({ ...l, avgEarnings: l.horseCount > 0 ? Math.round(l.totalEarnings / l.horseCount) : 0 })).sort((a, b) => b.totalEarnings - a.totalEarnings);
}

/***************************************************************************
 * SHARED: RECENT ACTIVITY (Activity-Log-Tabelle, Finance-Tab)
 ***************************************************************************/

function getRecentActivity(activities, limit) {
  const TYPE_LABELS = { conf: 'Conformation', comp: 'Competition', stud_in: 'Stud (in)', stud_out: 'Stud (out)', sale: 'Sale', chores: 'Chores', interest: 'Interest', show_fee: 'Show Fee', comp_fee: 'Comp Fee', colour_test: 'Colour Test', ultrasound: 'Ultrasound', foal_pastures: 'Foal Pastures', purchase: 'Purchase', exchange: 'Exchange', refund: 'Refund', savings_in: 'Savings In', savings_out: 'Savings Out' };
  return activities.slice().reverse().filter(a => a.net !== 0 || a.amount !== 0).slice(0, limit).map(a => ({ date: a.date instanceof Date ? a.date.toLocaleDateString('en-US') : String(a.date || ''), type: a.type, typeLabel: TYPE_LABELS[a.type] || a.type, horse: a.horseName, detail: [a.detail, a.detail2, a.detail3].filter(Boolean).join(' – '), amount: a.net }));
}

/***************************************************************************
 * MARKET / SALES (ehemals Elite-Tab)
 ***************************************************************************/

function getTopHorses(h, t, l) { return [...h].sort((a, b) => t === 'gp' ? b.gp - a.gp : b.maxConfo - a.maxConfo).slice(0, l).map(x => ({ name: x.name, breed: x.breed, gp: x.gp, confoMax: x.maxConfo })); }

function findHorseForSale_(horses, s) {
  if (s.horseId) { const byId = horses.find(h => h.id === s.horseId); if (byId) return byId; }
  const normTarget = normalizeHorseName_(s.horseName);
  if (!normTarget) return null;
  return horses.find(h => normalizeHorseName_(h.name) === normTarget) || null;
}

function calculateMarketValueEstimate(activities, horses) {
  const sales = activities.filter(a => a.type === 'sale' && a.amount > 0); const purchases = activities.filter(a => a.type === 'purchase' && a.amount !== 0);
  const byBreed = {};
  sales.forEach(s => { const horse = findHorseForSale_(horses, s); const breed = horse ? horse.breed : 'Unknown'; if (!byBreed[breed]) byBreed[breed] = { breed, totalPrice: 0, count: 0, prices: [] }; byBreed[breed].totalPrice += s.amount; byBreed[breed].count++; byBreed[breed].prices.push(s.amount); });
  const avgPrice = sales.length > 0 ? Math.round(sales.reduce((s, a) => s + a.amount, 0) / sales.length) : 0;
  const totalRevenue = sales.reduce((s, a) => s + a.amount, 0); const totalSpent = Math.abs(purchases.reduce((s, a) => s + a.net, 0));
  const recentSales = sales.slice().reverse().slice(0, 10).map(s => { const horse = findHorseForSale_(horses, s); return { date: s.date instanceof Date ? s.date.toLocaleDateString('en-US') : String(s.date), name: s.horseName, breed: horse ? horse.breed : '—', buyer: s.detail || '—', price: s.amount }; });
  return { avgPrice, totalSales: sales.length, totalPurchases: purchases.length, totalRevenue, totalSpent, netTradeBalance: totalRevenue - totalSpent, byBreed: Object.values(byBreed).map(b => ({ ...b, avgPrice: Math.round(b.totalPrice / b.count), minPrice: Math.min(...b.prices), maxPrice: Math.max(...b.prices) })).sort((a, b) => b.avgPrice - a.avgPrice), recentSales, highestSale: sales.length ? Math.max(...sales.map(s => s.amount)) : 0 };
}

/***************************************************************************
 * TAB 9: XP & BALANCE
 ***************************************************************************/

function getInterestRate_(amount) {
  for (const tier of INTEREST_TIERS) if (amount <= tier.max) return tier.rate;
  return 0;
}

function calculateInterestInfo_(xpEntry) {
  const savings = xpEntry ? Number(xpEntry.savings) || 0 : 0;
  if (savings <= 0) return { savings: 0, rate: 0, daily: 0, monthly: 0, tip: '' };

  const rate = getInterestRate_(savings);
  const daily = Math.floor(savings * (rate / 100));
  const monthly = daily * 30;

  let tip = '';
  if (rate === 0) {
    tip = 'Over 25M HRC earns no more interest. Keep the surplus as cash.';
  } else {
    const tierIndex = INTEREST_TIERS.findIndex(t => savings <= t.max);
    if (tierIndex > 0) {
      const prevCap = PROFITABLE_CAPS[tierIndex - 1];
      const prevRate = INTEREST_TIERS[tierIndex - 1].rate;
      const dailyAtPrevCap = Math.floor(prevCap * (prevRate / 100));
      tip = dailyAtPrevCap > daily
        ? 'Tip: ' + prevCap.toLocaleString('en-US') + ' HRC in Savings (' + prevRate + '%) yields ' + dailyAtPrevCap.toLocaleString('en-US') + '/day — more than now. Keep the rest as cash.'
        : 'Current savings amount is efficient for ' + rate + '%.';
    } else {
      tip = 'Optimal interest bracket (1.50%), up to ' + PROFITABLE_CAPS[0].toLocaleString('en-US') + ' HRC.';
    }
  }

  return { savings: savings, rate: rate, daily: daily, monthly: monthly, tip: tip };
}

function getLastXpEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('XP Tracker');
  if (!sheet) return null;
  const lastRow = lastDataRowByColumnA_(sheet);
  if (lastRow < 2) return null;
  const row = sheet.getRange(lastRow, 1, 1, 23).getValues()[0];
  return {
    lastDate: row[0] instanceof Date ? row[0].toLocaleDateString('en-US') : String(row[0] || ''),
    level: Number(row[1]) || 0, rank: '', xp: Number(row[3]) || 0,
    sinceLast: Number(row[4]) || 0, dayTotal: Number(row[5]) || 0,
    tilNextLevel: Number(row[6]) || 0, percent: Number(row[7]) || 0,
    hrc: Number(row[8]) || 0, savings: Number(row[9]) || 0,
    dp: Number(row[10]) || 0, ft: Number(row[11]) || 0, wt: Number(row[12]) || 0,
    netWorth: Number(row[13]) || 0,
    dollarChange: Number(row[14]) || 0, dollarChangeDay: Number(row[15]) || 0,
    cycle: { conf: Number(row[17]) || 0, comp: Number(row[18]) || 0, sales: Number(row[19]) || 0, stud: Number(row[20]) || 0 },
    totalStalls: Number(row[21]) || 0, horsesOnPasture: Number(row[22]) || 0
  };
}

function getXpHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('XP Tracker');
  if (!sheet) return [];
  const lastRow = lastDataRowByColumnA_(sheet);
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  const rows = data.map(row => ({
    date: row[0] instanceof Date ? row[0].toLocaleDateString('en-US') : String(row[0] || ''),
    level: Number(row[1]) || 0, xp: Number(row[3]) || 0,
    netWorth: Number(row[13]) || 0, changeDay: Number(row[15]) || 0
  }));
  // Pro Kalendertag nur EIN Datenpunkt (letzter Snapshot des Tages), um doppelte Einträge zu vermeiden
  const byDate = {};
  const order = [];
  rows.forEach(r => {
    if (!byDate[r.date]) order.push(r.date);
    byDate[r.date] = r; // letzter Eintrag pro Datum gewinnt
  });
  return order.map(d => byDate[d]);
}

function saveXpSnapshot(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('XP Tracker');
    if (!sheet) return { success: false, error: 'XP Tracker sheet not found' };
    const today = new Date();
    const lastRow = lastDataRowByColumnA_(sheet);
    const hrc = Number(data.hrc) || 0; const savings = Number(data.savings) || 0;
    const dp = Number(data.dp) || 0; const ft = Number(data.ft) || 0; const wt = Number(data.wt) || 0;
    const rates = loadCurrencyRates_(ss);
    const netWorth = hrc + savings + (dp * rates.DP) + (ft * rates.FT) + (wt * rates.WT);
    let dollarChange = 0, dollarChangeDay = 0;
    if (lastRow >= 2) { const prevNetWorth = Number(sheet.getRange(lastRow, 14).getValue()) || 0; dollarChange = netWorth - prevNetWorth; dollarChangeDay = dollarChange; }
    const cycleEarnings = calculateCycleEarnings_(ss);
    const newRow = [ today, Number(data.level) || 0, data.rank || '', Number(data.xp) || 0, Number(data.sinceLast) || 0, Number(data.dayTotal) || 0, Number(data.tilNext) || 0, (Number(data.percent) || 0) / 100, hrc, savings, dp, ft, wt, netWorth, dollarChange, dollarChangeDay, '', cycleEarnings.conf, cycleEarnings.comp, cycleEarnings.sales, cycleEarnings.stud, Number(data.totalStalls) || 0, Number(data.horsesOnPasture) || 0 ];
    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    updateSettingsBalance_(ss, hrc, savings, dp, ft, wt, netWorth);
    return { success: true, netWorth: netWorth };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * Einmalig ausführen: korrigiert bestehende $Change/Day-Werte (Spalte P) rückwirkend,
 * da diese bisher fälschlich gegen Spalte O statt Spalte N (Net Worth) berechnet wurden.
 */
function repairChangeDayHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('XP Tracker');
  if (!sheet) return;
  const lastRow = lastDataRowByColumnA_(sheet);
  if (lastRow < 3) return;
  const netWorths = sheet.getRange(2, 14, lastRow - 1, 1).getValues().map(r => Number(r[0]) || 0);
  const changes = netWorths.map((nw, i) => i === 0 ? 0 : nw - netWorths[i - 1]);
  sheet.getRange(2, 15, changes.length, 1).setValues(changes.map(c => [c])); // Spalte O
  sheet.getRange(2, 16, changes.length, 1).setValues(changes.map(c => [c])); // Spalte P
}

function calculateCycleEarnings_(ss) {
  const sheet = ss.getSheetByName('Activity Log');
  if (!sheet) return { conf: 0, comp: 0, sales: 0, stud: 0 };
  const data = sheet.getDataRange().getValues();
  let conf = 0, comp = 0, sales = 0, stud = 0;
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][1] || '').trim(); const net = Number(data[i][10]) || 0;
    if (type === 'conf') conf += net; else if (type === 'comp') comp += net;
    else if (type === 'sale') sales += net; else if (type === 'stud_in') stud += net;
  }
  return { conf, comp, sales, stud };
}

function updateSettingsBalance_(ss, hrc, savings, dp, ft, wt, netWorth) {
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return;
  const data = sheet.getRange('P1:Q20').getValues();
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]).trim().toUpperCase();
    if (key === 'HRC') sheet.getRange(i + 1, 17).setValue(hrc + savings);
    else if (key === 'DP') sheet.getRange(i + 1, 17).setValue(dp);
    else if (key === 'FT') sheet.getRange(i + 1, 17).setValue(ft);
    else if (key === 'WT') sheet.getRange(i + 1, 17).setValue(wt);
    else if (key === 'NET WORTH' || key === 'NETWORTH') sheet.getRange(i + 1, 17).setValue(netWorth);
  }
}

/***************************************************************************
 * XP MODAL & DAILY REMINDER
 ***************************************************************************/

function showXpBalanceModal() {
  const tpl = HtmlService.createTemplateFromFile('XpBalanceModal');
  const html = tpl.evaluate().setWidth(940).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function checkDailyXpReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('XP Tracker');
  if (!sheet) return;
  const lastRow = lastDataRowByColumnA_(sheet);
  if (lastRow < 2) { ss.toast('No XP snapshot yet. Menu → Horse Management → Daily Snapshot', '📊 Daily Snapshot', 10); return; }
  const lastDate = sheet.getRange(lastRow, 1).getValue();
  if (lastDate instanceof Date) {
    const today = new Date();
    if (lastDate.getFullYear() !== today.getFullYear() || lastDate.getMonth() !== today.getMonth() || lastDate.getDate() !== today.getDate()) {
      ss.toast('No snapshot logged today yet! Menu → Horse Management → Daily Snapshot', '📊 Daily Snapshot', 10);
    }
  }
}

/***************************************************************************
 * TAB 6: PERFORMANCE
 ***************************************************************************/

function calculateTopPerformers(activities) {
  const horses = {};
  activities.filter(a => a.type === 'comp' || a.type === 'conf').forEach(a => { const key = a.horseId || a.horseName; if (!horses[key]) horses[key] = { name: a.horseName, id: a.horseId, wins: 0, seconds: 0, thirds: 0, total: 0, earnings: 0 }; const h = horses[key]; h.total++; h.earnings += a.net; const place = a.detail3 || a.detail2; if (String(place).includes('1st') || String(place).includes('Day Champion')) h.wins++; else if (String(place).includes('2nd')) h.seconds++; else if (String(place).includes('3rd')) h.thirds++; });
  return Object.values(horses).sort((a, b) => b.earnings - a.earnings).slice(0, 50);
}

function calculateEventEfficiency(activities) {
  const compEarnings = activities.filter(a => a.type === 'comp').reduce((s, a) => s + a.net, 0); const compFees = Math.abs(activities.filter(a => a.type === 'comp_fee').reduce((s, a) => s + a.net, 0));
  const confEarnings = activities.filter(a => a.type === 'conf').reduce((s, a) => s + a.net, 0); const confFees = Math.abs(activities.filter(a => a.type === 'show_fee').reduce((s, a) => s + a.net, 0));
  return { comp: { earnings: compEarnings, fees: compFees, net: compEarnings - compFees, roi: compFees > 0 ? ((compEarnings / compFees) * 100).toFixed(0) : '—' }, conf: { earnings: confEarnings, fees: confFees, net: confEarnings - confFees, roi: confFees > 0 ? ((confEarnings / confFees) * 100).toFixed(0) : '—' } };
}

function calculateShowPerformanceIndex(activities) {
  const totalWinnings = activities.filter(a => a.type === 'comp').reduce((s, a) => s + a.net, 0) + activities.filter(a => a.type === 'conf').reduce((s, a) => s + a.net, 0);
  const totalFees = Math.abs(activities.filter(a => a.type === 'comp_fee').reduce((s, a) => s + a.net, 0) + activities.filter(a => a.type === 'show_fee').reduce((s, a) => s + a.net, 0));
  return { totalWinnings, totalFees, netProfit: totalWinnings - totalFees, index: totalFees > 0 ? (totalWinnings / totalFees).toFixed(2) : '—' };
}

function calculateDisciplineBreakdown(activities) {
  const map = {};
  activities.filter(a => a.type === 'comp').forEach(a => { const disc = a.detail || 'Unknown'; if (!map[disc]) map[disc] = { discipline: disc, count: 0, earnings: 0, wins: 0 }; map[disc].count++; map[disc].earnings += a.net; if (String(a.detail3).includes('1st')) map[disc].wins++; });
  return Object.values(map).sort((a, b) => b.earnings - a.earnings);
}

// Normalisiert Rassen-Kurzformen (z.B. "Finn") auf den vollen Rassenamen aus der Herde (z.B. "Finnhorse")
function normalizeBreedName_(raw, breeds) {
  const r = String(raw || '').trim();
  if (!r) return 'Unknown';
  if (breeds.indexOf(r) !== -1) return r;
  const rl = r.toLowerCase();
  const match = breeds.find(b => b.toLowerCase() === rl) ||
                breeds.find(b => b.toLowerCase().startsWith(rl) || rl.startsWith(b.toLowerCase())) ||
                breeds.find(b => b.toLowerCase().includes(rl) || rl.includes(b.toLowerCase()));
  return match || r;
}

function calculateEarningsOverview(activities, breeds) {
  const breedMap = {}; activities.filter(a => a.type === 'conf').forEach(a => { const breed = normalizeBreedName_(a.detail, breeds || []); if (!breedMap[breed]) breedMap[breed] = { breed, earnings: 0, count: 0 }; breedMap[breed].earnings += a.net; breedMap[breed].count++; });
  const discMap = {}; activities.filter(a => a.type === 'comp').forEach(a => { const disc = a.detail || 'Unknown'; if (!discMap[disc]) discMap[disc] = { discipline: disc, earnings: 0, count: 0 }; discMap[disc].earnings += a.net; discMap[disc].count++; });
  return { byBreed: Object.values(breedMap).sort((a, b) => b.earnings - a.earnings), byDiscipline: Object.values(discMap).sort((a, b) => b.earnings - a.earnings) };
}

function calculateGrowthEvents(horses) {
  const now = new Date(); const events = [];
  horses.forEach(h => { if (!h.birthDate || !(h.birthDate instanceof Date)) return; const birth = new Date(h.birthDate); const daysPassed = (now - birth) / (1000 * 60 * 60 * 24); let nextMilestone = null;
    if (daysPassed < 8) nextMilestone = { type: "Weaning (6 mo.)", daysLeft: (8 - daysPassed).toFixed(1), date: new Date(birth.getTime() + 8 * 86400000) };
    else if (daysPassed < 48) nextMilestone = { type: "Adult / Rideable (3 y.)", daysLeft: (48 - daysPassed).toFixed(1), date: new Date(birth.getTime() + 48 * 86400000) };
    else if (daysPassed < 112 && h.gender === 'Stallion') nextMilestone = { type: "Clinic / Approved (7 y.)", daysLeft: (112 - daysPassed).toFixed(1), date: new Date(birth.getTime() + 112 * 86400000) };
    if (nextMilestone) events.push({ name: h.name, breed: h.breed, type: nextMilestone.type, targetDate: nextMilestone.date.toLocaleDateString('en-US'), daysLeft: nextMilestone.daysLeft, isUrgent: parseFloat(nextMilestone.daysLeft) < 2 });
  });
  return events.sort((a, b) => parseFloat(a.daysLeft) - parseFloat(b.daysLeft));
}

function calculateFoalPastureEvents(activities) {
  const fpEvents = activities.filter(a => a.type === 'foal_pastures' && (a.detail === 'Returned' || a.detail === 'Notice' || a.detail === 'Left'));
  const latestByHorse = {};
  fpEvents.forEach(a => { const key = a.horseName; const date = toDate_(a.date); if (!latestByHorse[key] || toDate_(latestByHorse[key].date) <= date) { if (latestByHorse[key] && toDate_(latestByHorse[key].date).getTime() === date.getTime() && latestByHorse[key].detail === 'Returned') return; latestByHorse[key] = a; } });
  return Object.values(latestByHorse).filter(a => { if (a.detail === 'Notice') { return (parseInt(a.detail2) || 0) > 0; } return true; }).map(a => ({ date: a.date instanceof Date ? a.date.toLocaleDateString('en-US') : String(a.date), horseName: a.horseName, status: a.detail, daysLeft: a.detail2 || '' })).slice(0, 20);
}

function calculateBreedingCheckEvents(pregnancies) {
  const now = new Date();
  return pregnancies.filter(p => p.coverDateTime && !p.confirmed).map(p => {
    const hoursPassed = (now - p.coverDateTime) / (1000 * 60 * 60);
    const hoursLeft = 120 - hoursPassed;
    const targetDate = new Date(p.coverDateTime.getTime() + 120 * 60 * 60 * 1000);
    return {
      name: p.name,
      breed: p.breed,
      phase: p.phase,
      targetDate: targetDate.toLocaleDateString('en-US'),
      daysLeft: (hoursLeft / 24).toFixed(1)
    };
  }).filter(e => parseFloat(e.daysLeft) >= -30) // Mehr als 1 Monat überfällig: ausblenden
    .sort((a, b) => parseFloat(a.daysLeft) - parseFloat(b.daysLeft));
}

/***************************************************************************
 * GENERISCHE HELPER — dateiübergreifend genutzt
 ***************************************************************************/

function toDate_(d) {
  if (d instanceof Date) return d;
  const parsed = new Date(d);
  return isNaN(parsed.getTime()) ? new Date(0) : parsed;
}

function columnLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Entfernt Sonderzeichen/Symbole (☆, °, Geschlechts-Icons etc.) für robusten Namensabgleich
function normalizeHorseName_(name) {
  return String(name || '').replace(/[^\p{L}\p{N} ]/gu, '').trim().toLowerCase();
}
