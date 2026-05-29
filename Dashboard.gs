/*** ============================================
 * HERD DASHBOARD BACKEND — COMMUNITY VERSION
 * ============================================*/

const SHEETS = {
  TRACKER:    'HERD TRACKER',
  STATS:      'Horse Stats',
  STATS_ICE:  'ICE_Horse Stats',
  STATS_KATH: 'KATH_Horse Stats',
  BROODMARES: 'Broodmares',
  PEDIGREE:   'Pedigree'
};

const TRACKER_COLS = {
  STATUS: 2, ID: 3, NAME: 4, BREED: 5, GENDER: 6, BIRTH_DATE: 7,
  AGE: 9, COAT: 14, BREEDING_STATUS: 16, GP: 19, MAX_CONFO: 22,
  RECC_DISCIPLINE: 24, TRAINED_DISCIPLINE: 25
};

const STATS_COLS      = { ID: 2, NAME: 3, VG_COUNT: 28, TOTAL_GP: 44, CONFO_MAX: 63 };
const STATS_COLS_ICE  = { ID: 2, NAME: 3, VG_COUNT: 30, TOTAL_GP: 46, CONFO_MAX: 65 }; // +2 for Tölt/Pace
const STATS_COLS_KATH = { ID: 2, NAME: 3, VG_COUNT: 29, TOTAL_GP: 45, CONFO_MAX: 64 }; // +1 for Revaal

const PEDIGREE_COLS   = { ID: 2, NAME: 3, DAMLINE: 4, SIRELINE: 5 };
const BROODMARES_COLS = { CONFIRMED: 2, ID: 3, NAME: 4, BREED: 5, COVER_DATE: 6, COVER_TIME: 7, DUE_DATE: 8 };

function showDashboard() {
  const html = HtmlService.createTemplateFromFile('HerdDashboard')
    .evaluate()
    .setWidth(1200)
    .setHeight(900)
    .setTitle('Herd Tracker Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getDashboardData() {
  try {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const trackerSheet   = ss.getSheetByName(SHEETS.TRACKER);
    const statsSheet     = ss.getSheetByName(SHEETS.STATS);
    const statsIceSheet  = ss.getSheetByName(SHEETS.STATS_ICE);
    const statsKathSheet = ss.getSheetByName(SHEETS.STATS_KATH);
    const pedigreeSheet  = ss.getSheetByName(SHEETS.PEDIGREE);
    const breedingSheet  = ss.getSheetByName(SHEETS.BROODMARES);

    const trackerData  = trackerSheet   ? trackerSheet.getDataRange().getValues()   : [];
    const statsData    = statsSheet     ? statsSheet.getDataRange().getValues()     : [];
    const statsIceData = statsIceSheet  ? statsIceSheet.getDataRange().getValues()  : [];
    const statsKathData= statsKathSheet ? statsKathSheet.getDataRange().getValues() : [];
    const pedigreeData = pedigreeSheet  ? pedigreeSheet.getDataRange().getValues()  : [];
    const breedingData = breedingSheet  ? breedingSheet.getDataRange().getValues()  : [];

    const statsMap   = parseAllStatsData(statsData, statsIceData, statsKathData);
    const pedigreeMap= parsePedigreeData(pedigreeData);
    const pregnancies= parseBreedingData(breedingData);
    const horses     = parseTrackerData(trackerData, statsMap, pedigreeMap);
    const breedStats = calculateBreedStats(horses, pregnancies);

    return {
      totalHorses:           horses.length,
      stallions:             horses.filter(h => h.gender === 'Stallion').length,
      mares:                 horses.filter(h => h.gender === 'Mare').length,
      foals:                 horses.filter(h => h.isFoal).length,
      adults:                horses.filter(h => !h.isFoal).length,
      pregnant:              pregnancies.filter(p => p.confirmed).length,
      genderRatio:           calculateGenderRatio(horses),
      breeds:                getUniqueBreeds(horses),
      breedStats:            breedStats,
      roleStats:             calculateRoleStats(horses),
      vgStats:               calculateVGStats(horses),
      generationComparison:  calculateGenerationComparison(horses),
      generationByBreed:     calculateGenerationComparisonByBreed(horses),
      disciplineFocus:       calculateDisciplineFocus(horses),
      colorStats:            calculateColorStats(horses),
      sirelines:             calculateLineageStats(horses, 'sireline'),
      damlines:              calculateLineageStats(horses, 'damline'),
      topGP:                 getTopHorses(horses, 'gp', 6),
      topConfo:              getTopHorses(horses, 'confo', 6),
      upcomingBirths:        calculateUpcomingBirths(pregnancies),
      growthEvents:          calculateGrowthEvents(horses)
    };
  } catch (e) { return { error: e.toString() }; }
}

function parseTrackerData(data, statsMap, pedigreeMap) {
  const horses = [];
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const id   = String(row[TRACKER_COLS.ID - 1] || '').trim();
    if (!id) continue;
    const horseStats = statsMap[id] || { vgCount: 0, totalGP: 0, confoMax: 0 };
    horses.push({
      id, name: row[TRACKER_COLS.NAME - 1] || 'Unknown',
      breed:           row[TRACKER_COLS.BREED - 1]            || 'Unknown',
      gender:          row[TRACKER_COLS.GENDER - 1]           || '',
      birthDate:       row[TRACKER_COLS.BIRTH_DATE - 1],
      age:             row[TRACKER_COLS.AGE - 1],
      isFoal:          parseAge(row[TRACKER_COLS.AGE - 1]) < 3,
      breedingStatus:  row[TRACKER_COLS.BREEDING_STATUS - 1]  || '',
      coat:            row[TRACKER_COLS.COAT - 1]             || 'Unknown',
      gp:              parseFloat(row[TRACKER_COLS.GP - 1])   || horseStats.totalGP || 0,
      maxConfo:        parseFloat(row[TRACKER_COLS.MAX_CONFO - 1]) || horseStats.confoMax || 0,
      reccDiscipline:  row[TRACKER_COLS.RECC_DISCIPLINE - 1]  || '',
      trainedDiscipline: row[TRACKER_COLS.TRAINED_DISCIPLINE - 1] || '',
      stats:           horseStats,
      pedigree:        pedigreeMap[id] || { sireline: 'Unknown', damline: 'Unknown' }
    });
  }
  return horses;
}

function parseAllStatsData(standardData, iceData, kathData) {
  const map = {};
  // Kathiawari first, then Icelandic, then Standard (first match wins)
  [[kathData, STATS_COLS_KATH], [iceData, STATS_COLS_ICE], [standardData, STATS_COLS]]
    .forEach(([data, cols]) => {
      if (!data || !data.length) return;
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][cols.ID - 1]).trim();
        if (id && !map[id]) map[id] = {
          vgCount:  parseInt(data[i][cols.VG_COUNT - 1])  || 0,
          totalGP:  parseFloat(data[i][cols.TOTAL_GP - 1]) || 0,
          confoMax: parseFloat(data[i][cols.CONFO_MAX - 1])|| 0
        };
      }
    });
  return map;
}

function parsePedigreeData(data) {
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][PEDIGREE_COLS.ID - 1]).trim();
    if (id) map[id] = {
      damline:  data[i][PEDIGREE_COLS.DAMLINE - 1]  || 'Unknown',
      sireline: data[i][PEDIGREE_COLS.SIRELINE - 1] || 'Unknown'
    };
  }
  return map;
}

function parseBreedingData(data) {
  const pregnancies = [];
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][BROODMARES_COLS.ID - 1]).trim();
    if (!id) continue;

    const rawConfirmed = data[i][BROODMARES_COLS.CONFIRMED - 1];
    const isConfirmed  = rawConfirmed === true || String(rawConfirmed).toLowerCase() === 'true';

    let coverDate = data[i][BROODMARES_COLS.COVER_DATE - 1];
    if (typeof coverDate === 'string' && coverDate.includes('.')) {
      const parts = coverDate.split('.');
      coverDate   = new Date(parts[2], parts[1]-1, parts[0]);
    }
    let coverDateTime = null;
    if (coverDate instanceof Date && !isNaN(coverDate)) {
      coverDateTime = new Date(coverDate);
      const coverTime = data[i][BROODMARES_COLS.COVER_TIME - 1];
      if (coverTime instanceof Date) coverDateTime.setHours(coverTime.getHours(), coverTime.getMinutes());
    }

    let rawDueDate = data[i][BROODMARES_COLS.DUE_DATE - 1];
    let dueDate = null;
    if (rawDueDate instanceof Date && !isNaN(rawDueDate.getTime())) {
      dueDate = rawDueDate;
    } else if (rawDueDate) {
      let parts = String(rawDueDate).replace("EFD:", "").trim().split(" ")[0];
      parts = parts.includes('.') ? parts.split('.') : parts.split('/');
      if (parts.length === 3) dueDate = new Date(parts[2], parts[0]-1, parts[1]);
    }

    let phase = "Waiting";
    if (coverDateTime) {
      const hoursPassed = (now - coverDateTime) / (1000 * 60 * 60);
      if      (hoursPassed < 48)  phase = "Check: 48h (in "  + Math.round(48  - hoursPassed) + "h)";
      else if (hoursPassed < 96)  phase = "Check: 96h (in "  + Math.round(96  - hoursPassed) + "h)";
      else if (hoursPassed < 120) phase = "Ultrasound (in "  + Math.round(120 - hoursPassed) + "h)";
      else                        phase = "In foal";
    }

    pregnancies.push({
      id, confirmed: isConfirmed,
      name:  data[i][BROODMARES_COLS.NAME - 1],
      breed: data[i][BROODMARES_COLS.BREED - 1],
      dueDate, coverDateTime, phase
    });
  }
  return pregnancies;
}

function calculateBreedStats(horses, pregnancies) {
  const s = {};
  horses.forEach(h => {
    if (!s[h.breed]) s[h.breed] = {
      total: 0, stallions: 0, mares: 0, foals: 0, pregnant: 0,
      gpSum: 0, gpCount: 0, vgSum: 0, confoSum: 0, confoCount: 0,
      roles: { stud: 0, broodmare: 0, public: 0, none: 0 }
    };
    const b = s[h.breed];
    b.total++;
    if (h.gender === 'Stallion') b.stallions++;
    if (h.gender === 'Mare')     b.mares++;
    if (h.isFoal)                b.foals++;
    if (pregnancies.some(p => p.id === h.id && p.confirmed)) b.pregnant++;

    const st = h.breedingStatus.toLowerCase();
    if      (st.includes('stud') && !st.includes('public')) b.roles.stud++;
    else if (st.includes('broodmare')) b.roles.broodmare++;
    else if (st.includes('public'))    b.roles.public++;
    else                               b.roles.none++;

    if (h.gp > 0) { b.gpSum += h.gp; b.gpCount++; }
    b.vgSum += h.stats.vgCount;
    if (h.maxConfo > 0) { b.confoSum += h.maxConfo; b.confoCount++; }
  });
  for (let b in s) {
    const i   = s[b];
    i.avgGP    = i.gpCount    > 0 ? Math.round(i.gpSum / i.gpCount)          : 0;
    i.avgVG    = i.total      > 0 ? (i.vgSum / i.total).toFixed(2)           : 0;
    i.avgConfo = i.confoCount > 0 ? (i.confoSum / i.confoCount).toFixed(3)   : "0.000";
  }
  return s;
}

function calculateVGStats(horses) {
  const res = { all: { vg12: 0, vg11: 0, vg10: 0, totalWithStats: 0 } };
  horses.forEach(h => {
    if (!res[h.breed]) res[h.breed] = { vg12: 0, vg11: 0, vg10: 0, totalWithStats: 0 };
    [res.all, res[h.breed]].forEach(t => {
      if      (h.stats.vgCount >= 12) t.vg12++;
      else if (h.stats.vgCount === 11) t.vg11++;
      else if (h.stats.vgCount === 10) t.vg10++;
      t.totalWithStats++;
    });
  });
  return res;
}

function calculateColorStats(horses) {
  const patterns = ["tobiano","overo","sabino","rabicano","splash","white","leopard","blanket","varnish","spotted","appaloosa","snowflake"];
  const res = { all: { solid: {}, patterned: {} } };
  horses.forEach(h => {
    const cat = patterns.some(p => h.coat.toLowerCase().includes(p)) ? 'patterned' : 'solid';
    if (!res[h.breed]) res[h.breed] = { solid: {}, patterned: {} };
    [res.all, res[h.breed]].forEach(t => { t[cat][h.coat] = (t[cat][h.coat] || 0) + 1; });
  });
  for (let b in res) {
    const total = Object.values(res[b].solid).reduce((a,c)=>a+c,0) + Object.values(res[b].patterned).reduce((a,c)=>a+c,0);
    const map = (src) => Object.keys(src).map(c => ({ color: c, count: src[c], percentage: ((src[c]/total)*100).toFixed(1) })).sort((a,b)=>b.count-a.count);
    res[b] = { solid: map(res[b].solid), patterned: map(res[b].patterned) };
  }
  return res;
}

function calculateLineageStats(horses, lineType) {
  const res = { all: {} };
  horses.forEach(horse => {
    if (!horse.pedigree) return;
    const breed    = horse.breed || 'Unknown';
    const lineName = lineType === 'sireline' ? horse.pedigree.sireline : horse.pedigree.damline;
    if (!lineName) return;
    if (!res[breed]) res[breed] = {};
    [res.all, res[breed]].forEach(target => {
      if (!target[lineName]) target[lineName] = { name: lineName, total: 0, stallions: 0, mares: 0, gpSum: 0, validCount: 0 };
      const l = target[lineName];
      l.total++;
      if (horse.gender === 'Stallion') l.stallions++;
      if (horse.gender === 'Mare')     l.mares++;
      const gp = horse.gp || (horse.stats ? horse.stats.totalGP : 0);
      if (gp > 0) { l.gpSum += gp; l.validCount++; }
    });
  });
  const final = {};
  Object.keys(res).forEach(b => {
    const bTotal = Object.values(res[b]).reduce((acc,curr)=>acc+curr.total,0) || 1;
    final[b] = Object.values(res[b]).map(line => {
      const share        = (line.total / bTotal) * 100;
      const stallionRatio= (line.stallions / line.total) * 100;
      let domColor = '#2E7D32', domLabel = 'Stable';
      if      (share > 40) { domColor = '#C62828'; domLabel = 'Dominant'; }
      else if (share > 25) { domColor = '#F9A825'; domLabel = 'Notable';  }
      return {
        ...line, share: share.toFixed(1),
        avgGP: line.validCount > 0 ? Math.round(line.gpSum / line.validCount) : 0,
        domColor, domLabel, stallionRatio,
        balanceWarning: stallionRatio > 60 && line.total > 2
      };
    }).sort((a,b) => b.total - a.total);
  });
  return final;
}

function calculateDisciplineFocus(horses) {
  const res = { all: {} };
  horses.forEach(h => {
    const d = h.reccDiscipline || h.trainedDiscipline || 'None';
    if (!res[h.breed]) res[h.breed] = {};
    [res.all, res[h.breed]].forEach(t => {
      if (!t[d]) t[d] = { count: 0, gpSum: 0, confoSum: 0, valid: 0 };
      t[d].count++;
      if (h.gp > 0) { t[d].gpSum += h.gp; t[d].confoSum += h.maxConfo; t[d].valid++; }
    });
  });
  for (let b in res) {
    for (let d in res[b]) {
      const i = res[b][d];
      i.avgGP    = i.valid > 0 ? i.gpSum    / i.valid : 0;
      i.avgConfo = i.valid > 0 ? i.confoSum / i.valid : 0;
    }
  }
  return res;
}

function calculateGenerationComparison(arr) {
  const calc = (list) => ({
    avgGP:    list.length ? list.reduce((a,b)=>a+b.gp,0)              / list.length : 0,
    avgConfo: list.length ? list.reduce((a,b)=>a+b.maxConfo,0)        / list.length : 0,
    avgVG:    list.length ? list.reduce((a,b)=>a+b.stats.vgCount,0)   / list.length : 0
  });
  return { adults: calc(arr.filter(h=>!h.isFoal)), foals: calc(arr.filter(h=>h.isFoal)) };
}

function calculateGenerationComparisonByBreed(horses) {
  const res = {};
  getUniqueBreeds(horses).forEach(b => { res[b] = calculateGenerationComparison(horses.filter(h=>h.breed===b)); });
  return res;
}

function calculateGrowthEvents(horses) {
  const now = new Date();
  const events = [];
  horses.forEach(h => {
    if (!h.birthDate || !(h.birthDate instanceof Date)) return;
    const birth      = new Date(h.birthDate);
    const daysPassed = (now - birth) / (1000 * 60 * 60 * 24);
    let nextMilestone = null;

    if      (daysPassed < 8)                                { nextMilestone = { type: "Weaning (6 months)",         daysLeft: (8   - daysPassed).toFixed(1), date: new Date(birth.getTime() + 8   * 86400000) }; }
    else if (daysPassed < 48)                               { nextMilestone = { type: "Adult / Rideable (3 years)", daysLeft: (48  - daysPassed).toFixed(1), date: new Date(birth.getTime() + 48  * 86400000) }; }
    else if (daysPassed < 112 && h.gender === 'Stallion')   { nextMilestone = { type: "Clinic / Approved (7 years)",daysLeft: (112 - daysPassed).toFixed(1), date: new Date(birth.getTime() + 112 * 86400000) }; }

    if (nextMilestone) {
      events.push({
        name: h.name, breed: h.breed,
        type:       nextMilestone.type,
        targetDate: nextMilestone.date.toLocaleDateString('en-GB'),
        daysLeft:   nextMilestone.daysLeft,
        isUrgent:   parseFloat(nextMilestone.daysLeft) < 2
      });
    }
  });
  return events.sort((a,b) => parseFloat(a.daysLeft) - parseFloat(b.daysLeft));
}

function parseAge(s)           { const m = String(s).match(/(\d+)\s*y/i); return m ? parseInt(m[1]) : 0; }
function getUniqueBreeds(h)    { return [...new Set(h.map(x=>x.breed))].sort(); }
function calculateGenderRatio(h) {
  const s = h.filter(x=>x.gender==='Stallion').length, m = h.filter(x=>x.gender==='Mare').length;
  return m === 0 ? s+':0' : (s/m).toFixed(2)+':1';
}
function calculateRoleStats(h) {
  const r = { stud: 0, broodmare: 0, public: 0, none: 0 };
  h.forEach(x => {
    const s = x.breedingStatus.toLowerCase();
    if      (s.includes('stud') && !s.includes('public')) r.stud++;
    else if (s.includes('broodmare')) r.broodmare++;
    else if (s.includes('public'))    r.public++;
    else                              r.none++;
  });
  return r;
}
function getTopHorses(h, t, l) {
  return [...h].sort((a,b) => t==='gp' ? b.gp-a.gp : b.maxConfo-a.maxConfo).slice(0,l)
    .map(x => ({ name: x.name, breed: x.breed, gp: x.gp, confoMax: x.maxConfo }));
}
function calculateUpcomingBirths(p) {
  const now = new Date();
  now.setHours(0,0,0,0);
  const cutoff = new Date(now); cutoff.setDate(now.getDate()-5);
  return p.filter(x => x.dueDate && x.dueDate >= cutoff).map(x => {
    const diffDays = Math.ceil((x.dueDate - now) / (1000*60*60*24));
    let statusText = x.phase;
    if      (diffDays < 0) statusText = "⚠ Overdue (" + Math.abs(diffDays) + " days)";
    else if (diffDays === 0) statusText = "🚨 DUE TODAY!";
    return { mareName: x.name, breed: x.breed, dueDate: x.dueDate.toLocaleDateString('en-GB'), daysUntil: diffDays, phase: statusText };
  }).sort((a,b) => a.daysUntil - b.daysUntil);
}
