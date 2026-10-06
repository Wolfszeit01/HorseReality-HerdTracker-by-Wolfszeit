/* ══════════════════════════════════════════════════════════════
   ZIELABGLEICH / BREEDING EVALUATION — Zucht- & Kaufbewertung
   ══════════════════════════════════════════════════════════════
   Konzept:
   - Zuchtziel = 1..n Merkmale, jedes mit Stufe niedrig(1)/mittel(2)/hoch(3)
   - Pro Merkmal wird der Z-Score des Pferds gegen die (gefilterte) Herde berechnet
   - Gesamt-Score = additive Summe (Gewicht × Z-Score) → UND-Logik,
     kein Merkmal wird bei Fehlen bestraft (kein Nullsetzen, keine Multiplikation)
   - Kandidaten (Kaufpferde) werden per Copy-Paste eingefügt (wie Import.html)
     und gegen dieselbe Herd-Baseline bewertet → Ranking-Position
   - Rasse/Geschlecht-Filter: Auto-Vorschlag aus Kandidat, manuell überschreibbar

   Bekannte Grenzen (bewusste Design-Entscheidung, siehe Chat-Historie):
   - Colour-Match ist nur beim Kandidaten prüfbar (Ja/Nein-Bonus), fließt
     NICHT in den Herd-weiten Z-Score-Vergleich ein.
   - COI wird intern als NEGATIVER Wert geführt (coi: -x), damit die
     Konvention "höherer Rohwert = besser" für alle Merkmale einheitlich gilt.
     Anzeige nutzt coiRaw (positiver, echter COI-Wert).

   Disziplin-GP / Best Discipline (WICHTIG — aus ImportModal.html übernommen):
   Die Formel lebt NICHT als Google-Sheets-Formel, sondern in ImportModal.html
   (computeTopDisciplines). Pro Disziplin: gpTotal = Summe der relevanten
   GP-Attribute, avgGp = gpTotal / Anzahl Attribute, avgConfo = Ø der
   relevanten Confo-Grade-Werte, Score = 0.25×avgConfo + 0.75×avgGp.
   "Best Discipline" = höchster gpTotal (gleiche Sortierung wie im Import-
   Modal). Dadurch jetzt für Herd-Pferde UND Kandidaten identisch berechenbar
   — kein Sheet-Zugriff nötig, keine Einschränkung mehr beim Kandidaten.

   Datenquellen (bewusst wiederverwendet, keine Neu-Implementierung):
   - GP, Conf Score   → parseTrackerData() / parseAllStatsData()
   - Comp Score       → readScoreSheet_() + compScoreMap_()
   - VGs              → statsMap[id].vgCount
   - Disziplin-GP     → eigene Berechnung aus rohen Stats-Sheet-Spalten,
                        Formel identisch zu ImportModal.html
   - Health           → HERD_COLS.HEALTH_START (Herd Tracker, 6 Spalten, Grade-Text)
   - COI              → PEDIGREE_COLS.COI (Pedigree-Sheet)
   ══════════════════════════════════════════════════════════════ */

function openZielabgleichModal() {
  const html = HtmlService.createTemplateFromFile('ZielabgleichModal')
    .evaluate()
    .setWidth(980)
    .setHeight(980);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// Gewichtungsstufen (fest: keine Slider, keine freien Prozentwerte)
const ZG_WEIGHTS = { niedrig: 1, mittel: 2, hoch: 3 };

// Anzeige-Metadaten für die dynamischen KPI-Kacheln — eine Kachel wird NUR
// gebaut, wenn das Merkmal in DIESEM Durchlauf aktiv gewichtet ist,
// unabhängig von der Streuung der Vergleichsgruppe. "disziplin" läuft als
// "Best Discipline"-Kachel (Name + Score% + GP), keine einfache Zahl.
const ZG_KPI_META = {
  gp:     { label: 'Genetic Potential', icon: 'dna',    decimals: 0, negate: false, suffix: '' },
  confo:  { label: 'Conf Score',        icon: 'horse',  decimals: 1, negate: false, suffix: '' },
  comp:   { label: 'Comp Score',        icon: 'trophy', decimals: 1, negate: false, suffix: '' },
  vg:     { label: 'VGs',               icon: 'shield', decimals: 0, negate: false, suffix: '' },
  health: { label: 'Health',            icon: 'shield', decimals: 1, negate: false, suffix: '' },
  coi:    { label: 'COI',               icon: 'shield', decimals: 1, negate: true,  suffix: '%' }
};

function zg_buildKpiList_(candidate, baseline, activeMerkmale) {
  return activeMerkmale
    .filter(b => ZG_KPI_META[b.key])
    .map(b => {
      const meta = ZG_KPI_META[b.key];
      const rawValue = candidate[b.key];
      const base = baseline[b.key];
      const value = Number.isFinite(rawValue) ? (meta.negate ? -rawValue : rawValue) : null;
      const herdAvg = base && base.n > 0 ? (meta.negate ? -base.mean : base.mean) : null;
      return { key: b.key, label: meta.label, icon: meta.icon, decimals: meta.decimals, suffix: meta.suffix, value: value, herdAvg: herdAvg, diff: value !== null && herdAvg !== null ? value - herdAvg : null };
    });
}

// Merkmal-Definitionen: key = eindeutiger Bezeichner, label = UI-Text, higherIsBetter dokumentiert Konvention
const ZG_MERKMALE = [
  { key: 'gp',        label: 'GP' },
  { key: 'confo',     label: 'Conf Score' },
  { key: 'comp',      label: 'Comp Score' },
  { key: 'disziplin', label: 'Disziplin-GP' },
  { key: 'vg',        label: 'VGs' },
  { key: 'health',    label: 'Health' },
  { key: 'coi',       label: 'COI (invertiert)' } // intern negiert, siehe Kopf-Kommentar
];

// Echte Grade-Werte aus ImportModal.html (GRADE_VALUES) — Quelle der Wahrheit
const ZG_GRADE_VALUES = { 'P': 19.5, 'BA': 49.5, 'A': 64.5, 'G-': 74.5, 'G': 76.5, 'G+': 82, 'VG': 92.5 };
const ZG_HEALTH_GRADE_VALUES = { 'Excellent': 100, 'Good': 87.5, 'Average': 75, 'Fair': 62.5, 'Poor': 50 };
const ZG_HEALTH_KEYS = ['Colic resistance', 'Hoof quality', 'Back problems', 'Respiratory disease', 'Resistance to lameness', 'Fertility'];

// Disziplin-Definitionen — 1:1 aus ImportModal.html (DISCIPLINES) übernommen
const ZG_DISCIPLINES = [
  { name: 'Dressage',        confo: ['Walk', 'Trot', 'Canter', 'Posture'],
    gp: ['Agility', 'Balance', 'Strength'] },
  { name: 'Driving',         confo: ['Trot', 'Back', 'Shoulders', 'Hindquarters'],
    gp: ['Agility', 'Pulling power', 'Speed', 'Stamina', 'Strength'] },
  { name: 'Endurance',       confo: ['Walk', 'Trot', 'Canter', 'Head', 'Neck', 'Back'],
    gp: ['Speed', 'Stamina', 'Strength', 'Surefootedness'] },
  { name: 'Eventing',        confo: ['Walk', 'Trot', 'Canter', 'Posture', 'Head', 'Neck'],
    gp: ['Balance', 'Bascule', 'Speed', 'Strength', 'Surefootedness'] },
  { name: 'Flat Racing',     confo: ['Gallop', 'Posture', 'Neck', 'Back', 'Shoulders', 'Frontlegs', 'Hindquarters'],
    gp: ['Speed', 'Acceleration', 'Stamina', 'Sprint'] },
  { name: 'Show Jumping',    confo: ['Canter', 'Back', 'Shoulders', 'Frontlegs', 'Hindquarters'],
    gp: ['Acceleration', 'Agility', 'Bascule', 'Sprint', 'Strength'] },
  { name: 'Western Reining', confo: ['Head', 'Neck', 'Shoulders', 'Frontlegs', 'Hindquarters'],
    gp: ['Acceleration', 'Agility', 'Balance', 'Surefootedness'] }
];

// Reihenfolge der 10 GP-Attribute im Stats-Sheet (STATS_GP_COL, Import.gs)
const ZG_GP_ATTR_ORDER = ['Acceleration', 'Agility', 'Balance', 'Bascule', 'Pulling power', 'Speed', 'Sprint', 'Stamina', 'Strength', 'Surefootedness'];

// Confo-Spaltenreihenfolge je Rassevariante, wie in Import.gs (_writeStats) geschrieben
const ZG_CONFO_LAYOUT = {
  default: ['Walk', 'Trot', 'Canter', 'Gallop', 'Posture', 'Head', 'Neck', 'Back', 'Shoulders', 'Frontlegs', 'Hindquarters', 'Socks'],
  ICE:     ['Walk', 'Trot', 'Canter', 'Gallop', 'Tolt', 'Pace', 'Posture', 'Head', 'Neck', 'Back', 'Shoulders', 'Frontlegs', 'Hindquarters', 'Socks'],
  KATH:    ['Walk', 'Trot', 'Canter', 'Gallop', 'Revaal', 'Posture', 'Head', 'Neck', 'Back', 'Shoulders', 'Frontlegs', 'Hindquarters', 'Socks']
};

/* ──────────────────────────────────────────────────────────────
   1. HERDE LADEN — kombiniert bestehende Loader, keine Duplikat-Logik
   ────────────────────────────────────────────────────────────── */

function zg_loadHerdMetrics_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const trackerData  = readSheet_(ss, SHEETS.TRACKER);
  const standardData = readSheet_(ss, SHEETS.STATS);
  const iceData       = readSheet_(ss, SHEETS.STATS_ICE);
  const kathData      = readSheet_(ss, SHEETS.STATS_KATH);
  const pedigreeData  = readSheet_(ss, SHEETS.PEDIGREE);
  const compResults   = readScoreSheet_(ss, 'Comp. Results');

  const statsMap    = parseAllStatsData(standardData, iceData, kathData);
  const pedigreeMap = parsePedigreeData(pedigreeData);
  const horses      = parseTrackerData(trackerData, statsMap, pedigreeMap);
  const compMap     = compScoreMap_(compResults);

  // Rohe GP-Attribute + Confo-Grades je Pferd (für Disziplin-Berechnung, s.u.)
  const rawStatsMap = zg_buildRawStatsMap_(standardData, iceData, kathData);

  // Health + COI direkt aus den Rohdaten (nicht Teil von parseTrackerData)
  const healthMap = zg_buildHealthMap_(trackerData);
  const coiMap = zg_buildCoiMap_(pedigreeData);
  const lineMap = zg_buildLineMap_(pedigreeData);

  return horses.map(h => {
    const raw = rawStatsMap[h.id];
    const best = raw ? zg_pickBestDiscipline_(zg_computeDisciplines_(raw.gpAttrs, raw.confo)) : null;
    const coiRaw = coiMap[h.id];
    const lines = lineMap[h.id] || {};
    return {
      id: h.id,
      name: h.name,
      breed: h.breed,
      gender: h.gender,
      gp: h.gp || 0,
      confo: h.maxConfo || null,
      comp: compMap[h.id] || 0,
      vg: (h.stats && h.stats.vgCount) || 0,
      disziplin: best ? best.score : null,
      disziplinKey: best ? best.name : null, // ist bereits der Anzeigename (z.B. "Dressage")
      disziplinGpTotal: best ? best.gpTotal : null,
      health: healthMap[h.id] || 0,
      coiRaw: coiRaw != null ? coiRaw : null,
      coi: coiRaw != null ? -coiRaw : null, // negiert für einheitliche "höher=besser"-Konvention
      stallionline: lines.stallionline || '',
      damline: lines.damline || ''
    };
  });
}

// Baut je Pferd { gpAttrs: {Acceleration:.., ...}, confo: {Walk:.., ...} } aus den drei Stats-Sheets
function zg_buildRawStatsMap_(standardData, iceData, kathData) {
  const map = {};
  const sheets = [
    { data: standardData, variant: 'default' },
    { data: iceData,       variant: 'ICE' },
    { data: kathData,      variant: 'KATH' }
  ];

  sheets.forEach(({ data, variant }) => {
    const confoLayout = ZG_CONFO_LAYOUT[variant];
    const gpStartCol = STATS_GP_COL[variant];

    data.forEach(row => {
      const id = row[STATS_COLS.ID - 1];
      if (!id) return;

      const confo = {};
      confoLayout.forEach((trait, i) => { confo[trait] = row[STATS_COLS.CONFO_START - 1 + i]; });

      const gpAttrs = {};
      ZG_GP_ATTR_ORDER.forEach((attr, i) => { gpAttrs[attr] = parseFloat(row[gpStartCol - 1 + i]) || 0; });

      map[id] = { gpAttrs: gpAttrs, confo: confo };
    });
  });

  return map;
}

// Disziplin-Score je Disziplin — identische Formel zu ImportModal.html (computeTopDisciplines)
function zg_computeDisciplines_(gpAttrs, confo) {
  gpAttrs = gpAttrs || {};
  confo = confo || {};

  return ZG_DISCIPLINES.map(d => {
    const gpTotal = d.gp.reduce((sum, key) => sum + (parseFloat(gpAttrs[key]) || 0), 0);
    const avgGp = gpTotal / d.gp.length;

    const confoValues = d.confo
      .map(key => ZG_GRADE_VALUES[String(confo[key] || '').trim().toUpperCase()])
      .filter(v => v !== undefined);
    const avgConfo = confoValues.length ? confoValues.reduce((a, b) => a + b, 0) / confoValues.length : null;

    const score = avgConfo !== null ? (0.25 * avgConfo + 0.75 * avgGp) : null;

    return { name: d.name, gpTotal: gpTotal, avgGp: avgGp, avgConfo: avgConfo, score: score };
  });
}

// Beste Disziplin = höchster gpTotal (gleiche Sortierung wie "Top Disciplines" im Import-Modal)
function zg_pickBestDiscipline_(disciplines) {
  const valid = disciplines.filter(d => d.gpTotal > 0);
  if (!valid.length) return null;
  return valid.reduce((a, b) => (b.gpTotal > a.gpTotal ? b : a));
}

// Health-Score je Pferd (Ø der 6 Grade-Werte), direkt aus Herd-Tracker-Rohdaten
function zg_buildHealthMap_(trackerData) {
  const map = {};
  trackerData.forEach(row => {
    const id = row[HERD_COLS.ID - 1];
    if (!id) return;
    const grades = [];
    for (let i = 0; i < ZG_HEALTH_KEYS.length; i++) {
      const g = row[HERD_COLS.HEALTH_START - 1 + i];
      if (ZG_HEALTH_GRADE_VALUES[g] !== undefined) grades.push(ZG_HEALTH_GRADE_VALUES[g]);
    }
    map[id] = grades.length ? grades.reduce((a, b) => a + b, 0) / grades.length : 0;
  });
  return map;
}

// COI je Pferd direkt aus Pedigree-Rohdaten
function zg_buildCoiMap_(pedigreeData) {
  const map = {};
  pedigreeData.forEach(row => {
    const id = row[PEDIGREE_COLS.ID - 1];
    if (!id) return;
    const raw = row[PEDIGREE_COLS.COI - 1];
    const num = parseFloat(String(raw).replace('%', '').replace(',', '.'));
    if (!isNaN(num)) map[id] = num;
  });
  return map;
}

// Hengst-/Stutenlinie je Pferd direkt aus Pedigree-Rohdaten (nach ID). Für Kandidaten
// siehe zg_resolveCandidateLines_ (Auflösung über Vater-/Mutter-NAMEN).
function zg_buildLineMap_(pedigreeData) {
  const byId = {};
  pedigreeData.forEach(row => {
    const id = row[PEDIGREE_COLS.ID - 1];
    if (!id) return;
    byId[id] = {
      stallionline: row[PEDIGREE_COLS.STALLIONLINE - 1] || '',
      damline: row[PEDIGREE_COLS.DAMLINE - 1] || ''
    };
  });
  return byId;
}

// Wie copyParentPedigree() in Import.gs: Kandidat erbt Stallionline vom Vater,
// Damline von der Mutter — nur auflösbar, wenn Vater/Mutter bereits in der Herde sind.
function zg_resolveCandidateLines_(sireName, damName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pedigreeData = readSheet_(ss, SHEETS.PEDIGREE);
  const byName = {};
  pedigreeData.forEach(row => {
    const name = row[PEDIGREE_COLS.NAME - 1];
    if (!name) return;
    byName[name] = {
      stallionline: row[PEDIGREE_COLS.STALLIONLINE - 1] || '',
      damline: row[PEDIGREE_COLS.DAMLINE - 1] || ''
    };
  });

  const sire = sireName && sireName !== 'Unknown' ? byName[sireName] : null;
  const dam = damName && damName !== 'Unknown' ? byName[damName] : null;

  return {
    stallionline: sire ? sire.stallionline : '', // leer = unbekannt (Vater nicht in der Herde)
    damline: dam ? dam.damline : ''
  };
}

// Eindeutige Rassen/Geschlechter/Linien der Herde — für die Filter-Dropdowns im Modal
function getZielabgleichFilterOptions() {
  const herdMetrics = zg_loadHerdMetrics_();
  return {
    breeds: [...new Set(herdMetrics.map(h => h.breed).filter(Boolean))].sort(),
    genders: [...new Set(herdMetrics.map(h => h.gender).filter(Boolean))].sort(),
    stallionlines: [...new Set(herdMetrics.map(h => h.stallionline).filter(Boolean))].sort(),
    damlines: [...new Set(herdMetrics.map(h => h.damline).filter(Boolean))].sort()
  };
}

// Filtert die Herde auf Rasse/Geschlecht/Linien (leer/'Alle' = kein Filter auf der Dimension)
function zg_applyFilter_(herdMetrics, filter) {
  if (!filter) return herdMetrics;
  return herdMetrics.filter(h =>
    (!filter.breed || filter.breed === 'Alle' || h.breed === filter.breed) &&
    (!filter.gender || filter.gender === 'Alle' || h.gender === filter.gender) &&
    (!filter.stallionline || filter.stallionline === 'Alle' || h.stallionline === filter.stallionline) &&
    (!filter.damline || filter.damline === 'Alle' || h.damline === filter.damline)
  );
}

/* ──────────────────────────────────────────────────────────────
   2. BASELINE — Mittelwert + Std-Abweichung pro Merkmal
   ────────────────────────────────────────────────────────────── */

function zg_meanStdDev_(values) {
  const valid = values.filter(v => v !== null && v !== undefined && v !== 0);
  if (!valid.length) return { mean: 0, stdDev: 0, n: 0 };

  const mean = valid.reduce((a, b) => a + b, 0) / valid.length;
  const variance = valid.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / valid.length;
  return { mean: mean, stdDev: Math.sqrt(variance), n: valid.length };
}

function zg_buildBaseline_(herdMetrics) {
  const baseline = {};
  ZG_MERKMALE.forEach(m => {
    baseline[m.key] = zg_meanStdDev_(herdMetrics.map(h => h[m.key]));
  });
  return baseline;
}

/* ──────────────────────────────────────────────────────────────
   3. Z-SCORE + GEWICHTETER GESAMT-SCORE (additiv, UND-Logik)
   ────────────────────────────────────────────────────────────── */

// activeMerkmale: [{ key: 'gp', stufe: 'hoch' }, ...]
function zg_calculateScore_(horseMetrics, baseline, activeMerkmale) {
  let score = 0;
  let weightSum = 0;
  const breakdown = [];

  activeMerkmale.forEach(am => {
    const b = baseline[am.key];
    const weight = ZG_WEIGHTS[am.stufe] || 0;
    const value = horseMetrics[am.key];
    if (!b || b.stdDev === 0 || value === null || value === undefined) return; // kein Merkmal bestrafen bei fehlender Streuung/Daten

    const rawZScore = (value - b.mean) / b.stdDev;

    // Extreme Ausreißer begrenzen — ein einzelnes Merkmal darf den Gesamt-Score
    // nicht durch eine unrealistisch hohe Standardabweichung dominieren.
    const zScore = Math.max(-3, Math.min(3, rawZScore));

    const weighted = weight * zScore;
    score += weighted;
    weightSum += weight;
    breakdown.push({ key: am.key, value: value, mean: b.mean, zScore: zScore, weighted: weighted });
  });

  return { score: score, breakdown: breakdown, weightSum: weightSum, avgWeightedZ: weightSum ? score / weightSum : 0 };
}

/* ──────────────────────────────────────────────────────────────
   3b. KANDIDAT AUS COPY-PASTE-TEXT (Kaufpferd, noch nicht im Sheet)
   ────────────────────────────────────────────────────────────── */
   // Wiederverwendung von getPreviewData() aus Import.gs statt eigener Regex-Logik.
   //
   // Einschränkungen ggü. Herd-Pferden (Absprache mit Du):
   // - GP total   = Summe der 10 Basis-Attribute, identisch zu FoalCalculator.gs
   // - Conf Score = höchster Wert aus den Show-Results-Zahlen (Stats-Tab)
   // - Comp Score = Maximum der geparsten Wettkampf-Ergebnisse
   // - VGs        = Anzahl Merkmale mit Grade "VG"
   // - Health     = Ø der 6 Grade-Werte (Excellent=100 ... Poor=50)
   // - COI        = aus Pedigree-Text geparst, intern negiert
   // - Disziplin-GP / Best Discipline = identische Formel wie Import-Modal (computeTopDisciplines)

function zg_parseCandidateFromPaste_(inputs) {
  const d = getPreviewData(inputs); // Import.gs

  const gpValues = Object.values(d.gp).map(v => parseFloat(v) || 0);
  const gpTotal = gpValues.reduce((a, b) => a + b, 0);

  // "Conf Score" = höchster Wert aus den SHOW-RESULTS-Zahlen (Stats-Tab, landet
  // im Sheet "Conf. Results") — NICHT die Confo-Attribut-Noten (Walk/Trot/... als
  // VG/G+/G aus dem Info-Tab). Die Noten werden weiterhin separat für VG-Count
  // und die Disziplin-Berechnung gebraucht (siehe zg_computeDisciplines_ unten).
  const showScoreNums = (d.showScores || []).map(s => parseFloat(String(s).replace(',', '.')) || 0);
  const confoScore = showScoreNums.length ? Math.max(...showScoreNums) : null;

  const confoGrades = Object.values(d.confo).filter(g => ZG_GRADE_VALUES[g] !== undefined);
  const vgCount = confoGrades.filter(g => g === 'VG').length;

  const compScoreNums = (d.compScores || []).map(s => parseFloat(String(s).replace(',', '.')) || 0);
  const compMax = compScoreNums.length ? Math.max(...compScoreNums) : null;

  const healthGrades = ZG_HEALTH_KEYS
    .map(k => d.health && d.health[k])
    .filter(g => ZG_HEALTH_GRADE_VALUES[g] !== undefined)
    .map(g => ZG_HEALTH_GRADE_VALUES[g]);
  const healthScore = healthGrades.length ? healthGrades.reduce((a, b) => a + b, 0) / healthGrades.length : null;

  const coiRaw = d.coi ? parseFloat(String(d.coi).replace('%', '').replace(',', '.')) : null;

  const bestDiscipline = zg_pickBestDiscipline_(zg_computeDisciplines_(d.gp, d.confo));
  const lines = zg_resolveCandidateLines_(d.sire, d.dam);

  return {
    name: d.name || 'Kandidat',
    breed: d.breed || '',
    gender: d.gender || '',
    gp: gpTotal,
    confo: confoScore,
    comp: compMax,
    vg: vgCount,
    health: healthScore,
    coiRaw: isNaN(coiRaw) ? null : coiRaw,
    coi: (coiRaw != null && !isNaN(coiRaw)) ? -coiRaw : null,
    disziplin: bestDiscipline ? bestDiscipline.score : null,
    disziplinKey: bestDiscipline ? bestDiscipline.name : null,
    disziplinGpTotal: bestDiscipline ? bestDiscipline.gpTotal : null,
    stallionline: lines.stallionline, // leer = Vater nicht in der Herde, Linie unbekannt
    damline: lines.damline,
    genetics: d.genetics || {}
  };
}

/* ──────────────────────────────────────────────────────────────
   3c. COLOUR-MATCH — Ja/Nein-Bonus, nur beim Kandidaten
   ────────────────────────────────────────────────────────────── */
   // colourTarget: { gene, value } — z.B. { gene: 'Extension', value: 'ee' }
   // Vergleich ist case-insensitive, exakter Genotyp-Match.

function zg_checkColourMatch_(candidate, colourTarget) {
  if (!colourTarget || !colourTarget.gene || !colourTarget.value) return null;
  const actual = candidate.genetics && candidate.genetics[colourTarget.gene];
  if (!actual) return { matched: false, gene: colourTarget.gene, target: colourTarget.value, actual: null };
  const matched = String(actual).toLowerCase() === String(colourTarget.value).toLowerCase();
  return { matched: matched, gene: colourTarget.gene, target: colourTarget.value, actual: actual };
}

/* ──────────────────────────────────────────────────────────────
   3d. BREEDING GOAL MATCH % — eigene, dokumentierte Definition
   ────────────────────────────────────────────────────────────── */
   // Es gibt keine spielseitige Formel dafür — eigene, transparente Näherung:
   // Score wird linear von [-3*weightSum, +3*weightSum] (theoretisches Minimum/
   // Maximum bei komplett gecappten Z-Scores) auf [0, 100] abgebildet.
   // Colour-Match gibt +8 Punkte Bonus (gedeckelt auf 100).

function zg_calculateMatchPercent_(score, weightSum, colourMatch) {
  if (!weightSum) return 0;

  const minScore = -3 * weightSum;
  const maxScore =  3 * weightSum;

  let pct = ((score - minScore) / (maxScore - minScore)) * 100;

  if (colourMatch && colourMatch.matched) pct += 8;

  return Math.max(0, Math.min(100, Math.round(pct)));
}

/* ──────────────────────────────────────────────────────────────
   3e. EMPFEHLUNGS-BADGE — Schwellenwerte auf Basis Perzentil
   ────────────────────────────────────────────────────────────── */
   // percentile = Anteil der Herde, den der Kandidat schlägt (100 = bester)
   // Top 15% → Strong; unterste 15% → Not recommended; sonst neutral.

function zg_getRecommendation_(percentile, insufficientData) {
  if (insufficientData) return { level: 'insufficient', label: 'INSUFFICIENT DATA · REVIEW INPUT' };
  if (percentile >= 85) return { level: 'strong', label: 'STRONG RECOMMENDATION · KEEP / BUY' };
  if (percentile <= 15) return { level: 'pass', label: 'NOT RECOMMENDED · PASS / SELL' };
  return { level: 'neutral', label: 'NEUTRAL / HERD AVERAGE' };
}

/* ──────────────────────────────────────────────────────────────
   3f. BREEDING IMPACT PILLS — Stärken/Schwächen ggü. Herdendurchschnitt
   ────────────────────────────────────────────────────────────── */

// Zerlegt den Breeding-Goal-Match in "voll erfüllt" (zScore >= 0.5) /
// "teilweise erfüllt" (-0.5..0.5) / "verfehlt" (< -0.5) — fürs Tooltip.
function zg_buildMatchComposition_(breakdown) {
  let full = 0, partial = 0, missed = 0;
  breakdown.forEach(b => {
    if (b.zScore >= 0.5) full++;
    else if (b.zScore >= -0.5) partial++;
    else missed++;
  });
  return { full: full, partial: partial, missed: missed, total: breakdown.length };
}

// Impact-Pills zeigen die ROHE Abweichung vom Herdenmittel (value - mean), NICHT
// den gewichteten Z-Beitrag — das ist bewusst: die Pille soll erklären "wie weit
// weicht das Pferd hier vom Durchschnitt ab", nicht "wie stark zieht das am Score".
// zScore (inkl. Cap) entscheidet aber weiterhin, ob ein Merkmal überhaupt als
// Pille auftaucht (>0.4 / <-0.4) — Sortierung/Auswahl ist also schon Z-Score-basiert.
function zg_buildImpactPills_(candidateBreakdown) {
  const labelMap = Object.fromEntries(ZG_MERKMALE.map(m => [m.key, m.label]));
  const up = [];
  const down = [];

  candidateBreakdown.forEach(b => {
    const label = labelMap[b.key] || b.key;
    const diff = b.value - b.mean;
    // COI ist intern negiert — für die Anzeige zurückdrehen, damit "Diff" den echten COI-Unterschied zeigt
    const displayDiff = b.key === 'coi' ? -diff : diff;
    const displayValue = b.key === 'coi' ? -b.value : b.value;
    const displayMean = b.key === 'coi' ? -b.mean : b.mean;
    const pill = { label: label, diff: displayDiff, candidateValue: displayValue, herdAvg: displayMean, zScore: b.zScore, weighted: b.weighted };
    if (b.zScore > 0.4) up.push(pill);
    else if (b.zScore < -0.4) down.push(pill);
  });

  return { up: up, down: down };
}

/* ──────────────────────────────────────────────────────────────
   3g. SELL/REPLACEMENT-VORSCHLAG — volle Berechnung mit Bestandsschnitt-Delta
   ────────────────────────────────────────────────────────────── */
   // Vereinfachung: Baseline (Mittelwert/Std-Abw.) bleibt beim Tausch unverändert
   // (kein rekursives Neu-Berechnen). Delta = (Kandidat-Score − Score des
   // schwächsten Pferds) / Herdengröße — die Verschiebung des Gruppendurchschnitts,
   // wenn das schwächste Pferd durch den Kandidaten ersetzt wird.
   // Zusätzlich: "justAdd" — was passiert, wenn der Kandidat einfach zur Herde
   // dazukommt, OHNE dass jemand abgegeben wird (Herdengröße n → n+1).

function zg_buildSellSuggestion_(ranked, candidateScore) {
  if (!ranked.length) return null;
  const weakest = ranked[ranked.length - 1];
  const n = ranked.length;
  const currentAvg = ranked.reduce((a, r) => a + r.score, 0) / n;
  const newAvg = currentAvg + (candidateScore - weakest.score) / n;

  const justAddAvg = (currentAvg * n + candidateScore) / (n + 1);

  return {
    name: weakest.name,
    rank: n,
    score: weakest.score,
    currentAvg: currentAvg,
    newAvg: newAvg,
    avgDelta: newAvg - currentAvg,
    justAdd: {
      newAvg: justAddAvg,
      avgDelta: justAddAvg - currentAvg,
      newHerdSize: n + 1
    }
  };
}

/* ──────────────────────────────────────────────────────────────
   3h. LINIEN-BOTTLENECK — Warnung bei Häufung von Hengst-/Stutenlinie
   ────────────────────────────────────────────────────────────── */
   // Eigener Schwellenwert (nicht spielseitig vorgegeben): ab 20% Anteil der
   // Vergleichsgruppe mit derselben Linie wie der Kandidat wird gewarnt —
   // Hinweis auf steigendes künftiges COI-Risiko bei Zucht innerhalb der Linie.

function zg_buildLineWarnings_(candidate, herdMetrics) {
  const warnings = [];
  if (!herdMetrics.length) return warnings;

  [
    { key: 'stallionline', label: 'Hengstlinie' },
    { key: 'damline', label: 'Stutenlinie' }
  ].forEach(({ key, label }) => {
    const value = candidate[key];
    if (!value) return; // unbekannt (Vater/Mutter nicht in der Herde) — keine Warnung möglich
    const count = herdMetrics.filter(h => h[key] === value).length;
    const pct = Math.round((count / herdMetrics.length) * 100);
    if (pct >= 20) {
      warnings.push({ label: label, value: value, count: count, total: herdMetrics.length, pct: pct });
    }
  });

  return warnings;
}

/* ──────────────────────────────────────────────────────────────
   4. RANKING GEGEN HERDE (Haupteinstieg fürs Modal)
   ────────────────────────────────────────────────────────────── */

// activeMerkmale: [{ key, stufe }]
// candidateInputs optional: { info, stats, colour } — wie Import.html-Paste-Felder
// filter optional: { breed, gender, stallionline, damline } — 'Alle'/leer = kein Filter auf der Dimension
// colourTarget optional: { gene, value } — Ziel-Genotyp für Colour-Match-Bonus
// Reuse the import parser so preview and evaluation recognize the same values.
function getZielabgleichCandidateFilters(inputs) {
  const parsed = getPreviewData(inputs || {});
  return { breed: parsed.breed || '', gender: parsed.gender || '' };
}

function getZielabgleichRanking(activeMerkmale, candidateInputs, filter, colourTarget) {
  const herdMetricsAll = zg_loadHerdMetrics_();
  const candidate = candidateInputs ? zg_parseCandidateFromPaste_(candidateInputs) : null;

  const hasManualFilter = filter && (
    (filter.breed && filter.breed !== 'Alle') ||
    (filter.gender && filter.gender !== 'Alle') ||
    (filter.stallionline && filter.stallionline !== 'Alle') ||
    (filter.damline && filter.damline !== 'Alle')
  );
  // Linien werden bewusst NICHT automatisch aus dem Kandidaten vorgeschlagen —
  // sonst würde die Linien-Bottleneck-Warnung unten die eigene Linie ausblenden.
  const appliedFilter = (hasManualFilter || (filter && filter.explicitSelection))
    ? filter
    : (candidate ? { breed: candidate.breed, gender: candidate.gender, stallionline: 'Alle', damline: 'Alle' } : null);

  const herdMetrics = zg_applyFilter_(herdMetricsAll, appliedFilter);
  const baseline = zg_buildBaseline_(herdMetrics);

  const ranked = herdMetrics.map(h => {
    const result = zg_calculateScore_(h, baseline, activeMerkmale);
    return {
      id: h.id, name: h.name, breed: h.breed, gp: h.gp, confo: h.confo,
      disziplin: h.disziplin, disziplinKey: h.disziplinKey, disziplinGpTotal: h.disziplinGpTotal,
      score: result.score, breakdown: result.breakdown
    };
  }).sort((a, b) => b.score - a.score);

  let candidateResult = null;
  let summary = null;

  if (candidate) {
    const result = zg_calculateScore_(candidate, baseline, activeMerkmale);
    // 1-basierter Rang in der KOMBINIERTEN Liste (Herde + Kandidat), nicht nur
    // gegen die Herde allein — sonst weichen Hero-Card und Tabelle voneinander ab.
    const better = ranked.filter(r => r.score < result.score).length; // Herd-Pferde, die der Kandidat schlägt
    const totalHorses = ranked.length + 1; // Herde + Kandidat selbst
    const rank = (ranked.length - better) + 1;
    const percentile = ranked.length ? Math.round((better / ranked.length) * 100) : 0;

    // Insufficient-Data-Check: wie viele der AKTIVEN Merkmale konnten beim
    // Kandidaten gar nicht berechnet werden (breakdown enthält nur vorhandene)?
    const presentKeys = result.breakdown.map(b => b.key);
    const missingKeys = activeMerkmale.map(am => am.key).filter(k => !presentKeys.includes(k));
    const missingRatio = activeMerkmale.length ? missingKeys.length / activeMerkmale.length : 0;
    const insufficientData = missingRatio > 0.5; // mehr als die Hälfte der aktiven Merkmale fehlen
    const missingLabels = missingKeys.map(k => (ZG_MERKMALE.find(m => m.key === k) || {}).label || k);

    candidateResult = {
      name: candidate.name || 'Kandidat',
      score: result.score,
      breakdown: result.breakdown,
      rank: rank,
      totalHorses: totalHorses,
      percentile: percentile,
      insufficientData: insufficientData,
      missingFields: missingLabels,
      gp: candidate.gp,
      confo: candidate.confo,
      disziplinKey: candidate.disziplinKey,
      disziplin: candidate.disziplin,
      disziplinGpTotal: candidate.disziplinGpTotal
    };

    const colourMatch = zg_checkColourMatch_(candidate, colourTarget);
    const matchPercent = zg_calculateMatchPercent_(
      result.score,
      result.weightSum,
      colourMatch
    );
    const matchComposition = zg_buildMatchComposition_(result.breakdown);
    const recommendation = zg_getRecommendation_(percentile, insufficientData);
    const impactPills = zg_buildImpactPills_(result.breakdown);
    const sellSuggestion = zg_buildSellSuggestion_(ranked, result.score);
    const lineWarnings = zg_buildLineWarnings_(candidate, herdMetrics);
    const herdAvgScore = ranked.length ? ranked.reduce((a, r) => a + r.score, 0) / ranked.length : 0;

    const disziplinEntry = result.breakdown.find(b => b.key === 'disziplin');
    const kpiList = zg_buildKpiList_(candidate, baseline, activeMerkmale);

    summary = {
      matchPercent: matchPercent,
      matchComposition: matchComposition,
      recommendation: recommendation,
      colourMatch: colourMatch,
      impactPills: impactPills,
      sellSuggestion: sellSuggestion,
      lineWarnings: lineWarnings,
      insufficientData: insufficientData,
      missingFields: missingLabels,
      herdAvgScore: herdAvgScore,
      kpiList: kpiList,
      bestDiscipline: (disziplinEntry && candidate.disziplinKey) ? { name: candidate.disziplinKey, gpTotal: candidate.disziplinGpTotal, score: candidate.disziplin } : null
    };
  }

  return {
    ranking: ranked.map((r, i) => ({
      ...r,
      position: i + 1,
      topDisciplineName: r.disziplinKey || null
    })),
    candidate: candidateResult,
    summary: summary,
    baseline: baseline,
    appliedFilter: appliedFilter || { breed: 'Alle', gender: 'Alle', stallionline: 'Alle', damline: 'Alle' },
    herdSizeFiltered: herdMetrics.length,
    herdSizeTotal: herdMetricsAll.length
  };
}
