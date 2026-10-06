/* ══════════════════════════════════════════════════════════════
   Universial Import
   ══════════════════════════════════════════════════════════════ */

function openUniversalImportModal() {
  const tpl = HtmlService.createTemplateFromFile('ImportModal');
  tpl.horseId = '';
  tpl.horseName = '';
  tpl.sireName = '';
  tpl.damName = '';
  const html = tpl.evaluate().setWidth(1000).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/* ══════════════════════════════════════════════════════════════
   GENETIK-PARSING-TABELLE
   Eine Zeile pro Gen: "key" ist der Name im genetics-Objekt,
   "regex" ist das Suchmuster im eingefügten Text.

   Wie es funktioniert: Die meisten Regexes liefern 2 Gruppen
   (Allel 1, Allel 2), die zu "X/Y" zusammengesetzt werden.
   Bei wholeMatch:true steckt "X/Y" schon im Gesamt-Treffer
   (nur bei Agouti der Fall, weil dort beide Allele im selben
   Muster stecken statt in eigenen Klammern).
   ══════════════════════════════════════════════════════════════ */
const GENE_PATTERNS = [
  { key: 'Extension',     regex: /Extension[\s\S]{0,80}?([Ee])\s*[\/\\]\s*([Ee])/i },
  { key: 'Agouti',        regex: /(?:[Aa]\+|[Aa]t|[Aa]|aa)\s*[\/\\]\s*(?:[Aa]\+|[Aa]t|[Aa]|aa)/, wholeMatch: true },
  { key: 'Grey',          regex: /Grey[\s\S]{0,80}?([Gg])\s*[\/\\]\s*([Gg]|n)/i },
  { key: 'Silver',        regex: /Silver[\s\S]{0,80}?([Zz])\s*[\/\\]\s*([Zz])/i },
  { key: 'Creampearl',    regex: /Cream[\s\S]{0,80}?(Cr|prl)\s*[\/\\]\s*(Cr|prl|n)/i },
  { key: 'Dun',           regex: /Dun[\s\S]{0,80}?(nd[12]|[Dd])\s*[\/\\]\s*(nd[12]|[Dd])/i },
  { key: 'Champagne',     regex: /Champagne[\s\S]{0,80}?(Ch)\s*[\/\\]\s*(Ch|n)/i },
  { key: 'Mushroom',      regex: /Mushroom[\s\S]{0,80}?(mu|[Mm])\s*[\/\\]\s*(mu|[Mm]|n)/i },
  { key: 'Frame',         regex: /(?:\bFrame\b|\bOWL\b)[\s\S]{0,80}?(OWL|[Oo])\s*[\/\\]\s*(OWL|[Oo]|n)/i },
  { key: 'Appaloosa',     regex: /(?:Appaloosa|\bLp\b)[\s\S]{0,80}?(Lp)\s*[\/\\]\s*(Lp|n)/i },
  { key: 'PATN1',         regex: /PATN1[\s\S]{0,80}?(PATN1)\s*[\/\\]\s*(PATN1|n)/i },
  { key: 'MITF',          regex: /(?:SW[13]|MITF)[\s\S]{0,80}?(SW[13])\s*[\/\\]\s*(SW[13]|n)/i },
  { key: 'SW2',           regex: /SW2[\s\S]{0,80}?(SW2)\s*[\/\\]\s*(SW2|n)/i },
  { key: 'WhiteSpotting', regex: /(?:White\s*Spotting|\bWS\b)[\s\S]{0,80}?(WS)\s*[\/\\]\s*(WS|n)/i }
];

const KIT_ALLELES = ['TO', 'Rn', 'SB1', 'W3', 'W8', 'W10', 'W16', 'W19', 'W20', 'W21'];

function getPreviewData(inputs) {
  const d = {
    id: "", name: "", breed: "", dob: "", gender: "",
    agedWithDP: "0", predicates: "", sire: "Unknown", dam: "Unknown",
    sireExists: false, damExists: false,
    gs_p: "", gd_p: "", gs_m: "", gd_m: "",
    gp: {}, confo: {}, genetics: {}, health: {}, achieve: {},
    showScores: [], compScores: [], pregnancy: null, allHorses: []
  };

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const info = inputs.info || "";

  const idMatch = info.match(/Life\s*number\s*[\r\n]+#?(\d+)/i) || info.match(/Lifenumber\s*[:.]?\s*#?(\d+)/i);
  d.id = idMatch ? idMatch[1] : "";

  let rawName = "";
  const passportMatch = info.match(/Passport\s*[\r\n]+([^\r\n]+)/i);
  if (passportMatch) rawName = passportMatch[1];
  else rawName = (info.match(/Name\s*([\s\S]+?)(?=\n|Tagline|RC\d)/i) || [])[1] || "";
  d.name = rawName.replace(/Tagline.*/i, "").trim();

  const newDobMatch = info.match(/Date\s*of\s*Birth\s*[\r\n]+(\d{1,2}\s+\w+\s+\d{4})/i);
  if (newDobMatch) {
    const monthMap = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};
    const parts = newDobMatch[1].split(/\s+/);
    d.dob = (monthMap[parts[1].toLowerCase().slice(0,3)]||'01') + '/' + parts[0].padStart(2,'0') + '/' + parts[2];
  }

  if (!d.breed || d.breed.trim() === "") {
    const breedRegMatch = info.match(/Breed\s*registry\s*[\r\n]+([^\r\n]+)/i);
    if (breedRegMatch) {
      const br = breedRegMatch[1].trim();
      const breedMap = {
        'akhal':       'Akhal-Teke',
        'arabian':     'Arabian',
        'brabant':     'Brabant',
        'brumby':      'Brumby',
        'camargue':    'Camargue',
        'cleveland':   'Cleveland Bay',
        'exmoor':      'Exmoor Pony',
        'finnhorse':   'Finnhorse',
        'fjord':       'Fjord Horse',
        'friesian':    'Friesian',
        'haflinger':   'Haflinger',
        'icelandic':   'Icelandic Horse',
        'irish cob':   'Irish Cob',
        'kathiawari':  'Kathiawari',
        'kladruber':   'Kladruber',
        'knabstrupper':'Knabstrupper',
        'lipizzaner':  'Lipizzaner',
        'lusitano':    'Lusitano',
        'mongolian':   'Mongolian Horse',
        'mustang':     'Mustang',
        'namib':       'Namib Desert Horse',
        'noriker':     'Noriker',
        'norman cob':  'Norman Cob',
        'oldenburg':   'Oldenburg',
        'pantaneiro':  'Pantaneiro',
        'pre ':        'PRE',
        'quarter':     'Quarter Horse',
        'shetland':    'Shetland Pony',
        'shire':       'Shire Horse',
        'suffolk':     'Suffolk Punch',
        'thoroughbred':'Thoroughbred',
        'trakehner':   'Trakehner',
        'welsh':       'Welsh Pony'
      };
      const brLow = br.toLowerCase();
      let matched = false;
      for (const [key, name] of Object.entries(breedMap)) {
        if (brLow.includes(key)) { d.breed = name; matched = true; break; }
      }
      if (!matched) d.breed = br.replace(/\s*Horse\s*Society\s*/i, '').replace(/\s*Society\s*/i, '').trim();
    }
  }

  const genderMatch = info.match(/\b(Mare|Stallion|Gelding)\b/i) || (inputs.stats||"").match(/\b(Mare|Stallion|Gelding)\b/i);
  d.gender = genderMatch ? genderMatch[1] : "";

  const predMatch = info.match(/Predicates\s*[\r\n]+([^\r\n]+)/i);
  if (predMatch) {
    const rawPred = predMatch[1].trim();
    let found = "None";
    for (const pred of ['Star','Clinical Approved','Proven']) { if (rawPred.includes(pred)) { found = pred; break; } }
    d.predicates = found;
  } else { d.predicates = "None"; }

  // ─── PEDIGREE PARSING ───────────────────────────────────────────────────────
  const pedSplit = info.split(/Pedigree/i);
  const pedStart = pedSplit.length > 1 ? pedSplit[1] : null;
  if (pedStart) {
    const pedLines = pedStart.split(/Pregnancy/i)[0]
      .split(/\n/)
      .map(l => l.trim())
      .filter(l => l.length > 0 && !l.match(/^COI:/i));

    // Shared helpers defined in Pedigree.gs:
    // _isPlaceholder, _isStableName, _isTagline, _normalizeDecoName, _cleanName, _extractPedigreeSlots, _slotsToResult
    const result = _slotsToResult(_extractPedigreeSlots(pedLines));

    d.sire = result.sire || "Unknown";
    d.dam  = result.dam  || "Unknown";
    d.gs_p = result.gs_p; d.gd_p = result.gd_p;
    d.gs_m = result.gs_m; d.gd_m = result.gd_m;
    d.sireExists = checkHorseNameExists(d.sire);
    d.damExists  = checkHorseNameExists(d.dam);

    const coiMatch = pedStart.match(/COI:\s*([\d.,]+)\s*%/i);
    d.coi = coiMatch ? coiMatch[1].replace(',', '.') + "%" : "";
  }

  const pregMatch = info.match(/Pregnancy\s*[\r\n]+([\s\S]{0,300}?)(?=\n\n|\nHealth|\nGenetic|$)/i);
  if (pregMatch) {
    const pt = pregMatch[1];
    d.pregnancy = {
      foalGender:  (pt.match(/Unborn\s+(Colt|Filly)/i)||[])[1] || "",
      dueDate:     ((pt.match(/Due\s+on\s+([^\r\n]+)/i)||[])[1]||"").trim(),
      sire:        cleanHorseName(((pt.match(/Sire:\s*([^\r\n]+)/i)||[])[1]||"").trim()),
      coveredDate: ((pt.match(/Covered\s+(?:on\s+)?([^\r\n]+)/i)||pt.match(/(?:Breeding|Bred)\s+(?:on\s+)?([^\r\n]+)/i)||[])[1]||"").trim(),
      coveredTime: ""
    };
  }

  const colTxt = inputs.colour || "";
  ['Acceleration','Agility','Balance','Bascule','Pulling power','Speed','Sprint','Stamina','Strength','Surefootedness'].forEach(k => {
    d.gp[k] = (colTxt.match(new RegExp(k+"[^0-9]*(\\d+)","i"))||[])[1]||"";
  });

  // Alle Gene erstmal auf "n/n" (= unbekannt) setzen — Liste kommt aus der
  // Tabelle oben (GENE_PATTERNS) plus KIT, das separat behandelt wird.
  const geneKeys = GENE_PATTERNS.map(g => g.key).concat(['KIT']);
  geneKeys.forEach(gene=>{ d.genetics[gene]="n/n"; });
  const cleanG = (val) => val ? val.replace(/\s+/g,'').replace(/N$/i,'n') : "n/n";

  // Jede Zeile aus GENE_PATTERNS durchgehen und im Text suchen.
  GENE_PATTERNS.forEach(({ key, regex, wholeMatch }) => {
    const m = colTxt.match(regex);
    if (!m) return;
    d.genetics[key] = wholeMatch ? cleanG(m[0]) : cleanG(m[1] + '/' + m[2]);
  });

  // KIT ist ein Sonderfall: viele mögliche Allele (siehe KIT_ALLELES oben),
  // deshalb wird das Suchmuster dynamisch aus der Liste gebaut statt fest
  // eingetippt zu sein.
  const kitAllelePattern = KIT_ALLELES.join('|');
  const kitKeywordPattern = 'Tobiano|Roan|Sabino|' + KIT_ALLELES.filter(a => a.startsWith('W')).join('|') + '|\\bKIT\\b';
  const mKit = colTxt.match(new RegExp(
    '(?:' + kitKeywordPattern + ')[\\s\\S]{0,80}?(' + kitAllelePattern + ')\\s*[/\\\\]\\s*(' + kitAllelePattern + '|n)', 'i'
  ));
  if (mKit) d.genetics['KIT'] = cleanG(mKit[1] + '/' + mKit[2]);

  const statTxt = inputs.stats || "";
  const fullAnalysisText = colTxt + "\n" + statTxt;
  let rawConfo = {};
  ['Walk','Trot','Canter','Gallop','Posture','Head','Neck','Back','Shoulders','Frontlegs','Hindquarters','Socks'].forEach(k => {
    const reg = new RegExp((k==="Frontlegs"?"Front\\s*legs":k)+"[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)","i");
    const m = statTxt.match(reg); rawConfo[k] = (m&&m[1]) ? m[1].trim() : "";
  });
  if (d.breed&&d.breed.includes("Icelandic Horse")) {
    const toeltM=statTxt.match(/T[ö.]lt[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    const paceM=statTxt.match(/(?:Flying\s+pace|Pace)[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    rawConfo.Tolt=(toeltM&&toeltM[1])?toeltM[1].trim():""; rawConfo.Pace=(paceM&&paceM[1])?paceM[1].trim():"";
  }
  if (d.breed&&d.breed.includes("Kathiawari")) {
    const revaalM=statTxt.match(/Revaal[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    rawConfo.Revaal=(revaalM&&revaalM[1])?revaalM[1].trim():"";
  }
  d.confo = applyConfoLogic(rawConfo, fullAnalysisText);

  ['Fertility','Colic resistance','Hoof quality','Back problems','Respiratory disease','Resistance to lameness'].forEach(k => {
    d.health[k] = (statTxt.match(new RegExp(k+"[:\\s]+(Excellent|Good|Average|Fair|Poor)","i"))||[])[1]||"";
  });

  const achievementSection = statTxt.split(/Latest 25 show results/i)[0];
  const targetKeys = ['Day Champion','1st Premium','2nd Premium','3rd Premium','1st Prize','2nd Prize','3rd Prize'];
  const achLines = achievementSection.split(/\n/).map(l=>l.trim());
  targetKeys.forEach(k => {
    d.achieve[k] = "0";
    for (let i=0;i<achLines.length;i++) {
      if (achLines[i]===k) { for (let j=1;j<=3;j++) { if (achLines[i+j]&&/^\d+$/.test(achLines[i+j])) { d.achieve[k]=achLines[i+j]; break; } } }
    }
  });

  let showPart = statTxt.split(/show\s+results/i)[1] || "";
  showPart = showPart.split(/competition/i)[0];
  d.showScores = (showPart.match(/\d{1,3}[.,]\d{3}/g) || []);

  const compParts = statTxt.split(/competition\s+results/i);
  if (compParts.length > 1) {
    const compPart = compParts[1].split(/Health|Genetic/i)[0];
    d.compScores = (compPart.match(/\d{1,3}[.,]\d{3,4}/g) || []).filter(score => {
      const val = parseFloat(score.replace(',', '.'));
      return val > 10 && val < 150;
    }).slice(0, 25);
  }

  if (inputs.csv && inputs.csv.trim() !== "") {
    _mergeHrToolkitCsv(d, inputs.csv);
  }

  d.allHorses = getDropdownNames(ss);
  return d;
}

/* ══════════════════════════════════════════════════════════════
   HR TOOLKIT CSV — eigenständiges 4. Eingabefeld im Universal Import.
   Überschreibt gezielt die Felder, die das CSV strukturiert liefert
   (ID, Name, Breed, Sex, Birthdate, COI, GP-Werte, Konformation).
   Colour-Genetik, Health, Pedigree, Achievements, Pregnancy bleiben
   unberührt und kommen weiterhin aus den drei Paste-Feldern.
   ══════════════════════════════════════════════════════════════ */
const HRTOOLKIT_SEX_MAP = { C: 'Colt', F: 'Filly', S: 'Stallion', M: 'Mare', G: 'Gelding' };

const HRTOOLKIT_GP_FIELDS = {
  'Acceleration': 'Acceleration', 'Agility': 'Agility', 'Balance': 'Balance',
  'Bascule': 'Bascule', 'Pulling Power': 'Pulling power', 'Speed': 'Speed',
  'Sprint': 'Sprint', 'Stamina': 'Stamina', 'Strength': 'Strength',
  'Surefootedness': 'Surefootedness'
};

const HRTOOLKIT_CONFO_FIELDS = {
  'Walk': 'Walk', 'Trot': 'Trot', 'Canter': 'Canter', 'Gallop': 'Gallop',
  'Posture': 'Posture', 'Head': 'Head', 'Neck': 'Neck', 'Back': 'Back',
  'Shoulders': 'Shoulders', 'Frontlegs': 'Frontlegs', 'Hindquarters': 'Hindquarters',
  'Socks': 'Socks', 'Tölt': 'Tolt', 'Flying Pace': 'Pace', 'Revaal': 'Revaal'
};

function _mergeHrToolkitCsv(d, csvText) {
  const parsed = Utilities.parseCsv(csvText);
  if (parsed.length < 2) throw new Error("HR Toolkit CSV enthält keine Datenzeile.");

  const headers = parsed[0];
  const row = parsed[1];
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  const get = (name) => (idx[name] !== undefined ? String(row[idx[name]] || "").trim() : "");

  const rawId = get('ID');
  if (rawId) d.id = rawId;

  const rawName = get('Name');
  if (rawName) d.name = rawName.replace(/^[^\p{L}\p{N}]+\s*/u, '').trim();

  const rawBreed = get('Breed');
  if (rawBreed) d.breed = rawBreed;

  const rawSex = get('Sex');
  if (rawSex && HRTOOLKIT_SEX_MAP[rawSex]) d.gender = HRTOOLKIT_SEX_MAP[rawSex];

  const rawDob = get('Birthdate');
  let dobDay, dobMonth, dobYear, m;
  if ((m = rawDob.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/))) {
    dobDay = m[1]; dobMonth = m[2]; dobYear = m[3]; // DD-MM-YYYY (HR Toolkit)
  } else if ((m = rawDob.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/))) {
    dobYear = m[1]; dobMonth = m[2]; dobDay = m[3]; // YYYY-MM-DD Fallback
  }
  if (dobDay) {
    d.dob = dobMonth.padStart(2, '0') + '/' + dobDay.padStart(2, '0') + '/' + dobYear;
  }

  const rawCoi = get('COI');
  if (rawCoi && rawCoi.toLowerCase() !== 'n/a') {
    d.coi = rawCoi.replace(',', '.') + '%';
  }

  if (!d.gp) d.gp = {};
  for (const csvField in HRTOOLKIT_GP_FIELDS) {
    const val = get(csvField);
    if (val && val.toLowerCase() !== 'n/a') d.gp[HRTOOLKIT_GP_FIELDS[csvField]] = val;
  }

  if (!d.confo) d.confo = {};
  for (const csvField in HRTOOLKIT_CONFO_FIELDS) {
    const val = get(csvField);
    if (val && val.toLowerCase() !== 'n/a') d.confo[HRTOOLKIT_CONFO_FIELDS[csvField]] = val;
  }
}

function applyConfoLogic(base, text) {
  const final = {};
  const check = (key, phrase) => { if (base[key]==="Good") return text.includes(phrase)?"G+":"G-"; return shortenGrade(base[key]); };
  final.Walk=check('Walk',"amazing two-beat rhythm"); final.Trot=check('Trot',"amazing two-beat rhythm");
  final.Canter=check('Canter',"very elegant looking canter"); final.Gallop=check('Gallop',"balanced and very smooth gallop");
  final.Posture=check('Posture',"posture is perfectly balanced"); final.Head=check('Head',"shows some nice proportions");
  final.Back=check('Back',"back will be heaven for a rider"); final.Frontlegs=check('Frontlegs',"front legs are practically identical");
  final.Hindquarters=check('Hindquarters',"great engagement in the hindquarters"); final.Socks=check('Socks',"feathering is amazing");
  const pNS=text.includes("great elasticity in the shoulder"); const n=base.Neck, s=base.Shoulders;
  if (n==="Good"&&s==="Good") final.Neck=pNS?"G+":"G"; else if (n==="Good"&&s==="Very good") final.Neck=pNS?"G":"G-";
  else if (n==="Very good"&&s==="Good") final.Neck=pNS?"VG":"VG"; else final.Neck=shortenGrade(n);
  if (s==="Good"&&n==="Good") final.Shoulders=pNS?"G+":"G"; else if (s==="Good"&&n==="Very good") final.Shoulders=pNS?"G":"G-";
  else final.Shoulders=shortenGrade(s);
  final.Tolt=shortenGrade(base.Tolt); final.Pace=shortenGrade(base.Pace); final.Revaal=shortenGrade(base.Revaal);
  return final;
}

function shortenGrade(grade) {
  if (!grade) return "";
  const g = grade.toLowerCase();
  if (g.includes("very good")) return "VG";
  if (g.includes("below average")) return "BA";
  if (g.includes("good")) return "G";
  if (g.includes("average")) return "A";
  if (g.includes("poor")) return "P";
  return grade;
}

function getHerdsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of ['Herd Tracker','HERD TRACKER','HerdTracker','HERD']) { const s=ss.getSheetByName(name); if(s) return s; }
  return null;
}

function getDropdownNames(ss) {
  const herd = getHerdsheet(); if (!herd) return [];
  const lastRow = herd.getLastRow(); if (lastRow<2) return [];
  try { return herd.getRange(2,4,lastRow-1,1).getValues().map(r=>r[0]).filter(n=>n&&n.toString().trim()!=='').map(String).sort(); }
  catch(e) { return []; }
}

function findRow(sheet, id, col) {
  if (!sheet) return 2; const lastRow=sheet.getLastRow(); if (lastRow<1) return 2;
  const range=sheet.getRange(1,col,lastRow+10,1).getValues().map(r=>r[0].toString());
  const idx=range.indexOf(id.toString()); if(idx!==-1) return idx+1;
  for (let i=1;i<range.length;i++) { if(!range[i]||range[i]==="") return i+1; }
  return lastRow+1;
}

function cleanHorseName(name) {
  if (!name) return "Unknown";
  let n = name;

  // "!94.453 ℘ Name" → Score-Präfix strippen
  n = n.replace(/^[!$]?\d+[\.\,]\d+\s*[-–]?\s*\d*[\.\,]?\d*\s*[℘\s]*/,'').trim();

  // "95.029︱Name" → "Name"
  const scorePrefix = n.match(/^\d+[\.\,]\d+\s*[|︱]\s*(.+)$/);
  if (scorePrefix) n = scorePrefix[1];

  n = n.replace(/Deceased/gi, '');
  n = n.replace(/^[★☆ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃✶〔〕♟˚₊✧⋆°·•﹚ℙ♕⭐️♡♥*\-͟͞➳❥༅\s]+/,'').trim();
  n = n.replace(/\s*\|\|.*$/,'').trim();
  n = n.replace(/\s*\(\d+[\.\,]\d+\)\s*$/,'').trim();
  n = n.replace(/\s+\d+[\.\,]\d{3,}\s*$/,'').trim();
  n = n.replace(/\s*[|︱︲\/]\s*[\d\s].*$/,'').trim();
  n = n.split(/Tagline/i)[0];
  n = n.replace(/[★☆ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃✶〔〕♟˚₊✧⋆°·•﹚⭐️♡♥\s]+$/,'').trim();

  return n || "Unknown";
}

function checkHorseNameExists(horseName) {
  if (!horseName||horseName==="Unknown") return false;
  const herdSheet=getHerdsheet(); if(!herdSheet) return false;
  const lastRow=herdSheet.getLastRow(); if(lastRow<2) return false;
  const names=herdSheet.getRange(2,4,lastRow-1,1).getValues();
  const target=cleanHorseName(horseName);
  for (let i=0;i<names.length;i++) { if(names[i][0]&&cleanHorseName(names[i][0].toString())===target) return true; }
  return false;
}

/* ══════════════════════════════════════════════════════════════
   SPALTEN-KONSTANTEN — IMPORT-ZIELE
   Definiert, in welche Spalte processFinalImport() pro Sheet
   schreibt. Bei Spaltenverschiebungen im Sheet NUR HIER anpassen.
   (Analog zum SHEET_CONFIG-Muster in SheetConfig.gs, dort aber
   für Styling/Layout statt für Import-Schreibziele.)
   ══════════════════════════════════════════════════════════════ */

const HERD_COLS = {
  ID: 3, NAME: 4, BREED: 5, GENDER: 6, DOB: 7, PREDICATES: 18,
  HEALTH_START: 30 // 6 Spalten: Colic, Hoof, Back, Respiratory, Lameness, Fertility
};

const STATS_COLS = { ID: 2, NAME: 3, CONFO_START: 4 };

// Startspalte der 10 GP-Werte (Acceleration…Surefootedness), je nach Rasse
const STATS_GP_COL = { default: 35, ICE: 39, KATH: 37 };

// Zielspalten der Achievement-Zähler, je nach Rasse
const STATS_ACH_COLS = {
  default: { 'Day Champion':50,'1st Premium':51,'2nd Premium':52,'3rd Premium':53,'1st Prize':71,'2nd Prize':72,'3rd Prize':73 },
  ICE:     { 'Day Champion':54,'1st Premium':55,'2nd Premium':56,'3rd Premium':57,'1st Prize':75,'2nd Prize':76,'3rd Prize':77 },
  KATH:    { 'Day Champion':52,'1st Premium':53,'2nd Premium':54,'3rd Premium':55,'1st Prize':73,'2nd Prize':74,'3rd Prize':75 }
};

const COLOUR_COLS = {
  ID: 2, NAME: 3,
  EXTENSION: 4, AGOUTI: 5, GREY: 6, CREAMPEARL: 7, DUN: 8, CHAMPAGNE: 9,
  SILVER: 10, MUSHROOM: 11, FRAME: 12, APPALOOSA: 13, PATN1: 14, PATN2: 15,
  MITF: 16, SW2: 17,
  KIT_TO: 22, KIT_RN: 23, KIT_SB1: 24, KIT_W3: 27, KIT_W8: 32, KIT_W10: 34,
  KIT_W16: 40, KIT_W19: 43, KIT_W20: 44, KIT_W21: 45,
  FLAXEN: 48, SOOTY: 49, PANGARE: 50, WHITE_SPOTTING: 51, RABICANO: 52, MARKINGS: 53
};

const PEDIGREE_COLS = {
  ID: 2, NAME: 3, COI: 4, STALLIONLINE: 5, DAMLINE: 6,
  SIRE: 7, DAM: 8, GS_P: 9, GD_P: 10, GS_M: 11, GD_M: 12,
  GGS_PP: 13, GGD_PP: 14, GGS_PM: 15, GGD_PM: 16,
  GGS_MP: 17, GGD_MP: 18, GGS_MM: 19, GGD_MM: 20
};

const BROODMARE_COLS = {
  ID: 3, NAME: 4, BREED: 5, COVERED_DATE: 6, COVERED_TIME: 7,
  FOAL_GENDER: 12, SIRE_INSIDE: 13, DUE_DATE: 14, SIRE_OUTSIDE: 15, FAILINGS: 16
};

// ─── KLEINE KONVERTIERUNGS-HELFER ─────────────────────────────────────────────

function _getBreedVariant(breed) {
  if (!breed) return 'default';
  if (breed.includes('Icelandic Horse')) return 'ICE';
  if (breed.includes('Kathiawari')) return 'KATH';
  return 'default';
}

function _convertDOB(dob) {
  if (!dob) return "";
  if (dob.match(/^\d{2}\/\d{2}\/\d{4}$/)) { const p = dob.split('/'); return p[2] + '-' + p[0] + '-' + p[1]; }
  if (dob.match(/^\d{2}-\d{2}-\d{4}$/)) { const p = dob.split('-'); return p[2] + '-' + p[1] + '-' + p[0]; }
  return dob;
}

function _convertDueDate(raw) {
  if (!raw) return "";
  const mm = {jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};
  const cleaned = raw.replace(/^[A-Za-z]+,\s*/, "").trim();
  const parts = cleaned.split(/[\s,]+/);
  if (parts.length >= 2) {
    return (mm[parts[0].toLowerCase().slice(0,3)] || '01') + '/' + parts[1].padStart(2,'0') + '/' + (parts[2] || new Date().getFullYear().toString());
  }
  return raw;
}

// ─── PRO-SHEET SCHREIB-FUNKTIONEN ──────────────────────────────────────────────
// Jede Funktion kapselt genau ein Ziel-Sheet. Verhalten 1:1 identisch zur
// vorherigen processFinalImport()-Monolith-Funktion, nur aufgeteilt und mit
// benannten Spalten statt Magic Numbers.

function _writeHerdTracker(data, convertedDOB) {
  const herdSheet = getHerdsheet();
  if (!herdSheet) return null;

  const row = findRow(herdSheet, data.id, HERD_COLS.ID);
  herdSheet.getRange(row, HERD_COLS.ID).setValue(data.id);
  herdSheet.getRange(row, HERD_COLS.NAME).setValue(data.name);
  herdSheet.getRange(row, HERD_COLS.BREED).setValue(data.breed || "");
  herdSheet.getRange(row, HERD_COLS.GENDER).setValue(data.gender || "");
  herdSheet.getRange(row, HERD_COLS.DOB).setValue(convertedDOB);
  herdSheet.getRange(row, HERD_COLS.PREDICATES).setValue(
    (data.predicates && data.predicates.toLowerCase() !== "none") ? data.predicates : ""
  );
  herdSheet.getRange(row, HERD_COLS.HEALTH_START, 1, 6).setValues([[
    data.health['Colic resistance'] || "",
    data.health['Hoof quality'] || "",
    data.health['Back problems'] || "",
    data.health['Respiratory disease'] || "",
    data.health['Resistance to lameness'] || "",
    data.health['Fertility'] || ""
  ]]);

  return "✓ Herd Tracker updated";
}

function _writeStats(ss, data) {
  const variant = _getBreedVariant((data.breed || "").trim());
  const targetSheetName = variant === 'ICE' ? 'ICE_Stats' : variant === 'KATH' ? 'KATH_Stats' : 'Stats';
  const statSheet = ss.getSheetByName(targetSheetName);
  if (!statSheet) return null;

  const row = findRow(statSheet, data.id, STATS_COLS.ID);
  statSheet.getRange(row, STATS_COLS.ID).setValue(data.id);
  statSheet.getRange(row, STATS_COLS.NAME).setValue(data.name);

  let confo = [data.confo.Walk || "", data.confo.Trot || "", data.confo.Canter || "", data.confo.Gallop || ""];
  if (variant === 'ICE') confo.push(data.confo.Tolt || "", data.confo.Pace || "");
  else if (variant === 'KATH') confo.push(data.confo.Revaal || "");
  confo.push(
    data.confo.Posture || "", data.confo.Head || "", data.confo.Neck || "", data.confo.Back || "",
    data.confo.Shoulders || "", data.confo.Frontlegs || "", data.confo.Hindquarters || "", data.confo.Socks || ""
  );
  statSheet.getRange(row, STATS_COLS.CONFO_START, 1, confo.length).setValues([confo]);

  const gpCol = STATS_GP_COL[variant];
  statSheet.getRange(row, gpCol, 1, 10).setValues([[
    data.gp['Acceleration'] || "", data.gp['Agility'] || "", data.gp['Balance'] || "", data.gp['Bascule'] || "",
    data.gp['Pulling power'] || "", data.gp['Speed'] || "", data.gp['Sprint'] || "", data.gp['Stamina'] || "",
    data.gp['Strength'] || "", data.gp['Surefootedness'] || ""
  ]]);

  const achMap = STATS_ACH_COLS[variant];
  Object.entries(achMap).forEach(([key, col]) => statSheet.getRange(row, col).setValue(data.achieve[key] || "0"));

  return "✓ " + targetSheetName + " updated";
}

function _writeColourGenetics(ss, data, manualData) {
  const colSheet = ss.getSheetByName('Colour Genetics');
  if (!colSheet) return null;

  const row = findRow(colSheet, data.id, COLOUR_COLS.ID);
  const kitVal = data.genetics['KIT'] || "n/n";
  const kitA1 = (kitVal.split('/')[0] || '').trim();
  const kitA2 = (kitVal.split('/')[1] || '').trim();
  const getKitAllele = (allele) => {
    const a = allele.toLowerCase();
    const c = (kitA1.toLowerCase() === a ? 1 : 0) + (kitA2.toLowerCase() === a ? 1 : 0);
    if (c === 2) return allele + '/' + allele;
    if (c === 1) return allele + '/n';
    return 'n/n';
  };
  const wsValue = (data.genetics['WhiteSpotting'] && data.genetics['WhiteSpotting'] !== 'n/n')
    ? data.genetics['WhiteSpotting']
    : (manualData['HiddenSabino'] && manualData['HiddenSabino'].trim() !== '' ? manualData['HiddenSabino'].trim() : "n/n");

  colSheet.getRange(row, COLOUR_COLS.ID).setValue(data.id);
  colSheet.getRange(row, COLOUR_COLS.NAME).setValue(data.name);
  colSheet.getRange(row, COLOUR_COLS.EXTENSION).setValue(data.genetics['Extension'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.AGOUTI).setValue(data.genetics['Agouti'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.GREY).setValue(data.genetics['Grey'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.CREAMPEARL).setValue(data.genetics['Creampearl'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.DUN).setValue(data.genetics['Dun'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.CHAMPAGNE).setValue(data.genetics['Champagne'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.SILVER).setValue(data.genetics['Silver'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.MUSHROOM).setValue(data.genetics['Mushroom'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.FRAME).setValue(data.genetics['Frame'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.APPALOOSA).setValue(data.genetics['Appaloosa'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.PATN1).setValue(data.genetics['PATN1'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.PATN2).setValue(manualData['PATN2'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.MITF).setValue(data.genetics['MITF'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.SW2).setValue(data.genetics['SW2'] || "n/n");

  colSheet.getRange(row, COLOUR_COLS.KIT_TO).setValue(getKitAllele('TO'));
  colSheet.getRange(row, COLOUR_COLS.KIT_RN).setValue(getKitAllele('Rn'));
  colSheet.getRange(row, COLOUR_COLS.KIT_SB1).setValue(getKitAllele('SB1'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W3).setValue(getKitAllele('W3'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W8).setValue(getKitAllele('W8'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W10).setValue(getKitAllele('W10'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W16).setValue(getKitAllele('W16'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W19).setValue(getKitAllele('W19'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W20).setValue(getKitAllele('W20'));
  colSheet.getRange(row, COLOUR_COLS.KIT_W21).setValue(getKitAllele('W21'));

  colSheet.getRange(row, COLOUR_COLS.FLAXEN).setValue(manualData['Flaxen'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.SOOTY).setValue(manualData['Sooty'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.PANGARE).setValue(manualData['Pangare'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.WHITE_SPOTTING).setValue(wsValue);
  colSheet.getRange(row, COLOUR_COLS.RABICANO).setValue(manualData['Rabicano'] || "n/n");
  colSheet.getRange(row, COLOUR_COLS.MARKINGS).setValue(manualData['Markings'] || "");

  return "✓ Colour Genetics updated";
}

const SCORE_START_ROW = {
  'Conf. Results': 8,
  'Comp. Results': 7
};

function _writeScores(ss, data) {
  const results = [];
  ['Conf. Results', 'Comp. Results'].forEach(sName => {
    const s = ss.getSheetByName(sName);
    const scores = (sName === 'Conf. Results') ? data.showScores : data.compScores;
    if (!s || !scores || scores.length === 0) return;

    // Erste leere Spalte in Zeile 1 finden — EIN Call statt vieler
    const lastCol = s.getLastColumn();
    const headerRow = lastCol > 0 ? s.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    let col = headerRow.length + 1;
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === "" || headerRow[i] === null) { col = i + 1; break; }
    }

    s.getRange(1, col).setValue(data.id);
    s.getRange(2, col).setValue(data.name);

    const startRow = SCORE_START_ROW[sName];
    const values = scores.slice(0, 25).map(val => {
      const num = parseFloat(val.toString().replace(',', '.'));
      return [isNaN(num) ? "" : num];
    });
    if (values.length > 0) {
      s.getRange(startRow, col, values.length, 1).setValues(values);
    }
    refreshResultColumnStyle_(s, col);
    results.push("✓ " + sName + " updated");
  });
  return results;
}

function _writePedigree(ss, data, sireIsInside, damIsInside) {
  const pedSheet = ss.getSheetByName('Pedigree');
  if (!pedSheet) return null;

  const row = findRow(pedSheet, data.id, PEDIGREE_COLS.ID);
  pedSheet.getRange(row, PEDIGREE_COLS.ID).setValue(data.id);
  pedSheet.getRange(row, PEDIGREE_COLS.NAME).setValue(data.name);
  pedSheet.getRange(row, PEDIGREE_COLS.SIRE).setValue(data.sire || "");
  pedSheet.getRange(row, PEDIGREE_COLS.DAM).setValue(data.dam || "");
  pedSheet.getRange(row, PEDIGREE_COLS.COI).setValue(data.coi);

  if (!sireIsInside) {
    if (data.gs_p) pedSheet.getRange(row, PEDIGREE_COLS.GS_P).setValue(data.gs_p);
    if (data.gd_p) pedSheet.getRange(row, PEDIGREE_COLS.GD_P).setValue(data.gd_p);
  }
  if (!damIsInside) {
    if (data.gs_m) pedSheet.getRange(row, PEDIGREE_COLS.GS_M).setValue(data.gs_m);
    if (data.gd_m) pedSheet.getRange(row, PEDIGREE_COLS.GD_M).setValue(data.gd_m);
  }
  if (damIsInside) copyParentPedigree(pedSheet, row, data.dam, "dam");
  if (sireIsInside) copyParentPedigree(pedSheet, row, data.sire, "sire");

  return "✓ Pedigree updated";
}

function _writeBroodmares(ss, data) {
  if (!data.pregnancy || (!data.pregnancy.foalGender && !data.pregnancy.dueDate)) return null;
  const broodSheet = ss.getSheetByName('Broodmares');
  if (!broodSheet) return null;

  const row = findRow(broodSheet, data.id, BROODMARE_COLS.ID);
  broodSheet.getRange(row, BROODMARE_COLS.ID).setValue(data.id);
  broodSheet.getRange(row, BROODMARE_COLS.NAME).setValue(data.name);
  broodSheet.getRange(row, BROODMARE_COLS.BREED).setValue(data.breed || "");
  if (data.pregnancy.coveredDate) broodSheet.getRange(row, BROODMARE_COLS.COVERED_DATE).setValue(_convertGermanDate(data.pregnancy.coveredDate));
  if (data.pregnancy.coveredTime) broodSheet.getRange(row, BROODMARE_COLS.COVERED_TIME).setValue(data.pregnancy.coveredTime);
  broodSheet.getRange(row, BROODMARE_COLS.FOAL_GENDER).setValue(data.pregnancy.foalGender || "");
  if (data.pregnancy.sire) {
    const sireCol = checkHorseNameExists(data.pregnancy.sire) ? BROODMARE_COLS.SIRE_INSIDE : BROODMARE_COLS.SIRE_OUTSIDE;
    broodSheet.getRange(row, sireCol).setValue(data.pregnancy.sire);
  }
  if (data.pregnancy.dueDate) broodSheet.getRange(row, BROODMARE_COLS.DUE_DATE).setValue(_convertDueDate(data.pregnancy.dueDate));

  return "✓ Broodmares updated";
}

// ─── HAUPTFUNKTION ──────────────────────────────────────────────────────────
// Orchestriert nur noch: pro Sheet die passende _write...-Funktion aufrufen
// und die Ergebnis-Meldungen sammeln. Keine Spaltenlogik mehr hier.

function processFinalImport(data, opts, manual) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const manualData = manual || {};
    const convertedDOB = _convertDOB(data.dob);
    const sireIsInside = data.sireIsInside || false;
    const damIsInside = data.damIsInside || false;
    const needsPedigreeModal = (!damIsInside || !sireIsInside);
    
    let results = [];

    // ─── ALLES AUSSER PEDIGREE SOFORT SPEICHERN ───
    if (opts.herd) {
      const r = _writeHerdTracker(data, convertedDOB);
      if (r) results.push(r);
      updateStallionDropdown();
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

    const broodResult = _writeBroodmares(ss, data);
    if (broodResult) results.push(broodResult);

    // ─── PEDIGREE: NUR WENN INSIDE-ELTERN SOFORT, SONST WARTEN ───
    if (!needsPedigreeModal) {
      const pedResult = _writePedigree(ss, data, sireIsInside, damIsInside);
      if (pedResult) results.push(pedResult);
      
      return {
        message: "✓ Import complete!\n" + results.join("\n"),
        openPedigreeFirst: false
      };
    }

    const pedSheet = ss.getSheetByName('Pedigree');
    if (pedSheet && data.coi) {
      const row = findRow(pedSheet, data.id, PEDIGREE_COLS.ID);
      pedSheet.getRange(row, PEDIGREE_COLS.ID).setValue(data.id);
      pedSheet.getRange(row, PEDIGREE_COLS.NAME).setValue(data.name);
      pedSheet.getRange(row, PEDIGREE_COLS.COI).setValue(data.coi);
    }

    // OUTSIDE-ELTERN: Modal öffnen (backend-seitig)
    results.push("✓ Opening Pedigree Editor...");
    openPedigreeModalWithData({
      id: data.id,
      name: data.name,
      sire: data.sire || "Unknown",
      dam: data.dam || "Unknown",
      gs_p: data.gs_p || "",
      gd_p: data.gd_p || "",
      gs_m: data.gs_m || "",
      gd_m: data.gd_m || "",
      coi: data.coi || "",
      predicates: data.predicates || ""
    });
    
    return {
      message: results.join("\n"),
      openPedigreeFirst: true
    };
  } catch (e) {
    throw new Error("Import error: " + e.message);
  }
}

function copyParentPedigree(sheet, targetRow, parentName, type) {
  if (!parentName || parentName === "Unknown") return;

  const fullData = sheet.getDataRange().getValues();
  let pIdx = -1;
  for (let i = 0; i < fullData.length; i++) {
    if (fullData[i][PEDIGREE_COLS.NAME - 1] === parentName) { pIdx = i + 1; break; }
  }
  if (pIdx === -1) return;

  if (type === "dam") {
    sheet.getRange(targetRow, PEDIGREE_COLS.DAMLINE).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.DAMLINE).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GS_M).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.SIRE).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GD_M).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.DAM).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_MP).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GS_P).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_MP).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GD_P).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_MM).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GS_M).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_MM).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GD_M).getValue());
  } else {
    sheet.getRange(targetRow, PEDIGREE_COLS.STALLIONLINE).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.STALLIONLINE).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GS_P).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.SIRE).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GD_P).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.DAM).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_PP).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GS_P).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_PP).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GD_P).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGS_PM).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GS_M).getValue());
    sheet.getRange(targetRow, PEDIGREE_COLS.GGD_PM).setValue(sheet.getRange(pIdx, PEDIGREE_COLS.GD_M).getValue());
  }
}