/***********************
 * OUTSIDE STUDS IMPORT BACKEND
 ***********************/

function openOutsideStudsModal() {
  const html = HtmlService.createTemplateFromFile('outsideStudsModal')
    .evaluate()
    .setWidth(900)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Parses the three tab texts and returns preview data.
 * Reuses the same parsing logic as the main import.
 */
function getOutsideStudPreviewData(inputs) {
  const d = {
    id: '', name: '', breed: '', gender: '',
    confo: {}, gp: {}, genetics: {},
    confoMax: '', compMax: '', 
    coi: '',
    hcr: '',  // ✨ NEU: HCR Währung
    dp: '',   // ✨ NEU: DP Währung
    ft: '',   // ✨ NEU: FT Währung
    wt: ''    // ✨ NEU: WT Währung
  };

  const info    = inputs.info    || '';
  const colTxt  = inputs.colour  || '';
  const statTxt = inputs.stats   || '';

  // --- ID ---
  const idMatch = info.match(/Life\s*number\s*[\r\n]+#?(\d+)/i)
               || info.match(/Lifenumber\s*[:.]?\s*#?(\d+)/i);
  d.id = idMatch ? idMatch[1] : '';

  // --- NAME ---
  let rawName = '';
  const passportMatch = info.match(/Passport\s*[\r\n]+([^\r\n]+)/i);
  if (passportMatch) {
    rawName = passportMatch[1];
  } else {
    rawName = (info.match(/Name\s*([\s\S]+?)(?=\n|Tagline|RC\d)/i) || [])[1] || '';
  }
  d.name = rawName.replace(/Tagline.*/i, '').trim();

  // --- BREED ---
  const breedMatch = info.match(/Breed\s*registry\s*[\r\n]+([^\r\n]+)/i);
  if (breedMatch) {
    const br = breedMatch[1].trim();
    const breedMap = {
      'akhal':'Akhal-Teke','arabian':'Arabian','brabant':'Brabant','brumby':'Brumby',
      'camargue':'Camargue','cleveland':'Cleveland Bay','exmoor':'Exmoor Pony',
      'finnhorse':'Finnhorse','fjord':'Fjord Horse','friesian':'Friesian',
      'haflinger':'Haflinger','icelandic':'Icelandic Horse','irish cob':'Irish Cob',
      'kathiawari':'Kathiawari','kladruber':'Kladruber','knabstrupper':'Knabstrupper',
      'lipizzaner':'Lipizzaner','lusitano':'Lusitano','mongolian':'Mongolian Horse',
      'mustang':'Mustang','namib':'Namib Desert Horse','noriker':'Noriker',
      'norman cob':'Norman Cob','oldenburg':'Oldenburg','pantaneiro':'Pantaneiro',
      'pre ':'PRE','quarter':'Quarter Horse','shetland':'Shetland Pony',
      'shire':'Shire Horse','suffolk':'Suffolk Punch','thoroughbred':'Thoroughbred',
      'trakehner':'Trakehner','welsh':'Welsh Pony'
    };
    const brLow = br.toLowerCase();
    let matched = false;
    for (const [key, name] of Object.entries(breedMap)) {
      if (brLow.includes(key)) { d.breed = name; matched = true; break; }
    }
    if (!matched) d.breed = br.replace(/\s*Horse\s*Society\s*/i, '').replace(/\s*Society\s*/i, '').trim();
  }

  // --- GENDER ---
  const genderMatch = info.match(/\b(Mare|Stallion|Gelding)\b/i)
                   || statTxt.match(/\b(Mare|Stallion|Gelding)\b/i);
  d.gender = genderMatch ? genderMatch[1] : '';

  // --- COI ---
  const coiMatch = info.match(/COI[:\s]+(\d+(?:\.\d+)?)%?/i);
  if (coiMatch) {
    d.coi = coiMatch[1];
  }

  // --- CONFORMATION ---
  const fullAnalysisText = colTxt + '\n' + statTxt;
  let rawConfo = {};

  ['Walk','Trot','Canter','Gallop','Posture','Head','Neck','Back','Shoulders','Frontlegs','Hindquarters','Socks'].forEach(k => {
    const searchKey = k === 'Frontlegs' ? 'Front\\s*legs' : k;
    const reg = new RegExp(searchKey + '[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)', 'i');
    const m = statTxt.match(reg);
    rawConfo[k] = (m && m[1]) ? m[1].trim() : '';
  });

  // Icelandic gaits
  if (d.breed && d.breed.includes('Icelandic Horse')) {
    const toeltM = statTxt.match(/T[ö.]lt[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    const paceM  = statTxt.match(/(?:Flying\s+pace|Pace)[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    rawConfo.Tolt       = (toeltM && toeltM[1]) ? toeltM[1].trim() : '';
    rawConfo.FlyingPace = (paceM  && paceM[1])  ? paceM[1].trim()  : '';
  }

  // Kathiawari gait
  if (d.breed && d.breed.includes('Kathiawari')) {
    const revaalM = statTxt.match(/Revaal[^a-zA-Z0-9]*(Very good|Good|Average|Below average|Poor)/i);
    rawConfo.Revaal = (revaalM && revaalM[1]) ? revaalM[1].trim() : '';
  }

  d.confo = applyOutsideStudConfoLogic(rawConfo, fullAnalysisText);

  // --- GP ---
  ['Acceleration','Agility','Balance','Bascule','Pulling power','Speed','Sprint','Stamina','Strength','Surefootedness'].forEach(k => {
    const reg = new RegExp(k + '[^0-9]*(\\d+)', 'i');
    d.gp[k] = (colTxt.match(reg) || [])[1] || '';
  });

  // --- GENETICS ---
  const geneKeys = ['Extension','Agouti','Grey','Creampearl','Dun','Champagne',
    'Silver','Mushroom','Frame','Appaloosa','PATN1','MITF','SW2','KIT','WhiteSpotting'];
  geneKeys.forEach(gene => { d.genetics[gene] = 'n/n'; });

  const cleanG = (val) => val ? val.replace(/\s+/g, '').replace(/N$/i, 'n') : 'n/n';

  const mExt = colTxt.match(/Extension[\s\S]{0,80}?([Ee])\s*[\/\\]\s*([Ee])/i);
  if (mExt) d.genetics['Extension'] = cleanG(mExt[1] + '/' + mExt[2]);

  const mAgo = colTxt.match(/(?:[Aa]\+|[Aa]t|[Aa]|aa)\s*[\/\\]\s*(?:[Aa]\+|[Aa]t|[Aa]|aa)/);
  if (mAgo) d.genetics['Agouti'] = cleanG(mAgo[0]);

  const mGrey = colTxt.match(/Grey[\s\S]{0,80}?([Gg])\s*[\/\\]\s*([Gg]|n)/i);
  if (mGrey) d.genetics['Grey'] = cleanG(mGrey[1] + '/' + mGrey[2]);

  const mSilv = colTxt.match(/Silver[\s\S]{0,80}?([Zz])\s*[\/\\]\s*([Zz])/i);
  if (mSilv) d.genetics['Silver'] = cleanG(mSilv[1] + '/' + mSilv[2]);

  const mCr = colTxt.match(/Cream[\s\S]{0,80}?(Cr|prl)\s*[\/\\]\s*(Cr|prl|n)/i);
  if (mCr) d.genetics['Creampearl'] = cleanG(mCr[1] + '/' + mCr[2]);

  const mDun = colTxt.match(/\bDun\b[\s\S]{0,80}?(nd[12]|[Dd])\s*[\/\\]\s*(nd[12]|[Dd])/i);
  if (mDun) d.genetics['Dun'] = cleanG(mDun[1] + '/' + mDun[2]);

  const mChamp = colTxt.match(/Champagne[\s\S]{0,80}?(Ch)\s*[\/\\]\s*(Ch|n)/i);
  if (mChamp) d.genetics['Champagne'] = cleanG(mChamp[1] + '/' + mChamp[2]);

  const mMush = colTxt.match(/Mushroom[\s\S]{0,80}?(mu|[Mm])\s*[\/\\]\s*(mu|[Mm]|n)/i);
  if (mMush) d.genetics['Mushroom'] = cleanG(mMush[1] + '/' + mMush[2]);

  const mFrame = colTxt.match(/(?:\bFrame\b|\bOWL\b)[\s\S]{0,80}?(OWL|[Oo])\s*[\/\\]\s*(OWL|[Oo]|n)/i);
  if (mFrame) d.genetics['Frame'] = cleanG(mFrame[1] + '/' + mFrame[2]);

  const mLp = colTxt.match(/(?:Appaloosa|\bLp\b)[\s\S]{0,80}?(Lp)\s*[\/\\]\s*(Lp|n)/i);
  if (mLp) d.genetics['Appaloosa'] = cleanG(mLp[1] + '/' + mLp[2]);

  const mPatn1 = colTxt.match(/PATN1[\s\S]{0,80}?(PATN1)\s*[\/\\]\s*(PATN1|n)/i);
  if (mPatn1) d.genetics['PATN1'] = cleanG(mPatn1[1] + '/' + mPatn1[2]);

  const mMitf = colTxt.match(/(?:SW[13]|MITF)[\s\S]{0,80}?(SW[13])\s*[\/\\]\s*(SW[13]|n)/i);
  if (mMitf) d.genetics['MITF'] = cleanG(mMitf[1] + '/' + mMitf[2]);

  const mSw2 = colTxt.match(/SW2[\s\S]{0,80}?(SW2)\s*[\/\\]\s*(SW2|n)/i);
  if (mSw2) d.genetics['SW2'] = cleanG(mSw2[1] + '/' + mSw2[2]);

  const kitPat      = 'W21|W20|W19|W16|W10|W8|W3|SB1|Rn|TO';
  const kitKeywords = 'Tobiano|Roan|Sabino|W21|W20|W19|W16|W10|W8|W3|\\bKIT\\b';
  const mKit = colTxt.match(new RegExp('(?:' + kitKeywords + ')[\\s\\S]{0,80}?(' + kitPat + ')\\s*[/\\\\]\\s*(' + kitPat + '|n)', 'i'));
  if (mKit) d.genetics['KIT'] = cleanG(mKit[1] + '/' + mKit[2]);

  const mWs = colTxt.match(/(?:White\s*Spotting|\bWS\b)[\s\S]{0,80}?(WS)\s*[\/\\]\s*(WS|n)/i);
  if (mWs) d.genetics['WhiteSpotting'] = cleanG(mWs[1] + '/' + mWs[2]);

  // --- CONFO MAX & COMP MAX ---
  const showPart = statTxt.split(/show\s+results/i)[1] || '';
  const showScores = (showPart.split(/competition/i)[0].match(/\b\d{1,3}[.,]\d{3}\b/g) || [])
    .map(s => parseFloat(s.replace(',', '.')))
    .filter(v => !isNaN(v));
  if (showScores.length > 0) d.confoMax = Math.max(...showScores).toFixed(3);

  const compParts = statTxt.split(/competition\s+results/i);
  if (compParts.length > 1) {
    const compScores = (compParts[1].split(/Health|Genetic/i)[0].match(/\b\d{1,3}[.,]\d{3,4}\b/g) || [])
      .map(s => parseFloat(s.replace(',', '.')))
      .filter(v => v > 10 && v < 150);
    if (compScores.length > 0) d.compMax = Math.max(...compScores).toFixed(3);
  }

  return d;
}

/**
 * Applies G+/G- logic for Outside Stud confo — same as main import.
 */
function applyOutsideStudConfoLogic(base, text) {
  const final = {};
  const check = (key, phrase) => {
    if (base[key] === 'Good') return text.includes(phrase) ? 'G+' : 'G-';
    return shortenGradeOutside(base[key]);
  };

  final.Walk         = check('Walk',          'amazing two-beat rhythm');
  final.Trot         = check('Trot',          'amazing two-beat rhythm');
  final.Canter       = check('Canter',        'very elegant looking canter');
  final.Gallop       = check('Gallop',        'balanced and very smooth gallop');
  final.Posture      = check('Posture',       'posture is perfectly balanced');
  final.Head         = check('Head',          'shows some nice proportions');
  final.Back         = check('Back',          'back will be heaven for a rider');
  final.Frontlegs    = check('Frontlegs',     'front legs are practically identical');
  final.Hindquarters = check('Hindquarters',  'great engagement in the hindquarters');
  final.Socks        = check('Socks',         'feathering is amazing');

  const pNS = text.includes('great elasticity in the shoulder');
  const n = base.Neck, s = base.Shoulders;
  if      (n === 'Good'      && s === 'Good')      final.Neck = pNS ? 'G+'  : 'G';
  else if (n === 'Good'      && s === 'Very good') final.Neck = pNS ? 'G'   : 'G-';
  else if (n === 'Very good' && s === 'Good')      final.Neck = pNS ? 'VG'  : 'VG';
  else final.Neck = shortenGradeOutside(n);

  if      (s === 'Good'      && n === 'Good')      final.Shoulders = pNS ? 'G+' : 'G';
  else if (s === 'Good'      && n === 'Very good') final.Shoulders = pNS ? 'G'  : 'G-';
  else final.Shoulders = shortenGradeOutside(s);

  final.Tolt       = shortenGradeOutside(base.Tolt);
  final.FlyingPace = shortenGradeOutside(base.FlyingPace);
  final.Revaal     = shortenGradeOutside(base.Revaal);

  return final;
}

function shortenGradeOutside(grade) {
  if (!grade) return '';
  const g = grade.toLowerCase();
  if (g.includes('very good'))     return 'VG';
  if (g.includes('good'))          return 'G';
  if (g.includes('average'))       return 'A';
  if (g.includes('below average')) return 'BA';
  if (g.includes('poor'))          return 'P';
  return grade;
}

/**
 * Saves the outside stud to the Outside Studs sheet.
 *
 * Layout after inserting Breed at D: B ID, C Name, D Breed; E HRC, F DP, G FT, H WT, I COI.
 * K:X conformation, AT:BC GP attributes, BD total GP, BF conformation maximum,
 * BV competition maximum, BX genetics. Column numbers below are the current, post-insertion positions.
 */

function saveOutsideStud(data, manualGenes) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Outside Studs');
    if (!sheet) throw new Error('Sheet "Outside Studs" not found!');
    if (!/^(breed|rasse)$/i.test(String(sheet.getRange(1,4).getValue()).trim())) throw new Error('Insert column D named Breed before importing.');

    const isIce = (data.breed || '').includes('Icelandic Horse');
    const isKath = (data.breed || '').includes('Kathiawari');

    const lastRow = sheet.getLastRow();
    let targetRow = -1;

    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0].toString().trim() === data.id.toString().trim()) {
          targetRow = i + 2;
          break;
        }
      }
      if (targetRow === -1) {
        for (let i = 0; i < ids.length; i++) {
          if (!ids[i][0] || ids[i][0].toString().trim() === '') {
            targetRow = i + 2;
            break;
          }
        }
      }
    }
    if (targetRow === -1) targetRow = lastRow + 1;

    const set = (col, val) => sheet.getRange(targetRow, col).setValue(val !== undefined && val !== null ? val : '');
    sheet.getRange(targetRow, 4).setValue(data.breed || '');

    // ── Basic info ──────────────────────────────────────────
    set(2, data.id);
    set(3, data.name);
    set(5, data.hcr || '');
    set(6, data.dp || '');
    set(7, data.ft || '');
    set(8, data.wt || '');
    set(9, data.coi || '');
    // J = Link — skip

    // ── Conformation ────────────────────────────────────────
    const c = data.confo || {};
    set(11, c.Walk || '');
    set(12, c.Trot || '');
    set(13, c.Canter || '');
    set(14, c.Gallop || '');
    set(15, isIce ? (c.Tolt || '') : '');
    set(16, isIce ? (c.FlyingPace || '') : isKath ? (c.Revaal || '') : '');
    set(17, c.Posture || '');
    set(18, c.Head || '');
    set(19, c.Neck || '');
    set(20, c.Back || '');
    set(21, c.Shoulders || '');
    set(22, c.Frontlegs || '');
    set(23, c.Hindquarters || '');
    set(24, c.Socks || '');

    // ── GP base stats (AT–BC) ───────────────────────────────
    const gp = data.gp || {};
    set(46, gp['Acceleration']   || '');
    set(47, gp['Agility']        || '');
    set(48, gp['Balance']        || '');
    set(49, gp['Bascule']        || '');
    set(50, gp['Pulling power']  || '');
    set(51, gp['Speed']          || '');
    set(52, gp['Sprint']         || '');
    set(53, gp['Stamina']        || '');
    set(54, gp['Strength']       || '');
    set(55, gp['Surefootedness'] || '');

    // BD(56) = Total GP
    const gpVals = Object.values(gp).map(v => parseFloat(v)).filter(v => !isNaN(v));
    const totalGP = gpVals.length > 0 ? Math.round(gpVals.reduce((a, b) => a + b, 0)) : '';
    set(56, totalGP);

    // ── Results & Genetics (YOUR CORRECTIONS) ───────────────────────
    set(58, data.confoMax    || '');    
    set(74, data.compMax     || '');    
    set(76, data.geneticCode || '');    

    updateStallionDropdown();

    return '✓ ' + data.name + ' saved to Outside Studs (row ' + targetRow + ')';

  } catch (e) {
    throw new Error('Save error: ' + e.message);
  }
}