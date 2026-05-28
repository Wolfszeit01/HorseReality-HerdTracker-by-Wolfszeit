/*** MASTER SCRIPT: UNIVERSAL IMPORT — COMMUNITY VERSION */

function openUniversalImportModal() {
  const html = HtmlService.createTemplateFromFile('exportModal').evaluate().setWidth(1000).setHeight(950);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

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
      if      (br.match(/Icelandic/i))  d.breed = 'Icelandic Horse';
      else if (br.match(/Kathiawari/i)) d.breed = 'Kathiawari';
      else if (br.match(/Finnhorse/i))  d.breed = 'Finnhorse';
      else                              d.breed = br;
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

  const pedSplit = info.split(/Pedigree/i);
  const pedStart = pedSplit.length > 1 ? pedSplit[1] : null;
  if (pedStart) {
    const pedLines = pedStart.split(/Pregnancy/i)[0].split(/\n/).map(l=>l.trim()).filter(l=>l.length>0&&!l.match(/^COI:/i));
    const isStatLine = (l) => {
      if (/^\d+\s*[\|\/]/.test(l)) return true; if (/\bGP\d{3,}\b/i.test(l)) return true;
      if (/\d+VG\b/.test(l)) return true; if (/\d+:\d+/.test(l)) return true;
      if (/\d{2,}\.\d{1,}/.test(l)) return true; if (/\d{3,}\s+\d+\s+\d/.test(l)) return true;
      if (/FULLYTRAINED|BETA|Training/i.test(l)) return true;
      if (/(Estate|Stables|Gardens|Meadows|Stuteri|Stable|Farm|Ranch|Stud|Park|Acres)\s*('s\s*\w+)?\s*$/i.test(l)) return true;
      return false;
    };
    const normalizeDecoName = (l) => l
      .replace(/[ℬℌℛℐℑℒℓ]/g,m=>({'ℬ':'B','ℌ':'H','ℛ':'R','ℐ':'I','ℑ':'I','ℒ':'L','ℓ':'l'}[m]||m))
      .replace(/[ᴀʙᴄᴅᴇꜰɢʜɪᴊᴋʟᴍɴᴏᴘǫʀꜱᴛᴜᴠᴡxʏᴢ]/g,c=>c.normalize('NFKD')[0]||c).replace(/[σ]/g,'s');
    const horseNameLines = pedLines.filter(l=>!isStatLine(l)).map(normalizeDecoName);
    if (horseNameLines.length >= 2) {
      d.sire = cleanHorseName(horseNameLines[0]);
      const damIdx = horseNameLines.length >= 8 ? 4 : Math.ceil(horseNameLines.length/2);
      d.dam  = cleanHorseName(horseNameLines[damIdx]);
      if (horseNameLines.length > 1)          d.gd_p = cleanHorseName(horseNameLines[1]);
      if (horseNameLines.length > 2)          d.gs_p = cleanHorseName(horseNameLines[2]);
      if (horseNameLines.length > damIdx + 1) d.gd_m = cleanHorseName(horseNameLines[damIdx+1]);
      if (horseNameLines.length > damIdx + 2) d.gs_m = cleanHorseName(horseNameLines[damIdx+2]);
    } else if (horseNameLines.length === 1) { d.sire = cleanHorseName(horseNameLines[0]); }
    d.sireExists = checkHorseNameExists(d.sire); d.damExists = checkHorseNameExists(d.dam);
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

  const geneKeys = ['Extension','Agouti','Grey','Creampearl','Dun','Champagne','Silver','Mushroom','Frame','Appaloosa','PATN1','MITF','SW2','KIT','WhiteSpotting'];
  geneKeys.forEach(gene=>{ d.genetics[gene]="n/n"; });
  const cleanG = (val) => val ? val.replace(/\s+/g,'').replace(/N$/i,'n') : "n/n";
  const mExt=colTxt.match(/Extension[\s\S]{0,80}?([Ee])\s*[\/\\]\s*([Ee])/i); if(mExt) d.genetics['Extension']=cleanG(mExt[1]+'/'+mExt[2]);
  const mAgo=colTxt.match(/(?:[Aa]\+|[Aa]t|[Aa]|aa)\s*[\/\\]\s*(?:[Aa]\+|[Aa]t|[Aa]|aa)/); if(mAgo) d.genetics['Agouti']=cleanG(mAgo[0]);
  const mGrey=colTxt.match(/Grey[\s\S]{0,80}?([Gg])\s*[\/\\]\s*([Gg]|n)/i); if(mGrey) d.genetics['Grey']=cleanG(mGrey[1]+'/'+mGrey[2]);
  const mSilv=colTxt.match(/Silver[\s\S]{0,80}?([Zz])\s*[\/\\]\s*([Zz])/i); if(mSilv) d.genetics['Silver']=cleanG(mSilv[1]+'/'+mSilv[2]);
  const mCr=colTxt.match(/Cream[\s\S]{0,80}?(Cr|prl)\s*[\/\\]\s*(Cr|prl|n)/i); if(mCr) d.genetics['Creampearl']=cleanG(mCr[1]+'/'+mCr[2]);
  const mDun=colTxt.match(/\bDun\b[\s\S]{0,80}?(nd[12]|[Dd])\s*[\/\\]\s*(nd[12]|[Dd])/i); if(mDun) d.genetics['Dun']=cleanG(mDun[1]+'/'+mDun[2]);
  const mChamp=colTxt.match(/Champagne[\s\S]{0,80}?(Ch)\s*[\/\\]\s*(Ch|n)/i); if(mChamp) d.genetics['Champagne']=cleanG(mChamp[1]+'/'+mChamp[2]);
  const mMush=colTxt.match(/Mushroom[\s\S]{0,80}?(mu|[Mm])\s*[\/\\]\s*(mu|[Mm]|n)/i); if(mMush) d.genetics['Mushroom']=cleanG(mMush[1]+'/'+mMush[2]);
  const mFrame=colTxt.match(/(?:\bFrame\b|\bOWL\b)[\s\S]{0,80}?(OWL|[Oo])\s*[\/\\]\s*(OWL|[Oo]|n)/i); if(mFrame) d.genetics['Frame']=cleanG(mFrame[1]+'/'+mFrame[2]);
  const mLp=colTxt.match(/(?:Appaloosa|\bLp\b)[\s\S]{0,80}?(Lp)\s*[\/\\]\s*(Lp|n)/i); if(mLp) d.genetics['Appaloosa']=cleanG(mLp[1]+'/'+mLp[2]);
  const mPatn1=colTxt.match(/PATN1[\s\S]{0,80}?(PATN1)\s*[\/\\]\s*(PATN1|n)/i); if(mPatn1) d.genetics['PATN1']=cleanG(mPatn1[1]+'/'+mPatn1[2]);
  const mMitf=colTxt.match(/(?:SW[13]|MITF)[\s\S]{0,80}?(SW[13])\s*[\/\\]\s*(SW[13]|n)/i); if(mMitf) d.genetics['MITF']=cleanG(mMitf[1]+'/'+mMitf[2]);
  const mSw2=colTxt.match(/SW2[\s\S]{0,80}?(SW2)\s*[\/\\]\s*(SW2|n)/i); if(mSw2) d.genetics['SW2']=cleanG(mSw2[1]+'/'+mSw2[2]);
  const kitPat='W21|W20|W19|W16|W10|W8|W3|SB1|Rn|TO'; const kitKeywords='Tobiano|Roan|Sabino|W21|W20|W19|W16|W10|W8|W3|\\bKIT\\b';
  const mKit=colTxt.match(new RegExp('(?:'+kitKeywords+')[\\s\\S]{0,80}?('+kitPat+')\\s*[/\\\\]\\s*('+kitPat+'|n)','i')); if(mKit) d.genetics['KIT']=cleanG(mKit[1]+'/'+mKit[2]);
  const mWs=colTxt.match(/(?:White\s*Spotting|\bWS\b)[\s\S]{0,80}?(WS)\s*[\/\\]\s*(WS|n)/i); if(mWs) d.genetics['WhiteSpotting']=cleanG(mWs[1]+'/'+mWs[2]);

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

  let showPart = statTxt.split(/show\s+results/i)[1]||""; showPart=showPart.split(/competition/i)[0];
  d.showScores = [...new Set(showPart.match(/\b\d{1,3}[.,]\d{3}\b/g)||[])];
  const compParts = statTxt.split(/competition\s+results/i);
  if (compParts.length>1) {
    d.compScores = [...new Set((compParts[1].split(/Health|Genetic/i)[0].match(/\b\d{1,3}[.,]\d{3,4}\b/g)||[]))]
      .filter(s=>{const v=parseFloat(s.replace(',','.'));return v>10&&v<150;}).slice(0,25);
  }

  d.allHorses = getDropdownNames(ss);
  return d;
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
  if (g.includes("very good")) return "VG"; if (g.includes("good")) return "G";
  if (g.includes("average")) return "A"; if (g.includes("below average")) return "BA";
  if (g.includes("poor")) return "P"; return grade;
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
  let n = name.replace(/Deceased/gi,"").replace(/^[★☆ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃✶〔〕♟˚₊✧⋆°·•\-\s]+/,"").trim();
  n = n.replace(/\s+\d+[\.\,]\d{3,}\s*$/,'').trim();
  n = n.replace(/\s*\|\|.*$/,'').trim();
  n = n.split(/Tagline|\|/i)[0];
  n = n.replace(/[★☆ꋖꂵꀷ!❦↟ℛᨒ⤈⚜️⭃✶〔〕♟˚₊✧⋆°·•\s]+$/,"").trim();
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

function processFinalImport(data, opts, manual) {
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet(); const manualData=manual||{};
    function convertDOB(dob) {
      if(!dob) return "";
      if(dob.match(/^\d{2}\/\d{2}\/\d{4}$/)) { const p=dob.split('/'); return p[2]+'-'+p[0]+'-'+p[1]; }
      if(dob.match(/^\d{2}-\d{2}-\d{4}$/)) { const p=dob.split('-'); return p[2]+'-'+p[1]+'-'+p[0]; }
      return dob;
    }
    function convertDueDate(raw) {
      if(!raw) return "";
      const mm={jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'};
      const cleaned=raw.replace(/^[A-Za-z]+,\s*/,"").trim(); const parts=cleaned.split(/[\s,]+/);
      if(parts.length>=2) { return (mm[parts[0].toLowerCase().slice(0,3)]||'01')+'/'+parts[1].padStart(2,'0')+'/'+(parts[2]||new Date().getFullYear().toString()); }
      return raw;
    }
    const convertedDOB=convertDOB(data.dob); const results=[]; const sireIsInside=data.sireIsInside||false; const damIsInside=data.damIsInside||false;

    if (opts.herd) {
      const herdSheet=getHerdsheet();
      if (herdSheet) {
        const row=findRow(herdSheet,data.id,3);
        herdSheet.getRange(row,3).setValue(data.id); herdSheet.getRange(row,4).setValue(data.name);
        herdSheet.getRange(row,5).setValue(data.breed||""); herdSheet.getRange(row,6).setValue(data.gender||"");
        herdSheet.getRange(row,7).setValue(convertedDOB);
        herdSheet.getRange(row,17).setValue((data.predicates&&data.predicates.toLowerCase()!=="none")?data.predicates:"");
        herdSheet.getRange(row,30,1,6).setValues([[data.health['Colic resistance']||"",data.health['Hoof quality']||"",data.health['Back problems']||"",data.health['Respiratory disease']||"",data.health['Resistance to lameness']||"",data.health['Fertility']||""]]);
        results.push("✓ Herd Tracker updated");
      }
    }

    if (opts.stats) {
      const breed=(data.breed||"").trim(); const isIce=breed.includes('Icelandic Horse'); const isKath=breed.includes('Kathiawari');
      const targetSheetName=isIce?'ICE_Horse Stats':(isKath?'KATH_Horse Stats':'Horse Stats');
      const statSheet=ss.getSheetByName(targetSheetName);
      if (statSheet) {
        const row=findRow(statSheet,data.id,2); statSheet.getRange(row,2).setValue(data.id); statSheet.getRange(row,3).setValue(data.name);
        let confo=[data.confo.Walk||"",data.confo.Trot||"",data.confo.Canter||"",data.confo.Gallop||""];
        if(isIce) confo.push(data.confo.Tolt||"",data.confo.Pace||""); else if(isKath) confo.push(data.confo.Revaal||"");
        confo.push(data.confo.Posture||"",data.confo.Head||"",data.confo.Neck||"",data.confo.Back||"",data.confo.Shoulders||"",data.confo.Frontlegs||"",data.confo.Hindquarters||"",data.confo.Socks||"");
        statSheet.getRange(row,4,1,confo.length).setValues([confo]);
        const gpCol=isIce?37:(isKath?36:35);
        statSheet.getRange(row,gpCol,1,10).setValues([[data.gp['Acceleration']||"",data.gp['Agility']||"",data.gp['Balance']||"",data.gp['Bascule']||"",data.gp['Pulling power']||"",data.gp['Speed']||"",data.gp['Sprint']||"",data.gp['Stamina']||"",data.gp['Strength']||"",data.gp['Surefootedness']||""]]);
        const achMap=isIce?{'Day Champion':69,'1st Premium':70,'2nd Premium':71,'3rd Premium':72,'1st Prize':76,'2nd Prize':77,'3rd Prize':78}:isKath?{'Day Champion':67,'1st Premium':68,'2nd Premium':69,'3rd Premium':70,'1st Prize':74,'2nd Prize':75,'3rd Prize':76}:{'Day Champion':65,'1st Premium':66,'2nd Premium':67,'3rd Premium':68,'1st Prize':72,'2nd Prize':73,'3rd Prize':74};
        Object.entries(achMap).forEach(([key,col])=>statSheet.getRange(row,col).setValue(data.achieve[key]||"0"));
        results.push("✓ "+targetSheetName+" updated");
      }
    }

    if (opts.colour) {
      const colSheet=ss.getSheetByName('Colour Genetics');
      if (colSheet) {
        const row=findRow(colSheet,data.id,2); const kitVal=data.genetics['KIT']||"n/n";
        const kitA1=(kitVal.split('/')[0]||'').trim(); const kitA2=(kitVal.split('/')[1]||'').trim();
        const getKitAllele=(allele)=>{const a=allele.toLowerCase();const c=(kitA1.toLowerCase()===a?1:0)+(kitA2.toLowerCase()===a?1:0);if(c===2)return allele+'/'+allele;if(c===1)return allele+'/n';return 'n/n';};
        const wsValue=(data.genetics['WhiteSpotting']&&data.genetics['WhiteSpotting']!=='n/n')?data.genetics['WhiteSpotting']:(manualData['HiddenSabino']&&manualData['HiddenSabino'].trim()!==''?manualData['HiddenSabino'].trim():"n/n");
        colSheet.getRange(row,2).setValue(data.id); colSheet.getRange(row,3).setValue(data.name);
        colSheet.getRange(row,4).setValue(data.genetics['Extension']||"n/n"); colSheet.getRange(row,5).setValue(data.genetics['Agouti']||"n/n");
        colSheet.getRange(row,6).setValue(data.genetics['Grey']||"n/n"); colSheet.getRange(row,7).setValue(data.genetics['Creampearl']||"n/n");
        colSheet.getRange(row,8).setValue(data.genetics['Dun']||"n/n"); colSheet.getRange(row,9).setValue(data.genetics['Champagne']||"n/n");
        colSheet.getRange(row,10).setValue(data.genetics['Silver']||"n/n"); colSheet.getRange(row,11).setValue(data.genetics['Mushroom']||"n/n");
        colSheet.getRange(row,12).setValue(data.genetics['Frame']||"n/n"); colSheet.getRange(row,13).setValue(data.genetics['Appaloosa']||"n/n");
        colSheet.getRange(row,14).setValue(data.genetics['PATN1']||"n/n"); colSheet.getRange(row,15).setValue(manualData['PATN2']||"n/n");
        colSheet.getRange(row,16).setValue(data.genetics['MITF']||"n/n"); colSheet.getRange(row,17).setValue(data.genetics['SW2']||"n/n");
        colSheet.getRange(row,22).setValue(getKitAllele('TO')); colSheet.getRange(row,23).setValue(getKitAllele('Rn'));
        colSheet.getRange(row,24).setValue(getKitAllele('SB1')); colSheet.getRange(row,27).setValue(getKitAllele('W3'));
        colSheet.getRange(row,32).setValue(getKitAllele('W8')); colSheet.getRange(row,34).setValue(getKitAllele('W10'));
        colSheet.getRange(row,40).setValue(getKitAllele('W16')); colSheet.getRange(row,43).setValue(getKitAllele('W19'));
        colSheet.getRange(row,44).setValue(getKitAllele('W20')); colSheet.getRange(row,45).setValue(getKitAllele('W21'));
        colSheet.getRange(row,48).setValue(manualData['Flaxen']||"n/n"); colSheet.getRange(row,49).setValue(manualData['Sooty']||"n/n");
        colSheet.getRange(row,50).setValue(manualData['Pangare']||"n/n"); colSheet.getRange(row,51).setValue(wsValue);
        colSheet.getRange(row,52).setValue(manualData['Rabicano']||"n/n"); colSheet.getRange(row,53).setValue(manualData['Markings']||"");
        results.push("✓ Colour Genetics updated");
      }
    }

    if (opts.scores) {
      ['Conf. Results','Comp. Results'].forEach(sName=>{
        const s=ss.getSheetByName(sName); const scores=(sName==='Conf. Results')?data.showScores:data.compScores;
        if(s&&scores&&scores.length>0){let col=1;while(s.getRange(1,col).getValue()){col++;}s.getRange(1,col).setValue(data.id);s.getRange(2,col).setValue(data.name);scores.slice(0,25).forEach((val,i)=>{const num=parseFloat(val.toString().replace(',','.'));if(!isNaN(num))s.getRange(6+i,col).setValue(num);});results.push("✓ "+sName+" updated");}
      });
    }

    const pedSheet=ss.getSheetByName('Pedigree');
    if (pedSheet) {
      const row=findRow(pedSheet,data.id,2); pedSheet.getRange(row,2).setValue(data.id); pedSheet.getRange(row,3).setValue(data.name);
      pedSheet.getRange(row,6).setValue(data.sire||""); pedSheet.getRange(row,7).setValue(data.dam||"");
      if(!sireIsInside){if(data.gs_p)pedSheet.getRange(row,8).setValue(data.gs_p);if(data.gd_p)pedSheet.getRange(row,9).setValue(data.gd_p);}
      if(!damIsInside){if(data.gs_m)pedSheet.getRange(row,10).setValue(data.gs_m);if(data.gd_m)pedSheet.getRange(row,11).setValue(data.gd_m);}
      if(damIsInside) copyParentPedigree(pedSheet,row,data.dam,"dam");
      if(sireIsInside) copyParentPedigree(pedSheet,row,data.sire,"sire");
      results.push("✓ Pedigree updated");
    }

    if (data.pregnancy&&(data.pregnancy.foalGender||data.pregnancy.dueDate)) {
      const broodSheet=ss.getSheetByName('Broodmares');
      if(broodSheet){
        const row=findRow(broodSheet,data.id,3); broodSheet.getRange(row,3).setValue(data.id); broodSheet.getRange(row,4).setValue(data.name); broodSheet.getRange(row,5).setValue(data.breed||"");
        if(data.pregnancy.coveredDate) broodSheet.getRange(row,6).setValue(data.pregnancy.coveredDate);
        if(data.pregnancy.coveredTime) broodSheet.getRange(row,7).setValue(data.pregnancy.coveredTime);
        broodSheet.getRange(row,12).setValue(data.pregnancy.foalGender||"");
        if(data.pregnancy.sire) broodSheet.getRange(row,checkHorseNameExists(data.pregnancy.sire)?13:15).setValue(data.pregnancy.sire);
        if(data.pregnancy.dueDate) broodSheet.getRange(row,14).setValue(convertDueDate(data.pregnancy.dueDate));
        results.push("✓ Broodmares updated");
      }
    }

    if(!damIsInside||!sireIsInside) openPedigreeModalWithData({id:data.id,name:data.name,damName:data.dam,sireName:data.sire});
    return {message:"✓ Import complete!\n"+results.join("\n"),openPedigreeModal:(!damIsInside||!sireIsInside),pedigreeData:{id:data.id,name:data.name,damName:data.dam,sireName:data.sire}};
  } catch(e) { throw new Error("Import error: "+e.message); }
}

function copyParentPedigree(sheet,targetRow,parentName,type) {
  if(!parentName||parentName==="Unknown") return;
  const fullData=sheet.getDataRange().getValues(); let pIdx=-1;
  for(let i=0;i<fullData.length;i++){if(fullData[i][2]===parentName){pIdx=i+1;break;}}
  if(pIdx===-1) return;
  if(type==="dam"){sheet.getRange(targetRow,10).setValue(sheet.getRange(pIdx,6).getValue());sheet.getRange(targetRow,11).setValue(sheet.getRange(pIdx,7).getValue());sheet.getRange(targetRow,16).setValue(sheet.getRange(pIdx,8).getValue());sheet.getRange(targetRow,17).setValue(sheet.getRange(pIdx,9).getValue());sheet.getRange(targetRow,18).setValue(sheet.getRange(pIdx,10).getValue());sheet.getRange(targetRow,19).setValue(sheet.getRange(pIdx,11).getValue());}
  else{sheet.getRange(targetRow,8).setValue(sheet.getRange(pIdx,6).getValue());sheet.getRange(targetRow,9).setValue(sheet.getRange(pIdx,7).getValue());sheet.getRange(targetRow,12).setValue(sheet.getRange(pIdx,8).getValue());sheet.getRange(targetRow,13).setValue(sheet.getRange(pIdx,9).getValue());sheet.getRange(targetRow,14).setValue(sheet.getRange(pIdx,10).getValue());sheet.getRange(targetRow,15).setValue(sheet.getRange(pIdx,11).getValue());}
}

function openPedigreeModalWithData(pedigreeData) {
  PropertiesService.getScriptProperties().setProperty('pendingPedigreeData',JSON.stringify(pedigreeData));
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('pedigreeModal').evaluate().setWidth(600).setHeight(750),'Complete Pedigree');
}

function getPedigreeModalData() {
  try{const props=PropertiesService.getScriptProperties();const pendingData=props.getProperty('pendingPedigreeData');if(!pendingData) return {success:true,outsideDam:true,outsideSire:true};let data=JSON.parse(pendingData);props.deleteProperty('pendingPedigreeData');data.success=true;return data;}
  catch(e){return {success:false,error:e.message};}
}
