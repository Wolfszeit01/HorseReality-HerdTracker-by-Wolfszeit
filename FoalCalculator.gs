// ==========================================
// NORDRISKA BREEDING OPTIMIZER BACKEND — v2.0 FINAL
// ALLE FUNKTIONEN IN EINER DATEI
// ==========================================

function openBreedingOptimizer() {
  const html = HtmlService.createTemplateFromFile('BreedingOptimizer')
    .evaluate()
    .setTitle('Herd Tracker Foal Calculator')
    .setWidth(650);
  
  SpreadsheetApp.getUi().showSidebar(html);
}

function clearCacheAndReload() {
  Browser.msgBox('✓ Data cache cleared. Please reopen the Breeding Optimizer sidebar.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Data refreshed!', 'Herd Tracker');
}

function normalizeIdForComparison(id) {
  if (!id) return "";
  return String(id).trim().toLowerCase().replace(/\s+/g, '').replace(/[★☆]/g, '');
}

// Three-generation Wright path estimate, including stored COI of each shared ancestor.
// Reference: https://www.fao.org/4/x3840e/X3840E03.htm
function batchCheckInbreeding(mareId,stallionIds) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const read=n=>{
    const sh=ss.getSheetByName(n);if(!sh)return [];
    const rows=sh.getDataRange().getValues();
    const coiCol=n==='Pedigree'?4:n==='Outside Studs'?9:null;
    // Display values disambiguate 0.125 formatted as 12.5% from a plain 0.125 percent-point value.
    if(coiCol && rows.length){const displayed=sh.getRange(1,coiCol,rows.length,1).getDisplayValues();rows.forEach((r,i)=>r[coiCol-1]=displayed[i][0]);}
    return rows;
  };
  const herd=read('Herd'), outside=read('Outside Studs'), pedigree=read('Pedigree');
  return fcCheckPedigrees_(mareId,stallionIds,herd,outside,pedigree);
}
function fcCheckPedigrees_(mareId,stallionIds,herd,outside,pedigree) {
  const alias=new Map(),names=new Map(),graph=new Map(),storedCoi=new Map();
  const valid=v=>v && !/^(unknown|n\/a|[-–—]|foundation(?: breeder)?)$/i.test(String(v).trim());
  const register=(id,name)=>{id=normalizeIdForComparison(id);if(!id)return;alias.set(id,id);names.set(id,String(name||id));
    const n=normalizeIdForComparison(name);if(n)alias.set(n,alias.has(n)&&alias.get(n)!==id?null:id);};
  herd.slice(1).forEach(r=>register(r[2],r[3]));outside.slice(1).forEach(r=>register(r[1],r[2]));pedigree.slice(1).forEach(r=>register(r[1],r[2]));
  const resolve=v=>{if(!valid(v))return null;const n=normalizeIdForComparison(v);return alias.has(n)?alias.get(n):n;};
  const rememberCoi=(r,index,source)=>{
    const id=resolve(r[1]),fraction=fcParseStoredCoi_(r[index]);
    if(id && fraction!==null)storedCoi.set(id,{fraction,source});
  };
  // Pedigree is authoritative when both sources have a valid entry for the same ID.
  outside.slice(1).forEach(r=>rememberCoi(r,8,'Outside Studs · I'));
  pedigree.slice(1).forEach(r=>rememberCoi(r,3,'Pedigree · D'));
  pedigree.slice(1).forEach(r=>{
    const id=resolve(r[1]) || resolve(r[2]);if(!id)return;
    const nodes=[id,...r.slice(6,20).map(resolve)];
    for(let i=0;i<7;i++){
      const node=nodes[i];if(!node)continue;const parents=[nodes[2*i+1]||null,nodes[2*i+2]||null];
      if(!graph.has(node))graph.set(node,parents);
      else graph.set(node,graph.get(node).map((v,j)=>v||parents[j]));
    }
  });
  const paths=id=>{
    const found=new Map();
    function walk(node,path){if(!node)return;if(path.includes(node))throw new Error('Cycle in pedigree');
      const next=path.concat(node);if(!found.has(node))found.set(node,[]);found.get(node).push(next);
      if(path.length<3)(graph.get(node)||[]).forEach(parent=>walk(parent,next));}
    walk(id,[]);return found;
  };
  const results={};const mare=resolve(mareId);
  for(const rawId of stallionIds){
    const id=resolve(rawId), empty={isRelated:false,coi:null,coiRisk:'no_data',commonAncestors:[],reason:'Incomplete pedigree; COI cannot be estimated.'};
    try {
      if(!mare||!id){results[rawId]=empty;continue;}
      const mp=paths(mare),sp=paths(id);let total=0;const common=[];
      for(const [ancestor,ml] of mp){if(!sp.has(ancestor))continue;let contribution=0;
        for(const m of ml)for(const f of sp.get(ancestor)){
          if(m.slice(0,-1).some(n=>f.slice(0,-1).includes(n)))continue;
          contribution+=Math.pow(.5,m.length+f.length-1)*(1+(storedCoi.get(ancestor)?.fraction ?? 0));
        }
        if(contribution){total+=contribution;common.push({id:ancestor,name:names.get(ancestor)||ancestor,generationsFromMare:Math.min(...ml.map(p=>p.length-1)),generationsFromStallion:Math.min(...sp.get(ancestor).map(p=>p.length-1)),contribution:contribution*100,ancestorCoiPercent:storedCoi.has(ancestor)?storedCoi.get(ancestor).fraction*100:null,ancestorCoiSource:storedCoi.get(ancestor)?.source || null});}
      }
      const direct=mp.has(id)||sp.has(mare)||(graph.get(mare)||[]).some(p=>p&&(graph.get(id)||[]).includes(p));
      const hasParents=(graph.get(mare)||[]).filter(Boolean).length===2 && (graph.get(id)||[]).filter(Boolean).length===2;
      if(!hasParents && !total){results[rawId]=empty;continue;}
      const coi=total*100;
      results[rawId]={isRelated:direct,coi,coiRisk:getCoiRiskLevel(coi).level,commonAncestors:common,reason:'Three-generation estimate including stored common-ancestor COI; unknown ancestor COI is assumed zero.'};
    }catch(e){results[rawId]={...empty,coiRisk:'error',reason:e.message};}
  }return results;
}
// Input is a displayed percentage or a plain value in percentage points. Never guess by magnitude.
function fcParseStoredCoi_(value) {
  if(value==null || typeof value==='boolean')return null;
  const text=String(value).trim().replace(/\s/g,'');
  if(!/^\d+(?:[.,]\d+)?%?$/.test(text))return null;
  const percent=Number(text.replace('%','').replace(',','.'));
  return Number.isFinite(percent)&&percent>=0&&percent<=100?percent/100:null;
}

function getCoiRiskLevel(coi) {
  return {level:coi<2.5?'low':coi<6.25?'medium':coi<12.5?'high':'very-high'};
}

function loadAllDataOptimized() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const read=(...names)=> {const sh=names.map(n=>ss.getSheetByName(n)).find(Boolean); return sh ? sh.getDataRange().getValues() : [];};
  const output={tracker:{},stats:{},genetics:{},pedigree:{},stallionUsage:{},warnings:[]};
  const herd=read('Herd');
  herd.slice(1).forEach(r=>{
    const id=String(r[2] || '').trim(); if (!id) return;
    output.tracker[id]={name:String(r[3] || '').trim(),breed:String(r[4] || '').trim(),gender:normalizeGender(r[5]),age:fcParseAge_(r[8]),type:'owned'};
  });
  const process=(rows,extra)=>rows.slice(1).forEach(r=>{
    const id=String(r[1] || '').trim(); if (!id || !output.tracker[id] || output.stats[id]) return;
    const gp=Array.from({length:10},(_,i)=>fcNumber_(r[34+2*extra+i]));
    const grades=r.slice(3,15+extra).map(fcGrade_);
    const numeric=r.slice(15+extra,27+2*extra).map(fcNumber_);
    const confo=grades.map((g,i)=>Number.isFinite(numeric[i]) ? numeric[i] : g);
    if (gp.some(v=>v===null) || confo.some(v=>v===null)) {output.warnings.push('Incomplete stats: '+output.tracker[id].name);return;}
    const maxCol=fcColumn_(rows, ['conf max','confo max','conformation max','max confo score'],46+2*extra);
    const rangeCol=fcColumn_(rows,['conf range','confo range','conformation range'],47+2*extra);
    const max=fcNumber_(r[maxCol]);const range=String(r[rangeCol] || '').match(/(\d+(?:[.,]\d+)?)\s*[-–—]\s*(\d+(?:[.,]\d+)?)/);
    const avg=confo.reduce((a,b)=>a+b,0)/confo.length;
    output.stats[id]=fcStats_(gp,confo,range ? fcNumber_(range[1]) : (max ?? avg),max ?? avg);
  });
  process(read('KATH_Stats','KATH_Horse Stats'),1);
  process(read('ICE_Stats','ICE_Horse Stats'),2);
  process(read('Stats','Horse Stats'),0);
  const outside=read('Outside Studs','Public Studs');
  const hasBreed=outside.length && /^(breed|rasse)$/i.test(String(outside[0][3] || '').trim());
  if(outside.length>1 && !hasBreed) output.warnings.push('Outside Studs: column D must be named Breed; external candidates are excluded until this is fixed.');
  if(hasBreed) outside.slice(1).forEach(r=>{
    const id=String(r[1] || '').trim(); if(!id || output.tracker[id])return;
    const breed=String(r[3] || '').trim();
    if(!breed){output.warnings.push('Outside stud without breed: '+String(r[2] || id));return;}
    const extra=/icelandic/i.test(breed)?2:/kathiawari/i.test(breed)?1:0;
    const indices=[10,11,12,13].concat(extra===2?[14,15]:extra===1?[15]:[],[16,17,18,19,20,21,22,23]);
    const confo=indices.map(i=>fcGrade_(r[i]));const gp=Array.from({length:10},(_,i)=>fcNumber_(r[45+i]));
    output.tracker[id]={name:String(r[2] || id),breed,gender:'Stallion',age:3,type:'public',studFee:fcNumber_(r[4])};
    if(gp.every(v=>v!==null)&&confo.every(v=>v!==null)) {
      const max=fcNumber_(r[57]) ?? confo.reduce((a,b)=>a+b,0)/confo.length;
      output.stats[id]=fcStats_(gp,confo,max,max);
    } else output.warnings.push('Incomplete outside stud stats: '+output.tracker[id].name);
    const code=String(r[75] || '').trim();if(code)output.genetics[id]=code;
  });
  read('Colour Genetics').slice(1).forEach(r=>{const id=String(r[1] || '').trim(),code=String(r[79] || '').trim();if(id&&code)output.genetics[id]=code;});
  read('Broodmares').slice(1).forEach(r=>{const n=String(r[12] || '').trim();if(n)output.stallionUsage[n]=(output.stallionUsage[n] || 0)+1;});
  output.hrCoi=fcLoadHrCoi_();
  return output;
}

function fcNumber_(v) {
  if(v==null || String(v).trim()==='')return null;
  const n=Number(String(v).trim().replace(',','.'));return Number.isFinite(n)?n:null;
}
function fcGrade_(v) {
  const n=fcNumber_(v); if(n!==null)return n;
  return ({P:19.5,BA:49.5,A:64.5,'G-':74.5,G:76.5,'G+':82,VG:92.5,'Very good':92.5,Good:76.5,Average:64.5})[String(v).trim()] ?? null;
}
function fcColumn_(rows,names,fallback) {
  const normal=s=>String(s || '').toLowerCase().replace(/[^a-z0-9]/g,'');
  for(const row of rows.slice(0,2)){const i=row.findIndex(v=>names.some(n=>normal(n)===normal(v)));if(i>=0)return i;}
  return fallback;
}
function fcStats_(gp,confo,min,max) {
  const keys=['acceleration','agility','balance','bascule','pulling','speed','sprint','stamina','strength','surefoot'];
  const gpStats=Object.fromEntries(keys.map((k,i)=>[k,gp[i]]));
  const disciplines={dressage:[1,2,8],driving:[1,4,5,7,8],endurance:[5,7,8,9],eventing:[2,3,5,8,9],flat:[5,0,7,6],jumping:[0,1,3,6,8],reining:[0,1,2,9]};
  return {gpTotal:gp.reduce((a,b)=>a+b,0),gpStats,gpDisciplines:Object.fromEntries(Object.entries(disciplines).map(([k,idx])=>[k,idx.reduce((a,i)=>a+gp[i],0)])),confoRaw:confo,confoCount:confo.length,confoScoreMin:min,confoScoreMax:max};
}

function calculateFoalColors(mareId, studId) {
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const colour=ss.getSheetByName('Colour Genetics');
    const cache={};
    const m=getGeneticsForId(ss,mareId,colour,cache), f=getGeneticsForId(ss,studId,colour,cache);
    if (!m || !f) return {error:'Missing stored genetics for one or both parents.'};
    const result=fcColourDistribution_(m,f);
    let html='<div class="colour-assumption">Based on stored genotypes. Unlisted modifiers are treated as absent; hidden genes and markings can change the appearance.</div>';
    Object.entries(result).sort((a,b)=>b[1].probability-a[1].probability).forEach(([name,r])=>{
      html+='<p><b>'+(r.probability*100).toFixed(2)+'%</b> '+fcEscape_(name)+'</p>';
      html+='<small>Genetics: '+fcEscape_(r.genotypes.slice(0,2).join(' or '))+'</small>';
      if(r.genotypes.length>2)html+='<details class="colour-genotypes" data-mare="'+fcEscape_(mareId)+'" data-stud="'+fcEscape_(studId)+'" data-phenotype="'+fcEscape_(name)+'" data-offset="0" ontoggle="if(this.open &amp;&amp; !this.dataset.loaded) loadColourGenotypes(this)"><summary>Show all '+r.genotypes.length+' genotypes</summary><div class="genotype-list"></div><button type="button" class="btn btn-purple genotype-more" onclick="loadColourGenotypes(this.parentElement)">Load genotypes</button></details>';
    });
    return html;
  } catch(e) { return {error:e.message}; }
}

// Paged details avoid sending a potentially huge genotype list to the sidebar at once.
function getFoalColourGenotypes(mareId,studId,phenotype,offset) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),colour=ss.getSheetByName('Colour Genetics'),cache={};
  const m=getGeneticsForId(ss,mareId,colour,cache),f=getGeneticsForId(ss,studId,colour,cache);
  if(!m||!f)throw new Error('Missing stored genetics for one or both parents.');
  const distribution=fcColourDistribution_(m,f),group=distribution[String(phenotype)];
  if(!group)throw new Error('This colour result has changed. Reopen the calculator to refresh it.');
  const start=Math.max(0,Math.floor(Number(offset)||0)),size=50;
  return {genotypes:group.genotypes.slice(start,start+size),total:group.genotypes.length,nextOffset:Math.min(start+size,group.genotypes.length)};
}

function fcEscape_(value) { return String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

function parseGenetics(genString) {
  const genes = {};
  const kit = [];
  const pairs = String(genString || '').replace(/\s*\/\s*/g, '/').split(/[\s,;|]+/).filter(Boolean);
  pairs.forEach(token => {
    if (!token.includes('/')) throw new Error('Unrecognised genetics token: ' + token);
    const parts = token.split('/');
    if (parts.length !== 2) throw new Error('Invalid allele pair: ' + token);
    const loci = parts.map(identifyAllele).filter(Boolean);
    if (!loci.length) {
      if (parts.every(a => a === 'n' || a === 'N')) return;
      throw new Error('Unknown allele pair: ' + token);
    }
    const gene = loci[0];
    if (loci.some(g => g !== gene)) throw new Error('Mixed loci: ' + token);
    if (parts.some(a => !/^(n|N)$/.test(a) && !identifyAllele(a))) throw new Error('Unknown allele in: ' + token);
    const pair = parts.map(a => fcCanonicalAllele_(a, gene));
    if (gene === 'KIT') { kit.push(pair); return; }
    if (genes[gene]) throw new Error('Duplicate locus: ' + gene);
    genes[gene] = pair;
  });
  if (kit.length) {
    const active = kit.flat().filter(a => a !== 'n');
    if (active.length > 2) throw new Error('More than two KIT alleles; check the stored genetics code.');
    genes.KIT = active.concat(['n', 'n']).slice(0, 2);
  }
  return genes;
}

function fcCanonicalAllele_(a, gene) {
  if (a === 'n' || a === 'N') return 'n';
  const aliases = {Cr:'CR', Prl:'prl', Ch:'CH', Lp:'LP', OLW:'OLW', OWL:'OLW', O:'OLW', o:'n', owl:'n', olw:'n',
    To:'TO', to:'n', Rn:'RN', rn:'n', Sb:'SB1', SB:'SB1', sb:'n', sb1:'n', Sw:'SW1', SW:'SW1', sw:'n', sw1:'n', sw2:'n', sw3:'n',
    nz:'n', ng:'n', nch:'n', nmu:'n', nf:'n', nsty:'n', npa:'n', nws:'n',
    fl:'f', Rab:'rab', Ws:'WS', PANG:'PA'};
  return Object.prototype.hasOwnProperty.call(aliases,a) ? aliases[a] : a;
}

function identifyAllele(a) {
  if (/^(n|nn)$/i.test(a)) return null;
  if (/^[Ee]$/.test(a)) return 'E';
  if (/^(A\+?|At|a)$/.test(a)) return 'A';
  if (/^(CR|Cr|prl|Prl)$/.test(a)) return 'CR_Locus';
  if (/^(D|d|nd1|nd2)$/.test(a)) return 'DUN';
  if (/^(CH|Ch|ch|nch)$/.test(a)) return 'CH';
  if (/^(Z|z|nz)$/.test(a)) return 'Z';
  if (/^(Mu|mu|nmu)$/.test(a)) return 'MU';
  if (/^(G|g|ng)$/.test(a)) return 'G';
  if (/^(F|f|fl|nf)$/.test(a)) return 'F';
  if (/^(Sty|sty|nsty)$/.test(a)) return 'STY';
  if (/^(PA|pa|PANG|npa)$/.test(a)) return 'PA';
  if (/^(SW2|sw2)$/.test(a)) return 'SW2';
  if (/^(SW[13]?|Sw|sw[13]?)$/.test(a)) return 'MITF';
  if (/^(TO|To|to|RN|Rn|rn|SB1?|Sb|sb1?|W\d+)$/.test(a)) return 'KIT';
  if (/^(LP|Lp|lp)$/.test(a)) return 'LP';
  if (/^(PATN1|patn1)$/.test(a)) return 'PATN';
  if (/^(PATN2|patn2)$/.test(a)) return 'PATN2';
  if (/^(OLW|OWL|olw|owl|O|o)$/.test(a)) return 'OWL';
  if (/^(Rab|rab)$/.test(a)) return 'RAB';
  if (/^(WS|Ws|ws|nws)$/.test(a)) return 'WS';
  return null;
}

function calculateAllCombinations(mare, stud) {
  const names = [...new Set([...Object.keys(mare), ...Object.keys(stud)])];
  let combos = [{__probability: 1}];
  names.forEach(gene => {
    const m = mare[gene] || ['n','n'], f = stud[gene] || ['n','n'];
    const outcomes = new Map();
    for (const a of m) for (const b of f) {
      const pair = [a,b].sort(); const key = pair.join('/');
      if (!outcomes.has(key)) outcomes.set(key, {pair, probability: 0});
      outcomes.get(key).probability += .25;
    }
    if (combos.length * outcomes.size > 200000) throw new Error('Too many distinct genotypes for an exact calculation.');
    const next = [];
    for (const combo of combos) for (const outcome of outcomes.values()) {
      next.push({...combo, [gene]:outcome.pair, __probability:combo.__probability*outcome.probability});
    }
    combos = next;
  });
  return combos;
}

function fcColourDistribution_(mareCode, studCode) {
  const mare=parseGenetics(mareCode), stud=parseGenetics(studCode);
  for (const parent of [mare,stud]) {
    if (!parent.E || parent.E.some(a => !['E','e'].includes(a))) throw new Error('Missing or invalid Extension genotype.');
    if (parent.E.includes('E') && (!parent.A || parent.A.some(a => !['A','A+','At','a'].includes(a)))) throw new Error('Missing or invalid Agouti genotype.');
  }
  // Agouti cannot affect a chestnut x chestnut cross. For other crosses it must be known in both parents.
  if ((mare.E.includes('E') || stud.E.includes('E')) && (!mare.A || !stud.A)) throw new Error('Agouti is required in both parents for this cross.');
  const result = {};
  for (const combo of calculateAllCombinations(mare,stud)) {
    const phenotype=interpretPhenotype(combo);
    if (!result[phenotype]) result[phenotype]={probability:0,genotypes:[]};
    result[phenotype].probability += combo.__probability;
    result[phenotype].genotypes.push(Object.entries(combo).filter(([k,v])=>Array.isArray(v)).map(([k,v])=>v.join('/')).join(', '));
  }
  return result;
}

function interpretPhenotype(genetics) {
  if ((genetics.KIT || []).filter(a => /^W\d+$/.test(a)).length === 2) return 'Failed pregnancy (W/W)';
  if (countAllele(genetics.OWL, 'OLW') === 2) return 'Lethal White (OLW/OLW)';
  if (hasAllele(genetics['G'], 'G', 'G')) return "Grey";

  const hasE = hasAllele(genetics['E'], 'E', 'E');
  const isChestnut = !hasE;
  const agouti = genetics['A'] || ['n', 'n'];

  let baseColor = "";
  let isBlack = false;
  let isSealBrown = false;
  let isWildBay = false;

  if (isChestnut) {
    baseColor = "Chestnut";
  } else {
    // Dominanzreihenfolge: A+ > A > At > a
    if (hasAllele(agouti, 'A+', 'A')) {
      baseColor = "Wild Bay";
      isWildBay = true;
    } else if (hasAllele(agouti, 'A', 'A')) {
      baseColor = "Bay";
    } else if (hasAllele(agouti, 'At', 'A')) {
      baseColor = "Seal Brown";
      isSealBrown = true;
    } else {
      baseColor = "Black";
      isBlack = true;
    }
  }

  const dilPair = genetics['CR_Locus'] || ['n', 'n'];
  const cr = countAllele(dilPair, 'CR', 'CR_Locus');
  const prl = countAllele(dilPair, 'prl', 'CR_Locus');

  let color = baseColor;
  let isDoubleDilute = false;

  if (cr === 2) {
    isDoubleDilute = true;
    color = isChestnut ? "Cremello"
          : isBlack    ? "Smoky Cream"
          : isWildBay  ? "Wild Bay Perlino"
          : "Perlino";
  } else if (cr === 1 && prl === 1) {
    isDoubleDilute = true;
    if (isChestnut)      color = "Palomino Pearl";
    else if (isBlack)    color = "Smoky Black Pearl";
    else if (isWildBay)  color = "Wild Bay Buckskin Pearl";
    else                 color = isSealBrown ? "Seal Brown Buckskin Pearl" : "Buckskin Pearl";
  } else if (prl === 2) {
    isDoubleDilute = true;
    color = isChestnut ? "Chestnut Pearl"
          : isBlack    ? "Black Pearl"
          : isWildBay  ? "Wild Bay Pearl"
          : "Bay Pearl";
  } else if (cr === 1) {
    if (isChestnut)      color = "Palomino";
    else if (isBlack)    color = "Smoky Black";
    else if (isWildBay)  color = "Wild Bay Buckskin";
    else                 color = isSealBrown ? "Seal Brown Buckskin" : "Buckskin";
  } else if (prl === 1) {
    color = baseColor; // One pearl allele is a carrier, not a visible dilution.
  }

  // --- CHAMPAGNE (dominant, eine Kopie reicht, kein Dosis-Effekt) ---
  const CHAMPAGNE_MAP = {
    "Chestnut":            "Gold Champagne",
    "Bay":                 "Amber Champagne",
    "Wild Bay":            "Wild Bay Amber Champagne",
    "Seal Brown":          "Sable Champagne",
    "Black":               "Classic Champagne",
    "Palomino":            "Gold Cream Champagne",
    "Buckskin":            "Amber Cream Champagne",
    "Wild Bay Buckskin":   "Wild Bay Amber Cream Champagne",
    "Seal Brown Buckskin": "Sable Cream Champagne",
    "Smoky Black":         "Classic Cream Champagne"
  };
  const hasCH = hasAllele(genetics['CH'], 'CH');
  if (hasCH) color = CHAMPAGNE_MAP[color] || (color + ' Champagne');

  // --- SOOTY (nie auf Black) ---
  const styCount = countAllele(genetics['STY'], 'Sty', 'STY');
  let prefixMods = [];
  if (!isBlack) {
    if (styCount === 2) {
      if (isChestnut && color === "Chestnut") color = "Liver Chestnut";
      else prefixMods.push("Sooty");
    } else if (styCount === 1 && !isChestnut) {
      prefixMods.push("Sooty");
    }
  }

  // --- SILVER (maskiert bei Double-Dilute) ---
  if (hasAllele(genetics['Z'], 'Z', 'Z') && !isChestnut && !isDoubleDilute) {
    color = "Silver " + color;
  }
  if (hasAllele(genetics['DUN'], 'D', 'DUN')) color += " Dun";

  // --- HR APPALOOSA & PATN LOGIK ---
  const lpPair = genetics['LP'] || ['n', 'n'];
  const patnPair = genetics['PATN'] || ['n', 'n'];

  const hasLP = lpPair.some(a => String(a).trim() === 'LP');
  const isHomozygousLP = lpPair.filter(a => String(a).trim() === 'LP').length === 2;
  const patn1Count = patnPair.filter(a => String(a) === 'PATN1').length;

  let appaloosaPattern = "";

  if (hasLP) {
    if (!isHomozygousLP) {
      if (patn1Count === 0) {
        appaloosaPattern = "Varnish Roan";
      } else if (patn1Count === 1) {
        appaloosaPattern = "Spotted Blanket";
      } else if (patn1Count === 2) {
        appaloosaPattern = "Leopard";
      }
    } else {
      if (patn1Count === 0) {
        appaloosaPattern = "Varnish Roan";
      } else if (patn1Count === 1) {
        appaloosaPattern = "Snowcap";
      } else if (patn1Count === 2) {
        appaloosaPattern = "Fewspot";
      }
    }
  }

  if (!isBlack && countAllele(genetics.MU, 'mu') === 2) color += ' Mushroom';
  if (isChestnut && countAllele(genetics.F, 'f') === 2 && !isDoubleDilute) prefixMods.push('Flaxen');
  if (!isBlack && hasAllele(genetics.PA, 'PA')) prefixMods.push('Pangare');
  if (hasLP && patn1Count === 0 && hasAllele(genetics.PATN2, 'PATN2')) appaloosaPattern = isHomozygousLP ? 'Snowcap' : 'Spotted Blanket';
  let suffixMods = [];
  const kit = genetics.KIT || [];
  if (kit.includes('TO')) suffixMods.push('Tobiano');
  if (kit.includes('RN')) suffixMods.push('Roan');
  if (kit.includes('SB1')) suffixMods.push('Sabino');
  const whites = [...new Set(kit.filter(a => /^W\d+$/.test(a)))];
  if (whites.length) suffixMods.push('White Spotting (' + whites.join('/') + ')');
  if (hasPattern(genetics.MITF, 'MITF') || hasPattern(genetics.SW2, 'SW2')) suffixMods.push('Splash');
  if (hasAllele(genetics.OWL, 'OLW')) suffixMods.push('Frame Overo');
  if (appaloosaPattern) suffixMods.push(appaloosaPattern);

  if (hasPattern(genetics['TO'], 'TO'))  suffixMods.push("Tobiano");
  if (hasPattern(genetics['SW'], 'SW'))  suffixMods.push("Splash");
  if (hasPattern(genetics['SB'], 'SB'))  suffixMods.push("Sabino");
  if (hasPattern(genetics['RN'], 'RN'))  suffixMods.push("Roan");

  // --- RABICANO (case-insensitive, Rab/rab Mischschreibung erlaubt) ---
  const rabPair = genetics['RAB'] || ['n', 'n'];
  const isRabicano = rabPair.every(a => String(a).trim().toLowerCase() === 'rab');
  if (isRabicano) suffixMods.push("Rabicano");

  if (countAllele(genetics.WS, 'WS') === 2) suffixMods.push('Hidden Sabino');

  const finalName = prefixMods.length > 0 ? prefixMods.join(" ") + " " + color : color;
  return suffixMods.length > 0 ? finalName + " " + suffixMods.join(" ") : finalName;
}

function hasAllele(pair, allele) { return Array.isArray(pair) && pair.includes(allele); }

function countAllele(pair, allele) { return Array.isArray(pair) ? pair.filter(a => a === allele).length : 0; }

function hasPattern(pair, type) {
  const active={TO:['TO','To'], SW:['SW1','SW3','Sw'], MITF:['SW1','SW3'], SW2:['SW2'], SB:['SB1','Sb'], RN:['RN','Rn'], LP:['LP'], WS:['WS']};
  return Array.isArray(pair) && pair.some(a => (active[type] || []).includes(a));
}

function savePregnancy(mareId, studId) {
  const lock=LockService.getDocumentLock();
  if(!lock.tryLock(10000))throw new Error('Another booking is being saved. Please try again.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const broodSheet = ss.getSheetByName("Broodmares");
    const tracker = ss.getSheetByName("Herd");
    if (!broodSheet) throw new Error("Sheet 'Broodmares' not found!");
    const data = broodSheet.getDataRange().getValues();
    const today = new Date();
    const currentTime = Utilities.formatDate(today, Session.getScriptTimeZone(), "HH:mm");
    const mareData = findHorseByIdBreeding(tracker, mareId);
    const studData = findHorseByIdBreeding(tracker, studId);
    if(!mareData || !studData)throw new Error('Horse no longer found. Reopen the calculator.');
    if(mareData.gender!=='Mare' || studData.gender!=='Stallion' || mareData.age<3 || studData.age<3)throw new Error('Booking requires an adult mare and stallion.');
    if(String(mareData.breed).trim().toLowerCase()!==String(studData.breed).trim().toLowerCase())throw new Error('Parents must be the same breed.');

    let targetRow = -1;
    let firstEmptyRow = -1;
    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][2] || "").trim();
      if (currentId === String(mareId).trim()) {
        if(data[i][5] || data[i][12])throw new Error('This mare already has a booking. Review it in Broodmares before replacing it.');
        targetRow=i+1;break;
      }
      if (firstEmptyRow === -1 && currentId === "") firstEmptyRow = i + 1;
    }
    if (targetRow === -1) targetRow = firstEmptyRow !== -1 ? firstEmptyRow : broodSheet.getLastRow() + 1;

    broodSheet.getRange(targetRow, 3, 1, 5).setValues([[mareId, mareData ? mareData.name : mareId, mareData ? mareData.breed : "", today, currentTime]]);
    broodSheet.getRange(targetRow, 13).setValue(studData ? studData.name : studId);
    return "✓ Breeding booked for " + (mareData ? mareData.name : mareId) + " (row " + targetRow + ")";
  } catch (e) {
    throw new Error("Save failed: " + e.message);
  } finally { lock.releaseLock(); }
}

function fcParseAge_(ageStr) {
  if (ageStr === '' || ageStr == null) return 0;
  if (/^\d+(?:\.\d+)?$/.test(String(ageStr).trim())) return Number(ageStr);
  const str    = String(ageStr).toLowerCase().trim();
  const yMatch = str.match(/(\d+)\s*y/);
  const mMatch = str.match(/(\d+)\s*m/);
  return (yMatch ? Number(yMatch[1]) : 0) + (mMatch ? Number(mMatch[1]) / 12 : 0);
}

function normalizeGender(g) {
  const s=String(g || '').trim().toLowerCase();
  if (['m','f','mare','filly','stute','stutfohlen','female','♀'].includes(s)) return 'Mare';
  if (['s','h','c','stallion','colt','hengst','hengstfohlen','male','♂'].includes(s)) return 'Stallion';
  if (/\b(mare|filly|stute|female)\b/.test(s)) return 'Mare';
  if (/\b(stallion|colt|hengst|male)\b/.test(s)) return 'Stallion';
  return 'Other';
}

function getGeneticsForId(ss,id,sColor,cache={}) {
  if(sColor){if(!cache.colour)cache.colour=sColor.getDataRange().getValues();const row=cache.colour.slice(1).find(r=>String(r[1]).trim()===String(id).trim());if(row && row[79])return String(row[79]).trim();}
  const sh=ss.getSheetByName('Outside Studs') || ss.getSheetByName('Public Studs');
  if(sh){if(!cache.outside)cache.outside=sh.getDataRange().getValues();
    if(!/^(breed|rasse)$/i.test(String(cache.outside[0]?.[3] || '').trim()))throw new Error('Outside Studs needs the new Breed column D.');
    const row=cache.outside.slice(1).find(r=>String(r[1]).trim()===String(id).trim());if(row)return String(row[75] || '').trim();
  }return null;
}

function findHorseByIdBreeding(sheet, id) {
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(id).trim()) {
        return { name: data[i][3], breed: data[i][4], gender:normalizeGender(data[i][5]), age:fcParseAge_(data[i][8]) };
      }
    }
  }
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sPublic = ss.getSheetByName("Outside Studs") || ss.getSheetByName("Public Studs");
  if (sPublic) {
    const data = sPublic.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(id).trim()) {
        return { name: data[i][2], breed: data[i][3], gender:'Stallion', age:3 };
      }
    }
  }
  return null;
}

// Document-scoped, manually entered HR values are keyed by the exact mare/stallion IDs.
function fcHrCoiKey_(mareId,studId) {
  const ids=[mareId,studId].map(v=>String(v||'').trim());
  if(ids.some(id=>!/^[A-Za-z0-9_-]{1,64}$/.test(id)))throw new Error('Invalid horse ID.');
  return ids.join('|');
}
function fcLoadHrCoi_() {
  const props=PropertiesService.getDocumentProperties();
  if(!props)throw new Error('Document storage is unavailable.');
  const values=props.getProperties(),result={};
  Object.entries(values).forEach(([key,value])=>{
    if(!key.startsWith('fc_hr_coi_v1_'))return;
    const n=Number(value);if(value!==''&&Number.isFinite(n)&&n>=0&&n<=100)result[key.slice('fc_hr_coi_v1_'.length)]=n;
  });return result;
}
function saveFoalHrCoi(mareId,studId,value) {
  const pair=fcHrCoiKey_(mareId,studId),key='fc_hr_coi_v1_'+pair;
  const props=PropertiesService.getDocumentProperties();
  if(!props)throw new Error('Document storage is unavailable.');
  if(String(value??'').trim()===''){props.deleteProperty(key);return {pair,value:null};}
  const fraction=fcParseStoredCoi_(value);
  if(fraction===null)throw new Error('Enter an HR COI between 0 and 100%.');
  const percent=fraction*100;props.setProperty(key,String(percent));return {pair,value:percent};
}
