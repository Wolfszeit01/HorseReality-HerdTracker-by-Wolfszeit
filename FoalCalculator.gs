// ==========================================
// NORDRISKA BREEDING OPTIMIZER BACKEND — FINAL
// ==========================================

function openBreedingOptimizer() {
  const html = HtmlService.createTemplateFromFile('BreedingOptimizer').evaluate()
    .setTitle('Breeding Optimizer');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// ID-NORMALISIERUNG
// ==========================================
function normalizeIdForComparison(id) {
  if (!id) return "";
  return String(id).trim().toLowerCase().replace(/\s+/g, '').replace(/[★☆]/g, '');
}

// ==========================================
// BATCH INZUCHT-CHECK
// ==========================================
function batchCheckInbreeding(mareId, stallionIds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pedigreeSheet = ss.getSheetByName("Pedigree");
    const trackerSheet  = ss.getSheetByName("Herd Tracker");
    const publicSheet   = ss.getSheetByName("Outside Studs") || ss.getSheetByName("Public Studs");

    if (!pedigreeSheet) {
      Logger.log("❌ Pedigree Sheet not found!");
      const results = {};
      stallionIds.forEach(id => {
        results[id] = { isRelated: false, reason: "No Pedigree sheet available" };
      });
      return results;
    }

    // 1. NAME/ID LOOKUP & PUBLIC STUDS LIST
    const idToNameMap = {};
    const publicStudIds = new Set();

    if (trackerSheet) {
      const trackerData = trackerSheet.getDataRange().getValues();
      for (let i = 1; i < trackerData.length; i++) {
        const id   = String(trackerData[i][2] || "").trim();
        const name = String(trackerData[i][3] || "").trim();
        if (id && name) idToNameMap[id] = name;
      }
    }

    if (publicSheet) {
      const publicData = publicSheet.getDataRange().getValues();
      for (let i = 1; i < publicData.length; i++) {
        const id   = String(publicData[i][1] || "").trim();
        const name = String(publicData[i][2] || "").trim();
        if (id && name) {
          idToNameMap[id] = name;
          publicStudIds.add(id);
        }
      }
    }

    // 2. MARE PEDIGREE
    const marePed = getPedigreeForHorse(pedigreeSheet, mareId, idToNameMap);

    if (!marePed || marePed.length < 2) {
      const results = {};
      stallionIds.forEach(id => {
        results[id] = { isRelated: false, reason: "No pedigree data available for mare" };
      });
      return results;
    }

    const mareDam      = marePed[0] || "";
    const mareSire     = marePed[1] || "";
    const mareDamNorm  = normalizeIdForComparison(mareDam);
    const mareSireNorm = normalizeIdForComparison(mareSire);

    const mareAncestorsNorm = marePed
      .map(a => normalizeIdForComparison(a))
      .filter(n => n && n !== "unknown" && n !== "n/a" && n !== "-");

    const results = {};

    // 3. CHECK EACH STALLION
    stallionIds.forEach(studId => {
      const studName = idToNameMap[studId] || studId;

      // Outside studs — skip inbreeding check
      if (publicStudIds.has(studId)) {
        results[studId] = { isRelated: false, reason: "Outside stud — inbreeding check not available" };
        return;
      }

      const studIdNorm   = normalizeIdForComparison(studId);
      const studNameNorm = normalizeIdForComparison(studName);

      // Check 1: Stallion is direct parent of mare
      if (studIdNorm === mareDamNorm || studNameNorm === mareDamNorm) {
        results[studId] = { isRelated: true, reason: "Stallion is the dam of the mare" };
        return;
      }
      if (studIdNorm === mareSireNorm || studNameNorm === mareSireNorm) {
        results[studId] = { isRelated: true, reason: "Stallion is the sire of the mare" };
        return;
      }

      // Check 2: Stallion appears in mare's pedigree
      if (mareAncestorsNorm.includes(studIdNorm) || mareAncestorsNorm.includes(studNameNorm)) {
        results[studId] = { isRelated: true, reason: `"${studName}" is an ancestor of the mare` };
        return;
      }

      // Check 3: Load stallion pedigree
      const studPed = getPedigreeForHorse(pedigreeSheet, studId, idToNameMap);
      if (!studPed || studPed.length < 2) {
        results[studId] = { isRelated: false, reason: "No pedigree data available for stallion" };
        return;
      }

      const studDam      = studPed[0] || "";
      const studSire     = studPed[1] || "";
      const studDamNorm  = normalizeIdForComparison(studDam);
      const studSireNorm = normalizeIdForComparison(studSire);

      // Check 4: Mare is direct parent of stallion
      const mareIdNorm   = normalizeIdForComparison(mareId);
      const mareName     = idToNameMap[mareId] || mareId;
      const mareNameNorm = normalizeIdForComparison(mareName);

      if (mareIdNorm === studDamNorm || mareNameNorm === studDamNorm) {
        results[studId] = { isRelated: true, reason: "Mare is the dam of the stallion" };
        return;
      }
      if (mareIdNorm === studSireNorm || mareNameNorm === studSireNorm) {
        results[studId] = { isRelated: true, reason: "Mare is the sire of the stallion" };
        return;
      }

      // Check 5: Common ancestors
      const studAncestorsNorm = studPed
        .map(a => normalizeIdForComparison(a))
        .filter(n => n && n !== "unknown" && n !== "n/a" && n !== "-");

      const commonAncestors    = [];
      const commonAncestorsSet = new Set();

      mareAncestorsNorm.forEach((mareAnc, idx) => {
        if (studAncestorsNorm.includes(mareAnc)) {
          const originalName = marePed[idx] || mareAnc;
          if (!commonAncestorsSet.has(mareAnc)) {
            commonAncestorsSet.add(mareAnc);
            commonAncestors.push(originalName);
          }
        }
      });

      if (commonAncestors.length > 0) {
        const names  = commonAncestors.slice(0, 3).join(", ");
        const suffix = commonAncestors.length > 3 ? ` (+${commonAncestors.length - 3} more)` : "";
        results[studId] = {
          isRelated: true,
          reason: `${commonAncestors.length} common ancestor(s): ${names}${suffix}`
        };
        return;
      }

      // Check 6: Siblings
      if (studDamNorm && mareDamNorm && studDamNorm === mareDamNorm && mareDamNorm !== "") {
        if (studSireNorm && mareSireNorm && studSireNorm === mareSireNorm && mareSireNorm !== "") {
          results[studId] = { isRelated: true, reason: "Full siblings (same parents)" };
          return;
        }
        results[studId] = { isRelated: true, reason: "Half siblings (shared dam)" };
        return;
      }
      if (studSireNorm && mareSireNorm && studSireNorm === mareSireNorm && mareSireNorm !== "") {
        results[studId] = { isRelated: true, reason: "Half siblings (shared sire)" };
        return;
      }

      results[studId] = { isRelated: false, reason: "No known relation" };
    });

    return results;

  } catch (e) {
    Logger.log("❌ Error in batchCheckInbreeding: " + e.message);
    const results = {};
    stallionIds.forEach(id => {
      results[id] = { isRelated: false, reason: "Error during check: " + e.message };
    });
    return results;
  }
}

// ==========================================
// PEDIGREE LADEN
// Pedigree sheet: B=ID(1), C=Name(2), D=Sireline(3), E=Damline(4)
// F=Sire(5), G=Dam(6), H=GS(P)(7) ... S=GGD(MM)(18)
// ==========================================
function getPedigreeForHorse(sheet, horseId, idToNameMap) {
  try {
    const data     = sheet.getDataRange().getValues();
    const searchId = normalizeIdForComparison(horseId);

    // Try 1: ID in column B (index 1)
    for (let i = 1; i < data.length; i++) {
      if (normalizeIdForComparison(data[i][1]) === searchId) {
        return extractPedigree(data[i]);
      }
    }

    // Try 2: Name lookup
    const horseName = idToNameMap[horseId];
    if (horseName) {
      const searchName = normalizeIdForComparison(horseName);
      for (let i = 1; i < data.length; i++) {
        if (normalizeIdForComparison(data[i][2]) === searchName) {
          return extractPedigree(data[i]);
        }
      }
    }

    // Try 3: Direct name match
    for (let i = 1; i < data.length; i++) {
      if (normalizeIdForComparison(data[i][2]) === searchId) {
        return extractPedigree(data[i]);
      }
    }

    return null;
  } catch (e) {
    Logger.log("Error in getPedigreeForHorse: " + e.message);
    return null;
  }
}

function extractPedigree(row) {
  const pedigree = [];
  // Pedigree starts at column F (index 5): Sire, Dam, grandparents...
  for (let j = 5; j < Math.min(row.length, 25); j++) {
    const val = String(row[j] || "").trim();
    if (val && val.toLowerCase() !== "n/a") pedigree.push(val);
  }
  return pedigree;
}

// ==========================================
// ALLE DATEN LADEN
// ==========================================
function loadAllDataOptimized() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = {
      tracker:       ss.getSheetByName("Herd Tracker"),
      statsStandard: ss.getSheetByName("Horse Stats"),
      statsIce:      ss.getSheetByName("ICE_Horse Stats"),
      statsKath:     ss.getSheetByName("KATH_Horse Stats"),
      pedigree:      ss.getSheetByName("Pedigree"),
      color:         ss.getSheetByName("Colour Genetics"),
      public:        ss.getSheetByName("Outside Studs") || ss.getSheetByName("Public Studs"),
      bro:           ss.getSheetByName("Broodmares")
    };

    const output = { tracker: {}, stats: {}, pedigree: {}, genetics: {}, stallionUsage: {} };

    // 1. HERD TRACKER
    if (sheets.tracker) {
      const data = sheets.tracker.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][2]).trim();
        if (!id) continue;
        output.tracker[id] = {
          name:   String(data[i][3]).trim(),
          breed:  String(data[i][4]).trim(),
          gender: normalizeGender(data[i][5]),
          age:    parseAge(String(data[i][8] || "").trim()),
          type:   'owned',
          dam:    "",
          sire:   ""
        };
      }
    }

    // 2. PEDIGREE — Sire at index 5, Dam at index 6
    if (sheets.pedigree) {
      const data = sheets.pedigree.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const cleanId = String(data[i][1] || "").trim().replace(/[★☆\s]/g, "");
        if (!cleanId) continue;

        const sireName = String(data[i][5] || "").trim();
        const damName  = String(data[i][6] || "").trim();

        for (let trackerId in output.tracker) {
          if (trackerId.trim().replace(/[★☆\s]/g, "") === cleanId) {
            output.tracker[trackerId].sire = sireName;
            output.tracker[trackerId].dam  = damName;
          }
        }

        output.pedigree[cleanId] = data[i].slice(5, 25)
          .map(v => String(v).trim())
          .filter(v => v !== "" && v.toLowerCase() !== "n/a");
      }
    }

    // 3. HORSE STATS (Standard / ICE / KATH)
    const processStats = (sheet, offset) => {
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][1]).trim();
        if (!id || output.stats[id]) continue;

        const breed = output.tracker[id] ? output.tracker[id].breed : "";
        let confoCount = 12;
        if (breed.toLowerCase().includes("icelandic"))  confoCount = 14;
        else if (breed.toLowerCase().includes("kathiawari")) confoCount = 13;

        let gpSum = 0;
        for (let k = 34; k <= 43; k++) gpSum += Number(data[i][k + offset]) || 0;

        output.stats[id] = {
          gpTotal: gpSum,
          gpStats: {
            acceleration: Number(data[i][34 + offset]) || 0,
            agility:      Number(data[i][35 + offset]) || 0,
            balance:      Number(data[i][36 + offset]) || 0,
            bascule:      Number(data[i][37 + offset]) || 0,
            pulling:      Number(data[i][38 + offset]) || 0,
            speed:        Number(data[i][39 + offset]) || 0,
            sprint:       Number(data[i][40 + offset]) || 0,
            stamina:      Number(data[i][41 + offset]) || 0,
            strength:     Number(data[i][42 + offset]) || 0,
            surefoot:     Number(data[i][43 + offset]) || 0
          },
          gpDisciplines: {
            dressage:  Number(data[i][45 + offset]) || 0,
            driving:   Number(data[i][47 + offset]) || 0,
            endurance: Number(data[i][49 + offset]) || 0,
            eventing:  Number(data[i][51 + offset]) || 0,
            flat:      Number(data[i][53 + offset]) || 0,
            jumping:   Number(data[i][55 + offset]) || 0,
            reining:   Number(data[i][57 + offset]) || 0
          },
          confoRaw:      data[i].slice(15 + offset, 15 + offset + confoCount).map(v => Number(v) || 0),
          confoScoreMin: Number(data[i][61 + offset]) || 0,
          confoScoreMax: Number(data[i][63 + offset]) || 0,
          confoCount:    confoCount
        };
      }
    };

    if (sheets.statsKath)     processStats(sheets.statsKath, 1);
    if (sheets.statsIce)      processStats(sheets.statsIce, 2);
    if (sheets.statsStandard) processStats(sheets.statsStandard, 0);

    // 4. OUTSIDE STUDS
    // Column mapping (verified):
    // B(1)=ID, C(2)=Name, D(3)=Breed, E(4)=Stud Fee, F(5)=Link
    // G(6)–T(19) = Confo (WLK–Socks)
    // AQ(42)–AZ(51) = Base Stats (Acc–Srft)
    // BA(52) = Total GP
    // BB(53)–BN(65) = Discipline GPs (every other col = GP, in-between = HS)
    // BO(66) = rec. discipline, BP(67) = Pred. Conf. Score
    // BQ(68) = Conf. MAX, BR(69) = Comp MAX, BS(70) = Genetic Code
    if (sheets.public) {
      const data = sheets.public.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const id    = String(data[i][1]).trim();
        const name  = String(data[i][2]).trim();
        const breed = String(data[i][3]).trim();
        if (!id) continue;

        if (!output.tracker[id]) {
          output.tracker[id] = {
            name:     name,
            breed:    breed,
            gender:   "Stallion",
            type:     'public',
            age:      10,
            dam:      "",
            sire:     "",
            studFee:  String(data[i][4] || '').trim()  // E = Stud Fee
          };
        }

        // Confo: G(6)–T(19), convert grade text to numbers
        const confoRaw = [];
        for (let j = 6; j < 20; j++) {
          const val = data[i][j];
          if      (val === "VG") confoRaw.push(100);
          else if (val === "G+" || val === "G") confoRaw.push(87.5);
          else if (val === "G-") confoRaw.push(76.5);
          else if (val === "A")  confoRaw.push(75);
          else if (val === "BA") confoRaw.push(49.5);
          else                   confoRaw.push(Number(val) || 75);
        }

        // Base stats: AQ(42)–AZ(51)
        const gpStats = {
          acceleration: Number(data[i][42]) || 0,
          agility:      Number(data[i][43]) || 0,
          balance:      Number(data[i][44]) || 0,
          bascule:      Number(data[i][45]) || 0,
          pulling:      Number(data[i][46]) || 0,
          speed:        Number(data[i][47]) || 0,
          sprint:       Number(data[i][48]) || 0,
          stamina:      Number(data[i][49]) || 0,
          strength:     Number(data[i][50]) || 0,
          surefoot:     Number(data[i][51]) || 0
        };

        // Total GP: BA(52)
        const gpTotal = Number(data[i][52]) || 0;

        // Discipline GPs: BB(53)=DR, BD(55)=DV, BF(57)=EN, BH(59)=EV, BJ(61)=RC, BL(63)=JMP, BN(65)=RE
        const gpDisciplines = {
          dressage:  Number(data[i][53]) || 0,
          driving:   Number(data[i][55]) || 0,
          endurance: Number(data[i][57]) || 0,
          eventing:  Number(data[i][59]) || 0,
          flat:      Number(data[i][61]) || 0,
          jumping:   Number(data[i][63]) || 0,
          reining:   Number(data[i][65]) || 0
        };

        // Pred. Conf. Score: BP(67)
        const confoScore = Number(data[i][67]) || 0;

        output.stats[id] = {
          gpTotal:       gpTotal,
          gpStats:       gpStats,
          gpDisciplines: gpDisciplines,
          confoRaw:      confoRaw,
          confoScoreMin: confoScore,
          confoScoreMax: confoScore,
          confoCount:    12
        };

        // Genetic Code: BS(70)
        const gencode = String(data[i][70] || "").trim();
        if (gencode && gencode.includes("/")) output.genetics[id] = gencode;
      }
    }

    // 5. COLOUR GENETICS — genetic code at index 57
    if (sheets.color) {
      const data = sheets.color.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const id   = String(data[i][1]).trim();
        const code = String(data[i][57]).trim();
        if (id && code && !output.genetics[id]) output.genetics[id] = code;
      }
    }

    // 6. STALLION USAGE from Broodmares (col M = index 12)
    if (sheets.bro) {
      const data = sheets.bro.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const stallionName = String(data[i][12] || "").trim();
        if (stallionName) {
          output.stallionUsage[stallionName] = (output.stallionUsage[stallionName] || 0) + 1;
        }
      }
    }

    return output;

  } catch (e) {
    console.error("Error in loadAllDataOptimized:", e);
    return { error: "Load failed: " + e.message, tracker: {}, stats: {}, pedigree: {}, genetics: {}, stallionUsage: {} };
  }
}

// ==========================================
// FOAL COLOUR CALCULATION
// ==========================================
function calculateFoalColors(mareId, studId) {
  try {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const sColor       = ss.getSheetByName("Colour Genetics");
    const mareGenetics = getGeneticsForId(ss, mareId, sColor);
    const studGenetics = getGeneticsForId(ss, studId, sColor);

    if (!mareGenetics || mareGenetics === '') return { error: "No genetics code for mare" };
    if (!studGenetics || studGenetics === '') return { error: "No genetics code for stallion" };

    const mareAlleles        = parseGenetics(mareGenetics);
    const studAlleles        = parseGenetics(studGenetics);
    const possibleCombos     = calculateAllCombinations(mareAlleles, studAlleles);
    const phenotypes         = possibleCombos.map(combo => interpretPhenotype(combo));
    const phenotypeCounts    = {};
    phenotypes.forEach(p => { phenotypeCounts[p] = (phenotypeCounts[p] || 0) + 1; });

    const total = phenotypes.length;
    let output  = "<div style='line-height:1.6;'>";
    Object.keys(phenotypeCounts).sort().forEach(pheno => {
      const pct = ((phenotypeCounts[pheno] / total) * 100).toFixed(2);
      output += `<b>${pct}%</b> - ${pheno}<br>`;
    });
    output += "</div>";
    return output || "No colours calculated";

  } catch (e) {
    return { error: "Error: " + e.message };
  }
}

function parseGenetics(genString) {
  const alleles = {};
  genString.trim().split(/\s+/).forEach(token => {
    const parts = token.split('/');
    if (parts.length === 2) {
      const a1         = parts[0].trim();
      const a2         = parts[1].trim();
      const targetGene = identifyAllele(a1) || identifyAllele(a2);
      if (targetGene) alleles[targetGene] = [a1, a2];
    }
  });
  return alleles;
}

function identifyAllele(allele) {
  const upper = allele.toUpperCase();
  if (upper === 'N' || upper === 'NN')                return null;
  if (upper === 'E' || allele === 'e')                return 'E';
  if (upper === 'A' || upper === 'A+' || upper === 'AT' || allele === 'a') return 'A';
  if (upper.includes('CR') || upper.includes('PRL'))  return 'CR_Locus';
  if (upper === 'D' || upper === 'ND1' || upper === 'ND2') return 'DUN';
  if (upper === 'CH' || upper === 'NCH')              return 'CH';
  if (upper === 'Z'  || upper === 'NZ')               return 'Z';
  if (upper === 'MU' || upper === 'NMU')              return 'MU';
  if (upper === 'G'  || upper === 'NG')               return 'G';
  if (upper === 'F'  || upper === 'NF')               return 'F';
  if (upper === 'STY'|| upper === 'NSTY')             return 'STY';
  if (upper === 'PA' || upper === 'PANG' || upper === 'NPA') return 'PA';
  if (upper.startsWith('SW'))   return 'SW';
  if (upper.startsWith('TO'))   return 'TO';
  if (upper.startsWith('SB'))   return 'SB';
  if (upper.startsWith('RN'))   return 'RN';
  if (upper.startsWith('LP'))   return 'LP';
  if (upper.startsWith('PATN')) return 'PATN';
  if (upper === 'OLW' || upper === 'OWL') return 'OWL';
  if (upper.startsWith('RAB')) return 'RAB';
  if (upper === 'WS' || upper === 'NWS' || upper.match(/^W\d+/)) return 'W';
  return null;
}

function calculateAllCombinations(mare, stud) {
  const geneNames = [...new Set([...Object.keys(mare), ...Object.keys(stud)])];
  let combos = [{}];
  geneNames.forEach(gene => {
    const mA = mare[gene] || ['n', 'n'];
    const sA = stud[gene] || ['n', 'n'];
    const newCombos = [];
    combos.forEach(combo => {
      mA.forEach(mAllele => {
        sA.forEach(sAllele => {
          newCombos.push({ ...combo, [gene]: [mAllele, sAllele] });
        });
      });
    });
    combos = newCombos;
  });
  return combos;
}

function interpretPhenotype(genetics) {
  if (hasAllele(genetics['G'], 'G')) return "Grey";
  const hasE       = hasAllele(genetics['E'], 'E');
  const isChestnut = !hasE;
  const agouti     = genetics['A'] || ['n', 'n'];
  let baseColor    = "";
  let isBlack      = false;
  let isSealBrown  = false;

  if (isChestnut) {
    baseColor = "Chestnut";
  } else {
    if (hasAllele(agouti, 'A') || hasAllele(agouti, 'A+'))  { baseColor = "Bay"; }
    else if (hasAllele(agouti, 'At'))                        { baseColor = "Seal Brown"; isSealBrown = true; }
    else                                                     { baseColor = "Black"; isBlack = true; }
  }

  const dilPair = genetics['CR_Locus'] || ['n', 'n'];
  const cr  = countAllele(dilPair, 'CR');
  const prl = countAllele(dilPair, 'prl');
  let color = baseColor;

  if (cr === 2) {
    color = isChestnut ? "Cremello" : (isBlack ? "Smoky Cream" : "Perlino");
  } else if (cr === 1 && prl === 1) {
    if (isChestnut)    color = "Palomino Pearl";
    else if (isBlack)  color = "Smoky Black Pearl";
    else               color = isSealBrown ? "Seal Brown Buckskin Pearl" : "Buckskin Pearl";
  } else if (prl === 2) {
    color = isChestnut ? "Gold Pearl" : (isBlack ? "Classic Pearl" : "Amber Pearl");
  } else if (cr === 1) {
    if (isChestnut)    color = "Palomino";
    else if (isBlack)  color = "Smoky Black";
    else               color = isSealBrown ? "Seal Brown Buckskin" : "Buckskin";
  } else if (prl === 1) {
    color = baseColor + " Pearl";
  }

  const styCount = countAllele(genetics['STY'], 'Sty');
  let prefixMods = [];
  if (styCount === 2) {
    if (isChestnut && color === "Chestnut") color = "Liver Chestnut";
    else prefixMods.push("Sooty");
  } else if (styCount === 1 && !isChestnut) {
    prefixMods.push("Sooty");
  }

  if (hasAllele(genetics['Z'], 'Z') && !isChestnut) color = "Silver " + color;
  if (hasAllele(genetics['DUN'], 'D')) color += " Dun";

  let suffixMods = [];
  if (hasPattern(genetics['LP']))  suffixMods.push("Appaloosa");
  if (hasPattern(genetics['TO']))  suffixMods.push("Tobiano");
  if (hasPattern(genetics['SW']))  suffixMods.push("Splash");
  if (hasPattern(genetics['SB']))  suffixMods.push("Sabino");
  if (hasPattern(genetics['RN']))  suffixMods.push("Roan");
  if (hasPattern(genetics['RAB'])) suffixMods.push("Rabicano");
  if (hasPattern(genetics['WS']))  suffixMods.push("White Pattern");

  const finalName = prefixMods.length > 0 ? prefixMods.join(" ") + " " + color : color;
  return suffixMods.length > 0 ? finalName + " " + suffixMods.join(" ") : finalName;
}

function hasAllele(genePair, targetAllele) {
  if (!genePair || !Array.isArray(genePair)) return false;
  const caseSensitive = ['E','e','A','a','A+','At','a+','at'];
  if (caseSensitive.includes(targetAllele)) return genePair.some(a => String(a) === targetAllele);
  return genePair.some(a => String(a).toLowerCase() === targetAllele.toLowerCase());
}

function countAllele(genePair, targetAllele) {
  if (!genePair || !Array.isArray(genePair)) return 0;
  const caseSensitive = ['E','e','A','a','A+','At','a+','at'];
  if (caseSensitive.includes(targetAllele)) return genePair.filter(a => String(a) === targetAllele).length;
  return genePair.filter(a => String(a).toLowerCase() === targetAllele.toLowerCase()).length;
}

function hasPattern(genePair) {
  if (!genePair) return false;
  return genePair.some(a => String(a).toLowerCase() !== 'n' && String(a).toLowerCase() !== 'nd2');
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================
function savePregnancy(mareId, studId) {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const broodSheet = ss.getSheetByName("Broodmares");
    const tracker    = ss.getSheetByName("Herd Tracker");

    if (!broodSheet) throw new Error("Sheet 'Broodmares' not found!");

    const data        = broodSheet.getDataRange().getValues();
    const today       = new Date();
    const currentTime = Utilities.formatDate(today, Session.getScriptTimeZone(), "HH:mm");
    const mareData    = findHorseByIdBreeding(tracker, mareId);
    const studData    = findHorseByIdBreeding(tracker, studId);

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][2] || "").trim();
      if (currentId === String(mareId).trim()) { targetRow = i + 1; break; }
      if (targetRow === -1 && currentId === "") targetRow = i + 1;
    }
    if (targetRow === -1) targetRow = broodSheet.getLastRow() + 1;

    // Write columns C–G, preserve formula columns H–K
    broodSheet.getRange(targetRow, 3, 1, 5).setValues([[
      mareId,
      mareData ? mareData.name  : mareId,
      mareData ? mareData.breed : "",
      today,
      currentTime
    ]]);

    // Stallion name in column M (13)
    broodSheet.getRange(targetRow, 13).setValue(studData ? studData.name : studId);

    return "✓ Breeding booked for " + (mareData ? mareData.name : mareId) + " (row " + targetRow + ")";

  } catch (e) {
    throw new Error("Save failed: " + e.message);
  }
}

function parseAge(ageStr) {
  if (!ageStr) return 0;
  const str    = String(ageStr).toLowerCase().trim();
  const yMatch = str.match(/(\d+)\s*y/);
  const mMatch = str.match(/(\d+)\s*m/);
  return (yMatch ? Number(yMatch[1]) : 0) + (mMatch ? Number(mMatch[1]) / 12 : 0);
}

function normalizeGender(g) {
  const s = String(g).toLowerCase().trim();
  if (s.includes("stallion") || s === "m") return "Stallion";
  if (s.includes("mare")     || s === "f") return "Mare";
  return "Other";
}

function getGeneticsForId(ss, id, sColor) {
  // 1. Check Colour Genetics sheet (index 57)
  if (sColor) {
    const data = sColor.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(id).trim()) {
        const code = String(data[i][57]).trim();
        if (code) return code;
      }
    }
  }
  // 2. Check Outside Studs sheet — Genetic Code at BS (index 70)
  const sPublic = ss.getSheetByName("Outside Studs") || ss.getSheetByName("Public Studs");
  if (sPublic) {
    const data = sPublic.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(id).trim()) {
        return data[i].length > 70 ? String(data[i][70] || "").trim() : "";
      }
    }
  }
  return null;
}

function findHorseByIdBreeding(sheet, id) {
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(id).trim()) {
        return { name: data[i][3], breed: data[i][4] };
      }
    }
  }
  // Fallback: Outside Studs
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sPublic = ss.getSheetByName("Outside Studs") || ss.getSheetByName("Public Studs");
  if (sPublic) {
    const data = sPublic.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(id).trim()) {
        return { name: data[i][2], breed: data[i][3] };
      }
    }
  }
  return null;
}
