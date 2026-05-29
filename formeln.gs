/***********************
 * CUSTOM SHEET FUNCTIONS — COMMUNITY VERSION
 ***********************/

/**
 * Calculates HorseReality age based on birth date and aging items.
 * @customfunction
 */
function HR_AGE(birthDate, agingItems) {
  if (!birthDate) return agingItems ? (agingItems * 0.5) + " years" : "";
  const now        = new Date();
  const birth      = new Date(birthDate);
  const rlDays     = (now.getTime() - birth.getTime()) / (1000 * 3600 * 24);
  const gameYears  = rlDays / 16;
  const agingYears = agingItems ? agingItems * 0.5 : 0;
  const total      = gameYears + agingYears;
  const years      = Math.floor(total);
  const months     = Math.floor((total - years) * 12);
  return years + "y " + months + "m";
}

function getSheetInfo(breed) {
  if (!breed) return { name: "Horse Stats", offset: 0 };
  if (breed.includes("Icelandic"))  return { name: "ICE_Horse Stats",  offset: 2 };
  if (breed.includes("Kathiawari")) return { name: "KATH_Horse Stats", offset: 1 };
  return { name: "Horse Stats", offset: 0 };
}

function GET_HR_LINK(id) {
  return id ? "https://www.horsereality.com/horses/" + id + "/" : "";
}

function GET_PEDIGREE_INFO(id, type) {
  if (!id) return "";
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pedigree");
  const data  = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] == id) {
      const name      = (type === "Dam") ? data[i][4] : data[i][5];
      const pedigreeId= (type === "Dam") ? data[i][2] : data[i][3];
      return pedigreeId + " (" + name + ")";
    }
  }
  return "Not found";
}

function GET_CONFO_SUMMARY(id, breed) {
  if (!id) return "";
  const info  = getSheetInfo(breed);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(info.name);
  const data  = sheet.getDataRange().getValues();
  const off   = info.offset;
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] == id) {
      let res = "";
      if (data[i][27+off]) res += data[i][27+off] + " VG ";
      if (data[i][28+off]) res += data[i][28+off] + " G+ ";
      if (data[i][29+off]) res += data[i][29+off] + " G ";
      if (data[i][30+off]) res += data[i][30+off] + " G- ";
      if (data[i][31+off]) res += data[i][31+off] + " A ";
      if (data[i][32+off]) res += data[i][32+off] + " BA ";
      return res.trim() || "No stats";
    }
  }
  return "Not found";
}

function GET_HORSE_STAT(id, breed, baseCol) {
  if (!id) return "";
  const info  = getSheetInfo(breed);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(info.name);
  const data  = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] == id) return data[i][baseCol - 1 + info.offset];
  }
  return "";
}

function GET_RESULTS_STRING(id, breed, type) {
  if (!id) return "";
  const info  = getSheetInfo(breed);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(info.name);
  const data  = sheet.getDataRange().getValues();
  const off   = info.offset;
  const col   = (type === "Confo") ? 64 + off : 71 + off;
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] == id) {
      if (type === "Confo") return data[i][col] + " | " + data[i][col+1] + " | " + data[i][col+2] + " | " + data[i][col+3];
      return data[i][col] + " | " + data[i][col+1] + " | " + data[i][col+2];
    }
  }
  return "";
}

function GET_HEALTH_SCORE(range, mode) {
  const flat   = range.flat();
  const exc    = flat.filter(v => v === "Excellent").length;
  const avgFair= flat.filter(v => v === "Average" || v === "Fair").length;
  if (mode === 1) return exc;
  if (mode === 2) return avgFair;
  return exc - avgFair;
}

function GET_GENETICS(id, mode) {
  if (!id) return "";
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Colour Genetics');
  if (!sheet) return "Sheet 'Colour Genetics' not found";
  const data  = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] == id) return mode === "Color" ? data[i][55] : data[i][56];
  }
  return "Not found";
}

function GET_RESULT_HLOOKUP(id, sheetName, row) {
  if (!id) return "";
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Sheet " + sheetName + " not found";
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let col = 1; col < headers.length; col++) {
    if (headers[col] == id) return data[row - 1][col] || "";
  }
  return "";
}

function CONFO_VALUE(rating) {
  const map = { "P": 19.5, "BA": 49.5, "A": 64.5, "G-": 74.5, "G": 76.5, "G+": 82, "VG": 92.5 };
  return map[rating] || 0;
}

function CALC_HS(statsRange, gpValues, type) {
  const s = statsRange[0];
  switch(type) {
    case 'DR':  return (0.25*(s[0]+s[1]+s[2]+s[4])/4)               + (0.75*gpValues/3);
    case 'DV':  return (0.25*(s[1]+s[7]+s[8]+s[10])/4)              + (0.75*gpValues/5);
    case 'EN':  return (0.25*(s[0]+s[1]+s[2]+s[5]+s[6]+s[7])/6)     + (0.75*gpValues/4);
    case 'EV':  return (0.25*(s[0]+s[1]+s[2]+s[4]+s[5]+s[6])/6)     + (0.75*gpValues/5);
    case 'RC':  return (0.25*(s[3]+s[4]+s[6]+s[7]+s[8]+s[9]+s[10])/7)+(0.75*gpValues/4);
    case 'JPM': return (0.25*(s[2]+s[7]+s[8]+s[9]+s[10])/5)         + (0.75*gpValues/5);
    case 'RE':  return (0.25*(s[5]+s[6]+s[8]+s[9]+s[10])/5)         + (0.75*gpValues/4);
    default:    return 0;
  }
}

function GET_RECC_DISCIPLINE(scores) {
  const names   = ["Dressage","Driving","Endurance","Eventing","Reining","Show Jumping","Racing"];
  const maxScore= Math.max(...scores);
  const idx     = scores.indexOf(maxScore);
  return names[idx] || "";
}

function GET_LOOKUP_VAL(id, sheetName, row) {
  if (!id) return "";
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "";
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let col = 1; col < headers.length; col++) {
    if (headers[col] == id) return data[row - 1][col] || 0;
  }
  return 0;
}
