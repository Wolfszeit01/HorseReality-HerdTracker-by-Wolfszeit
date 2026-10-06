/***************************************************************************
 * NOTIFICATION & BANK LOG PARSER — UNIFIED BACKEND
 *
 * Schreibt alles in ein einziges "Activity Log" Sheet.
 * Ersetzt die alten separaten Log-Tabs (Comp/Conf/Stud/Sales/Bank Log).
 *
 * Activity Log Spalten (1-indexiert):
 *   A: Date | B: Type | C: Horse ID | D: Horse Name | E: Detail
 *   F: Detail 2 | G: Detail 3 | H: Currency | I: Amount | J: Fee
 *   K: Netto HRC | L: Raw Text
 *
 * Namespace: el_ (Emily Logs) — vermeidet Kollisionen mit Wolf's Herd Manager.
 ***************************************************************************/


/***************************************************************************
 * CONSTANTS
 ***************************************************************************/

const AL = {
  LOG:        'Activity Log',
  REVIEW:     'Needs Review',
  HERD:       'Herd Tracker',
  ARCHIVE:    'Archive',
  BROODMARES: 'Broodmares'
};


/***************************************************************************
 * MODAL
 ***************************************************************************/

function el_showNotificationParserModal() {
  const html = HtmlService.createTemplateFromFile('NotificationParserModal')
    .evaluate()
    .setWidth(800)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}


/***************************************************************************
 * LOG TAB VISIBILITY
 ***************************************************************************/

const EL_HIDEABLE_TABS = ['Activity Log', 'Needs Review', 'Notifications Inbox'];

function fd_hideLogSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  EL_HIDEABLE_TABS.forEach(n => { const s = ss.getSheetByName(n); if (s) s.hideSheet(); });
  SpreadsheetApp.getUi().alert('Log-Tabs ausgeblendet.');
}

function fd_showLogSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  EL_HIDEABLE_TABS.forEach(n => { const s = ss.getSheetByName(n); if (s) s.showSheet(); });
  SpreadsheetApp.getUi().alert('Log-Tabs eingeblendet.');
}


/***************************************************************************
 * PROCESS NOTIFICATIONS (from Modal paste)
 ***************************************************************************/

function el_processNotificationHtml(htmlBlob) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return el_processNotificationHtml_(htmlBlob); }
  finally { lock.releaseLock(); }
}

function el_processNotificationHtml_(htmlBlob) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = el_splitNotificationHtml_(htmlBlob);
  if (!lines.length) return { parsed: 0, review: 0, skipped: 0, errored: 0, counts: {}, reviewDetails: [], empty: true };

  const logSheet = ss.getSheetByName(AL.LOG);
  if (!logSheet || !ss.getSheetByName(AL.REVIEW)) throw new Error('Activity Log and Needs Review sheets are required.');
  const seen = el_loadExistingHashes_(logSheet);
  const ctx = el_buildContext_(ss);
  let parsed = 0, review = 0, skipped = 0, errored = 0;
  const rows = [], reviewRows = [], reviewDetails = [];

  lines.forEach(line => {
    const rawText = line.text.trim();
    if (!rawText) return;

    let eventDate;
    try { eventDate = el_parseEventDate_(rawText, ''); }
    catch (e) { reviewRows.push([new Date(), 'Date parse failed', e.message, false, rawText]); errored++; return; }

    const type = el_classifyNotification_(rawText);

    try {
      if (type === 'noise') { skipped++; return; }

      const result = el_parseNotification_(type, rawText, eventDate, line.links, ctx);
      if (result.row) {
        const key = el_rowHash_(result.row);
        if (seen.has(key)) { skipped++; return; }
        seen.add(key);
      }

      if (result.needsReview) {
        reviewRows.push([new Date(), result.reason, result.suggestion || '', false, rawText]);
        reviewDetails.push({ reason: result.reason, text: rawText });
        review++;
        if (result.row) rows.push(result.row); // Write anyway if row exists
      } else {
        rows.push(result.row);
        parsed++;
      }

      // Side effects
      if (type === 'covering_failed' && result.row) {
        el_incrementCoveringFailed_(ss, result.row[2], result.row[3]);
      }
      if (type === 'birth' && result.row) {
        el_resetMareOnBirth_(ss, result.row[3], result.row[5]);
      }
      if (type === 'deloryan' && result.genderLabel) {
        el_writeDeloryanGender_(ss, result.mareId, result.mareName, result.genderLabel);
      }
    } catch (e) {
      reviewRows.push([new Date(), 'Parser exception: ' + e.message, '', false, rawText]);
      errored++;
    }
  });

  const written = el_writeResults_(ss, rows, reviewRows);

  return {
    parsed, review, skipped, errored,
    counts: { total: written },
    reviewDetails
  };
}


/***************************************************************************
 * PROCESS BANK LOG (from Modal paste)
 ***************************************************************************/

function el_processBankLogText(text) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return el_processBankLogText_(text); }
  finally { lock.releaseLock(); }
}

function el_processBankLogText_(text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lines = String(text).split('\n').map(l => l.trim()).filter(Boolean);
  if (!lines.length) return { parsed: 0, review: 0, skipped: 0, errored: 0, counts: {}, reviewDetails: [], empty: true };

  const logSheet = ss.getSheetByName(AL.LOG);
  if (!logSheet || !ss.getSheetByName(AL.REVIEW)) throw new Error('Activity Log and Needs Review sheets are required.');
  const seen = el_loadExistingHashes_(logSheet);
  const ctx = el_buildContext_(ss);
  let parsed = 0, review = 0, skipped = 0, errored = 0;
  const rows = [], reviewRows = [], reviewDetails = [];

  lines.forEach(line => {
    // Skip header row
    if (/^Type\s/i.test(line)) { skipped++; return; }

    // Split by tab: BankType | Amount | Description | Date
    const parts = line.split('\t');
    if (parts.length < 3) { skipped++; return; }

    const bankType = parts[0].trim().toUpperCase();
    const bankAmount = el_parseAmount_(parts[1]);
    const description = (parts[2] || '').trim();
    const dateStr = (parts[3] || '').trim();

    if (!description) { skipped++; return; }

    let eventDate;
    try { eventDate = el_parseEventDate_(description, dateStr); }
    catch (e) { eventDate = el_getGameDate_(0); }

    // Strip "Horse Reality: " prefix
    const desc = description.replace(/^Horse Reality:\s*/i, '');

    const type = el_classifyBankEntry_(desc, bankType);

    try {
      if (type === 'skip' || type === 'noise') { skipped++; return; }

      const result = el_parseBankEntry_(type, desc, bankAmount, bankType, eventDate, ctx, description);
      if (result.row) {
        const key = el_rowHash_(result.row);
        if (seen.has(key)) { skipped++; return; }
        seen.add(key);
      }

      if (result.needsReview) {
        reviewRows.push([new Date(), result.reason, result.suggestion || '', false, description]);
        reviewDetails.push({ reason: result.reason, text: description });
        review++;
        if (result.row) rows.push(result.row);
      } else {
        rows.push(result.row);
        parsed++;
      }

      // Side effect: ultrasound → auto-confirm Stute in Broodmares
      if (type === 'ultrasound' && result.row) {
        el_autoConfirmOnUltrasound_(ss, result.row[2], result.row[3]); // horseId, horseName
      }
    } catch (e) {
      reviewRows.push([new Date(), 'Bank parser: ' + e.message, '', false, description]);
      errored++;
    }
  });

  const written = el_writeResults_(ss, rows, reviewRows);

  return { parsed, review, skipped, errored, counts: { total: written }, reviewDetails };
}


/***************************************************************************
 * HARDCODED GAME CONSTANTS (ändern sich nie)
 ***************************************************************************/

// Disziplinen — längste zuerst für endsWith-Matching
const EL_DISCIPLINES = ['Western Reining', 'Show Jumping', 'Flat Racing', 'Dressage', 'Driving', 'Endurance', 'Eventing'];

/** Extrahiert Discipline aus "{Breed} {Discipline}" Strings */
function el_extractDiscipline_(fullName) {
  const s = fullName.replace(/^All Breeds\s+/i, '');
  for (const d of EL_DISCIPLINES) {
    if (s.endsWith(d)) return d;
  }
  return s;
}


/***************************************************************************
 * CONTEXT BUILDER — loads lookups once
 ***************************************************************************/

function el_buildContext_(ss) {
  return {
    herd:     el_buildHerdLookup_(ss),
    breedMap: el_loadBreedMapFromSettings_(ss),
    rates:    el_loadCurrencyRatesFromSettings_(ss)
  };
}


/***************************************************************************
 * NOTIFICATION CLASSIFIER
 ***************************************************************************/

function el_classifyNotification_(text) {
  const t = text.replace(/^Horse Reality:\s*/i, '');

  // --- NOISE (kein $ Anker — Timestamp hängt am Ende) ---
  // Bewusst nicht protokollierte Meldungen; keine Sheet-Nebenwirkungen.
  if (/^(?:Bank:\s*)?You['’]ve received your daily interest of\s+[\d\s,.]+\s*HRC!/i.test(t)) return 'noise';
  if (/^Your stallion .+ after the clinical check by veterinarian Deloryan\./i.test(t)) return 'noise';
  if (/^Y?our horse .+ passed away peacefully in its sleep last night\./i.test(t)) return 'noise';
  if (/^You won the bidding on .+!\s*You paid /i.test(t)) return 'noise';
  if (/^You have received \d+ orbs? due to your premium membership!/i.test(t)) return 'noise';
  if (/^The .+ show has just ended/i.test(t)) return 'noise';
  if (/^The .+ competition has just ended/i.test(t)) return 'noise';
  if (/wasn.t sold on the market/i.test(t)) return 'noise';
  if (/created a Private trade/i.test(t)) return 'noise';
  if (/is here!.*we gifted you/i.test(t)) return 'noise';
  if (/^You can now breed with a private stud/i.test(t)) return 'noise';

  // --- BIRTH ---
  if (/^Congratulations!.+gave birth to /i.test(t)) return 'birth';

  // --- DELORYAN (Geschlechtsvorhersage) ---
  if (/^Deloryan determined the sex of the unborn foal from /i.test(t)) return 'deloryan';

  // --- COVERING FAILED ---
  if (/^The covering between .+ has failed/i.test(t)) return 'covering_failed';

  // --- FOAL PASTURES ---
  if (/^Kai:\s*Last night .+ at my Foal Pastures/i.test(t)) return 'foal_return';
  if (/will return from Kai.s Foal Pastures in /i.test(t)) return 'foal_notice';

  // --- EXCHANGE ORDER ---
  if (/^You received [\d\s,]+ hrc for order/i.test(t)) return 'exchange';

  // --- SALE ---
  if (/^Y?our horse .+ was sold to .+ on the market for /i.test(t)) return 'sale';

  // --- STUD INCOMING ---
  if (/^You received [\d\s,]+\s*(HRC|DP|FT|WT)?\s+as .+ used your stud service/i.test(t)) return 'stud_in';

  // --- STUD OUTGOING (with transport) ---
  if (/^You paid [\d\s,]+ HRC \+ [\d\s,]+ HRC for transport to /i.test(t)) return 'stud_out';

  // --- STUD OUTGOING (simple, without transport) ---
  if (/^You paid [\d\s,]+ HRC to .+ to breed .+ with the stud /i.test(t)) return 'stud_out_simple';

  // --- CONF DAY CHAMPION ---
  if (/became Day Champion in the /i.test(t)) return 'conf_daychamp';

  // --- CONF PLACEMENT ---
  if (/became (1st|2nd|3rd) in the .+ \(category:/i.test(t)) return 'conf';

  // --- COMP PLACEMENT ---
  if (/won (1st|2nd|3rd) Prize in the .+ of our .+ competition/i.test(t)) return 'comp';

  return 'unknown';
}


/***************************************************************************
 * NOTIFICATION ROUTER
 ***************************************************************************/

function el_parseNotification_(type, rawText, eventDate, horseLinks, ctx) {
  const t = rawText.replace(/^Horse Reality:\s*/i, '');

  switch (type) {
    case 'comp':             return el_parseComp_(t, eventDate, horseLinks, ctx, rawText);
    case 'conf':             return el_parseConf_(t, eventDate, horseLinks, ctx, rawText);
    case 'conf_daychamp':    return el_parseConfDayChamp_(t, eventDate, horseLinks, ctx, rawText);
    case 'stud_in':          return el_parseStudIn_(t, eventDate, horseLinks, ctx, rawText);
    case 'stud_out':         return el_parseStudOut_(t, eventDate, horseLinks, ctx, rawText);
    case 'stud_out_simple':  return el_parseStudOutSimple_(t, eventDate, horseLinks, ctx, rawText);
    case 'sale':             return el_parseSaleN_(t, eventDate, horseLinks, ctx, rawText);
    case 'birth':            return el_parseBirthN_(t, eventDate, rawText);
    case 'deloryan':         return el_parseDeloryan_(t, eventDate, horseLinks, ctx, rawText);
    case 'covering_failed':  return el_parseCoveringFailed_(t, eventDate, horseLinks, rawText);
    case 'foal_return':      return el_parseFoalReturn_(t, eventDate, horseLinks, ctx, rawText);
    case 'foal_notice':      return el_parseFoalNotice_(t, eventDate, horseLinks, ctx, rawText);
    case 'exchange':         return el_parseExchangeN_(t, eventDate, rawText);
    default:
      return { needsReview: true, reason: 'Unrecognized notification', suggestion: rawText };
  }
}


/***************************************************************************
 * NOTIFICATION PARSERS — each returns {row, needsReview?, reason?}
 * Row format: [Date, Type, HorseID, HorseName, Detail, Detail2, Detail3, Currency, Amount, Fee, Net, RawText]
 ***************************************************************************/

function el_parseComp_(t, date, links, ctx, raw) {
  const m = t.match(/^Your horse (.+?) won (1st|2nd|3rd) Prize in the (.+?) of our (.+?) competition\.\s*You received ([\d\s,]+)!/i);
  if (!m) return { needsReview: true, reason: 'Comp regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  const amount = el_parseAmount_(m[5]);
  const discipline = el_extractDiscipline_(m[4].trim());
  return { needsReview: !horse.id, reason: horse.id ? '' : 'Horse not found: ' + m[1],
    row: [date, 'comp', horse.id, horse.name || m[1].trim(), discipline, m[3].trim(), m[2], 'HRC', amount, 0, amount, raw] };
}

function el_parseConf_(t, date, links, ctx, raw) {
  const m = t.match(/^Your horse (.+?) became (1st|2nd|3rd) in the (.+?) \(category:\s*(Foals|Mares|Stallions|Geldings)\)\.\s*You received ([\d\s,]+)!/i);
  if (!m) return { needsReview: true, reason: 'Conf regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  const amount = el_parseAmount_(m[5]);
  const breed = el_resolveBreedFromSociety_(m[3].trim(), ctx.breedMap);
  return { needsReview: !horse.id, reason: horse.id ? '' : 'Horse not found: ' + m[1],
    row: [date, 'conf', horse.id, horse.name || m[1].trim(), breed, m[4], m[2], 'HRC', amount, 0, amount, raw] };
}

function el_parseConfDayChamp_(t, date, links, ctx, raw) {
  const m = t.match(/^Your horse (.+?) became Day Champion in the (.+?)\.\s*You received ([\d\s,]+)!/i);
  if (!m) return { needsReview: true, reason: 'DayChamp regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  const amount = el_parseAmount_(m[3]);
  const breed = el_resolveBreedFromSociety_(m[2].trim(), ctx.breedMap);
  // Day Champion Rewards (Spielkonstante):
  //   1000 → HRC | 100 → DP | 5 → FT oder WT (mehrdeutig)
  let currency, needsReview = false, reason = '';
  if (amount === 1000)    { currency = 'HRC'; }
  else if (amount === 100) { currency = 'DP'; }
  else if (amount === 5)  { currency = 'FT'; needsReview = true; reason = 'DayChamp: 5 FT oder WT — manuell prüfen'; }
  else                     { currency = 'HRC'; needsReview = true; reason = 'DayChamp: unbekannter Betrag ' + amount; }
  const earningsHRC = el_toHRC_(amount, currency, ctx.rates);
  return { needsReview, reason,
    row: [date, 'conf', horse.id, horse.name || m[1].trim(), breed, '', 'Day Champion', currency, amount, 0, earningsHRC, raw] };
}

function el_parseStudIn_(t, date, links, ctx, raw) {
  const m = t.match(/^You received ([\d\s,]+)\s*(HRC|DP|FT|WT)?\s+as (.+?) used your stud service to breed (.+?) with your stud (.+?)\./i);
  if (!m) return { needsReview: true, reason: 'Stud-in regex failed' };
  const amount = el_parseAmount_(m[1]);
  const currency = (m[2] || 'HRC').toUpperCase();
  let studName = el_normalizeHorseName_(m[5].trim());
  let studId = '';
  if (links && links.length >= 2) {
    const resolved = el_lookupHorseById_(links[1].id, ctx.herd);
    studId = resolved.id || links[1].id;
    studName = resolved.name || el_normalizeHorseName_(links[1].text);
  }
  const net = el_toHRC_(amount, currency, ctx.rates);
  return { needsReview: false,
    row: [date, 'stud_in', studId, studName, m[3].trim(), el_normalizeHorseName_(m[4].trim()), '', currency, amount, 0, net, raw] };
}

function el_parseStudOut_(t, date, links, ctx, raw) {
  const m = t.match(/^You paid ([\d\s,]+) HRC \+ ([\d\s,]+) HRC for transport to (.+?) to breed (.+?) with the stud (.+?)\./i);
  if (!m) return { needsReview: true, reason: 'Stud-out regex failed' };
  const fee = el_parseAmount_(m[1]);
  const transport = el_parseAmount_(m[2]);
  let mareName = el_normalizeHorseName_(m[4].trim()), mareId = '';
  if (links && links.length >= 1) {
    const resolved = el_lookupHorseById_(links[0].id, ctx.herd);
    mareId = resolved.id || links[0].id;
    mareName = resolved.name || el_normalizeHorseName_(links[0].text);
  }
  const total = -(fee + transport);
  return { needsReview: false,
    row: [date, 'stud_out', mareId, mareName, m[3].trim(), el_normalizeHorseName_(m[5].trim()), transport, 'HRC', total, 0, total, raw] };
}

function el_parseStudOutSimple_(t, date, links, ctx, raw) {
  const m = t.match(/^You paid ([\d\s,]+) HRC to (.+?) to breed (.+?) with the stud (.+?)\./i);
  if (!m) return { needsReview: true, reason: 'Stud-out-simple regex failed' };
  const fee = el_parseAmount_(m[1]);
  let mareName = el_normalizeHorseName_(m[3].trim()), mareId = '';
  if (links && links.length >= 1) {
    const resolved = el_lookupHorseById_(links[0].id, ctx.herd);
    mareId = resolved.id || links[0].id;
    mareName = resolved.name || el_normalizeHorseName_(links[0].text);
  }
  return { needsReview: false,
    row: [date, 'stud_out', mareId, mareName, m[2].trim(), el_normalizeHorseName_(m[4].trim()), 0, 'HRC', -fee, 0, -fee, raw] };
}

function el_parseSaleN_(t, date, links, ctx, raw) {
  const m = t.match(/^Y?our horse (.+?) was sold to (.+?) on the market for ([\d\s,]+) HRC/i);
  if (!m) return { needsReview: true, reason: 'Sale regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  const price = el_parseAmount_(m[3]);
  return { needsReview: false,
    row: [date, 'sale', horse.id, horse.name || m[1].trim(), m[2].trim(), '', '', 'HRC', price, 0, price, raw] };
}

function el_parseBirthN_(t, date, raw) {
  const m = t.match(/^Congratulations! A new (mare|stallion) was born\.\s*(.+?) gave birth to Foal Doe (\d+)/i);
  if (!m) return { needsReview: true, reason: 'Birth regex failed' };
  return { needsReview: false,
    row: [date, 'birth', '', el_normalizeHorseName_(m[2].trim()), m[3], m[1], '', '', 0, 0, 0, raw] };
}

function el_parseDeloryan_(t, date, links, ctx, raw) {
  const m = t.match(/^Deloryan determined the sex of the unborn foal from (.+?)\.\s*It is going to be a (Mare|Stallion)/i);
  if (!m) return { needsReview: true, reason: 'Deloryan regex failed' };

  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  const predictedSex = m[2].trim();
  // Umwandlung: "Mare" → "Filly", "Stallion" → "Colt"
  const genderLabel = predictedSex.toLowerCase() === 'mare' ? 'Filly' : 'Colt';

  return {
    needsReview: false,
    row: [date, 'deloryan', horse.id, horse.name || el_normalizeHorseName_(m[1].trim()), genderLabel, predictedSex, '', '', 0, 0, 0, raw],
    sideEffect: 'deloryan',
    mareId: horse.id,
    mareName: horse.name || el_normalizeHorseName_(m[1].trim()),
    genderLabel: genderLabel
  };
}

function el_parseCoveringFailed_(t, date, links, raw) {
  const m = t.match(/^The covering between (.+?) and (.+?) has failed/i);
  if (!m) return { needsReview: true, reason: 'Covering regex failed' };
  const mare = el_normalizeHorseName_(m[1].trim());
  const stallion = el_normalizeHorseName_(m[2].trim());
  // Try to resolve mare ID from links (first link is typically the mare)
  let mareId = '';
  if (links && links.length >= 1) mareId = links[0].id || '';
  return { needsReview: false,
    row: [date, 'covering_failed', mareId, mare, stallion, '', '', '', 0, 0, 0, raw] };
}

function el_parseFoalReturn_(t, date, links, ctx, raw) {
  const m = t.match(/^Kai:\s*Last night (.+?) became \d+ years? old at my Foal Pastures/i);
  if (!m) return { needsReview: true, reason: 'Foal return regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  return { needsReview: false,
    row: [date, 'foal_pastures', horse.id, horse.name || el_normalizeHorseName_(m[1].trim()), 'Returned', '', '', '', 0, 0, 0, raw] };
}

function el_parseFoalNotice_(t, date, links, ctx, raw) {
  const m = t.match(/^Your horse (.+?) will return from Kai.s Foal Pastures in (\d+) day/i);
  if (!m) return { needsReview: true, reason: 'Foal notice regex failed' };
  const horse = el_resolveHorse_(links && links[0], m[1].trim(), ctx.herd);
  return { needsReview: false,
    row: [date, 'foal_pastures', horse.id, horse.name || el_normalizeHorseName_(m[1].trim()), 'Notice', m[2] + ' days', '', '', 0, 0, 0, raw] };
}

function el_parseExchangeN_(t, date, raw) {
  const m = t.match(/^You received ([\d\s,]+) hrc for order \(([^)]+)\)/i);
  if (!m) return { needsReview: true, reason: 'Exchange regex failed' };
  const amount = el_parseAmount_(m[1]);
  return { needsReview: false,
    row: [date, 'exchange', '', '', m[2], '', '', 'HRC', amount, 0, amount, raw] };
}


/***************************************************************************
 * BANK LOG CLASSIFIER
 ***************************************************************************/

function el_classifyBankEntry_(desc, bankType) {
  // --- SKIP: covered by notification parser ---
  if (/You won [\d\s,]+ HRC in the .+ conformation show/i.test(desc)) return 'skip';
  if (/You won [\d\s,]+ HRC in the .+ competition/i.test(desc)) return 'skip';
  if (/was sold to .+ on the market for/i.test(desc)) return 'skip';
  if (/used your stud service to breed/i.test(desc)) return 'skip';
  if (/became Day Champion/i.test(desc)) return 'skip';
  if (/You paid [\d\s,]+ HRC to .+ to breed/i.test(desc)) return 'skip';
  if (/You paid [\d\s,]+ HRC \+ [\d\s,]+ HRC for transport to .+ to breed/i.test(desc)) return 'skip';

  // --- PROCESS ---
  if (/earned .+ HRC for the chores/i.test(desc)) return 'chores';
  if (/daily interest of/i.test(desc)) return 'skip'; // Zinsen werden bewusst nicht gespeichert.
  if (/paid a fee to enter .+ in the show of/i.test(desc)) return 'show_fee';
  if (/paid a fee to enter .+ in the .+ competition/i.test(desc)) return 'comp_fee';
  if (/colour test\(s\) on/i.test(desc)) return 'colour_test';
  if (/left .+ at Kai.s Foal Pastures/i.test(desc)) return 'foal_pastures';
  if (/made a deposit to your savings/i.test(desc)) return 'savings_in';
  if (/made a withdrawal from your savings/i.test(desc)) return 'savings_out';
  if (/applied the following task: Ultrasound on/i.test(desc)) return 'ultrasound';
  if (/You have bought .+ from/i.test(desc)) return 'purchase';
  if (/was refunded because/i.test(desc)) return 'refund';
  if (/Received from exchange order/i.test(desc)) return 'exchange';
  if (/You received .+ hrc for order/i.test(desc)) return 'exchange';

  return 'noise';
}


/***************************************************************************
 * BANK LOG PARSER ROUTER
 ***************************************************************************/

function el_parseBankEntry_(type, desc, bankAmount, bankType, date, ctx, rawDesc) {
  const isOut = bankType === 'OUT' || bankType === 'INT OUT';
  const signedAmount = isOut ? -Math.abs(bankAmount) : Math.abs(bankAmount);

  switch (type) {
    case 'chores':
      return { needsReview: false,
        row: [date, 'chores', '', '', '', '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };

    case 'interest':
      return { needsReview: false,
        row: [date, 'interest', '', '', '', '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };

    case 'show_fee': {
      const m = desc.match(/enter (\d+) horses? in the show of the (.+?)\.?$/i);
      return { needsReview: false,
        row: [date, 'show_fee', '', '', m ? m[2] : '', m ? m[1] : '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'comp_fee': {
      const m = desc.match(/enter (\d+) horses? in the (.+?) competition\.?$/i);
      return { needsReview: false,
        row: [date, 'comp_fee', '', '', m ? m[2] : '', m ? m[1] : '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'colour_test': {
      const m = desc.match(/done (\d+) colour test\(s\) on (.+?) for ([\d\s,]+) HRC/i);
      return { needsReview: false,
        row: [date, 'colour_test', '', m ? el_normalizeHorseName_(m[2]) : '', '', m ? m[1] : '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'foal_pastures': {
      const m = desc.match(/left (.+?) at Kai.s Foal Pastures for ([\d\s,]+) HRC/i);
      const horse = m ? el_resolveHorse_(null, m[1].trim(), ctx.herd) : { id: '', name: '' };
      return { needsReview: false,
        row: [date, 'foal_pastures', horse.id, horse.name || (m ? el_normalizeHorseName_(m[1]) : ''), 'Left', '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'savings_in':
      return { needsReview: false,
        row: [date, 'savings_in', '', '', '', '', '', 'HRC', Math.abs(bankAmount), 0, 0, rawDesc] }; // Net=0, internal transfer

    case 'savings_out':
      return { needsReview: false,
        row: [date, 'savings_out', '', '', '', '', '', 'HRC', Math.abs(bankAmount), 0, 0, rawDesc] }; // Net=0, internal transfer

    case 'ultrasound': {
      const m = desc.match(/(.+?) applied the following task: Ultrasound on (.+?)\.?$/i);
      const horse = m ? el_resolveHorse_(null, m[2].trim(), ctx.herd) : { id: '', name: '' };
      return { needsReview: false,
        row: [date, 'ultrasound', horse.id, horse.name || (m ? el_normalizeHorseName_(m[2]) : ''), m ? m[1] : '', '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'purchase': {
      // "You have bought X from Y for N HRC." or "...for N HRC and paid an additional M HRC for transport."
      const m1 = desc.match(/bought (.+?) from (.+?) for ([\d\s,]+) HRC and paid an additional ([\d\s,]+) HRC for transport/i);
      const m2 = desc.match(/bought (.+?) from (.+?) for ([\d\s,]+) HRC/i);
      const m = m1 || m2;
      if (!m) return { needsReview: true, reason: 'Purchase regex failed' };
      const transport = m1 ? el_parseAmount_(m1[4]) : 0;
      return { needsReview: false,
        row: [date, 'purchase', '', el_normalizeHorseName_(m[1].trim()), m[2].trim(), '', transport || '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'refund': {
      const m = desc.match(/fee to enter (.+?) in the (.+?) of the (.+?) competition was refunded/i);
      return { needsReview: false,
        row: [date, 'refund', '', m ? el_normalizeHorseName_(m[1]) : '', m ? m[3] : '', m ? m[2] : '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    case 'exchange': {
      const m = desc.match(/exchange order \(([^)]+)\)/i);
      return { needsReview: false,
        row: [date, 'exchange', '', '', m ? m[1] : '', '', '', 'HRC', signedAmount, 0, signedAmount, rawDesc] };
    }

    default:
      return { needsReview: true, reason: 'Unrecognized bank entry: ' + type };
  }
}


/***************************************************************************
 * COVERING FAILED — increment counter in Broodmares
 ***************************************************************************/

function el_incrementCoveringFailed_(ss, mareId, mareName) {
  const sheet = ss.getSheetByName(AL.BROODMARES);
  if (!sheet) return;

  // Find "Failed Coverings" column header — flexible position
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let failCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().includes('failed')) { failCol = i + 1; break; }
  }
  if (failCol === -1) return; // Column doesn't exist yet — user needs to add it

  // Find mare row by ID (col 3) or name (col 4)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 3, lastRow - 1, 2).getValues(); // C:D = ID, Name

  for (let i = 0; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    const rowName = String(data[i][1]).trim();
    if ((mareId && rowId === String(mareId)) || el_normalizeHorseName_(rowName).toLowerCase() === el_normalizeHorseName_(mareName).toLowerCase()) {
      const cell = sheet.getRange(i + 2, failCol);
      const current = Number(cell.getValue()) || 0;
      cell.setValue(current + 1);
      return;
    }
  }
}


/***************************************************************************
 * BROODMARES AUTOMATION — Ultrasound & Birth Side Effects
 ***************************************************************************/

/**
 * Ultrasound geparst → Stute automatisch auf "Confirmed" setzen (Spalte B).
 */
function el_autoConfirmOnUltrasound_(ss, horseId, horseName) {
  const sheet = ss.getSheetByName(AL.BROODMARES);
  if (!sheet) return;

  const row = el_findMareRow_(sheet, horseId, horseName);
  if (row === -1) return;

  sheet.getRange(row, 2).setValue(true); // Spalte B = Confirmed
}

/**
 * Geburt geparst → Stute zurücksetzen:
 * - "Foaled on" (Spalte N/14) = heutiges Datum
 * - "Gender" (Spalte L/12) = tatsächliches Geschlecht des Fohlens
 * - Confirmed (B/2) = FALSE
 * - Cover Date (F/6) = leer
 * - Cover Time (G/7) = leer
 * - Due Date (H/8) = leer
 * - Stud (M/13) = leer
 */
function el_resetMareOnBirth_(ss, mareName, foalGender) {
  const sheet = ss.getSheetByName(AL.BROODMARES);
  if (!sheet) return;

  const row = el_findMareRow_(sheet, '', mareName);
  if (row === -1) return;

  // Geschlecht umwandeln: "mare" → "Filly", "stallion" → "Colt"
  let genderLabel = '';
  if (foalGender) {
    const g = foalGender.toLowerCase().trim();
    if (g === 'mare') genderLabel = 'Filly';
    else if (g === 'stallion') genderLabel = 'Colt';
  }

  // Foaled on + Gender eintragen
  sheet.getRange(row, 14).setValue(new Date()); // N = Foaled on
  if (genderLabel) sheet.getRange(row, 12).setValue(genderLabel); // L = Gender

  // Reset: Confirmed, Cover Date, Cover Time, Due Date, Stud
  sheet.getRange(row, 2).setValue(false);  // B = Confirmed
  sheet.getRange(row, 6).setValue('');     // F = Cover Date
  sheet.getRange(row, 7).setValue('');     // G = Cover Time
  sheet.getRange(row, 8).setValue('');     // H = Due Date
  sheet.getRange(row, 13).setValue('');    // M = Stud
}

/**
 * Hilfsfunktion: Stute in Broodmares finden (gibt Zeilennummer zurück, -1 wenn nicht gefunden).
 */
function el_findMareRow_(sheet, horseId, horseName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const data = sheet.getRange(2, 3, lastRow - 1, 2).getValues(); // C:D = ID, Name

  for (let i = 0; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    const rowName = String(data[i][1]).trim();
    if ((horseId && rowId === String(horseId).trim()) ||
        el_normalizeHorseName_(rowName).toLowerCase() === el_normalizeHorseName_(horseName).toLowerCase()) {
      return i + 2; // Sheet row (1-indexed, skip header)
    }
  }
  return -1;
}


/**
 * Deloryan Geschlechtsvorhersage → Gender (Spalte L/12) in Broodmares eintragen.
 */
function el_writeDeloryanGender_(ss, horseId, horseName, genderLabel) {
  const sheet = ss.getSheetByName(AL.BROODMARES);
  if (!sheet) return;

  const row = el_findMareRow_(sheet, horseId, horseName);
  if (row === -1) return;

  sheet.getRange(row, 12).setValue(genderLabel); // Spalte L = Gender
}


/***************************************************************************
 * WRITE RESULTS
 ***************************************************************************/

function el_writeResults_(ss, rows, reviewRows) {
  let written = 0;
  if (rows.length) {
    const sheet = ss.getSheetByName(AL.LOG);
    if (sheet) {
      // Duplikat-Check: Hash aus Date+Type+HorseID+Amount+Raw
      const existingHashes = el_loadExistingHashes_(sheet);
      const uniqueRows = rows.filter(r => {
        const hash = el_rowHash_(r);
        if (existingHashes.has(hash)) return false;
        existingHashes.add(hash); // auch neue Rows untereinander deduplizieren
        return true;
      });
      if (uniqueRows.length) el_appendRows_(sheet, uniqueRows);
      written = uniqueRows.length;
    }
  }
  if (reviewRows.length) {
    const sheet = ss.getSheetByName(AL.REVIEW);
    if (sheet) el_appendRows_(sheet, reviewRows);
  }
  return written;
}

/**
 * Ereignisschlüssel: Datum, Typ, Pferd, Betrag und vollständiger Meldungstext.
 * Relative Datumswörter werden normalisiert; die Uhrzeit bleibt erhalten.
 */
function el_stableRaw_(raw) {
  return String(raw || '').replace(/(?:Today|Yesterday|\d{2}-\d{2}-\d{4})\s+at\s+(\d{1,2}:\d{2})/gi, 'at $1').replace(/\s+/g, ' ').trim();
}

function el_rowHash_(row) {
  const date = row[0] instanceof Date
    ? row[0].getFullYear() + '-' + (row[0].getMonth()+1) + '-' + row[0].getDate()
    : String(row[0]);
  const type = String(row[1] || '');
  const horseId = String(row[2] || '');
  const amount = String(row[8] || 0);
  const rawSnippet = el_stableRaw_(row[11]);
  return date + '|' + type + '|' + horseId + '|' + amount + '|' + rawSnippet;
}

/**
 * Alle vorhandenen Einträge berücksichtigen, auch bei älteren Paste-Blöcken.
 */
function el_loadExistingHashes_(sheet) {
  const hashes = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return hashes;

  const startRow = 2;
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 12).getValues();

  data.forEach(r => hashes.add(el_rowHash_(r)));
  return hashes;
}


/***************************************************************************
 * HELPERS — date, name, lookups, currency
 ***************************************************************************/

function el_parseEventDate_(text, dateString) {
  const source = (dateString && dateString.trim()) ? dateString : text;
  const dateMatch = source.match(/(\d{2})-(\d{2})-(\d{4})\s+at\s+(\d{1,2}):(\d{2})/);
  if (dateMatch) return new Date(parseInt(dateMatch[3]), parseInt(dateMatch[2]) - 1, parseInt(dateMatch[1]));
  if (/\bToday at /i.test(source)) return el_getGameDate_(0);
  if (/\bYesterday at /i.test(source)) return el_getGameDate_(-1);
  return el_getGameDate_(0);
}

function el_getGameDate_(dayOffset) {
  const tz = 'Europe/Berlin';
  const formatted = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const parts = formatted.split('-').map(n => parseInt(n, 10));
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  d.setDate(d.getDate() + dayOffset);
  return d;
}

function el_parseAmount_(s) {
  return parseInt(String(s).replace(/[\s,\.]/g, ''), 10) || 0;
}

function el_normalizeHorseName_(name) {
  let s = String(name);
  s = s.replace(/\s*\|\s*\d+(\.\d+)?\s*$/g, '');
  s = s.replace(/^[\s✤ʬ☆★•·✯⭐♦♠♥♣⚜︎˚₊✧⋆☽☾℘ℙ⚚]+/, '');
  s = s.replace(/[\s✤ʬ☆★•·✯⭐♦♠♥♣⚜︎˚₊✧⋆☽☾℘ℙ⚚]+$/, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function el_buildHerdLookup_(ss) {
  const byName = {}, byId = {};
  function register(rawName, id) {
    const clean = el_normalizeHorseName_(rawName);
    if (!clean) return;
    const stripped = clean.replace(/\s*\([^)]*\)\s*$/, '').trim();
    const display = stripped || clean;
    const idKey = String(id).trim();
    const entry = { id: idKey, name: display };
    const fullKey = clean.toLowerCase();
    const strippedKey = display.toLowerCase();
    if (!byName[fullKey]) byName[fullKey] = entry;
    if (!byName[strippedKey]) byName[strippedKey] = entry;
    if (idKey && !byId[idKey]) byId[idKey] = entry;
  }
  function ingest(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const last = sheet.getLastRow();
    if (last < 3) return;
    const data = sheet.getRange(3, 3, last - 2, 2).getValues();
    data.forEach(r => { if (r[0] && r[1]) register(r[1], r[0]); });
  }
  ingest(AL.HERD);
  ingest(AL.ARCHIVE);
  return { byName, byId };
}

function el_resolveHorse_(linkEntry, rawName, herdLookup) {
  if (linkEntry && linkEntry.id) {
    const found = herdLookup.byId[String(linkEntry.id).trim()];
    if (found) return found;
    return { id: linkEntry.id, name: el_normalizeHorseName_(linkEntry.text) };
  }
  const clean = el_normalizeHorseName_(rawName);
  return herdLookup.byName[clean.toLowerCase()] || { id: '', name: clean };
}

function el_lookupHorseById_(id, herdLookup) {
  if (!id) return { id: '', name: '' };
  return herdLookup.byId[String(id).trim()] || { id: '', name: '' };
}

function el_resolveBreedFromSociety_(society, confShowMap) {
  for (const key in confShowMap) {
    if (society.indexOf(key) !== -1) return confShowMap[key];
  }
  return society.replace(/\s*Horse Society.*$/i, '').replace(/\s*Society.*$/i, '').trim();
}

function el_toHRC_(amount, currency, rates) {
  if (currency === 'HRC') return amount;
  if (currency === 'DP') return amount * (rates.DP || 0);
  if (currency === 'FT') return amount * (rates.FT || 0);
  if (currency === 'WT') return amount * (rates.WT || 0);
  return amount;
}


/***************************************************************************
 * SETTINGS LOADERS — liest aus dem "Settings" Sheet
 ***************************************************************************/

/**
 * Breed Map aus Rasseliste: Spalte D (Rasse Verband / Society) → Spalte C (Rasse / Breed)
 * Contains-Match: "Finnhorse Society" → "Finnhorse"
 */
function el_loadBreedMapFromSettings_(ss) {
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  const map = {};
  const last = Math.min(sheet.getLastRow(), 40);
  if (last < 3) return map;
  const data = sheet.getRange(3, 3, last - 2, 2).getValues(); // C3:D(last)
  data.forEach(r => {
    const breed = String(r[0]).trim();
    const society = String(r[1]).trim();
    if (breed && society) map[society] = breed;
  });
  return map;
}

/**
 * Currency Rates aus Settings: Spalte M (Currency) → Spalte N (HRC Rate)
 * Sucht dynamisch nach DP/FT/WT in M1:N20
 */
function el_loadCurrencyRatesFromSettings_(ss) {
  const sheet = ss.getSheetByName('Settings');
  const rates = { DP: 0, FT: 0, WT: 0 };
  if (!sheet) return rates;
  const data = sheet.getRange('M1:N20').getValues();
  data.forEach(r => {
    const key = String(r[0]).trim().toUpperCase();
    if (rates.hasOwnProperty(key)) rates[key] = Number(r[1]) || 0;
  });
  return rates;
}


/***************************************************************************
 * HTML PARSING — notifications from pasted HTML
 ***************************************************************************/

function el_splitNotificationHtml_(htmlBlob) {
  const TS_RE = /((?:Today|Yesterday)\s+at\s+\d{1,2}:\d{2}|\d{2}-\d{2}-\d{4}\s+at\s+\d{1,2}:\d{2})/gi;
  const withDelims = String(htmlBlob).replace(TS_RE, '$1\u0001');
  return withDelims.split('\u0001').map(s => s.trim()).filter(Boolean)
    .map(chunk => ({ text: el_stripTags_(chunk).replace(/^\s*(?:Notifications\s*)?(?:Description\s*Date\s*)?(?=Horse Reality:|Your horse |Congratulations!|Bank:)/i, ''), links: el_extractHorseLinksFromHtml_(chunk) }))
    .filter(l => l.text && l.text.trim().length > 5);
}

function el_extractHorseLinksFromHtml_(htmlLine) {
  const results = [];
  const re = /<a\s+[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(htmlLine)) !== null) {
    const idMatch = m[1].match(/horsereality\.com\/horses\/(\d+)/i);
    if (idMatch) results.push({ id: idMatch[1], text: el_stripTags_(m[2]).trim() });
  }
  return results;
}

function el_stripTags_(html) {
  return String(html).replace(/<\/(?:td|th|tr|p|div)>|<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>').replace(/&#39;/gi, "'").replace(/&quot;/gi, '"').replace(/\s+/g, ' ').trim();
}


/***************************************************************************
 * SHEET HELPERS
 ***************************************************************************/

function el_appendRows_(sheet, rows) {
  if (!sheet || !rows.length) return;
  const startRow = el_lastDataRowByColumnA_(sheet) + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function el_lastDataRowByColumnA_(sheet) {
  const numRows = sheet.getMaxRows();
  const colA = sheet.getRange(1, 1, numRows, 1).getValues();
  for (let r = colA.length - 1; r >= 0; r--) {
    const v = colA[r][0];
    if (v !== '' && v !== null && typeof v !== 'undefined') return r + 1;
  }
  return 1;
}


/***************************************************************************
 * XP CYCLE EARNINGS — reads from Activity Log
 *
 * Schreibt Conf/Comp/Sales/Stud-Totals in XP Tracker (Q:T = col 17–20).
 * Bookkeeping: col BB (54) = letzte verarbeitete Activity Log Zeile.
 *
 * Breed/Discipline-Breakdown dynamisch aus Activity Log Daten.
 * XP Tracker Spalten U:AH (21–34) = Breed-Breakdown (Conf) + Disc-Breakdown (Comp).
 * Reihenfolge wird aus den Headers in Zeile 1 gelesen, nicht hardcoded.
 ***************************************************************************/

function el_fillXpCycleEarnings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const xp = ss.getSheetByName('XP Tracker');
  if (!xp) { SpreadsheetApp.getUi().alert('XP Tracker nicht gefunden.'); return; }

  const lastRow = el_lastDataRowByColumnA_(xp);
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Keine XP-Zeilen — erst Date/Level/Rank/XP eintragen.'); return; }

  const logSheet = ss.getSheetByName(AL.LOG);
  if (!logSheet) { SpreadsheetApp.getUi().alert('Activity Log nicht gefunden.'); return; }

  // Bookkeeping: letzte verarbeitete Activity-Log-Zeile (col BB = 54)
  const prevLogRow = (lastRow <= 2) ? 1 : (Number(xp.getRange(lastRow - 1, 54).getValue()) || 1);
  const curLogRow = el_lastDataRowByColumnA_(logSheet);

  let confTotal = 0, compTotal = 0, salesTotal = 0, studTotal = 0;
  const confByBreed = {};
  const compByDisc = {};

  if (curLogRow > prevLogRow) {
    // Activity Log: B=Type(1), E=Detail(4), I=Amount(8), K=Net(10)
    const data = logSheet.getRange(prevLogRow + 1, 1, curLogRow - prevLogRow, 11).getValues();
    data.forEach(r => {
      const type = String(r[1]).trim();
      const detail = String(r[4]).trim();
      const amount = Number(r[8]) || 0;
      const net = Number(r[10]) || 0;

      if (type === 'conf')    { confTotal += net; confByBreed[detail] = (confByBreed[detail] || 0) + net; }
      if (type === 'comp')    { compTotal += net; compByDisc[detail] = (compByDisc[detail] || 0) + net; }
      if (type === 'sale')    { salesTotal += amount; }
      if (type === 'stud_in') { studTotal += net; }
    });
  }

  // Totals → Q:T (col 17–20)
  xp.getRange(lastRow, 17, 1, 4).setValues([[confTotal, compTotal, salesTotal, studTotal]]);

  // Optional: Breed/Discipline breakdown → col 21 onward (U:AH)
  // Reads header labels from row 1 to match dynamically
  const breakdownStart = 21; // col U
  const breakdownEnd = 34;   // col AH
  const headerRange = xp.getRange(1, breakdownStart, 1, breakdownEnd - breakdownStart + 1).getValues()[0];
  const breakdownValues = headerRange.map(h => {
    const label = String(h).trim();
    if (!label) return 0;
    // Check breed map first, then discipline map
    if (confByBreed[label] !== undefined) return confByBreed[label];
    if (compByDisc[label] !== undefined) return compByDisc[label];
    return 0;
  });
  xp.getRange(lastRow, breakdownStart, 1, breakdownValues.length).setValues([breakdownValues]);

  // Bookkeeping update
  xp.getRange(lastRow, 54).setValue(curLogRow);

  SpreadsheetApp.getUi().alert(
    '✓ Cycle-Earnings eingetragen:\n\n' +
    'Conf: ' + confTotal + ' HRC\n' +
    'Comp: ' + compTotal + ' HRC\n' +
    'Sales: ' + salesTotal + ' HRC\n' +
    'Stud: ' + studTotal + ' HRC'
  );
}
