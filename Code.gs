/**
 * JOB APPLICATION TRACKER  (Google Apps Script + Gmail + Google Sheets)
 * --------------------------------------------------------------------
 * Reads your Gmail, finds application confirmations, rejections, interview
 * invitations and offers, and keeps a Google Sheet up to date automatically.
 *
 * SETUP (one time, about 5 minutes)
 *  1. Create a new Google Sheet (e.g. "Job Tracker").
 *  2. Extensions > Apps Script. Delete the sample code, paste this whole file, save.
 *  3. Reload the Sheet. A "Job Tracker" menu appears.
 *  4. Job Tracker > 1. Set up sheets            (approve the Gmail/Sheets permissions)
 *  5. Job Tracker > 2. Import history (backfill) (imports everything since BACKFILL_START;
 *     it runs in 5-minute chunks every 10 minutes until done, then stops by itself)
 *  6. Job Tracker > 3. Turn on daily auto-update (checks new emails every morning)
 *
 * WHAT IT WRITES
 *  - Applications : one row per application. Edit only the yellow columns
 *                   (Status override, Notes); everything else is rewritten on each run.
 *  - Events       : one row per email that was classified (audit trail, Gmail links).
 *  - Summary      : live counts and conversion rates (formulas).
 *
 * PRIVACY: read-only on Gmail. Nothing is sent, labelled or deleted unless you set
 * WEEKLY_DIGEST = true, which emails a short summary to yourself only.
 */

// ============================ CONFIG ==================================
const CONFIG = {
  BACKFILL_START: '2023/01/01',      // yyyy/mm/dd
  TIMEZONE: 'Europe/Berlin',
  DAILY_RUN_HOUR: 7,                 // 0-23, local time
  NO_REPLY_DAYS: 42,                 // after this many days with no reply = "No reply"
  MATCH_WINDOW_DAYS: 120,            // a rejection/interview is linked to an application seen within this window
  MAX_RUNTIME_MS: 5 * 60 * 1000,     // Apps Script hard limit is 6 min
  WEEKLY_DIGEST: false,              // true = email yourself a summary every Monday
  // Senders that are never applications (job alerts, newsletters). Add your own.
  EXCLUDE_SENDERS: [
    'jobalerts-noreply@linkedin.com', 'jobs-listings@linkedin.com', 'messages-noreply@linkedin.com',
    'jobagent.stepstone.de', 'jobalert.indeed.com', 'glassdoor.com', 'jobleads.com', 'instaffo.com',
    'careerfairy.io', 'mail.xing.com', 'news.xing.com', 'coursera.org', 'respondent.io'
  ],
  // Words that mark admissions/scholarships (not jobs). Matched against subject.
  EXCLUDE_SUBJECT: /(stipendium|scholarship|admission|zulassung|master'?s? program|m\.sc\.|financial aid|bestellung|order confirmation)/i
};

// German cities used to tag location/country when the email names a city.
const DE_CITIES = ['Berlin','Hamburg','München','Munich','Frankfurt','Köln','Cologne','Düsseldorf','Stuttgart',
  'Heilbronn','Neckarsulm','Mannheim','Heidelberg','Ludwigshafen','Karlsruhe','Freiburg','Saarbrücken','Darmstadt',
  'Nürnberg','Nuremberg','Leipzig','Dresden','Hannover','Bonn','Essen','Dortmund','Kiel','Bremen','Wiesbaden',
  'Mainz','Walldorf','Konstanz','Bielefeld','Münster','Augsburg','Regensburg','Ulm','Potsdam','Erlangen','Wolfsburg'];

// Applicant-tracking-system domains: the real employer is NOT this domain.
const ATS_DOMAINS = ['myworkday.com','successfactors.eu','successfactors.com','smartrecruiters.com','greenhouse.io',
  'greenhouse-mail.io','lever.co','personio.de','softgarden.io','recruitee.com','recruitee-email.com',
  'recruitee-mailbox.com','workablemail.com','join.com','onlyfy.jobs','teamtailor-mail.com','teamtailor.com',
  'ashbyhq.com','icims.com','icims.eu','jobvite.com','dvinci-hr.com','beesite.de','concludis.de','talentry.com',
  'recruitmail.com','hrworks.de','bewerbermanagement.net','mein-check-in.de','brassring.com','rippling.com',
  'pinpoint.email','hirehive.io','applytojob.com','linkedin.com','stepstone.de','indeed.com'];

// ============================ PATTERNS ================================
const RX = {
  offer: /(we are (pleased|happy|delighted) to offer|offer letter|job offer|vertragsangebot|wir freuen uns,? ihnen (eine|die) (stelle|position) anzubieten|zusage für die stelle)/i,
  interview: /(invite you (to|for) (an? |the |your )?(first |second |final |personal |video |online )?(interview|conversation|call|meeting)|invitation (to|for) (an? )?(interview|conversation)|interview invitation|schedule (an?|your) (interview|call)|einladung zu(m|r)? (vorstellungsgespräch|kennenlerngespräch|interview|gespräch)|laden (sie|dich) (gerne )?(zu einem|zum) (vorstellungsgespräch|kennenlerngespräch|interview|gespräch)|kennenlerntermin|home ?task|take-?home|case study|coding challenge|online assessment|video ?interview|next step(s)? in (the|our) (process|recruiting))/i,
  rejection: /(leider (müssen|können|haben) wir|leider mitteilen|leider nicht (berücksichtigen|weiter)|nicht weiter berücksichtigen|absage|anderen? kandidat|unfortunately|regret to inform|not (to )?(move|moving) forward|decided not to (proceed|pursue|move)|pursue other candidates|other candidates|not been successful|was not successful|position has been filled|nicht geklappt|haben uns für (eine|einen) andere)/i,
  confirmation: /(thank you for (your application|applying)|thanks for applying|vielen dank für (ihre|deine) bewerbung|danke für (ihre|deine) bewerbung|eingangsbestätigung|application (received|confirmed|submitted)|we (have )?received your application|your application was sent|bewerbung (ist )?(bei uns )?eingegangen|wir haben (ihre|deine) bewerbung erhalten|the first step is made|der erste schritt ist gemacht|your application (for|as|to|at)\b|(ihre|deine) bewerbung (als|auf|für|bei))/i,
  viewed: /your application (for .+ )?was viewed/i
};

// ============================ MENU ====================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Job Tracker')
    .addItem('1. Set up sheets', 'setupSheets')
    .addItem('2. Import history (backfill)', 'startBackfill')
    .addItem('3. Turn on daily auto-update', 'enableDailyTrigger')
    .addSeparator()
    .addItem('Run update now', 'runIncremental')
    .addItem('Stop all automatic runs', 'removeAllTriggers')
    .addToUi();
}

// ============================ SETUP ===================================
const APP_HEADERS = ['Key','Company','Job title','Location','Country','Channel','Applied on','Last update',
  'Status (auto)','Status override','Final status','Days since update','Emails','Gmail link','Notes'];
const EVT_HEADERS = ['Message ID','Thread ID','Date','From','Subject','Type','Company','Job title','Location',
  'Country','Channel','App key','Gmail link'];

function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const apps = getOrCreate_(ss, 'Applications', APP_HEADERS);
  getOrCreate_(ss, 'Events', EVT_HEADERS);
  apps.setFrozenRows(1);
  apps.getRange('J:J').setBackground('#FFF2CC');
  apps.getRange('O:O').setBackground('#FFF2CC');
  apps.getRange(1, 1, 1, APP_HEADERS.length).setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(
    ['Applied','No reply','Interviewing','Rejected','Rejected after interview','Offer','Hired','Withdrawn'], true)
    .setAllowInvalid(true).build();
  apps.getRange('J2:J').setDataValidation(rule);
  buildSummary_(ss);
  SpreadsheetApp.getActive().toast('Sheets ready. Next: Job Tracker > 2. Import history.');
}

function getOrCreate_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}

function buildSummary_(ss) {
  let sh = ss.getSheetByName('Summary');
  if (!sh) sh = ss.insertSheet('Summary', 0);
  sh.clear();
  const S = "Applications!$K$2:$K", C = "Applications!$E$2:$E", D = "Applications!$L$2:$L",
        Y = "Applications!$G$2:$G", CH = "Applications!$F$2:$F", B = "Applications!$B$2:$B";
  const rows = [
    ['Job application tracker', ''],
    ['Last updated', '=NOW()'],
    ['', ''],
    ['Metric', 'All', 'Germany', 'Other countries'],
    ['Applications', `=COUNTA(${B})`, `=COUNTIF(${C},"Germany*")`, `=B5-C5`],
    ['Interviewing (active)', `=COUNTIF(${S},"Interviewing")`, `=COUNTIFS(${S},"Interviewing",${C},"Germany*")`, '=B6-C6'],
    ['Offers / hired', `=COUNTIF(${S},"Offer")+COUNTIF(${S},"Hired")`,
      `=COUNTIFS(${S},"Offer",${C},"Germany*")+COUNTIFS(${S},"Hired",${C},"Germany*")`, '=B7-C7'],
    ['Rejected (any stage)', `=COUNTIF(${S},"Rejected*")`, `=COUNTIFS(${S},"Rejected*",${C},"Germany*")`, '=B8-C8'],
    ['  of which after interview', `=COUNTIF(${S},"Rejected after interview")`,
      `=COUNTIFS(${S},"Rejected after interview",${C},"Germany*")`, '=B9-C9'],
    ['Pending (no reply, < ' + CONFIG.NO_REPLY_DAYS + ' days)', `=COUNTIFS(${S},"Applied",${D},"<${CONFIG.NO_REPLY_DAYS}")`,
      `=COUNTIFS(${S},"Applied",${D},"<${CONFIG.NO_REPLY_DAYS}",${C},"Germany*")`, '=B10-C10'],
    ['No reply (' + CONFIG.NO_REPLY_DAYS + '+ days)', `=COUNTIFS(${S},"Applied",${D},">=${CONFIG.NO_REPLY_DAYS}")+COUNTIF(${S},"No reply")`,
      `=COUNTIFS(${S},"Applied",${D},">=${CONFIG.NO_REPLY_DAYS}",${C},"Germany*")+COUNTIFS(${S},"No reply",${C},"Germany*")`, '=B11-C11'],
    ['Interview conversion', '=IFERROR((B6+B7+B9)/B5,0)', '=IFERROR((C6+C7+C9)/C5,0)', '=IFERROR((D6+D7+D9)/D5,0)'],
    ['', ''],
    ['By year', 'Applications', 'Reached interview', 'Conversion']
  ];
  sh.getRange(1, 1, rows.length, 4).setValues(rows.map(r => r.concat(['', '', '']).slice(0, 4)));
  const yr0 = parseInt(CONFIG.BACKFILL_START.slice(0, 4), 10), yrNow = new Date().getFullYear() + 1;
  let r = rows.length + 1;
  for (let y = yr0; y <= yrNow; y++, r++) {
    const a = `DATE(${y},1,1)`, b = `DATE(${y + 1},1,1)`;
    sh.getRange(r, 1, 1, 4).setValues([[String(y),
      `=COUNTIFS(${Y},">="&${a},${Y},"<"&${b})`,
      `=COUNTIFS(${Y},">="&${a},${Y},"<"&${b},${S},"Interviewing")+COUNTIFS(${Y},">="&${a},${Y},"<"&${b},${S},"Rejected after interview")+COUNTIFS(${Y},">="&${a},${Y},"<"&${b},${S},"Offer")+COUNTIFS(${Y},">="&${a},${Y},"<"&${b},${S},"Hired")`,
      `=IFERROR(C${r}/B${r},0)`]]);
    sh.getRange(r, 4).setNumberFormat('0.0%');
  }
  r++;
  sh.getRange(r++, 1, 1, 2).setValues([['By channel', 'Applications']]);
  ['Company portal / ATS', 'LinkedIn Easy Apply', 'StepStone', 'Indeed', 'XING', 'Recruitment agency'].forEach(ch => {
    sh.getRange(r++, 1, 1, 2).setValues([[ch, `=COUNTIF(${CH},"${ch}")`]]);
  });
  sh.getRange('B12:D12').setNumberFormat('0.0%');
  sh.getRange('B2').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  sh.getRange('A4:D4').setBackground('#1F3A5F').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(rows.length, 1, 1, 4).setFontWeight('bold');
  sh.setColumnWidth(1, 280);
}

// ============================ TRIGGERS ================================
function enableDailyTrigger() {
  removeTriggersFor_('runIncremental');
  ScriptApp.newTrigger('runIncremental').timeBased().everyDays(1).atHour(CONFIG.DAILY_RUN_HOUR)
    .inTimezone(CONFIG.TIMEZONE).create();
  if (CONFIG.WEEKLY_DIGEST) {
    removeTriggersFor_('sendWeeklyDigest');
    ScriptApp.newTrigger('sendWeeklyDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(CONFIG.DAILY_RUN_HOUR + 1).inTimezone(CONFIG.TIMEZONE).create();
  }
  SpreadsheetApp.getActive().toast('Daily auto-update is on (' + CONFIG.DAILY_RUN_HOUR + ':00).');
}

function startBackfill() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('backfillCursor', CONFIG.BACKFILL_START);
  removeTriggersFor_('runBackfill');
  ScriptApp.newTrigger('runBackfill').timeBased().everyMinutes(10).create();
  runBackfill();
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getActive().toast('All automatic runs stopped.');
}

function removeTriggersFor_(fn) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === fn).forEach(t => ScriptApp.deleteTrigger(t));
}

// ============================ RUNNERS =================================
/** Imports history month by month; resumes where it stopped. */
function runBackfill() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const t0 = Date.now(), p = PropertiesService.getScriptProperties();
    let cursor = p.getProperty('backfillCursor');
    if (!cursor) { removeTriggersFor_('runBackfill'); return; }
    const state = loadState_();
    const today = new Date();
    while (Date.now() - t0 < CONFIG.MAX_RUNTIME_MS) {
      const start = parseYmd_(cursor);
      if (start > today) {
        p.deleteProperty('backfillCursor');
        p.setProperty('lastIncremental', fmt_(today));
        removeTriggersFor_('runBackfill');
        SpreadsheetApp.getActive().toast('History import finished.');
        break;
      }
      const end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
      const done = processWindow_(state, fmt_(start), fmt_(end), t0);
      if (!done) break;                        // out of time: redo this month next run (duplicates are skipped)
      cursor = fmt_(end);
      p.setProperty('backfillCursor', cursor);
    }
    saveState_(state);
  } finally { lock.releaseLock(); }
}

/** Daily run: emails since the last run (2-day overlap; duplicates are skipped). */
function runIncremental() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('lastIncremental')) {            // history not imported yet
    if (!props.getProperty('backfillCursor')) startBackfill();
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const t0 = Date.now(), p = PropertiesService.getScriptProperties();
    const last = p.getProperty('lastIncremental') || CONFIG.BACKFILL_START;
    const from = new Date(parseYmd_(last).getTime() - 2 * 864e5);
    const to = new Date(Date.now() + 864e5);
    const state = loadState_();
    const done = processWindow_(state, fmt_(from), fmt_(to), t0);
    saveState_(state);
    if (done) p.setProperty('lastIncremental', fmt_(new Date()));
  } finally { lock.releaseLock(); }
}

// ============================ CORE ====================================
function buildQueries_(after, before) {
  const excl = CONFIG.EXCLUDE_SENDERS.map(s => '-from:' + s).join(' ');
  const base = `after:${after} before:${before} -in:spam -in:trash -category:promotions -from:me ${excl}`;
  return [
    `${base} ("thank you for your application" OR "thank you for applying" OR "thanks for applying" OR "vielen Dank für Ihre Bewerbung" OR "danke für deine Bewerbung" OR "Eingangsbestätigung" OR "application received" OR "your application" OR "Ihre Bewerbung" OR "deine Bewerbung" OR "your application was sent")`,
    `${base} (leider OR unfortunately OR Absage OR "other candidates" OR "nicht berücksichtigen" OR "not to move forward" OR "nicht geklappt") (Bewerbung OR application OR position OR Stelle)`,
    `${base} (interview OR Vorstellungsgespräch OR Kennenlerngespräch OR Kennenlerntermin OR "case study" OR "home task" OR assessment OR offer OR Vertragsangebot) (Bewerbung OR application OR position OR Stelle OR role)`
  ];
}

/** Returns true if the whole window was processed, false if time ran out. */
function processWindow_(state, after, before, t0) {
  const seenThreads = {};
  for (const q of buildQueries_(after, before)) {
    for (let start = 0; ; start += 100) {
      if (Date.now() - t0 > CONFIG.MAX_RUNTIME_MS) return false;
      const threads = GmailApp.search(q, start, 100);
      if (!threads.length) break;
      for (const th of threads) {
        if (seenThreads[th.getId()]) continue;
        seenThreads[th.getId()] = true;
        for (const msg of th.getMessages()) {
          if (state.seen[msg.getId()]) continue;
          if (Date.now() - t0 > CONFIG.MAX_RUNTIME_MS) return false;
          const ev = classify_(msg);
          state.seen[msg.getId()] = true;
          if (ev) applyEvent_(state, ev);
        }
      }
      if (threads.length < 100) break;
    }
  }
  return true;
}

let ME_ = null;
function classify_(msg) {
  const from = msg.getFrom() || '';
  const fromLc = from.toLowerCase();
  if (ME_ === null) ME_ = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (ME_ && fromLc.indexOf(ME_) >= 0) return null;          // skip your own replies
  if (CONFIG.EXCLUDE_SENDERS.some(s => fromLc.indexOf(s) >= 0)) return null;
  const subject = msg.getSubject() || '';
  if (CONFIG.EXCLUDE_SUBJECT.test(subject)) return null;
  let body = '';
  try { body = (msg.getPlainBody() || '').slice(0, 4000); } catch (e) {}
  const text = subject + '\n' + body;

  let type = null;
  if (RX.offer.test(text)) type = 'OFFER';
  else if (RX.interview.test(text)) type = 'INTERVIEW';
  else if (RX.rejection.test(text)) type = 'REJECTION';
  else if (RX.confirmation.test(text)) type = 'APPLIED';
  else if (RX.viewed.test(subject)) type = 'VIEWED';
  if (!type) return null;
  // Confirmations often say "we may invite you to an interview": only trust interview wording outside boilerplate.
  if (type === 'INTERVIEW' && RX.confirmation.test(subject)) type = 'APPLIED';

  const isLinkedIn = fromLc.indexOf('linkedin.com') >= 0;
  const info = isLinkedIn ? parseLinkedIn_(subject, body) : parseGeneric_(subject, body, from);
  const country = countryOf_(info.location, text);
  return {
    id: msg.getId(), threadId: msg.getThread().getId(), date: msg.getDate(), from: from, subject: subject,
    type: type, company: info.company || 'Unknown', role: info.role || '', location: info.location || '',
    country: country, channel: channelOf_(fromLc, subject)
  };
}

function parseLinkedIn_(subject, body) {
  let company = '', role = '', location = '';
  let m = subject.match(/application was sent to (.+)$/i) || subject.match(/application (?:for|to) .+? (?:was viewed by|at) (.+?)\.?$/i);
  if (m) company = m[1].trim();
  m = subject.match(/application (?:for|to) (.+?) (?:was viewed by|at) /i);
  if (m) role = m[1].trim();
  // Body usually contains: "<Job title>\n<Company> · <City, Region, Country> (Hybrid)"
  const lines = body.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  for (let i = 1; i < lines.length; i++) {
    if (lines[i].indexOf(' · ') > 0) {
      const parts = lines[i].split(' · ');
      if (!company) company = parts[0];
      location = parts.slice(1).join(' · ').trim();
      if (!role && !/^(view|apply|applied)/i.test(lines[i - 1])) role = lines[i - 1];
      break;
    }
  }
  return { company: company, role: cleanRole_(role), location: location };
}

function parseGeneric_(subject, body, from) {
  let role = '', company = '';
  const roleRx = [
    /(?:bewerbung|application)\s+[-–:]\s+(?!.*bewerbung)(.+?)(?:\s+[-–]\s+\d{5}.*|\s*,\s*\d{5}.*)?$/i,
    /(?:application|bewerbung)\s+(?:for|as|als|auf)\s+(?:the\s+)?(?:position|role|stelle)?\s*(?:of|as)?\s*[„"']?(.+?)[“"']?(?:\s+(?:at|bei|in|with|to)\s+|\s+-\s+|\s*\||\s*\(\d|$)/i,
    /(?:position|stelle|role)[:\s]+[„"']?(.+?)[“"']?(?:\s+(?:at|bei)\s+|\s*\||$)/i,
    /applying (?:for|to) (?:the )?(.+?) (?:at|position|role)/i
  ];
  for (const rx of roleRx) { const m = subject.match(rx); if (m && m[1].length > 3) { role = m[1].split(' / ')[0]; break; } }
  const compRx = [
    /(?:applying to|application to|application at|bewerbung bei|applying at|thank you for applying to|interest in|join)\s+(?:the\s+)?(.+?)(?:[!.,|]|\s+-\s+|\s+as\s+|$)/i,
    /\s(?:at|bei)\s+(?:der\s+|the\s+)?(.+?)(?:[!.,|]|\s+-\s+|\s*\(|$)/i,
    /^([^-|]{2,40}?)\s+(?:-|\||\/\/)\s+/
  ];
  for (const rx of compRx) {
    const m = subject.match(rx);
    if (m && m[1].length > 1 && m[1].split(' ').length <= 5 &&
        !/(your|ihre|deine|position|application|bewerbung|eingangsbest|nicht|thank|danke|update|vielen|first step)/i.test(m[1])) {
      company = m[1].replace(/^join\s+/i, ''); break;
    }
  }
  if (!company) company = companyFromSender_(from);
  let location = '';
  const loc = (subject + ' ' + body.slice(0, 1500)).match(/(?:location|standort|arbeitsort|einsatzort)\s*[:\-]\s*([^\n,;]+)/i);
  if (loc) location = loc[1].trim();
  if (!location) {
    const c = DE_CITIES.find(x => new RegExp('\\b' + x + '\\b', 'i').test(subject + ' ' + role));
    if (c) location = c;
  }
  return { company: tidyCompany_(company), role: cleanRole_(role), location: location };
}

function companyFromSender_(from) {
  const name = (from.match(/^"?([^"<]+?)"?\s*</) || [])[1] || '';
  const email = ((from.match(/<([^>]+)>/) || [])[1] || from).toLowerCase();
  const domain = email.split('@')[1] || '';
  const local = email.split('@')[0] || '';
  const isAts = ATS_DOMAINS.some(d => domain.endsWith(d));
  if (isAts) {
    if (domain.endsWith('myworkday.com') && !/^(no-?reply|notifications?)$/.test(local)) return local; // e.g. acme@myworkday.com -> "acme"
    if (name && !/(no-?reply|notification|recruit|talent|careers?|jobs?|team|hr\b)/i.test(name)) return name;
    const m = domain.match(/^([a-z0-9-]+)-jobs\./) || local.match(/^([a-z0-9]+)-jobs$/);
    return m ? m[1] : '';
  }
  if (name && !/(no-?reply|notification|recruit|talent|careers?|jobs?|karriere|bewerbung|hr\b)/i.test(name)) return name;
  const parts = domain.split('.');
  return parts.length >= 2 ? parts[parts.length - 2] : '';
}

function tidyCompany_(c) {
  return (c || '').replace(/[​‌‍​]/g, '').replace(/\s+/g, ' ')
    .replace(/^(the|die|der)\s+/i, '').replace(/\s*(careers?|recruiting|talent acquisition|jobs)$/i, '').trim();
}

function cleanRole_(r) {
  return (r || '').replace(/[​​]/g, '')
    .replace(/\((?:m|w|f|d|x|nb|gn\*?|div\.?|all genders|alle geschlechter)(?:[\/|,]\s*(?:m|w|f|d|x|nb|div\.?))*\)/gi, '')
    .replace(/\s+/g, ' ').replace(/^[\s\-–:„"']+|[\s\-–:“"'.!]+$/g, '').trim();
}

function countryOf_(location, text) {
  const loc = location || '';
  if (/germany|deutschland/i.test(loc) || DE_CITIES.some(c => new RegExp('\\b' + c + '\\b', 'i').test(loc))) return 'Germany';
  if (loc && loc.indexOf(',') > 0) return loc.split(',').pop().replace(/\(.*?\)/g, '').trim();
  if (loc) return loc;
  if (/\((?:m\/w\/d|w\/m\/d|m\/f\/d|f\/m\/d|d\/f\/m|gn\*?)\)|bewerbung|vielen dank/i.test(text)) return 'Germany (likely)';
  return 'Unknown';
}

function channelOf_(fromLc, subject) {
  if (fromLc.indexOf('linkedin.com') >= 0) return 'LinkedIn Easy Apply';
  if (fromLc.indexOf('stepstone') >= 0) return 'StepStone';
  if (fromLc.indexOf('indeed') >= 0) return 'Indeed';
  if (fromLc.indexOf('xing') >= 0) return 'XING';
  if (/(michaelpage|hays|robert ?half|brunel|randstad|harnham|recruitment|personalberatung)/.test(fromLc + subject.toLowerCase())) return 'Recruitment agency';
  return 'Company portal / ATS';
}

// ============================ MATCHING ================================
function norm_(s) {
  return (s || '').toLowerCase()
    .replace(/\b(gmbh|ag|se|kg|co|inc|ltd|limited|llc|plc|group|gruppe|holding|b\.?v\.?|s\.?a\.?|deutschland|germany)\b/g, '')
    .replace(/[^a-z0-9äöüß]+/g, ' ').trim();
}

function applyEvent_(state, ev) {
  const nc = norm_(ev.company), nr = norm_(ev.role);
  let app = null;
  if (nc && nc !== 'unknown') {
    const cands = state.apps.filter(a => a.nc === nc && Math.abs(ev.date - a.lastDate) <= CONFIG.MATCH_WINDOW_DAYS * 864e5);
    if (ev.type === 'APPLIED' || ev.type === 'VIEWED') {
      // a new confirmation is the same application only if the role matches (or one side has no role) and it is recent
      app = cands.find(a => (nr && a.nr === nr) || ((!nr || !a.nr) && Math.abs(ev.date - a.firstDate) <= 21 * 864e5)) || null;
    } else {
      app = cands.find(a => nr && a.nr === nr) ||
            cands.sort((x, y) => y.lastDate - x.lastDate)[0] || null;
    }
  } else {
    app = state.apps.find(a => a.threads.indexOf(ev.threadId) >= 0) || null;   // same thread, unknown company
  }
  if (!app) {
    app = { key: Utilities.getUuid().slice(0, 8), company: ev.company, nc: nc, role: ev.role, nr: nr,
      location: ev.location, country: ev.country, channel: ev.channel, firstDate: ev.date, lastDate: ev.date,
      interview: false, rejected: false, offer: false, emails: 0, threads: [], lastThread: ev.threadId,
      override: '', notes: '', isNew: true };
    state.apps.push(app);
  }
  if (!app.role && ev.role) { app.role = ev.role; app.nr = nr; }
  if (!app.location && ev.location) { app.location = ev.location; app.country = ev.country; }
  if ((app.country === 'Unknown' || app.country === 'Germany (likely)') && ev.country !== 'Unknown') app.country = ev.country;
  if (app.company === 'Unknown' && ev.company !== 'Unknown') { app.company = ev.company; app.nc = nc; }
  if (ev.date < app.firstDate) app.firstDate = ev.date;
  if (ev.date >= app.lastDate) { app.lastDate = ev.date; app.lastThread = ev.threadId; }
  if (ev.type === 'INTERVIEW') app.interview = true;
  if (ev.type === 'REJECTION') app.rejected = true;
  if (ev.type === 'OFFER') app.offer = true;
  app.emails++;
  if (app.threads.indexOf(ev.threadId) < 0) app.threads.push(ev.threadId);
  ev.appKey = app.key;
  state.newEvents.push(ev);
}

function statusOf_(a) {
  if (a.offer) return 'Offer';
  if (a.rejected) return a.interview ? 'Rejected after interview' : 'Rejected';
  if (a.interview) return 'Interviewing';
  return 'Applied';
}

// ============================ STATE I/O ===============================
function loadState_() {
  const ss = SpreadsheetApp.getActive();
  const appSh = ss.getSheetByName('Applications'), evtSh = ss.getSheetByName('Events');
  if (!appSh || !evtSh) throw new Error('Run "Job Tracker > 1. Set up sheets" first.');
  const state = { apps: [], seen: {}, newEvents: [] };
  const ev = evtSh.getLastRow() > 1 ? evtSh.getRange(2, 1, evtSh.getLastRow() - 1, EVT_HEADERS.length).getValues() : [];
  ev.forEach(r => { state.seen[r[0]] = true; });
  const byKey = {};
  const av = appSh.getLastRow() > 1 ? appSh.getRange(2, 1, appSh.getLastRow() - 1, APP_HEADERS.length).getValues() : [];
  av.forEach(r => {
    const a = { key: String(r[0]), company: r[1], nc: norm_(r[1]), role: r[2], nr: norm_(r[2]), location: r[3],
      country: r[4], channel: r[5], firstDate: new Date(r[6]), lastDate: new Date(r[7]),
      interview: false, rejected: false, offer: false, emails: Number(r[12]) || 0, threads: [], lastThread: '',
      override: r[9], notes: r[14], isNew: false };
    const st = String(r[8]);
    a.offer = st === 'Offer'; a.rejected = st.indexOf('Rejected') === 0;
    a.interview = st === 'Interviewing' || st === 'Rejected after interview';
    byKey[a.key] = a; state.apps.push(a);
  });
  ev.forEach(r => { const a = byKey[String(r[11])]; if (a && a.threads.indexOf(r[1]) < 0) { a.threads.push(r[1]); a.lastThread = r[1]; } });
  return state;
}

function saveState_(state) {
  const ss = SpreadsheetApp.getActive();
  const evtSh = ss.getSheetByName('Events'), appSh = ss.getSheetByName('Applications');
  if (state.newEvents.length) {
    const rows = state.newEvents.map(e => [e.id, e.threadId, e.date, e.from, e.subject, e.type, e.company, e.role,
      e.location, e.country, e.channel, e.appKey, 'https://mail.google.com/mail/u/0/#all/' + e.threadId]);
    evtSh.getRange(evtSh.getLastRow() + 1, 1, rows.length, EVT_HEADERS.length).setValues(rows);
  }
  state.apps.sort((a, b) => b.firstDate - a.firstDate);
  const out = state.apps.map((a, i) => {
    const r = i + 2;
    return [a.key, a.company, a.role, a.location, a.country, a.channel, a.firstDate, a.lastDate, statusOf_(a),
      a.override || '', `=IF(J${r}<>"",J${r},IF(AND(I${r}="Applied",L${r}>=${CONFIG.NO_REPLY_DAYS}),"No reply",I${r}))`,
      `=INT(TODAY()-H${r})`, a.emails,
      a.lastThread ? `=HYPERLINK("https://mail.google.com/mail/u/0/#all/${a.lastThread}","Open")` : '', a.notes || ''];
  });
  if (appSh.getLastRow() > 1) appSh.getRange(2, 1, appSh.getLastRow() - 1, APP_HEADERS.length).clearContent();
  if (out.length) {
    appSh.getRange(2, 1, out.length, APP_HEADERS.length).setValues(out);
    appSh.getRange(2, 7, out.length, 2).setNumberFormat('yyyy-mm-dd');
  }
}

// ============================ DIGEST (optional) =======================
function sendWeeklyDigest() {
  if (!CONFIG.WEEKLY_DIGEST) return;
  const sh = SpreadsheetApp.getActive().getSheetByName('Summary');
  const v = sh.getRange('A4:D12').getDisplayValues();
  const lines = v.map(r => r.filter(String).join('  |  ')).join('\n');
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Weekly job application summary',
    lines + '\n\nOpen the tracker: ' + SpreadsheetApp.getActive().getUrl());
}

// ============================ UTILS ===================================
function fmt_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy/MM/dd'); }
function parseYmd_(s) { const p = s.split('/').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
