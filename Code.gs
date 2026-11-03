/**
 * NUDGE - auto follow-ups for Gmail + Google Sheets
 *
 * WHAT IT NEEDS: a date column, an email column, and (optionally) a follow-up column.
 * Columns are found by name, so order doesn't matter and other columns are ignored.
 * If there is no follow-up column, one called NEXT_FOLLOW is created for you.
 *
 * ONE-TIME SETUP: Services (+) > Gmail API > Add, then run `setup` once.
 * Start with MODE = 'DRAFT' to test. Switch to 'SEND' when you're happy.
 *
 * WHAT IT DOES
 *  1. New row (DATE today/yesterday) and follow-up cell empty -> fills it with DATE + FOLLOW_UP_DAYS
 *  2. When the follow-up date == today it opens your sent thread in Gmail:
 *       - anyone other than you in that thread  -> writes REPLIED, no follow-up
 *       - a delivery-failure notice in it       -> writes BOUNCED, no follow-up
 *       - nobody else                           -> replies inside the thread, writes DONE
 *  3. Passed dates are ignored. One send at a time, random 3-5 min wait between sends.
 */

// ================== CONFIG ==================
const SHEET_NAME     = 'Sheet1';         // tab name. If not found, the first tab is used
const TIMEZONE       = 'Asia/Karachi';   // CHANGE THIS to your timezone, e.g. 'America/New_York'
const FOLLOW_UP_DAYS = 4;                // gap after DATE for auto-filled dates
const MODE           = 'DRAFT';          // 'DRAFT' = only creates Gmail drafts (testing), 'SEND' = sends for real
const DATE_ORDER     = 'DMY';            // for ambiguous dates like 12/10/2026: 'DMY' = 12 Oct, 'MDY' = Dec 10
const GAP_MINUTES    = [3, 5];           // random wait between sends
const SCAN_HOURS     = 3;                // how often to check upcoming leads for replies (so REPLIED shows early)

// ============ COLUMN DETECTION ============
// Optional: force a column by its letter if auto-detect ever fails, e.g. { DATE: 'A', EMAIL: 'C', NEXT_FOLLOW: 'F' }
const COLUMNS = { DATE: '', EMAIL: '', NEXT_FOLLOW: '' };

const ALIASES = {
  DATE:        ['date', 'datesent', 'sentdate', 'senton', 'dateemailed', 'emaildate', 'contacted', 'contactdate'],
  EMAIL:       ['email', 'emailaddress', 'websiteemail', 'emailwebsite', 'contactemail', 'mail'],
  NEXT_FOLLOW: ['nextfollow', 'nextfollowup', 'followup', 'followupdate', 'nextfollowupdate', 'followupon']
};

const FOLLOWUPS = [
"Hey, just checking in on this. Any thoughts?",
"Hey! Wanted to follow up in case my last message got buried.",
"Hey, hope your week’s going well! Just wanted to see if you had a chance to think about this.",
"Hey, no rush at all. Just wanted to bring this back to your attention.",
"Hey, curious if this is still something you’re looking into.",
"Just wanted to check if this is still on your radar.",
"Hey! Would love to hear what you think whenever you get a chance.",
"Hey, I know things get busy. Just bumping this up in your inbox.",
"Hey, circling back on this. Let me know what you think!",
"Hi! Just floating this back to the top of your inbox.",
"Hey, wanted to make sure this didn’t slip through the cracks.",
"Hope you’re having a good week! Any thoughts on my last message?",
"Hey, just a quick nudge on this. Would love your take.",
"Hi, following up on my previous email. Does this sound relevant to you?",
"Hey! No pressure, just wanted to see if you had a chance to look at this.",
"Hey, just resurfacing this in case it got lost in the shuffle.",
"Hi, checking back in. Even a quick reply would be super helpful!",
"Hey, wanted to circle back before this gets too buried. Thoughts?",
"Hey, hope all is well! Just checking if you got a chance to see my note.",
"Hi! Last quick nudge from me. Would love to hear your thoughts either way.",
"Hey, just popping back into your inbox. Any thoughts on this?",
"Hi! Wanted to see if you had a moment to think about my last note.",
"Hey, totally understand if you’re swamped. Just wanted to keep this on your radar.",
"Hey there, checking in to see if this is something worth exploring.",
"Hi, just a friendly nudge on my previous message.",
"Hey! Hoping you had a chance to read my last email.",
"Hey, just wanted to loop back on this. Would love to know your thoughts.",
"Hi, no worries if it’s not a fit. Just wanted to hear back either way.",
"Hey, bringing this back up in case it slipped past you.",
"Hi! Quick check-in on my last message. Let me know what you think.",
"Hey, hope your day’s going great! Any update on your end?",
"Hey, just making sure my last email didn’t end up in the spam folder.",
"Hi, I’d really appreciate a quick reply whenever you have a minute.",
"Hey! Just following up on this. Open to hearing any feedback.",
"Hey, I know your inbox is probably packed. Just giving this a gentle nudge.",
"Hi there, wanted to check whether you’ve had time to look this over.",
"Hey, still curious to hear your thoughts on this when you get a chance.",
"Hi! Circling back once more. A short reply would mean a lot.",
"Hey, just a quick reminder about my previous email.",
"Hope all’s well! Wanted to check in and see if you’d seen my note.",
"Hey, no pressure at all. Just wanted to keep the conversation going.",
"Hi, following up to see if this caught your interest.",
"Hey! Just touching base on my earlier message.",
"Hey, sorry to bug you again. Just wanted to see if you had any thoughts.",
"Hi, wanted to give this one more nudge before I stop bothering you!",
"Hey, just checking if now is a better time to chat about this.",
"Hi there! Any chance you’ve had a moment to think this over?",
"Hey, hoping to hear from you whenever it’s convenient.",
"Hey! Just a quick ping in case my last message got lost.",
"Hi, one last check-in from me. Would love to hear your thoughts either way."
];
// ============================================

function tick() {
  if (!shouldRun_()) return;                   // cheap exit when idle (saves daily runtime quota)
  try { runLocked_(); } catch (e) { alertOnce_(e); }
}

function runLocked_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return [];
  try { return run_(); } finally { lock.releaseLock(); }
}

// Wakes the real logic only: once an hour (to fill new dates), or when follow-ups are due and the random wait is over
function shouldRun_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const now = Date.now();
  if (now - Number(p.lastRun || 0) > 3600000) return true;
  return Number(p.dueLeft || 0) > 0 && now >= Number(p.nextSendAt || 0);
}

function run_() {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();
  if (ss.getSpreadsheetTimeZone() !== TIMEZONE) ss.setSpreadsheetTimeZone(TIMEZONE);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const { data, hRow, cDate, cEmail, cNext } = prepare_(sh);
  const log = [];
  if (data.length <= hRow + 1) return log;

  const props = PropertiesService.getScriptProperties();
  const state = JSON.parse(props.getProperty('state') || '{}');   // email -> 'filled' | 'sent' | 'drafted'
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  props.setProperty('lastRun', String(Date.now()));
  const scanNow = Date.now() - Number(props.getProperty('lastScan') || 0) > SCAN_HOURS * 3600000;
  if (scanNow) props.setProperty('lastScan', String(Date.now()));
  let canSend = Date.now() >= Number(props.getProperty('nextSendAt') || 0);
  let dueLeft = 0, completed = true;

  for (let i = hRow + 1; i < data.length; i++) {
    const rowNum = i + 1;
    if (Date.now() - t0 > 270000) { completed = false; break; }       // stay under Google's 6-min limit
    try {
      const row = data[i];
      const email = extractEmail_(row[cEmail]);
      if (!email) continue;                                           // "Filled Form", links, etc.
      const rawNext = String(row[cNext]).trim();
      if (/^(DONE|REPLIED|BOUNCED)$/i.test(rawNext)) continue;

      // 1) auto-fill the follow-up date for NEW rows only
      let next = dateKey_(row[cNext], TIMEZONE);
      const sentOn = dateKey_(row[cDate], TIMEZONE);
      if (!next && rawNext === '' && !state[email] && sentOn && sentOn >= addDays_(today, -1)) {
        next = addDays_(sentOn, FOLLOW_UP_DAYS);
        sh.getRange(rowNum, cNext + 1).setValue(Utilities.parseDate(next, TIMEZONE, 'yyyy-MM-dd'));
        state[email] = 'filled';
        log.push('Row ' + rowNum + ' (' + email + '): follow-up date set to ' + next);
      }

      // early reply check for upcoming leads (every few hours): mark REPLIED now instead of on the due date
      if (scanNow && next && next > today) {
        const t = findThread_(email, sentOn);
        const st = t && threadStatus_(t);
        if (st) {
          sh.getRange(rowNum, cNext + 1).setValue(st);
          log.push('Row ' + rowNum + ' (' + email + '): ' + st + ' (found early)');
        }
        continue;
      }

      // 2) only act when the date is exactly today (future = wait, passed = ignore)
      if (next !== today) continue;
      if (state[email] === 'sent' || (MODE !== 'SEND' && state[email] === 'drafted')) continue;
      dueLeft++;
      if (!canSend) continue;

      if (MODE === 'SEND' && MailApp.getRemainingDailyQuota() < 1) {  // Gmail's own daily limit reached
        props.setProperty('nextSendAt', String(Date.now() + 3600000));
        canSend = false; continue;
      }

      const thread = findThread_(email, sentOn);
      if (!thread) {
        sh.getRange(rowNum, cEmail + 1).setNote('No sent thread found - follow-up skipped');
        log.push('Row ' + rowNum + ' (' + email + '): no email TO this address in Gmail Sent -> skipped');
        dueLeft--; continue;
      }

      // someone other than you in the thread = replied; a failure notice = bounced. Either way: no follow-up
      const st = threadStatus_(thread);
      if (st) {
        sh.getRange(rowNum, cNext + 1).setValue(st);
        log.push('Row ' + rowNum + ' (' + email + '): ' + st + ' -> skipped');
        dueLeft--; continue;
      }

      // 3) send the follow-up as a reply in the thread
      const res = followUp_(thread, email, FOLLOWUPS[Math.floor(Math.random() * FOLLOWUPS.length)]);
      if (res === 'sent') {
        state[email] = 'sent';
        props.setProperty('state', JSON.stringify(state));            // save immediately: no double sends
        sh.getRange(rowNum, cNext + 1).setValue('DONE');
        log.push('Row ' + rowNum + ' (' + email + '): follow-up SENT in the thread');
      } else {
        state[email] = 'drafted';
        log.push('Row ' + rowNum + ' (' + email + '): DRAFT created (test mode)');
      }
      canSend = false; dueLeft--;                                     // one per tick, then a random wait
      const gapMs = (GAP_MINUTES[0] + Math.random() * (GAP_MINUTES[1] - GAP_MINUTES[0])) * 60000;
      props.setProperty('nextSendAt', String(Date.now() + Math.round(gapMs)));
    } catch (rowErr) {
      props.setProperty('lastError', 'Row ' + rowNum + ': ' + rowErr);
      log.push('Row ' + rowNum + ' ERROR: ' + rowErr);                // one bad row never stops the rest
    }
  }
  props.setProperty('state', JSON.stringify(state));
  props.setProperty('dueLeft', String(completed ? dueLeft : 1));
  return log;
}

// ---------- column detection ----------
const norm_ = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
const letterToIdx_ = L => String(L).toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

function findCol_(heads, key) {
  const exact = heads.map(norm_).findIndex(h => ALIASES[key].indexOf(h) > -1);
  if (exact > -1) return exact;
  const loose = {                                   // fallback: header looks like the right kind of column
    DATE:        h => /\bdate\b/.test(h) && !/follow|next/.test(h),
    EMAIL:       h => /e-?mail/.test(h) && !/subject|message|body|sent|status|date|follow|problem|draft/.test(h),
    NEXT_FOLLOW: h => /follow/.test(h) && !/message|text|body/.test(h)
  };
  return heads.map(h => h.toLowerCase()).findIndex(loose[key]);
}

// Finds the header row (first 10 rows) and the three columns. Adds a NEXT_FOLLOW column if it's missing.
function prepare_(sh) {
  let data = sh.getDataRange().getValues();
  for (let r = 0; r < Math.min(data.length, 10); r++) {
    const heads = data[r].map(h => String(h).trim());
    const pick = k => COLUMNS[k] ? letterToIdx_(COLUMNS[k]) : findCol_(heads, k);
    const cDate = pick('DATE'), cEmail = pick('EMAIL');
    let cNext = pick('NEXT_FOLLOW');
    if (cDate > -1 && cEmail > -1) {
      if (cNext === -1) {                           // no follow-up column: create one at the end
        cNext = heads.length;
        sh.getRange(r + 1, cNext + 1).setValue('NEXT_FOLLOW');
        data = sh.getDataRange().getValues();
      }
      return { data: data, hRow: r, cDate: cDate, cEmail: cEmail, cNext: cNext };
    }
  }
  const seen = data[0] ? data[0].map(h => String(h).trim()).filter(Boolean).join(', ') : '(empty sheet)';
  throw new Error('Could not find a date column and an email column in the first 10 rows. Headers seen in row 1: ' + seen +
                  '. Fix: rename them, or set COLUMNS at the top of the script (e.g. EMAIL: "C").');
}

// ---------- Gmail ----------
// The sent thread to this lead that is closest to the DATE you logged (so an older chat with the same address is never used)
function findThread_(email, sentOn) {
  const threads = GmailApp.search('in:sent to:' + email, 0, 10);
  if (!threads.length || !sentOn) return threads[0] || null;
  const target = new Date(sentOn + 'T12:00:00Z').getTime();
  let best = null, bestGap = Infinity;
  threads.forEach(t => {
    const gap = Math.abs(t.getMessages()[0].getDate().getTime() - target);
    if (gap < bestGap) { best = t; bestGap = gap; }
  });
  return best;
}

// 'BOUNCED' = delivery failure notice in the thread, 'REPLIED' = someone other than you wrote in it, else null
function threadStatus_(thread) {
  const me = myEmails_();
  const msgs = thread.getMessages();
  const failure = /delivery status notification \(failure\)|undeliverable|delivery has failed|delivery failure|returned mail|failure notice|address not found/i;
  for (let j = 0; j < msgs.length; j++) {
    const f = extractEmail_(msgs[j].getFrom());
    if (f && /mailer-daemon|postmaster/.test(f) && failure.test(msgs[j].getSubject())) return 'BOUNCED';
  }
  for (let j = 0; j < msgs.length; j++) {
    const f = extractEmail_(msgs[j].getFrom());
    if (f && me.indexOf(f) === -1 && !/mailer-daemon|postmaster|no-?reply/.test(f)) return 'REPLIED';
  }
  return null;
}

// True reply INSIDE the original thread: To = the lead, In-Reply-To/References = your original message, same threadId.
// No fallback: if it can't reply in-thread it throws and sends nothing.
function followUp_(thread, email, body) {
  if (typeof Gmail === 'undefined')
    throw new Error('Enable the Gmail API: Apps Script left menu > Services (+) > Gmail API > Add.');
  const me = myEmails_();
  const msgs = thread.getMessages();
  let target = msgs[msgs.length - 1];
  for (let j = msgs.length - 1; j >= 0; j--) {                        // reply to YOUR last message
    if (me.indexOf(extractEmail_(msgs[j].getFrom())) !== -1) { target = msgs[j]; break; }
  }
  const msgId = target.getHeader('Message-ID') || target.getHeader('Message-Id') || '';
  const refs = ((target.getHeader('References') || '') + ' ' + msgId).trim();
  const subj = 'Re: ' + thread.getFirstMessageSubject().replace(/^\s*(re:\s*)+/i, '');
  const b64 = str => Utilities.base64Encode(Utilities.newBlob(str).getBytes());
  const lines = [
    'To: ' + email,
    'Subject: =?UTF-8?B?' + b64(subj) + '?=',
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64'
  ];
  if (msgId) lines.push('In-Reply-To: ' + msgId);
  if (refs) lines.push('References: ' + refs);
  const raw = Utilities.base64EncodeWebSafe(Utilities.newBlob(lines.join('\r\n') + '\r\n\r\n' + b64(body)).getBytes());
  if (MODE === 'SEND') {
    Gmail.Users.Messages.send({ raw: raw, threadId: thread.getId() }, 'me');
    return 'sent';
  }
  Gmail.Users.Drafts.create({ message: { raw: raw, threadId: thread.getId() } }, 'me');
  return 'drafted';
}

function myEmails_() {
  return [Session.getEffectiveUser().getEmail()].concat(GmailApp.getAliases()).map(x => String(x).toLowerCase());
}

// ---------- helpers ----------
function extractEmail_(v) { const m = String(v || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i); return m ? m[0].toLowerCase() : null; }

// DATE INTELLIGENCE: turns almost any date format into 'yyyy-MM-dd', or null if it isn't a date.
// Understands: real date cells, sheet serial numbers, "12 Oct, 2026", "12 october 2026", "12th Oct. 2026", "Oct 12, 2026",
// "12/10/2026", "12-10-2026", "12.10.26", "2026-10-12", "2026/10/12". For ambiguous 12/10/2026 it uses DATE_ORDER
// (but 25/10/2026 or 10/25/2026 are unambiguous and always read correctly).
const MONTH_NAMES = ['january','february','march','april','may','june','july','august','september','october','november','december'];
function monthIdx_(w) {
  if (!w || w.length < 3) return -1;
  for (let i = 0; i < 12; i++) if (MONTH_NAMES[i].indexOf(w) === 0) return i;
  return -1;
}
function dateKey_(v, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (typeof v === 'number' && v > 20000 && v < 80000)                       // Google Sheets serial date
    return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
  const s = String(v || '').trim().toLowerCase().replace(/(\d)(st|nd|rd|th)\b/g, '$1');
  if (!s) return null;
  let y, m, d, x;
  if ((x = s.match(/\b(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/))) { y = +x[1]; m = +x[2]; d = +x[3]; }                       // 2026-10-12
  else if ((x = s.match(/\b(\d{1,2})[\s\-\/.,]*([a-z]{3,9})\.?[\s\-\/.,]*(\d{2,4})\b/)) && monthIdx_(x[2]) > -1) {          // 12 oct 2026
    d = +x[1]; m = monthIdx_(x[2]) + 1; y = +x[3];
  }
  else if ((x = s.match(/\b([a-z]{3,9})\.?[\s\-\/.,]*(\d{1,2})[\s\-\/.,]+(\d{2,4})\b/)) && monthIdx_(x[1]) > -1) {          // oct 12, 2026
    m = monthIdx_(x[1]) + 1; d = +x[2]; y = +x[3];
  }
  else if ((x = s.match(/\b(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})\b/))) {                                                  // 12/10/2026
    const a = +x[1], b = +x[2]; y = +x[3];
    if (a > 12) { d = a; m = b; } else if (b > 12) { m = a; d = b; }
    else if (DATE_ORDER === 'MDY') { m = a; d = b; } else { d = a; m = b; }
  }
  else return null;
  if (y < 100) y += 2000;
  const t = new Date(Date.UTC(y, m - 1, d));
  if (isNaN(t) || t.getUTCFullYear() !== y || t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return null;   // e.g. 31 Feb
  return y + '-' + pad_(m) + '-' + pad_(d);
}
function addDays_(key, n) { const d = new Date(key + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function pad_(n) { return n < 10 ? '0' + n : String(n); }

function say_(title, text) {                    // Execution log always; pop-up when used from the sheet menu
  Logger.log(title + '\n' + text);
  try { SpreadsheetApp.getUi().alert(title, text, SpreadsheetApp.getUi().ButtonSet.OK); return; } catch (e) {}
  try { SpreadsheetApp.getActive().toast(String(text).slice(0, 400), title, 30); } catch (e2) {}
}

function alertOnce_(e) {                        // emails you at most once a day if the script breaks
  const props = PropertiesService.getScriptProperties();
  const k = 'alert_' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  if (props.getProperty(k)) return;
  props.setProperty(k, '1');
  props.setProperty('lastError', String(e && e.message || e));
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Follow-up script error', String(e && e.stack || e));
}

// ---------- menu / setup ----------
function setup() {                              // run ONCE: creates the every-minute trigger and starts immediately
  const p = PropertiesService.getScriptProperties();
  ['lastRun', 'dueLeft', 'nextSendAt', 'lastScan', 'lastError'].forEach(k => p.deleteProperty(k));
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  let msg;
  if (typeof Gmail === 'undefined') {
    msg = 'Trigger created, BUT the Gmail API is not enabled yet.\nApps Script left menu > Services (+) > Gmail API > Add. Then run setup again.';
  } else {
    try {
      const log = runLocked_();
      msg = 'Automation is ON (runs every minute by itself).\nMode: ' + MODE + '\n\n' + (log.length ? log.join('\n') : 'Nothing due right now.');
    } catch (e) { msg = 'Trigger created, but the first run failed:\n' + e; }
  }
  say_('Follow-up automation', msg);
}

function runNow() {                             // forces one pass right now
  try {
    const log = runLocked_();
    say_('Run now', log.length ? log.join('\n') : 'Ran fine. Nothing is due right now.');
  } catch (e) { say_('Run now - ERROR', String(e)); }
}

function status() {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActive().getSheets()[0];
    const { data, hRow, cNext, cEmail } = prepare_(sh);
    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    let dueToday = 0, upcoming = 0;
    for (let i = hRow + 1; i < data.length; i++) {
      if (!extractEmail_(data[i][cEmail])) continue;
      const k = dateKey_(data[i][cNext], TIMEZONE);
      if (k === today) dueToday++; else if (k && k > today) upcoming++;
    }
    say_('Status', [
      'Today: ' + today + ' (' + TIMEZONE + ')',
      'MODE: ' + MODE,
      'Header row: ' + (hRow + 1),
      'Gmail API enabled: ' + (typeof Gmail !== 'undefined' ? 'yes' : 'NO - add it under Services'),
      'Trigger installed: ' + (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tick') ? 'yes' : 'NO - run setup'),
      'Due today (not yet handled): ' + dueToday,
      'Upcoming: ' + upcoming,
      'Last error: ' + (PropertiesService.getScriptProperties().getProperty('lastError') || 'none')
    ].join('\n'));
  } catch (e) { say_('Status - ERROR', String(e)); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Follow-ups')
    .addItem('Run now', 'runNow')
    .addItem('Status', 'status')
    .addItem('Start automation (one time)', 'setup')
    .addToUi();
}
