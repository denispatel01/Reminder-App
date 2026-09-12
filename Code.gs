// ===== REMINDER APP - BACKEND (Google Sheets storage, with recurring reminders + edit) =====

// Config lives in Script Properties (Project Settings > Script Properties), NOT in code,
// so this file can be shared publicly without leaking your Sheet ID or email.
// Run setupConfig() once (below) OR add the properties manually, then deploy.
const TRIGGER_FLAG = 'TRIGGER_INSTALLED';
const REMINDERS_SHEET = 'Reminders';
const LOGS_SHEET = 'Logs';
const CONTACTS_SHEET = 'Contacts';

/**
 * ONE-TIME SETUP: fill in your values, run this function once from the editor,
 * then clear the values again (they're safely stored in Script Properties).
 * Alternatively set SHEET_ID / DEFAULT_EMAIL under Project Settings > Script Properties.
 */
function setupConfig() {
  PropertiesService.getScriptProperties().setProperties({
    SHEET_ID: '',        // <-- paste your Google Sheet ID here, run once, then clear
    DEFAULT_EMAIL: ''    // <-- paste the default recipient email here
  });
}

function getSheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('SHEET_ID is not set. Run setupConfig() once, or add it under Project Settings > Script Properties.');
  return id;
}

// repeat: 'none' | 'daily' | 'weekly' | 'monthly' | 'annually' | 'custom'
const REM_HEADERS = ['id', 'title', 'description', 'datetime', 'email', 'sent', 'createdAt', 'repeat', 'customDays', 'recurrence', 'occurrenceCount', 'category', 'priority', 'done', 'doneAt'];
const LOG_HEADERS = ['timestamp', 'reminderId', 'title', 'email', 'status', 'message'];
const CONTACT_HEADERS = ['id', 'label', 'email', 'createdAt'];

/**
 * Serves the web app UI
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('✨ Reminder App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Opens the fixed spreadsheet and ensures both sheets exist with headers.
 */
function getSpreadsheet_() {
  const ss = SpreadsheetApp.openById(getSheetId_());
  ensureSheet_(ss, REMINDERS_SHEET, REM_HEADERS);
  ensureSheet_(ss, LOGS_SHEET, LOG_HEADERS);
  ensureSheet_(ss, CONTACTS_SHEET, CONTACT_HEADERS);
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#764ba2').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, headers.length);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach((h, idx) => {
      if (existingHeaders[idx] !== h) {
        sheet.getRange(1, idx + 1).setValue(h).setFontWeight('bold').setBackground('#764ba2').setFontColor('#ffffff');
      }
    });
  }
  return sheet;
}

function getSpreadsheetUrl() {
  return getSpreadsheet_().getUrl();
}

function getDefaultEmail() {
  return PropertiesService.getScriptProperties().getProperty('DEFAULT_EMAIL')
    || Session.getActiveUser().getEmail()
    || '';
}

/* ===================== SAVED CONTACTS (permanent, cross-device) ===================== */

/**
 * Returns all saved contacts, sorted by label/email.
 */
function getContacts() {
  const sheet = getSpreadsheet_().getSheetByName(CONTACTS_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const rows = data.slice(1)
    .filter(r => r[0])
    .map(r => ({ id: r[0], label: r[1] || '', email: r[2] || '' }));
  rows.sort((a, b) => (a.label || a.email).toLowerCase().localeCompare((b.label || b.email).toLowerCase()));
  return rows;
}

/**
 * Adds a contact. label is optional; email is required and must be valid.
 * De-duplicates on email (case-insensitive). Returns the fresh contact list.
 */
function addContact(label, email) {
  const clean = String(email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean)) {
    throw new Error('Please enter a valid email address.');
  }
  const sheet = getSpreadsheet_().getSheetByName(CONTACTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim().toLowerCase() === clean.toLowerCase()) {
      // update the label if a new one was supplied
      if (label && data[i][1] !== label) sheet.getRange(i + 1, 2).setValue(label);
      return getContacts();
    }
  }
  sheet.appendRow([Utilities.getUuid(), String(label || '').trim(), clean, new Date()]);
  return getContacts();
}

/**
 * Deletes a contact by id. Returns the fresh contact list.
 */
function deleteContact(id) {
  const sheet = getSpreadsheet_().getSheetByName(CONTACTS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getContacts();
}

/**
 * Get all reminders, sorted by date (soonest first)
 */
function getReminders() {
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const rows = data.slice(1).map(row => rowToReminder_(row));
  rows.sort((a, b) => new Date(a.datetime) - new Date(b.datetime));
  return rows;
}

function rowToReminder_(row) {
  return {
    id: row[0],
    title: row[1],
    description: row[2],
    datetime: row[3] instanceof Date ? row[3].toISOString() : row[3],
    email: row[4],
    sent: row[5] === true || row[5] === 'TRUE',
    createdAt: row[6] instanceof Date ? row[6].toISOString() : row[6],
    repeat: row[7] || 'none',
    customDays: row[8] || '',
    recurrence: parseJson_(row[9]),
    occurrenceCount: Number(row[10]) || 0,
    category: row[11] || '',
    priority: row[12] || 'normal',
    done: row[13] === true || row[13] === 'TRUE',
    doneAt: row[14] instanceof Date ? row[14].toISOString() : (row[14] || '')
  };
}

function parseJson_(value) {
  if (!value) return null;
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(value);
  } catch (e) {
    return null;
  }
}

/**
 * Add a new reminder
 * reminder = { title, description, datetime, email, repeat, customDays }
 * email can be a single address or a comma-separated list of addresses
 */
function addReminder(reminder) {
  validateReminderInput_(reminder);

  const repeat = reminder.repeat || 'none';
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const id = Utilities.getUuid();

  sheet.appendRow([
    id,
    reminder.title,
    reminder.description || '',
    new Date(reminder.datetime),
    reminder.email,
    false,
    new Date(),
    repeat,
    repeat === 'custom' ? Number(reminder.customDays) || '' : '',
    repeat === 'custom' && reminder.recurrence ? JSON.stringify(reminder.recurrence) : '',
    0,
    reminder.category || '',
    reminder.priority || 'normal',
    false,
    ''
  ]);

  createTrigger();

  return getReminders();
}

/**
 * Update an existing reminder by id.
 * reminder = { title, description, datetime, email, repeat, customDays }
 * Resets "sent" to false so an edited reminder becomes active again (e.g. rescheduled or amended).
 */
function updateReminder(id, reminder) {
  if (!id) throw new Error('Missing reminder id.');
  validateReminderInput_(reminder);

  const repeat = reminder.repeat || 'none';
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 2).setValue(reminder.title);
      sheet.getRange(rowIndex, 3).setValue(reminder.description || '');
      sheet.getRange(rowIndex, 4).setValue(new Date(reminder.datetime));
      sheet.getRange(rowIndex, 5).setValue(reminder.email);
      sheet.getRange(rowIndex, 6).setValue(false); // reactivate on edit
      sheet.getRange(rowIndex, 8).setValue(repeat);
      sheet.getRange(rowIndex, 9).setValue(repeat === 'custom' ? Number(reminder.customDays) || '' : '');
      sheet.getRange(rowIndex, 10).setValue(repeat === 'custom' && reminder.recurrence ? JSON.stringify(reminder.recurrence) : '');
      sheet.getRange(rowIndex, 11).setValue(0); // reset occurrence count on edit
      sheet.getRange(rowIndex, 12).setValue(reminder.category || '');
      sheet.getRange(rowIndex, 13).setValue(reminder.priority || 'normal');
      // done (col 14) & doneAt (col 15) intentionally left unchanged here
      return getReminders();
    }
  }

  throw new Error('Reminder not found — it may have already been deleted.');
}

function validateReminderInput_(reminder) {
  if (!reminder.title || !reminder.datetime || !reminder.email) {
    throw new Error('Title, date/time, and email are required.');
  }
  const repeat = reminder.repeat || 'none';
  if (repeat === 'custom' && !reminder.recurrence && (!reminder.customDays || Number(reminder.customDays) <= 0)) {
    throw new Error('Please configure the custom repeat.');
  }
}

/**
 * Delete a reminder by id
 */
function deleteReminder(id) {
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return getReminders();
}

/**
 * Toggle the manual "done" flag for a reminder.
 * done=true  -> set doneAt=now, sent=true  (so the trigger won't email a completed one-time item)
 * done=false -> clear doneAt, sent=false    (reactivate)
 */
function toggleDone(id) {
  if (!id) throw new Error('Missing reminder id.');
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowIndex = i + 1;
      const currentlyDone = data[i][13] === true || data[i][13] === 'TRUE';
      const newDone = !currentlyDone;
      sheet.getRange(rowIndex, 14).setValue(newDone);          // done
      sheet.getRange(rowIndex, 15).setValue(newDone ? new Date() : ''); // doneAt
      sheet.getRange(rowIndex, 6).setValue(newDone);           // sent mirrors done
      return getReminders();
    }
  }

  throw new Error('Reminder not found — it may have already been deleted.');
}

/**
 * Push a reminder forward by `minutes` from now and reactivate it.
 * minutes is a Number.
 */
function snoozeReminder(id, minutes) {
  if (!id) throw new Error('Missing reminder id.');
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();
  const mins = Number(minutes) || 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, 4).setValue(new Date(Date.now() + mins * 60000)); // datetime
      sheet.getRange(rowIndex, 6).setValue(false);  // sent
      sheet.getRange(rowIndex, 14).setValue(false); // done
      sheet.getRange(rowIndex, 15).setValue('');    // doneAt
      createTrigger();
      return getReminders();
    }
  }

  throw new Error('Reminder not found — it may have already been deleted.');
}

/**
 * Duplicate a reminder into a fresh row with a new id.
 */
function duplicateReminder(id) {
  if (!id) throw new Error('Missing reminder id.');
  const sheet = getSpreadsheet_().getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const src = data[i];
      sheet.appendRow([
        Utilities.getUuid(),
        src[1],                                              // title
        src[2] || '',                                        // description
        src[3] instanceof Date ? src[3] : new Date(src[3]),  // datetime
        src[4],                                              // email
        false,                                               // sent
        new Date(),                                          // createdAt
        src[7] || 'none',                                    // repeat
        src[8] || '',                                        // customDays
        src[9] || '',                                        // recurrence
        0,                                                   // occurrenceCount
        src[11] || '',                                       // category
        src[12] || 'normal',                                 // priority
        false,                                               // done
        ''                                                   // doneAt
      ]);
      createTrigger();
      return getReminders();
    }
  }

  throw new Error('Reminder not found — it may have already been deleted.');
}

/**
 * Send a reminder's email immediately without changing its datetime/sent/done.
 * Logs it with status 'TEST'. Returns true on success; throws on failure.
 */
function sendTestEmail(id) {
  if (!id) throw new Error('Missing reminder id.');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const reminder = rowToReminder_(data[i]);
      try {
        MailApp.sendEmail({
          to: reminder.email,
          subject: '⏰ Reminder: ' + reminder.title,
          htmlBody: buildEmailBody(reminder)
        });
        logEvent_(ss, reminder.id, reminder.title, reminder.email, 'TEST', 'Test email sent manually');
        return true;
      } catch (err) {
        logEvent_(ss, reminder.id, reminder.title, reminder.email, 'FAILED', String(err));
        throw err;
      }
    }
  }

  throw new Error('Reminder not found — it may have already been deleted.');
}

/**
 * Computes the next occurrence date based on repeat type
 */
function getNextOccurrence_(currentDate, repeat, customDays) {
  const next = new Date(currentDate);
  switch (repeat) {
    case 'daily':
      next.setDate(next.getDate() + 1);
      break;
    case 'weekly':
      next.setDate(next.getDate() + 7);
      break;
    case 'monthly':
      next.setMonth(next.getMonth() + 1);
      break;
    case 'annually':
      next.setFullYear(next.getFullYear() + 1);
      break;
    case 'custom':
      next.setDate(next.getDate() + (Number(customDays) || 1));
      break;
    default:
      return null; // not recurring
  }
  return next;
}

/**
 * Days in a given month (m is 0-based).
 */
function daysInMonth_(y, m) {
  return new Date(y, m + 1, 0).getDate();
}

/**
 * Returns the date of the nth (1-5) `weekday` in month m (0-based) of year y.
 * If the nth occurrence doesn't exist, falls back to the last one that does.
 */
function nthWeekdayOfMonth_(y, m, weekday, nth) {
  const firstWd = new Date(y, m, 1).getDay();
  let day = 1 + ((weekday - firstWd + 7) % 7) + (nth - 1) * 7;
  if (day > daysInMonth_(y, m)) day -= 7; // "5th" that doesn't exist -> last
  return new Date(y, m, day);
}

/**
 * Computes the next occurrence for a Google-Calendar-style custom recurrence.
 * rec = { freq, interval, byWeekday:[0-6], monthlyMode:'dayOfMonth'|'dayOfWeek' }
 */
function computeNextCustom_(current, rec) {
  if (!rec) return null;
  const freq = rec.freq || 'daily';
  const interval = Math.max(1, Number(rec.interval) || 1);
  const h = current.getHours(), min = current.getMinutes();

  if (freq === 'daily') {
    const next = new Date(current);
    next.setDate(next.getDate() + interval);
    return next;
  }

  if (freq === 'weekly') {
    const days = (Array.isArray(rec.byWeekday) && rec.byWeekday.length
      ? rec.byWeekday.slice() : [current.getDay()]).sort((a, b) => a - b);
    const cur = current.getDay();
    const later = days.filter(d => d > cur);
    const next = new Date(current);
    if (later.length) {
      next.setDate(next.getDate() + (later[0] - cur));
    } else {
      next.setDate(next.getDate() + (7 * interval) - cur + days[0]);
    }
    return next;
  }

  if (freq === 'monthly') {
    const targetMonth = current.getMonth() + interval;
    const y = current.getFullYear() + Math.floor(targetMonth / 12);
    const m = ((targetMonth % 12) + 12) % 12;
    if (rec.monthlyMode === 'dayOfWeek') {
      const nth = Math.ceil(current.getDate() / 7);
      const date = nthWeekdayOfMonth_(y, m, current.getDay(), nth);
      date.setHours(h, min, 0, 0);
      return date;
    }
    const d = Math.min(current.getDate(), daysInMonth_(y, m));
    return new Date(y, m, d, h, min, 0, 0);
  }

  if (freq === 'yearly') {
    const y = current.getFullYear() + interval;
    const m = current.getMonth();
    const d = Math.min(current.getDate(), daysInMonth_(y, m));
    return new Date(y, m, d, h, min, 0, 0);
  }

  return null;
}

/**
 * Resolves the next date for any reminder, honoring custom recurrence + end rules.
 * Returns a Date to reschedule to, or null to stop (mark delivered).
 */
function resolveNextDate_(dueDate, repeat, customDays, recurrence, occurrenceCount) {
  if (repeat !== 'custom') {
    return getNextOccurrence_(dueDate, repeat, customDays);
  }

  // Backward compat: old custom rows stored only a day interval.
  let rec = recurrence;
  if (!rec && customDays) rec = { freq: 'daily', interval: Number(customDays) || 1 };
  if (!rec) return null;

  let next = computeNextCustom_(dueDate, rec);
  if (!next) return null;

  const end = rec.end || { type: 'never' };
  if (end.type === 'afterCount') {
    // occurrenceCount is delivered-so-far; this send makes it +1.
    if ((Number(occurrenceCount) || 0) + 1 >= Number(end.count)) return null;
  }
  if (end.type === 'onDate' && end.date) {
    if (next > new Date(end.date)) return null;
  }
  return next;
}

/**
 * Runs every 5 minutes (via trigger). Sends email for due reminders,
 * reschedules recurring ones, and logs every attempt.
 * MailApp.sendEmail's "to" field natively accepts a comma-separated list,
 * so multi-recipient reminders work with no extra handling.
 */
function checkReminders() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(REMINDERS_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sent = row[5] === true || row[5] === 'TRUE';
    const done = row[13] === true || row[13] === 'TRUE';
    const dueDate = row[3] instanceof Date ? row[3] : new Date(row[3]);
    const repeat = row[7] || 'none';
    const customDays = row[8];

    if (done) continue; // skip manually completed reminders

    if (!sent && dueDate <= now) {
      const reminder = rowToReminder_(row);
      try {
        MailApp.sendEmail({
          to: reminder.email, // comma-separated list works natively here
          subject: '⏰ Reminder: ' + reminder.title,
          htmlBody: buildEmailBody(reminder)
        });

        const nextDate = resolveNextDate_(dueDate, repeat, customDays, reminder.recurrence, reminder.occurrenceCount);

        if (nextDate) {
          sheet.getRange(i + 1, 4).setValue(nextDate); // datetime column
          sheet.getRange(i + 1, 6).setValue(false);     // stays unsent for next round
          sheet.getRange(i + 1, 11).setValue((Number(reminder.occurrenceCount) || 0) + 1);
          logEvent_(ss, reminder.id, reminder.title, reminder.email, 'SENT',
            'Recurring (' + repeat + ') — next run: ' + nextDate.toLocaleString());
        } else {
          sheet.getRange(i + 1, 6).setValue(true);
          logEvent_(ss, reminder.id, reminder.title, reminder.email, 'SENT', 'One-time reminder delivered');
        }
      } catch (err) {
        logEvent_(ss, reminder.id, reminder.title, reminder.email, 'FAILED', String(err));
      }
    }
  }
}

function logEvent_(ss, reminderId, title, email, status, message) {
  const logSheet = ss.getSheetByName(LOGS_SHEET);
  logSheet.appendRow([new Date(), reminderId, title, email, status, message]);
}

function buildEmailBody(r) {
  const desc = r.description
    ? '<p style="margin:16px 0;color:#444;font-size:15px;line-height:1.6;">' + escapeHtml(r.description) + '</p>'
    : '';
  const repeatLabel = r.repeat && r.repeat !== 'none' ? ' · Repeats ' + r.repeat : '';
  return `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width:520px; margin:0 auto; background:#f4f5f9; padding:32px;">
      <div style="background:linear-gradient(135deg,#667eea,#764ba2); border-radius:16px; padding:32px; color:#fff; text-align:center;">
        <div style="font-size:40px; margin-bottom:8px;">⏰</div>
        <h1 style="margin:0; font-size:22px;">${escapeHtml(r.title)}</h1>
      </div>
      <div style="background:#fff; border-radius:16px; padding:24px; margin-top:16px; box-shadow:0 4px 20px rgba(0,0,0,0.08);">
        ${desc}
        <p style="color:#888; font-size:13px; margin-top:20px; border-top:1px solid #eee; padding-top:16px;">
          Scheduled for: ${new Date(r.datetime).toLocaleString()}${repeatLabel}
        </p>
      </div>
    </div>
  `;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * Installs a time-driven trigger (runs once, safe to call repeatedly)
 */
function createTrigger() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(TRIGGER_FLAG) === 'true') return;

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('checkReminders')
    .timeBased()
    .everyMinutes(5)
    .create();

  props.setProperty(TRIGGER_FLAG, 'true');
}