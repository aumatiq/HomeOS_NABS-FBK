// ══════════════════════════════════════════════════════════
// NABS&FBK HOME OS — Code.gs
// Google Apps Script Backend v2.0
// Functions: Password verify, OTP reset, Email reports,
//            Budget alerts, Schedule triggers, Drive upload,
//            Sheets sync, Ping
// ══════════════════════════════════════════════════════════

// ─── CONFIGURATION (EDIT THESE BEFORE DEPLOY) ───────────
const CONFIG = {
  SHEET_ID:         '1RLoBAewJ6GbgAm_NpF9i4ZDD3LKFsqVlbintntWgwaM',          // Google Sheets ID (from URL)
  APP_PASS_HASH:    'fh4u6k',    // Fallback ONLY — real password check now reads from Settings sheet first
  ADMIN_PASS_HASH:  '-ol2yc0',    // Fallback ONLY — real password check now reads from Settings sheet first
  RESET_EMAIL:      '',          // Password reset OTP email
  REPORT_EMAIL:     '',          // Weekly/monthly report email
  DRIVE_FOLDER_ID:  '1_wGbs0aMfmxEKgU6ev9mXgmchZe_9z_x',          // Master Drive folder ID for documents
  TIMEZONE:         'Asia/Dhaka'
};

// ─── SHEET NAMES ─────────────────────────────────────────
const SHEETS = {
  ENTRIES:   'Entries',
  PROJECTS:  'Projects',
  COSTS:     'ProjectCosts',
  TASKS:     'Tasks',
  SHOPPING:  'Shopping',
  DOCS:      'Documents',
  SETTINGS:  'Settings',
  OTP_LOG:   'OTP_Log',
  SYNC_LOG:  'Sync_Log'
};

// ══════════════════════════════════════════════════════════
// ENTRY POINT — doGet / doPost
// ══════════════════════════════════════════════════════════
function doGet(e) {
  const action = e.parameter.action || '';
  try {
    if (action === 'ping')        return jsonResp({ status: 'ok', ts: new Date().toISOString() });
    if (action === 'verifyPass')  return handleVerifyPass(e);
    if (action === 'getData')     return handleGetData(e);
    if (action === 'sync')        return jsonResp({ status: 'ok', msg: 'Use POST for sync' });
    return jsonResp({ status: 'error', msg: 'Unknown GET action: ' + action });
  } catch(err) {
    return jsonResp({ status: 'error', msg: err.message });
  }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(ex) {}
  const action = e.parameter.action || body.action || '';
  try {
    if (action === 'addEntry')       return handleAddEntry(body);
    if (action === 'sync')           return handleSync(body);
    if (action === 'sendOTP')        return handleSendOTP(body);
    if (action === 'budgetAlert')    return handleBudgetAlert(body);
    if (action === 'uploadFile')     return handleUploadFile(body);
    if (action === 'updateSchedule') return handleUpdateSchedule(body);
    if (action === 'updatePassword') return handleUpdatePassword(body);
    if (action === 'saveDoc')        return handleSaveDoc(body);
    if (action === 'deleteDoc')      return handleDeleteDoc(body);
    return jsonResp({ status: 'error', msg: 'Unknown POST action: ' + action });
  } catch(err) {
    return jsonResp({ status: 'error', msg: err.message });
  }
}

// ══════════════════════════════════════════════════════════
// PASSWORD VERIFICATION
// ══════════════════════════════════════════════════════════

/**
 * Verifies hashed password sent from frontend.
 * GET: ?action=verifyPass&hash=XXXX&type=app|admin
 * IMPORTANT FIX: Previously this only checked the hardcoded CONFIG hash,
 * so changing the password inside the app (Settings tab) never actually
 * changed what the backend accepted. Now it checks the Settings sheet
 * first (live value), and falls back to the CONFIG default only if the
 * Settings sheet has no value yet (fresh install).
 */
function handleVerifyPass(e) {
  const hash   = e.parameter.hash || '';
  const type   = e.parameter.type || 'app';
  const key    = type === 'admin' ? 'admin_pass_hash' : 'app_pass_hash';
  const stored = getSettingValue(key) || (type === 'admin' ? CONFIG.ADMIN_PASS_HASH : CONFIG.APP_PASS_HASH);
  return jsonResp({ ok: hash === stored });
}

/**
 * Updates the app-lock or admin-lock password hash.
 * POST body: { type: 'app'|'admin', newHash: '...' }
 * Called by the frontend whenever the user changes a password from
 * Settings, so the backend and the app never go out of sync again.
 */
function handleUpdatePassword(body) {
  const type = body.type === 'admin' ? 'admin' : 'app';
  const key  = type === 'admin' ? 'admin_pass_hash' : 'app_pass_hash';
  if (!body.newHash) return jsonResp({ status: 'error', msg: 'Missing newHash' });
  setSettingValue(key, body.newHash);
  return jsonResp({ status: 'ok', msg: 'Password updated' });
}

/**
 * Reads a single Key/Value from the Settings sheet. Returns null if not found.
 */
function getSettingValue(key) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.SETTINGS, ['Key','Value','UpdatedAt']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

/**
 * Writes (upserts) a single Key/Value into the Settings sheet.
 */
function setSettingValue(key, value) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.SETTINGS, ['Key','Value','UpdatedAt']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[value, new Date()]]);
      return;
    }
  }
  sheet.appendRow([key, value, new Date()]);
}

// ══════════════════════════════════════════════════════════
// OTP — Password Reset
// ══════════════════════════════════════════════════════════

/**
 * Sends OTP to reset email.
 * POST body: { email, otp, type: 'app'|'admin' }
 */
function handleSendOTP(body) {
  const { email, otp, type } = body;
  if (!email || !otp) return jsonResp({ status: 'error', msg: 'Missing email or OTP' });

  // Log OTP
  logOTP(email, otp, type);

  const subject = '🔐 NABS&FBK Home OS — Password Reset OTP';
  const htmlBody = buildOTPEmailHTML(otp, type);

  GmailApp.sendEmail(email, subject, `Your OTP is: ${otp}`, { htmlBody });
  return jsonResp({ status: 'ok', msg: 'OTP sent to ' + email });
}

function logOTP(email, otp, type) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.OTP_LOG, ['Timestamp','Email','OTP','Type','Used']);
  sheet.appendRow([new Date(), email, otp, type, 'false']);
}

function buildOTPEmailHTML(otp, type) {
  const lockType = type === 'admin' ? 'Settings / Admin' : 'App Lock';
  return `
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8">
    <style>
      body { font-family: Arial, Helvetica, sans-serif; background: #0A0A0F; color: #F8F9FF; margin: 0; padding: 0; }
      .wrap { max-width: 500px; margin: 0 auto; background: #0D1117; border: 1px solid rgba(79,70,229,0.3); border-radius: 16px; overflow: hidden; }
      .header { background: linear-gradient(135deg, #4F46E5, #6366F1); padding: 28px 32px; text-align: center; }
      .header h1 { font-size: 22px; margin: 0; letter-spacing: 1px; }
      .header p { font-size: 13px; opacity: 0.8; margin: 6px 0 0; }
      .body { padding: 32px; text-align: center; }
      .otp-box { background: #111118; border: 2px solid #F5A623; border-radius: 12px; padding: 20px; margin: 24px 0; display: inline-block; min-width: 200px; }
      .otp-code { font-size: 42px; font-weight: 900; letter-spacing: 10px; color: #F5A623; font-family: monospace; }
      .info { font-size: 14px; color: #A0A0B0; line-height: 1.7; }
      .footer { background: #0A0A0F; padding: 16px 32px; text-align: center; font-size: 12px; color: #606070; border-top: 1px solid rgba(79,70,229,0.15); }
      .warning { background: rgba(239,68,68,0.1); border: 1px solid rgba(239,68,68,0.3); border-radius: 8px; padding: 12px 16px; margin-top: 20px; font-size: 13px; color: #EF4444; }
    </style>
    </head>
    <body>
    <div class="wrap">
      <div class="header">
        <h1>🏠 NABS&amp;FBK Home OS</h1>
        <p>Password Reset Request</p>
      </div>
      <div class="body">
        <p class="info">আপনার <strong>${lockType}</strong> পাসওয়ার্ড রিসেট করতে নিচের OTP কোড ব্যবহার করুন।</p>
        <div class="otp-box">
          <div class="otp-code">${otp}</div>
        </div>
        <p class="info">এই কোডটি <strong>10 মিনিট</strong> পর্যন্ত কার্যকর।</p>
        <div class="warning">⚠️ এই কোড কাউকে দেবেন না। NABS&amp;FBK Home OS কখনো আপনার কাছে OTP চাইবে না।</div>
      </div>
      <div class="footer">NABS&amp;FBK Home OS · Family Management System · Dhaka, Bangladesh</div>
    </div>
    </body>
    </html>`;
}

// ══════════════════════════════════════════════════════════
// ENTRIES — Add single entry from PWA
// ══════════════════════════════════════════════════════════
function handleAddEntry(body) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.ENTRIES, [
    'ID','Date','Type','Category','Amount','Member','PaymentMethod','Note','CreatedAt'
  ]);
  sheet.appendRow([
    body.id || '', body.date || '', body.type || '', body.category || '',
    parseFloat(body.amount) || 0, body.member || '', body.payment || '',
    body.note || '', body.createdAt || new Date().toISOString()
  ]);
  return jsonResp({ status: 'ok', id: body.id });
}

// ══════════════════════════════════════════════════════════
// FULL SYNC — Bulk write all data to sheets
// (FIXED: previously only 'entries' and 'projects' were synced.
//  Tasks, Shopping and ProjectCosts were silently never pushed
//  to Google Sheets even though they exist as sheets.)
// ══════════════════════════════════════════════════════════
function handleSync(body) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  syncArrayToSheet(ss, SHEETS.ENTRIES, ['ID','Date','Type','Category','Amount','Member','PaymentMethod','Note','CreatedAt'],
    body.entries, e => [e.id||'', e.date||'', e.type||'', e.category||'', parseFloat(e.amount)||0, e.member||'', e.payment||'', e.note||'', e.createdAt||'']);

  syncArrayToSheet(ss, SHEETS.PROJECTS, ['ID','Name','Category','Status','StartDate','EndDate','Budget','ActualCost','Priority','Notes','CreatedAt'],
    body.projects, p => [p.id||'', p.name||'', p.category||'', p.status||'', p.startDate||'', p.endDate||'', p.budget||0, p.actualCost||0, p.priority||'medium', p.notes||'', p.createdAt||'']);

  syncArrayToSheet(ss, SHEETS.COSTS, ['ID','ProjectID','Name','Amount','Date'],
    body.costs, c => [c.id||'', c.projectId||'', c.name||'', parseFloat(c.amount)||0, c.date||'']);

  syncArrayToSheet(ss, SHEETS.TASKS, ['ID','Text','Due','Priority','Member','Category','Done','Notes','CreatedAt'],
    body.tasks, t => [t.id||'', t.text||'', t.due||'', t.priority||'', t.member||'', t.category||'', t.done?'true':'false', t.notes||'', t.createdAt||'']);

  syncArrayToSheet(ss, SHEETS.SHOPPING, ['ID','Name','Qty','Price','Category','List','Bought','Note'],
    body.shopping, s => [s.id||'', s.name||'', s.qty||'', s.price||0, s.category||'', s.list||'', s.bought?'true':'false', s.note||'']);

  syncArrayToSheet(ss, SHEETS.DOCS, ['ID','Title','Category','Type','Tags','Content','DriveFileId','DriveFileName','UpdatedAt'],
    body.docs, d => [d.id||'', d.title||'', d.category||'', d.type||'', (d.tags||[]).join(','), d.content||'', d.driveFileId||'', d.driveFileName||'', d.updatedAt||'']);

  const logSheet = getOrCreateSheet(ss, SHEETS.SYNC_LOG, ['Timestamp','EntriesCount','Status']);
  logSheet.appendRow([new Date(), (body.entries||[]).length, 'ok']);

  return jsonResp({ status: 'ok', synced: new Date().toISOString() });
}

/**
 * Helper: clears a sheet's data rows and rewrites them from an array,
 * only touching the sheet if an array was actually sent (so partial
 * syncs from the frontend never wipe data it didn't send).
 */
function syncArrayToSheet(ss, sheetName, headers, arr, rowMapper) {
  if (!arr || !arr.length) return;
  const sheet   = getOrCreateSheet(ss, sheetName, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  const rows = arr.map(rowMapper);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// ══════════════════════════════════════════════════════════
// PULL — Frontend calls this on load / manual refresh so data
// is restored from Google Sheets (e.g. new device, reinstalled
// app, or someone else on the family added something directly
// in the Sheet). Previously there was NO way to pull data back —
// only push existed, so the Sheet was write-only.
// GET: ?action=getData
// ══════════════════════════════════════════════════════════
function handleGetData(e) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return jsonResp({
    status:   'ok',
    entries:  sheetRowsToObjects(ss, SHEETS.ENTRIES,  ['id','date','type','category','amount','member','payment','note','createdAt']),
    projects: sheetRowsToObjects(ss, SHEETS.PROJECTS, ['id','name','category','status','startDate','endDate','budget','actualCost','priority','notes','createdAt']),
    costs:    sheetRowsToObjects(ss, SHEETS.COSTS,    ['id','projectId','name','amount','date']),
    tasks:    sheetRowsToObjects(ss, SHEETS.TASKS,    ['id','text','due','priority','member','category','done','notes','createdAt']).map(t => ({...t, done: t.done === 'true' || t.done === true})),
    shopping: sheetRowsToObjects(ss, SHEETS.SHOPPING, ['id','name','qty','price','category','list','bought','note']).map(s => ({...s, bought: s.bought === 'true' || s.bought === true})),
    docs:     sheetRowsToObjects(ss, SHEETS.DOCS,     ['id','title','category','type','tags','content','driveFileId','driveFileName','updatedAt']).map(d => ({...d, tags: d.tags ? String(d.tags).split(',').filter(Boolean) : []})),
    ts: new Date().toISOString()
  });
}

function sheetRowsToObjects(ss, sheetName, keys) {
  const rows = getSheetData(ss, sheetName);
  return rows.map(row => {
    const obj = {};
    keys.forEach((k, i) => { obj[k] = row[i]; });
    return obj;
  });
}

// ══════════════════════════════════════════════════════════
// BUDGET ALERT EMAIL
// ══════════════════════════════════════════════════════════
function handleBudgetAlert(body) {
  const { email, pct, spent, budget } = body;
  if (!email) return jsonResp({ status: 'error', msg: 'No email' });

  const subject = `⚠️ NABS&FBK Home OS — বাজেট সতর্কতা: ${pct}% ব্যবহৃত`;
  const html = buildBudgetAlertHTML(pct, spent, budget);
  GmailApp.sendEmail(email, subject, `বাজেট সতর্কতা: ${pct}% ব্যয় হয়েছে।`, { htmlBody: html });
  return jsonResp({ status: 'ok' });
}

function buildBudgetAlertHTML(pct, spent, budget) {
  const color = pct >= 100 ? '#EF4444' : '#F5A623';
  return `
    <!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
      body{font-family:Arial,sans-serif;background:#0A0A0F;color:#F8F9FF;margin:0;padding:0}
      .wrap{max-width:480px;margin:0 auto;background:#0D1117;border:1px solid rgba(79,70,229,0.3);border-radius:16px;overflow:hidden}
      .header{background:linear-gradient(135deg,${color},#f97316);padding:24px 28px;text-align:center}
      .header h1{font-size:20px;margin:0}
      .body{padding:28px;text-align:center}
      .pct{font-size:64px;font-weight:900;color:${color};margin:16px 0}
      .row{display:flex;justify-content:space-between;background:#111118;border-radius:8px;padding:12px 16px;margin:8px 0;font-size:14px}
      .label{color:#A0A0B0} .val{font-weight:700;color:#F8F9FF}
      .bar-track{background:#1a1a2e;border-radius:6px;height:12px;overflow:hidden;margin:16px 0}
      .bar-fill{height:100%;background:${color};width:${Math.min(pct,100)}%;border-radius:6px}
      .footer{background:#0A0A0F;padding:14px 28px;text-align:center;font-size:12px;color:#606070;border-top:1px solid rgba(79,70,229,0.12)}
    </style>
    </head><body><div class="wrap">
      <div class="header"><h1>⚠️ বাজেট সতর্কতা</h1></div>
      <div class="body">
        <p style="color:#A0A0B0;font-size:14px">এই মাসের ব্যয় বাজেটের সীমা অতিক্রম করছে</p>
        <div class="pct">${pct}%</div>
        <div class="bar-track"><div class="bar-fill"></div></div>
        <div class="row"><span class="label">মোট ব্যয়</span><span class="val">৳ ${Number(spent).toLocaleString()}</span></div>
        <div class="row"><span class="label">মাসিক বাজেট</span><span class="val">৳ ${Number(budget).toLocaleString()}</span></div>
        <div class="row"><span class="label">বাকি বাজেট</span><span class="val" style="color:${pct>=100?'#EF4444':'#22C55E'}">৳ ${Math.max(0,Number(budget)-Number(spent)).toLocaleString()}</span></div>
      </div>
      <div class="footer">NABS&amp;FBK Home OS · Family Management System</div>
    </div></body></html>`;
}

// ══════════════════════════════════════════════════════════
// FILE UPLOAD TO GOOGLE DRIVE
// FIXED: now auto-creates a subfolder per document category
// inside the master Drive folder, so files are organised
// automatically (e.g. Master Folder/Warranty/, Master Folder/Bills/)
// ══════════════════════════════════════════════════════════
function handleUploadFile(body) {
  const { fileName, mimeType, base64, folderId, category } = body;
  if (!fileName || !base64) return jsonResp({ status: 'error', msg: 'Missing file data' });

  const masterFolderId = folderId || CONFIG.DRIVE_FOLDER_ID;
  if (!masterFolderId) return jsonResp({ status: 'error', msg: 'Drive folder ID not configured' });

  try {
    const targetFolder = category ? getOrCreateSubfolder(masterFolderId, category) : DriveApp.getFolderById(masterFolderId);
    const decoded  = Utilities.base64Decode(base64);
    const blob     = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
    const file     = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    return jsonResp({ status: 'ok', fileId: file.getId(), fileName: file.getName(), url: file.getUrl(), folder: targetFolder.getName() });
  } catch(err) {
    return jsonResp({ status: 'error', msg: 'Drive upload failed: ' + err.message });
  }
}

/**
 * Finds a subfolder by name inside a parent folder, creating it
 * if it doesn't exist yet. Used to auto-organise Documents by category.
 */
function getOrCreateSubfolder(parentFolderId, subfolderName) {
  const parent = DriveApp.getFolderById(parentFolderId);
  const safeName = String(subfolderName).trim() || 'Others';
  const existing = parent.getFoldersByName(safeName);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(safeName);
}

// ══════════════════════════════════════════════════════════
// DOCUMENT METADATA — Save / Delete a single document row
// FIXED: previously Document titles/notes/tags typed in the app
// were NEVER sent to the Documents sheet at all (only the raw
// uploaded file went to Drive, with no record in Sheets). Now
// every save/delete keeps the Documents sheet in sync too.
// ══════════════════════════════════════════════════════════
function handleSaveDoc(body) {
  const d = body.doc || {};
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.DOCS, ['ID','Title','Category','Type','Tags','Content','DriveFileId','DriveFileName','UpdatedAt']);
  const data  = sheet.getDataRange().getValues();
  const row   = [d.id||'', d.title||'', d.category||'', d.type||'', (d.tags||[]).join(','), d.content||'', d.driveFileId||'', d.driveFileName||'', d.updatedAt||new Date().toISOString()];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.id) { sheet.getRange(i + 1, 1, 1, row.length).setValues([row]); return jsonResp({ status: 'ok', id: d.id }); }
  }
  sheet.appendRow(row);
  return jsonResp({ status: 'ok', id: d.id });
}

function handleDeleteDoc(body) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.DOCS, ['ID','Title','Category','Type','Tags','Content','DriveFileId','DriveFileName','UpdatedAt']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) { sheet.deleteRow(i + 1); break; }
  }
  return jsonResp({ status: 'ok' });
}

// ══════════════════════════════════════════════════════════
// UPDATE SCHEDULE (from Settings tab)
// Saves schedule preferences and re-installs triggers
// ══════════════════════════════════════════════════════════
function handleUpdateSchedule(body) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    weeklyDay:     String(body.weeklyDay || '0'),
    weeklyHour:    String((body.weeklyTime || '08:00').split(':')[0]),
    monthlyDate:   String(body.monthlyDate || '1'),
    monthlyHour:   String((body.monthlyTime || '09:00').split(':')[0]),
    reportEmail:   body.email || '',
    toggleWeekly:  String(body.toggleWeekly  !== false),
    toggleMonthly: String(body.toggleMonthly !== false)
  });
  // Reinstall triggers with new schedule
  installTriggers();
  return jsonResp({ status: 'ok', msg: 'Schedule updated and triggers reinstalled' });
}

// ══════════════════════════════════════════════════════════
// TRIGGER INSTALLER
// Call this manually once from Apps Script editor (Run → installTriggers)
// ══════════════════════════════════════════════════════════
function installTriggers() {
  // Remove all existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  const props = PropertiesService.getScriptProperties();
  const toggleWeekly  = props.getProperty('toggleWeekly')  !== 'false';
  const toggleMonthly = props.getProperty('toggleMonthly') !== 'false';

  // Weekly report trigger
  if (toggleWeekly) {
    const weeklyHour = parseInt(props.getProperty('weeklyHour') || '8');
    const weeklyDay  = parseInt(props.getProperty('weeklyDay')  || '0'); // 0=Sun
    const dayMap     = [ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
                        ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY, ScriptApp.WeekDay.SATURDAY];
    ScriptApp.newTrigger('sendWeeklyReport')
      .timeBased().onWeekDay(dayMap[weeklyDay]).atHour(weeklyHour).create();
  }

  // Monthly report trigger
  if (toggleMonthly) {
    const monthlyDate = parseInt(props.getProperty('monthlyDate') || '1');
    const monthlyHour = parseInt(props.getProperty('monthlyHour') || '9');
    ScriptApp.newTrigger('sendMonthlyReport')
      .timeBased().onMonthDay(monthlyDate).atHour(monthlyHour).create();
  }

  Logger.log('Triggers installed successfully');
}

// ══════════════════════════════════════════════════════════
// WEEKLY REPORT EMAIL
// ══════════════════════════════════════════════════════════
function sendWeeklyReport() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('reportEmail') || CONFIG.REPORT_EMAIL;
  if (!email) { Logger.log('No report email configured'); return; }

  const ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const entries = getSheetData(ss, SHEETS.ENTRIES);
  const now     = new Date();
  const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  // Filter this week's entries
  const weekEntries = entries.filter(row => {
    const d = new Date(row[1]); return d >= weekAgo && d <= now;
  });

  const income  = weekEntries.filter(r=>r[2]==='income' ).reduce((s,r)=>s+Number(r[4]),0);
  const expense = weekEntries.filter(r=>r[2]==='expense').reduce((s,r)=>s+Number(r[4]),0);
  const savings = weekEntries.filter(r=>r[2]==='savings').reduce((s,r)=>s+Number(r[4]),0);

  const html = buildWeeklyReportHTML(income, expense, savings, weekEntries.length, now);
  const subject = `📊 NABS&FBK Home OS — সাপ্তাহিক রিপোর্ট (${formatDate(weekAgo)} — ${formatDate(now)})`;
  GmailApp.sendEmail(email, subject, `এই সপ্তাহে: আয় ৳${income} | ব্যয় ৳${expense} | সঞ্চয় ৳${savings}`, { htmlBody: html });
  Logger.log('Weekly report sent to ' + email);
}

// ══════════════════════════════════════════════════════════
// MONTHLY SUMMARY EMAIL
// ══════════════════════════════════════════════════════════
function sendMonthlyReport() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('reportEmail') || CONFIG.REPORT_EMAIL;
  if (!email) return;

  const ss      = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const entries = getSheetData(ss, SHEETS.ENTRIES);
  const now     = new Date();
  const ym      = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;

  const monthEntries = entries.filter(r => String(r[1]).startsWith(ym));
  const income  = monthEntries.filter(r=>r[2]==='income' ).reduce((s,r)=>s+Number(r[4]),0);
  const expense = monthEntries.filter(r=>r[2]==='expense').reduce((s,r)=>s+Number(r[4]),0);
  const savings = monthEntries.filter(r=>r[2]==='savings').reduce((s,r)=>s+Number(r[4]),0);

  // Category breakdown
  const catMap = {};
  monthEntries.filter(r=>r[2]==='expense').forEach(r => {
    catMap[r[3]] = (catMap[r[3]]||0) + Number(r[4]);
  });
  const topCats = Object.entries(catMap).sort((a,b)=>b[1]-a[1]).slice(0,5);

  const html    = buildMonthlyReportHTML(income, expense, savings, topCats, now);
  const MONTHS  = ['জানুয়ারি','ফেব্রুয়ারি','মার্চ','এপ্রিল','মে','জুন','জুলাই','আগস্ট','সেপ্টেম্বর','অক্টোবর','নভেম্বর','ডিসেম্বর'];
  const subject = `📅 NABS&FBK Home OS — ${MONTHS[now.getMonth()]} ${now.getFullYear()} মাসিক রিপোর্ট`;
  GmailApp.sendEmail(email, subject, `মাসিক সারসংক্ষেপ: আয় ৳${income} | ব্যয় ৳${expense} | সঞ্চয় ৳${savings}`, { htmlBody: html });
  Logger.log('Monthly report sent for ' + ym);
}

// ══════════════════════════════════════════════════════════
// EMAIL HTML BUILDERS
// ══════════════════════════════════════════════════════════
function buildWeeklyReportHTML(income, expense, savings, count, now) {
  const net = income - expense;
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:Arial,Helvetica,sans-serif;background:#0A0A0F;color:#F8F9FF;padding:20px}
      .wrap{max-width:560px;margin:0 auto;background:#0D1117;border:1px solid rgba(79,70,229,0.25);border-radius:16px;overflow:hidden}
      .top-bar{height:4px;background:linear-gradient(90deg,#4F46E5,#F5A623)}
      .header{padding:28px 32px;border-bottom:1px solid rgba(79,70,229,0.15)}
      .header h1{font-size:18px;font-weight:700;letter-spacing:0.5px;color:#F8F9FF}
      .header p{font-size:13px;color:#A0A0B0;margin-top:4px}
      .body{padding:24px 32px}
      .stats-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:24px}
      .stat{background:#111118;border-radius:10px;padding:16px;text-align:center}
      .stat .lbl{font-size:11px;color:#A0A0B0;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px}
      .stat .val{font-size:22px;font-weight:800;letter-spacing:-0.5px}
      .income{color:#22C55E}.expense{color:#EF4444}.savings{color:#4F46E5}.net{color:${net>=0?'#22C55E':'#EF4444'}}
      .info-row{display:flex;justify-content:space-between;padding:10px 14px;background:#111118;border-radius:8px;margin-bottom:8px;font-size:14px}
      .info-row .lbl{color:#A0A0B0} .info-row .val{font-weight:600;color:#F8F9FF}
      .footer{background:#0A0A0F;padding:14px 32px;text-align:center;font-size:12px;color:#606070;border-top:1px solid rgba(79,70,229,0.12)}
    </style></head><body>
    <div class="wrap">
      <div class="top-bar"></div>
      <div class="header">
        <h1>🏠 NABS&amp;FBK Home OS — সাপ্তাহিক রিপোর্ট</h1>
        <p>${formatDate(new Date(now.getTime()-7*24*60*60*1000))} থেকে ${formatDate(now)} পর্যন্ত</p>
      </div>
      <div class="body">
        <div class="stats-grid">
          <div class="stat"><div class="lbl">আয়</div><div class="val income">৳${income.toLocaleString()}</div></div>
          <div class="stat"><div class="lbl">ব্যয়</div><div class="val expense">৳${expense.toLocaleString()}</div></div>
          <div class="stat"><div class="lbl">সঞ্চয়</div><div class="val savings">৳${savings.toLocaleString()}</div></div>
        </div>
        <div class="info-row"><span class="lbl">নিট (আয় − ব্যয়)</span><span class="val net">৳${net.toLocaleString()}</span></div>
        <div class="info-row"><span class="lbl">মোট এন্ট্রি</span><span class="val">${count}টি</span></div>
      </div>
      <div class="footer">NABS&amp;FBK Home OS · Family Management System · Powered by Google Apps Script</div>
    </div></body></html>`;
}

function buildMonthlyReportHTML(income, expense, savings, topCats, now) {
  const MONTHS = ['জানুয়ারি','ফেব্রুয়ারি','মার্চ','এপ্রিল','মে','জুন','জুলাই','আগস্ট','সেপ্টেম্বর','অক্টোবর','নভেম্বর','ডিসেম্বর'];
  const net    = income - expense;
  const catRows = topCats.map(c=>`
    <div style="display:flex;justify-content:space-between;padding:8px 14px;background:#111118;border-radius:6px;margin-bottom:6px;font-size:13px">
      <span style="color:#A0A0B0">${c[0]}</span><span style="font-weight:700;color:#F5A623">৳${Number(c[1]).toLocaleString()}</span>
    </div>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:Arial,Helvetica,sans-serif;background:#0A0A0F;color:#F8F9FF;padding:20px}
      .wrap{max-width:560px;margin:0 auto;background:#0D1117;border:1px solid rgba(79,70,229,0.25);border-radius:16px;overflow:hidden}
      .top-bar{height:4px;background:linear-gradient(90deg,#F5A623,#4F46E5)}
      .header{padding:28px 32px;border-bottom:1px solid rgba(79,70,229,0.15)}
      .header h1{font-size:18px;font-weight:700;color:#F8F9FF}
      .header p{font-size:13px;color:#A0A0B0;margin-top:4px}
      .hero{background:linear-gradient(135deg,#4F46E5,#6366F1);padding:28px 32px;text-align:center;margin:0}
      .hero .month{font-size:14px;opacity:0.8;letter-spacing:2px;text-transform:uppercase;margin-bottom:8px}
      .hero .net-val{font-size:40px;font-weight:900;letter-spacing:-1px}
      .body{padding:24px 32px}
      .grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:20px}
      .stat{background:#111118;border-radius:10px;padding:14px;text-align:center}
      .stat .lbl{font-size:11px;color:#A0A0B0;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px}
      .stat .val{font-size:20px;font-weight:800}
      .section-title{font-size:12px;text-transform:uppercase;letter-spacing:1.5px;color:#606070;margin-bottom:10px}
      .footer{background:#0A0A0F;padding:14px 32px;text-align:center;font-size:12px;color:#606070;border-top:1px solid rgba(79,70,229,0.12)}
    </style></head><body>
    <div class="wrap">
      <div class="top-bar"></div>
      <div class="header">
        <h1>📅 ${MONTHS[now.getMonth()]} ${now.getFullYear()} — মাসিক রিপোর্ট</h1>
        <p>NABS&amp;FBK পরিবারের আর্থিক সারসংক্ষেপ</p>
      </div>
      <div class="hero">
        <div class="month">NET THIS MONTH</div>
        <div class="net-val" style="color:${net>=0?'#22C55E':'#EF4444'}">৳${net.toLocaleString()}</div>
      </div>
      <div class="body">
        <div class="grid3">
          <div class="stat"><div class="lbl">আয়</div><div class="val" style="color:#22C55E">৳${income.toLocaleString()}</div></div>
          <div class="stat"><div class="lbl">ব্যয়</div><div class="val" style="color:#EF4444">৳${expense.toLocaleString()}</div></div>
          <div class="stat"><div class="lbl">সঞ্চয়</div><div class="val" style="color:#4F46E5">৳${savings.toLocaleString()}</div></div>
        </div>
        <p class="section-title">শীর্ষ ব্যয়ের ক্যাটাগরি</p>
        ${catRows || '<p style="font-size:13px;color:#606070">কোনো ব্যয় নেই</p>'}
      </div>
      <div class="footer">NABS&amp;FBK Home OS · Family Management System · Powered by Google Apps Script</div>
    </div></body></html>`;
}

// ══════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ══════════════════════════════════════════════════════════
function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight('bold').setBackground('#0D1117').setFontColor('#F5A623');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSheetData(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
}

function formatDate(date) {
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'dd/MM/yyyy');
}

// Simple hash matching frontend (XOR)
function hashPass(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) {
    h = ((h << 5) - h) + str.charCodeAt(i);
    h |= 0;
  }
  return h.toString(36);
}

// ══════════════════════════════════════════════════════════
// SETUP FUNCTION — Run manually ONCE after deploy
// ══════════════════════════════════════════════════════════
function setup() {
  // 1. Create all sheets
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  getOrCreateSheet(ss, SHEETS.ENTRIES,   ['ID','Date','Type','Category','Amount','Member','PaymentMethod','Note','CreatedAt']);
  getOrCreateSheet(ss, SHEETS.PROJECTS,  ['ID','Name','Category','Status','StartDate','EndDate','Budget','ActualCost','Priority','Notes','CreatedAt']);
  getOrCreateSheet(ss, SHEETS.COSTS,     ['ID','ProjectID','Name','Amount','Date']);
  getOrCreateSheet(ss, SHEETS.TASKS,     ['ID','Text','Due','Priority','Member','Category','Done','Notes','CreatedAt']);
  getOrCreateSheet(ss, SHEETS.SHOPPING,  ['ID','Name','Qty','Price','Category','List','Bought','Note']);
  getOrCreateSheet(ss, SHEETS.DOCS,      ['ID','Title','Category','Type','Tags','Content','DriveFileId','DriveFileName','UpdatedAt']);
  getOrCreateSheet(ss, SHEETS.SETTINGS,  ['Key','Value','UpdatedAt']);
  getOrCreateSheet(ss, SHEETS.OTP_LOG,   ['Timestamp','Email','OTP','Type','Used']);
  getOrCreateSheet(ss, SHEETS.SYNC_LOG,  ['Timestamp','EntriesCount','Status']);

  // 2. Set default password hashes in Settings sheet
  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  settingsSheet.appendRow(['app_pass_hash',   hashPass('NABSFBK2025'), new Date()]);
  settingsSheet.appendRow(['admin_pass_hash', hashPass('ADMIN2025'),   new Date()]);

  // 3. Log setup
  Logger.log('Setup complete. Hash for NABSFBK2025: ' + hashPass('NABSFBK2025'));
  Logger.log('Hash for ADMIN2025: ' + hashPass('ADMIN2025'));
  Logger.log('Copy these hashes to CONFIG at the top of Code.gs');
}
