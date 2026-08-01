/**
 * THE LEDGER — Google Sheets Sync Backend (v4: real app push notifications
 * via Firebase Cloud Messaging, and receipt photo storage via Google Drive)
 * ------------------------------------------------------
 * SETUP — same as before if you've already deployed this once:
 * 1. sheets.new to create a Sheet (or reuse your existing one).
 * 2. Extensions > Apps Script. Replace all code with this file.
 * 3. Script Properties: ADMIN_CODE (your owner password), APP_URL (optional).
 * 4. Deploy > New deployment > Web app > Execute as Me > Access: Anyone.
 * 5. Nothing destructive happens to an existing sheet — new tabs/columns are
 *    created automatically on first request.
 *
 * PUSH NOTIFICATIONS (Firebase Cloud Messaging — free tier):
 * 1. Go to console.firebase.google.com > Add project (free, no card needed).
 * 2. In the new project: Project settings (gear) > Cloud Messaging tab.
 *    Confirm Cloud Messaging API (v1) is enabled (it is by default).
 * 3. Project settings > Service accounts > Generate new private key.
 *    This downloads a JSON file — open it and copy its ENTIRE contents.
 * 4. Back in Apps Script: Script Properties > add FCM_SERVICE_ACCOUNT_JSON,
 *    paste that whole JSON as the value (it's one line, that's fine).
 * 5. In the same Firebase project: Project settings > General > "Your apps" >
 *    add a Web app. Copy the firebaseConfig object it gives you — you'll
 *    paste that into index.html (see firebase-init.js / SETUP.md).
 * 6. Also grab your "Web Push certificate" VAPID key: Cloud Messaging tab >
 *    Web configuration > Generate key pair. Paste that into index.html too.
 * 7. Apps Script > Triggers (clock icon) > + Add Trigger:
 *      Function: checkAndNotify | Event source: Time-driven | Day timer
 *    Save — Google's servers now run this daily on their own, no button,
 *    no phone needing to be on, no app needing to be open.
 * Full copy-paste steps are in SETUP.md.
 */

const SHEETS = {
  transactions: ['id','date','type','account','category','amount','description','tags','recurring','nextDueDate','receiptUrl','createdAt'],
  budgets: ['id','category','account','monthlyLimit','notes'],
  goals: ['id','name','account','kind','targetAmount','currentAmount','deadline','notes'],
  investments: ['id','name','account','assetType','amountInvested','currentValue','date','notes'],
  preferences: ['id','item','category','rank','estimatedCost','notes'],
  accounts: ['id','name','type','currency','openingBalance','status','createdAt'],
  categories: ['id','name','type','isDefault'],
  loans: ['id','direction','counterparty','principal','currency','interestRate','startDate','dueDate','status','notes'],
  loanPayments: ['id','loanId','date','amount','notes'],
  exchangeRates: ['id','fromCurrency','toCurrency','rate','updatedAt']
};
const USERS_HEADERS = ['id','name','email','code','role','fcmToken','notifyEnabled','invitedAt','lastSeen'];
const LOG_HEADERS = ['id','email','action','detail','timestamp'];
const ROLE_RANK = { 'Viewer':1, 'Editor':2, 'Owner':3 };
const RECEIPTS_FOLDER_NAME = 'Ledger Receipts';

const DEFAULT_ACCOUNTS = [
  { name:'Personal', type:'Personal', currency:'GHS', openingBalance:0, status:'Active' },
  { name:'Business', type:'Business', currency:'GHS', openingBalance:0, status:'Active' }
];
const DEFAULT_CATEGORIES = [
  ...['Salary','Business Income','Interest','Investment','Gift','Commission','Bonus','Rental Income','Other Income'].map(n=>({name:n,type:'Income'})),
  ...['Food','Fuel','Transport','Utilities','Medical','Education','Entertainment','Shopping','Insurance','Rent','Maintenance','Tax','Travel','Other Expenses'].map(n=>({name:n,type:'Expense'}))
];


function props_(){ return PropertiesService.getScriptProperties(); }

function getOrCreateSheet_(ss, name, headers){
  let sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); seedDefaults_(sh, name, headers); }
  else if (sh.getLastRow() === 0){ sh.appendRow(headers); sh.setFrozenRows(1); seedDefaults_(sh, name, headers); }
  return sh;
}
function seedDefaults_(sh, name, headers){
  if (name === 'accounts'){
    DEFAULT_ACCOUNTS.forEach(a=>{
      const row = Object.assign({ id: Utilities.getUuid(), createdAt: new Date().toISOString() }, a);
      sh.appendRow(headers.map(h=>row[h]!==undefined?row[h]:''));
    });
  }
  if (name === 'categories'){
    DEFAULT_CATEGORIES.forEach(c=>{
      const row = Object.assign({ id: Utilities.getUuid(), isDefault: true }, c);
      sh.appendRow(headers.map(h=>row[h]!==undefined?row[h]:''));
    });
  }
}
function usersSheet_(ss){ return getOrCreateSheet_(ss, 'Users', USERS_HEADERS); }
function logSheet_(ss){ return getOrCreateSheet_(ss, 'ActivityLog', LOG_HEADERS); }
function logActivity_(ss, email, action, detail){
  try{
    const sh = logSheet_(ss);
    sh.appendRow([Utilities.getUuid(), email, action, detail||'', new Date().toISOString()]);
    // keep log bounded to last 300 rows
    const last = sh.getLastRow();
    if (last > 301) sh.deleteRows(2, last-301);
  }catch(e){ /* logging is best-effort */ }
}

function sheetToObjects_(sh){
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => { const o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
}
function writeObjectsToSheet_(sh, headers, objects){
  sh.clear();
  sh.appendRow(headers);
  sh.setFrozenRows(1);
  if (objects.length === 0) return;
  const rows = objects.map(obj => headers.map(h => obj[h] !== undefined ? obj[h] : ''));
  sh.getRange(2,1,rows.length,headers.length).setValues(rows);
}
function genCode_(){
  return Math.random().toString(36).slice(2,6).toUpperCase() + '-' + Math.random().toString(36).slice(2,6).toUpperCase();
}

function authenticate_(ss, email, code){
  const sh = usersSheet_(ss);
  const users = sheetToObjects_(sh);
  email = (email||'').trim().toLowerCase();
  code = (code||'').trim();

  if (users.length === 0){
    const adminCode = (props_().getProperty('ADMIN_CODE')||'').trim();
    if (!adminCode) throw new Error('Server not configured: set ADMIN_CODE in Script Properties.');
    if (code !== adminCode) throw new Error('Invalid code.');
    const row = { id: Utilities.getUuid(), name: email.split('@')[0], email, code: adminCode, role:'Owner',
                  fcmToken:'', notifyEnabled:false,
                  invitedAt: new Date().toISOString(), lastSeen: new Date().toISOString() };
    sh.appendRow(USERS_HEADERS.map(h=>row[h]!==undefined?row[h]:''));
    return { role:'Owner', match: row };
  }

  const match = users.find(u => String(u.email).toLowerCase() === email && String(u.code).trim() === code);
  if (!match) throw new Error('Invalid email or access code.');

  const rowIdx = users.findIndex(u => String(u.email).toLowerCase() === email && String(u.code).trim() === code);
  sh.getRange(rowIdx+2, USERS_HEADERS.indexOf('lastSeen')+1).setValue(new Date().toISOString());

  return { role: match.role, match };
}
function requireRole_(role, minRole){
  if ((ROLE_RANK[role]||0) < (ROLE_RANK[minRole]||99)) throw new Error('You need ' + minRole + ' access for this action.');
}

function doGet(e){
  try{
    const email = e.parameter.email, code = e.parameter.code;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auth = authenticate_(ss, email, code);

    if (e.parameter.action === 'rates'){
      const base = (e.parameter.base || 'GHS').toUpperCase();
      try{
        const res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/' + base, { muteHttpExceptions:true });
        const json = JSON.parse(res.getContentText());
        if (json.result === 'success'){
          return jsonOut_({ ok:true, base, rates: json.rates, fetchedAt: new Date().toISOString() });
        }
        return jsonOut_({ ok:false, error:'Rate provider error.' });
      }catch(err){ return jsonOut_({ ok:false, error:'Could not reach exchange rate provider.' }); }
    }

    const data = {};
    Object.keys(SHEETS).forEach(name=>{
      const sh = getOrCreateSheet_(ss, name, SHEETS[name]);
      data[name] = sheetToObjects_(sh);
    });

    let users = [], log = [];
    if (auth.role === 'Owner'){
      users = sheetToObjects_(usersSheet_(ss)).map(u=>({name:u.name,email:u.email,role:u.role,invitedAt:u.invitedAt,lastSeen:u.lastSeen}));
      log = sheetToObjects_(logSheet_(ss)).slice(-50).reverse();
    }

    return jsonOut_({ ok:true, role: auth.role, name: auth.match.name, notifyEnabled: !!auth.match.notifyEnabled, data, users, log });
  }catch(err){
    return jsonOut_({ ok:false, error: err.message });
  }
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auth = authenticate_(ss, body.email, body.code);

    if (body.action === 'sync'){
      requireRole_(auth.role, 'Editor');
      Object.keys(SHEETS).forEach(name=>{
        if (body.data && body.data[name]){
          const sh = getOrCreateSheet_(ss, name, SHEETS[name]);
          writeObjectsToSheet_(sh, SHEETS[name], body.data[name]);
        }
      });
      logActivity_(ss, body.email, 'sync', 'Pushed data update');
      return jsonOut_({ ok:true, savedAt:new Date().toISOString() });
    }

    if (body.action === 'invite'){
      requireRole_(auth.role, 'Owner');
      const sh = usersSheet_(ss);
      const users = sheetToObjects_(sh);
      const inviteeEmail = String(body.inviteeEmail||'').trim().toLowerCase();
      if (!inviteeEmail) throw new Error('Email required.');
      if (users.some(u=>String(u.email).toLowerCase()===inviteeEmail)) throw new Error('That person is already invited.');
      const role = (body.role === 'Editor') ? 'Editor' : 'Viewer';
      const code = genCode_();
      const row = { id: Utilities.getUuid(), name: body.inviteeName||inviteeEmail.split('@')[0], email: inviteeEmail,
                    code, role, fcmToken:'', notifyEnabled:false, invitedAt: new Date().toISOString(), lastSeen:'' };
      sh.appendRow(USERS_HEADERS.map(h=>row[h]!==undefined?row[h]:''));

      const appUrl = props_().getProperty('APP_URL');
      const scriptUrl = ScriptApp.getService().getUrl();
      const link = appUrl ? (appUrl + (appUrl.includes('?')?'&':'?') + 'u='+encodeURIComponent(scriptUrl)+'&e='+encodeURIComponent(inviteeEmail)+'&c='+encodeURIComponent(code)) : null;
      const emailBody = 'You have been invited to a shared finance workspace (' + role + ' access).\n\n' +
        (link ? ('Open this link to get in automatically:\n'+link+'\n\n') : '') +
        'Or enter these details manually in the app:\n' +
        'Sheet URL: ' + scriptUrl + '\n' + 'Email: ' + inviteeEmail + '\n' + 'Access code: ' + code + '\n';
      try{ MailApp.sendEmail(inviteeEmail, 'You\'ve been invited to a shared Ledger workspace', emailBody); }catch(mailErr){}
      logActivity_(ss, body.email, 'invite', inviteeEmail + ' as ' + role);
      return jsonOut_({ ok:true, code, emailed:true });
    }

    if (body.action === 'updateRole'){
      requireRole_(auth.role, 'Owner');
      const sh = usersSheet_(ss);
      const users = sheetToObjects_(sh);
      const idx = users.findIndex(u=>String(u.email).toLowerCase()===String(body.targetEmail).toLowerCase());
      if (idx===-1) throw new Error('User not found.');
      const newRole = body.newRole==='Editor'?'Editor':(body.newRole==='Owner'?'Owner':'Viewer');
      sh.getRange(idx+2, USERS_HEADERS.indexOf('role')+1).setValue(newRole);
      logActivity_(ss, body.email, 'role change', body.targetEmail + ' -> ' + newRole);
      return jsonOut_({ ok:true });
    }

    if (body.action === 'removeUser'){
      requireRole_(auth.role, 'Owner');
      const sh = usersSheet_(ss);
      const users = sheetToObjects_(sh);
      const idx = users.findIndex(u=>String(u.email).toLowerCase()===String(body.targetEmail).toLowerCase());
      if (idx===-1) throw new Error('User not found.');
      sh.deleteRow(idx+2);
      logActivity_(ss, body.email, 'remove user', body.targetEmail);
      return jsonOut_({ ok:true });
    }

    if (body.action === 'registerPushToken'){
      // Self-service: any authenticated user can register their own device token.
      const sh = usersSheet_(ss);
      const users = sheetToObjects_(sh);
      const idx = users.findIndex(u=>String(u.email).toLowerCase()===String(body.email).toLowerCase());
      if (idx===-1) throw new Error('User not found.');
      sh.getRange(idx+2, USERS_HEADERS.indexOf('fcmToken')+1).setValue(body.fcmToken||'');
      sh.getRange(idx+2, USERS_HEADERS.indexOf('notifyEnabled')+1).setValue(!!body.notifyEnabled);
      return jsonOut_({ ok:true });
    }

    if (body.action === 'uploadImage'){
      requireRole_(auth.role, 'Editor');
      if (!body.base64 || !body.filename) throw new Error('Missing image data.');
      const folder = getOrCreateReceiptsFolder_();
      const bytes = Utilities.base64Decode(body.base64);
      const blob = Utilities.newBlob(bytes, body.mimeType||'image/jpeg', body.filename);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      logActivity_(ss, body.email, 'upload receipt', body.filename);
      return jsonOut_({ ok:true, url });
    }

    throw new Error('Unknown action.');
  }catch(err){
    return jsonOut_({ ok:false, error: err.message });
  }
}
function getOrCreateReceiptsFolder_(){
  const folders = DriveApp.getFoldersByName(RECEIPTS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(RECEIPTS_FOLDER_NAME);
}
function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================
   FIREBASE CLOUD MESSAGING (HTTP v1 — the legacy FCM API was
   retired by Google in 2024, so this uses a service-account JWT)
========================================================= */
function getFcmAccessToken_(){
  const raw = props_().getProperty('FCM_SERVICE_ACCOUNT_JSON');
  if (!raw) throw new Error('FCM_SERVICE_ACCOUNT_JSON not set in Script Properties.');
  const sa = JSON.parse(raw);
  const header = { alg:'RS256', typ:'JWT' };
  const now = Math.floor(Date.now()/1000);
  const claim = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  };
  const b64 = (obj)=> Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/,'');
  const toSign = b64(header) + '.' + b64(claim);
  const sigBytes = Utilities.computeRsaSha256Signature(toSign, sa.private_key);
  const sig = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/,'');
  const jwt = toSign + '.' + sig;

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method:'post',
    contentType:'application/x-www-form-urlencoded',
    payload: { grant_type:'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions:true
  });
  const json = JSON.parse(res.getContentText());
  if (!json.access_token) throw new Error('Could not get FCM access token: ' + res.getContentText());
  return { token: json.access_token, projectId: sa.project_id };
}
function sendPush_(fcmToken, title, body){
  const { token, projectId } = getFcmAccessToken_();
  const payload = {
    message: {
      token: fcmToken,
      notification: { title, body },
      webpush: { fcm_options: { link: props_().getProperty('APP_URL') || '/' } }
    }
  };
  const res = UrlFetchApp.fetch('https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send', {
    method:'post',
    contentType:'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions:true
  });
  return JSON.parse(res.getContentText());
}

/**
 * Runs on a time-driven trigger (see setup notes at top of file) — no
 * button, no app needing to be open. Computes the same kind of alerts as
 * the in-app Insights tab, and pushes a native notification to anyone
 * who has notifications enabled.
 */
function checkAndNotify(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {};
  Object.keys(SHEETS).forEach(name=>{ data[name] = sheetToObjects_(getOrCreateSheet_(ss, name, SHEETS[name])); });
  const users = sheetToObjects_(usersSheet_(ss)).filter(u=>u.notifyEnabled && u.fcmToken);
  if (users.length === 0) return;

  const alerts = [];
  const today = new Date(); today.setHours(0,0,0,0);

  // Budgets exceeded this month
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  data.budgets.forEach(b=>{
    const spent = data.transactions.filter(t=>t.type==='Expense' && t.category===b.category && t.account===b.account && new Date(t.date) >= monthStart)
      .reduce((s,t)=>s+Number(t.amount||0),0);
    if (Number(b.monthlyLimit) > 0 && spent >= Number(b.monthlyLimit)) alerts.push(b.category + ' budget exceeded (' + b.account + ')');
  });
  // Loans due within 3 days
  data.loans.forEach(l=>{
    if (l.status === 'Paid' || !l.dueDate) return;
    const paid = data.loanPayments.filter(p=>p.loanId===l.id).reduce((s,p)=>s+Number(p.amount||0),0);
    const bal = Number(l.principal||0) - paid;
    if (bal <= 0) return;
    const days = Math.ceil((new Date(l.dueDate) - today) / 86400000);
    if (days <= 3) alerts.push('Loan ' + (l.direction==='Lent'?'to':'from') + ' ' + l.counterparty + (days<0?' is overdue':' due in '+days+'d'));
  });
  // No transaction logged today
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const hasToday = data.transactions.some(t=>String(t.date).slice(0,10) === todayStr);
  if (!hasToday) alerts.push('No transaction logged today yet');

  if (alerts.length === 0) return;
  const title = 'The Ledger — ' + alerts.length + ' alert' + (alerts.length>1?'s':'');
  const body = alerts.slice(0,3).join(' · ');

  users.forEach(u=>{
    try{ sendPush_(u.fcmToken, title, body); }catch(err){ /* one bad token shouldn't stop the rest */ }
  });
  logActivity_(ss, 'system', 'push notify', users.length + ' user(s), ' + alerts.length + ' alert(s)');
}
