/**
 * THE LEDGER — Backend (Google Apps Script + Google Sheets)
 * Deploy as Web App (Execute as: Me, Access: Anyone with the link).
 * The bound Spreadsheet is the database. No other backend is used.
 *
 * ARCHITECTURE
 * Workspace (Personal / Business / Family / Church / …)
 *   └── Wallet (Cash / Bank / Mobile Money / Card / …)
 *         └── Transactions, Transfers, Budgets, Goals, Loans, Investments
 * Every financial record belongs to exactly one Workspace (directly or,
 * for Transactions, via its Wallet). Users are granted access to specific
 * Workspaces individually — the Owner always has access to all of them
 * and can grant/revoke access to everyone else.
 *
 * SETUP
 * 1. Create a new Google Sheet.
 * 2. Extensions > Apps Script, paste this file as Code.gs.
 * 3. Project Settings > Script Properties: add ADMIN_CODE = <a secret you choose>.
 * 4. Deploy > New deployment > Web app. Execute as "Me", Access "Anyone".
 * 5. Copy the /exec URL into the frontend's login screen as the "Server URL".
 * 6. First login uses that URL + your name + your email + the ADMIN_CODE
 *    to bootstrap you as Owner (with access to every workspace you create).
 */

// ----------------------------------------------------------------------
// SHEET SCHEMA
// ----------------------------------------------------------------------

const SHEETS = {
  Workspaces: ['id','name','icon','color','createdAt'],
  Wallets: ['id','workspaceId','name','type','currency','openingBalance','status','createdAt'],
  Transactions: ['id','workspaceId','wallet','type','category','amount','currency','date','description','tags','recurring','nextDueDate','receiptUrl','transferId','transferPeer','createdBy','createdAt'],
  Categories: ['id','workspaceId','name','type','isDefault'],
  Budgets: ['id','workspaceId','category','wallet','monthlyLimit','notes'],
  Goals: ['id','workspaceId','name','wallet','kind','targetAmount','currentAmount','deadline','notes'],
  Loans: ['id','workspaceId','direction','counterparty','principal','currency','interestRate','startDate','dueDate','status','notes'],
  LoanPayments: ['id','loanId','date','amount','notes'],
  ExchangeRates: ['id','fromCurrency','toCurrency','rate','updatedAt'],
  Preferences: ['id','workspaceId','item','category','rank','estimatedCost','notes'],
  Investments: ['id','workspaceId','name','assetType','wallet','amountInvested','currentValue','date','notes'],
  Users: ['id','name','email','code','role','invitedAt','lastSeen'],
  UserWorkspaceAccess: ['id','userId','workspaceId'],
  ActivityLog: ['id','email','action','detail','timestamp']
};

const DEFAULT_CATEGORIES = [
  ['Salary','Income',true], ['Business Income','Income',true], ['Gifts Received','Income',true], ['Other Income','Income',true],
  ['Food & Dining','Expense',true], ['Transport','Expense',true], ['Housing','Expense',true], ['Utilities','Expense',true],
  ['Health','Expense',true], ['Shopping','Expense',true], ['Entertainment','Expense',true], ['Education','Expense',true],
  ['Business Expense','Expense',true], ['Savings & Investment','Expense',true], ['Debt Repayment','Expense',true], ['Other Expense','Expense',true]
];

// ----------------------------------------------------------------------
// ENTRY POINTS
// ----------------------------------------------------------------------

function doGet(e) {
  return jsonOut({ ok: true, message: 'The Ledger API is running.' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ ok: false, error: 'Bad request body' }); }

  const action = body.action;
  try {
    ensureSchema_();
    switch (action) {
      case 'bootstrap': return jsonOut(bootstrap_(body));
      case 'login': return jsonOut(login_(body));
      case 'invite': return jsonOut(withAuth_(body, 'Owner', invite_));
      case 'updateUserRole': return jsonOut(withAuth_(body, 'Owner', updateUserRole_));
      case 'setWorkspaceAccess': return jsonOut(withAuth_(body, 'Owner', setWorkspaceAccess_));
      case 'removeUser': return jsonOut(withAuth_(body, 'Owner', removeUser_));
      case 'listAll': return jsonOut(withAuth_(body, 'Viewer', listAll_));
      case 'createWorkspace': return jsonOut(withAuth_(body, 'Editor', createWorkspace_));
      case 'updateWorkspace': return jsonOut(withAuth_(body, 'Editor', updateWorkspace_));
      case 'deleteWorkspace': return jsonOut(withAuth_(body, 'Owner', deleteWorkspace_));
      case 'create': return jsonOut(withAuth_(body, 'Editor', createRecord_));
      case 'update': return jsonOut(withAuth_(body, 'Editor', updateRecord_));
      case 'delete': return jsonOut(withAuth_(body, 'Editor', deleteRecord_));
      case 'createTransfer': return jsonOut(withAuth_(body, 'Editor', createTransfer_));
      case 'deleteTransfer': return jsonOut(withAuth_(body, 'Editor', deleteTransfer_));
      case 'uploadReceipt': return jsonOut(withAuth_(body, 'Editor', uploadReceipt_));
      case 'fetchRates': return jsonOut(withAuth_(body, 'Viewer', fetchLiveRates_));
      case 'processRecurring': return jsonOut(withAuth_(body, 'Editor', processRecurring_));
      default: return jsonOut({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------
// SCHEMA / SHEET HELPERS
// ----------------------------------------------------------------------

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSchema_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(function(name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(SHEETS[name]);
      sh.setFrozenRows(1);
    }
  });
  const trash = ss.getSheetByName('Sheet1');
  if (trash && trash.getLastRow() === 0 && trash.getLastColumn() <= 1) ss.deleteSheet(trash);
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Unknown sheet: ' + name);
  return sh;
}

function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift() || SHEETS[name];
  return values.filter(function(r) { return r.some(function(c) { return c !== '' && c !== null; }); })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function appendRow_(name, obj) {
  const sh = sheet_(name);
  const headers = SHEETS[name];
  const row = headers.map(function(h) { return (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ''; });
  sh.appendRow(row);
  return obj;
}

function findRowIndexById_(sh, headers, id) {
  const idCol = headers.indexOf('id');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) return i + 1;
  }
  return -1;
}

function updateRowById_(name, id, patch) {
  const sh = sheet_(name);
  const headers = SHEETS[name];
  const rowIdx = findRowIndexById_(sh, headers, id);
  if (rowIdx === -1) throw new Error('Record not found: ' + id);
  const current = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  const merged = headers.map(function(h, i) { return patch[h] !== undefined ? patch[h] : current[i]; });
  sh.getRange(rowIdx, 1, 1, headers.length).setValues([merged]);
  const obj = {};
  headers.forEach(function(h, i) { obj[h] = merged[i]; });
  return obj;
}

function deleteRowById_(name, id) {
  const sh = sheet_(name);
  const headers = SHEETS[name];
  const rowIdx = findRowIndexById_(sh, headers, id);
  if (rowIdx === -1) return false;
  sh.deleteRow(rowIdx);
  return true;
}

function logActivity_(email, action, detail) {
  appendRow_('ActivityLog', { id: Utilities.getUuid(), email: email, action: action, detail: detail, timestamp: new Date().toISOString() });
}

// ----------------------------------------------------------------------
// AUTH
// ----------------------------------------------------------------------

function bootstrap_(body) {
  const users = readAll_('Users');
  const props = PropertiesService.getScriptProperties();
  const adminCode = props.getProperty('ADMIN_CODE');
  if (!adminCode) return { ok: false, error: 'Server not configured: set ADMIN_CODE in Script Properties.' };
  if (users.length > 0) return { ok: false, error: 'Owner already exists. Ask them to invite you instead.' };
  if (String(body.code) !== String(adminCode)) return { ok: false, error: 'Invalid admin code.' };
  if (!body.name || !body.email) return { ok: false, error: 'Name and email are required.' };

  const userCode = generateCode_();
  const user = appendRow_('Users', {
    id: Utilities.getUuid(), name: body.name, email: String(body.email).toLowerCase(),
    code: userCode, role: 'Owner', invitedAt: new Date().toISOString(), lastSeen: new Date().toISOString()
  });
  // Seed a default "Personal" workspace so the app isn't empty on first login.
  const ws = appendRow_('Workspaces', { id: Utilities.getUuid(), name: 'Personal', icon: 'user', color: '#2563EB', createdAt: new Date().toISOString() });
  seedDefaultCategories_(ws.id);
  logActivity_(user.email, 'bootstrap', 'Created as Owner');
  return { ok: true, user: sanitizeUser_(user), accessCode: userCode };
}

function login_(body) {
  const email = String(body.email || '').toLowerCase().trim();
  const code = String(body.code || '').trim();
  const users = readAll_('Users');
  const user = users.find(function(u) { return String(u.email).toLowerCase() === email && String(u.code) === code; });
  if (!user) return { ok: false, error: 'Invalid email or access code.' };
  updateRowById_('Users', user.id, { lastSeen: new Date().toISOString() });
  logActivity_(user.email, 'login', '');
  return { ok: true, user: sanitizeUser_(user) };
}

function sanitizeUser_(u) {
  return { id: u.id, name: u.name, email: u.email, role: u.role };
}

function generateCode_() {
  return Math.random().toString(36).slice(2, 6).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
}

function withAuth_(body, minRole, fn) {
  const auth = body.auth || {};
  const users = readAll_('Users');
  const user = users.find(function(u) {
    return String(u.email).toLowerCase() === String(auth.email || '').toLowerCase() && String(u.code) === String(auth.code || '');
  });
  if (!user) return { ok: false, error: 'Not authenticated.' };
  if (!roleAtLeast_(user.role, minRole)) return { ok: false, error: 'You do not have permission to do that.' };
  updateRowById_('Users', user.id, { lastSeen: new Date().toISOString() });
  return fn(body, user);
}

function roleAtLeast_(role, min) {
  const order = { Viewer: 1, Editor: 2, Owner: 3 };
  return (order[role] || 0) >= (order[min] || 99);
}

// ----------------------------------------------------------------------
// WORKSPACE ACCESS
// ----------------------------------------------------------------------

function accessibleWorkspaceIds_(user) {
  const all = readAll_('Workspaces').map(function(w){ return w.id; });
  if (roleAtLeast_(user.role, 'Owner')) return all;
  const grants = readAll_('UserWorkspaceAccess').filter(function(g){ return g.userId === user.id; }).map(function(g){ return g.workspaceId; });
  return all.filter(function(id){ return grants.indexOf(id) > -1; });
}

function assertWorkspaceAccess_(user, workspaceId) {
  if (!workspaceId) throw new Error('workspaceId is required.');
  if (accessibleWorkspaceIds_(user).indexOf(workspaceId) === -1) throw new Error('No access to that workspace.');
}

// ----------------------------------------------------------------------
// PEOPLE (Owner only)
// ----------------------------------------------------------------------

function invite_(body, actingUser) {
  const email = String(body.email || '').toLowerCase();
  if (!email || !body.name || !body.role) return { ok: false, error: 'Name, email, and role are required.' };
  const users = readAll_('Users');
  if (users.some(function(u) { return String(u.email).toLowerCase() === email; })) {
    return { ok: false, error: 'That email is already a user.' };
  }
  const code = generateCode_();
  const user = appendRow_('Users', {
    id: Utilities.getUuid(), name: body.name, email: email, code: code,
    role: body.role, invitedAt: new Date().toISOString(), lastSeen: ''
  });
  const workspaceIds = body.workspaceIds || [];
  workspaceIds.forEach(function(wid) {
    appendRow_('UserWorkspaceAccess', { id: Utilities.getUuid(), userId: user.id, workspaceId: wid });
  });
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'You have been invited to The Ledger',
      body: actingUser.name + ' invited you to The Ledger as ' + body.role + '.\n\n' +
        'Server URL: ' + ScriptApp.getService().getUrl() + '\n' +
        'Email: ' + email + '\n' +
        'Access code: ' + code + '\n\n' +
        'Open the app, choose "Log in", and enter these three values.'
    });
  } catch (err) { /* mail failure shouldn't block invite creation */ }
  logActivity_(actingUser.email, 'invite', 'Invited ' + email + ' as ' + body.role);
  return { ok: true, user: sanitizeUser_(user), accessCode: code };
}

function updateUserRole_(body, actingUser) {
  if (!body.id || !body.role) return { ok: false, error: 'id and role required.' };
  const updated = updateRowById_('Users', body.id, { role: body.role });
  logActivity_(actingUser.email, 'updateUserRole', updated.email + ' -> ' + body.role);
  return { ok: true, user: sanitizeUser_(updated) };
}

function setWorkspaceAccess_(body, actingUser) {
  if (!body.userId || !body.workspaceIds) return { ok: false, error: 'userId and workspaceIds required.' };
  const existing = readAll_('UserWorkspaceAccess').filter(function(g){ return g.userId === body.userId; });
  existing.forEach(function(g){ deleteRowById_('UserWorkspaceAccess', g.id); });
  body.workspaceIds.forEach(function(wid){
    appendRow_('UserWorkspaceAccess', { id: Utilities.getUuid(), userId: body.userId, workspaceId: wid });
  });
  logActivity_(actingUser.email, 'setWorkspaceAccess', body.userId + ' -> ' + body.workspaceIds.join(','));
  return { ok: true };
}

function removeUser_(body, actingUser) {
  if (!body.id) return { ok: false, error: 'id required.' };
  if (body.id === actingUser.id) return { ok: false, error: 'You cannot remove yourself.' };
  deleteRowById_('Users', body.id);
  readAll_('UserWorkspaceAccess').filter(function(g){ return g.userId === body.id; }).forEach(function(g){ deleteRowById_('UserWorkspaceAccess', g.id); });
  logActivity_(actingUser.email, 'removeUser', body.id);
  return { ok: true };
}

// ----------------------------------------------------------------------
// WORKSPACES
// ----------------------------------------------------------------------

function createWorkspace_(body, user) {
  const d = body.data || {};
  if (!d.name) return { ok: false, error: 'Workspace name is required.' };
  const ws = appendRow_('Workspaces', { id: Utilities.getUuid(), name: d.name, icon: d.icon || 'briefcase', color: d.color || '#2563EB', createdAt: new Date().toISOString() });
  appendRow_('UserWorkspaceAccess', { id: Utilities.getUuid(), userId: user.id, workspaceId: ws.id });
  seedDefaultCategories_(ws.id);
  logActivity_(user.email, 'createWorkspace', ws.name);
  return { ok: true, workspace: ws };
}
function seedDefaultCategories_(workspaceId) {
  DEFAULT_CATEGORIES.forEach(function(c) {
    appendRow_('Categories', { id: Utilities.getUuid(), workspaceId: workspaceId, name: c[0], type: c[1], isDefault: c[2] });
  });
}
function updateWorkspace_(body, user) {
  assertWorkspaceAccess_(user, body.id);
  const updated = updateRowById_('Workspaces', body.id, body.data || {});
  logActivity_(user.email, 'updateWorkspace', body.id);
  return { ok: true, workspace: updated };
}
function deleteWorkspace_(body, user) {
  const wallets = readAll_('Wallets').filter(function(w){ return w.workspaceId === body.id; });
  if (wallets.length) return { ok: false, error: 'Delete or move this workspace\'s wallets first.' };
  deleteRowById_('Workspaces', body.id);
  ['Categories','Budgets','Goals','Loans','Preferences','Investments'].forEach(function(entity){
    readAll_(entity).filter(function(r){ return r.workspaceId === body.id; }).forEach(function(r){ deleteRowById_(entity, r.id); });
  });
  readAll_('UserWorkspaceAccess').filter(function(g){ return g.workspaceId === body.id; }).forEach(function(g){ deleteRowById_('UserWorkspaceAccess', g.id); });
  logActivity_(user.email, 'deleteWorkspace', body.id);
  return { ok: true };
}

// ----------------------------------------------------------------------
// GENERIC CRUD — every entity here carries (or resolves to) a workspaceId
// ----------------------------------------------------------------------

const CRUD_ENTITIES = ['Wallets','Transactions','Categories','Budgets','Goals','Loans','LoanPayments','Preferences','Investments','ExchangeRates'];
const WORKSPACE_SCOPED = ['Wallets','Transactions','Categories','Budgets','Goals','Loans','Preferences','Investments'];

function assertEntity_(entity) {
  if (CRUD_ENTITIES.indexOf(entity) === -1) throw new Error('Not a CRUD entity: ' + entity);
}

// LoanPayments has no workspaceId of its own — resolve via its Loan.
function workspaceIdForRecord_(entity, data) {
  if (entity === 'LoanPayments') {
    const loan = readAll_('Loans').find(function(l){ return l.id === data.loanId; });
    return loan ? loan.workspaceId : null;
  }
  return data.workspaceId;
}

function listAll_(body, user) {
  const wsIds = accessibleWorkspaceIds_(user);
  const out = { ok: true, users: readAll_('Users').map(sanitizeUser_), workspaces: readAll_('Workspaces').filter(function(w){ return wsIds.indexOf(w.id) > -1; }) };
  WORKSPACE_SCOPED.forEach(function(name) {
    out[name] = readAll_(name).filter(function(r) { return wsIds.indexOf(r.workspaceId) > -1; });
  });
  const visibleLoanIds = (out.Loans || []).map(function(l){ return l.id; });
  out.LoanPayments = readAll_('LoanPayments').filter(function(p){ return visibleLoanIds.indexOf(p.loanId) > -1; });
  out.ExchangeRates = readAll_('ExchangeRates'); // currency rates are global, not workspace-scoped
  if (roleAtLeast_(user.role, 'Owner')) {
    out.activityLog = readAll_('ActivityLog').slice(-200);
    out.userWorkspaceAccess = readAll_('UserWorkspaceAccess');
  }
  return out;
}

function createRecord_(body, user) {
  assertEntity_(body.entity);
  const obj = body.data || {};
  obj.id = obj.id || Utilities.getUuid();
  const wsId = workspaceIdForRecord_(body.entity, obj);
  if (WORKSPACE_SCOPED.indexOf(body.entity) > -1 || body.entity === 'LoanPayments') assertWorkspaceAccess_(user, wsId);
  if (body.entity === 'Transactions') { obj.createdBy = user.email; obj.createdAt = new Date().toISOString(); }
  const created = appendRow_(body.entity, obj);
  logActivity_(user.email, 'create:' + body.entity, obj.id);
  return { ok: true, record: created };
}

function updateRecord_(body, user) {
  assertEntity_(body.entity);
  const existingList = readAll_(body.entity);
  const existing = existingList.find(function(r){ return r.id === body.id; });
  if (!existing) return { ok: false, error: 'Record not found.' };
  const wsId = workspaceIdForRecord_(body.entity, existing);
  assertWorkspaceAccess_(user, wsId);
  if (body.entity === 'Transactions' && existing.transferId) {
    return { ok: false, error: 'Editing a Transfer is disabled. Delete it and recreate instead.' };
  }
  const updated = updateRowById_(body.entity, body.id, body.data || {});
  logActivity_(user.email, 'update:' + body.entity, body.id);
  return { ok: true, record: updated };
}

function deleteRecord_(body, user) {
  assertEntity_(body.entity);
  const existing = readAll_(body.entity).find(function(r){ return r.id === body.id; });
  if (!existing) return { ok: true };
  const wsId = workspaceIdForRecord_(body.entity, existing);
  assertWorkspaceAccess_(user, wsId);
  if (body.entity === 'Transactions' && existing.transferId) {
    return deleteTransfer_({ transferId: existing.transferId }, user);
  }
  const ok = deleteRowById_(body.entity, body.id);
  logActivity_(user.email, 'delete:' + body.entity, body.id);
  return { ok: ok };
}

// ----------------------------------------------------------------------
// TRANSFERS (linked pair of transactions — same or different Workspaces)
// ----------------------------------------------------------------------

function createTransfer_(body, user) {
  const d = body.data || {};
  if (!d.fromWallet || !d.toWallet || d.fromWallet === d.toWallet) {
    return { ok: false, error: 'Pick two different wallets.' };
  }
  const wallets = readAll_('Wallets');
  const fromW = wallets.find(function(w){ return w.id === d.fromWallet; });
  const toW = wallets.find(function(w){ return w.id === d.toWallet; });
  if (!fromW || !toW) return { ok: false, error: 'Wallet not found.' };
  assertWorkspaceAccess_(user, fromW.workspaceId);
  assertWorkspaceAccess_(user, toW.workspaceId);

  const transferId = Utilities.getUuid();
  const now = new Date().toISOString();
  const legOut = {
    id: Utilities.getUuid(), workspaceId: fromW.workspaceId, wallet: d.fromWallet, type: 'Transfer',
    category: 'Transfer', amount: -Math.abs(d.amount), currency: fromW.currency, date: d.date,
    description: d.description || ('Transfer to ' + toW.name), tags: '', recurring: '',
    nextDueDate: '', receiptUrl: '', transferId: transferId, transferPeer: d.toWallet,
    createdBy: user.email, createdAt: now
  };
  const legIn = {
    id: Utilities.getUuid(), workspaceId: toW.workspaceId, wallet: d.toWallet, type: 'Transfer',
    category: 'Transfer', amount: Math.abs(d.amount), currency: toW.currency, date: d.date,
    description: d.description || ('Transfer from ' + fromW.name), tags: '', recurring: '',
    nextDueDate: '', receiptUrl: '', transferId: transferId, transferPeer: d.fromWallet,
    createdBy: user.email, createdAt: now
  };
  appendRow_('Transactions', legOut);
  appendRow_('Transactions', legIn);
  logActivity_(user.email, 'createTransfer', transferId);
  return { ok: true, legs: [legOut, legIn] };
}

function deleteTransfer_(body, user) {
  const all = readAll_('Transactions');
  const legs = all.filter(function(t) { return t.transferId === body.transferId; });
  legs.forEach(function(t) { assertWorkspaceAccess_(user, t.workspaceId); });
  legs.forEach(function(t) { deleteRowById_('Transactions', t.id); });
  logActivity_(user.email, 'deleteTransfer', body.transferId);
  return { ok: true, deleted: legs.length };
}

// ----------------------------------------------------------------------
// RECURRING TRANSACTIONS
// ----------------------------------------------------------------------

function processRecurring_(body, user) {
  const wsIds = accessibleWorkspaceIds_(user);
  const all = readAll_('Transactions').filter(function(t){ return wsIds.indexOf(t.workspaceId) > -1; });
  const templates = all.filter(function(t) { return t.recurring && t.nextDueDate; });
  const today = new Date(); today.setHours(0,0,0,0);
  const created = [];
  templates.forEach(function(t) {
    let due = new Date(t.nextDueDate);
    let guard = 0;
    while (due <= today && guard < 24) {
      const copy = {
        id: Utilities.getUuid(), workspaceId: t.workspaceId, date: Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        type: t.type, wallet: t.wallet, category: t.category, amount: t.amount, currency: t.currency,
        description: t.description, tags: t.tags, recurring: '', nextDueDate: '', receiptUrl: '',
        transferId: '', transferPeer: '', createdBy: 'recurring:' + t.id, createdAt: new Date().toISOString()
      };
      appendRow_('Transactions', copy);
      created.push(copy);
      due = advanceDate_(due, t.recurring);
      guard++;
    }
    updateRowById_('Transactions', t.id, { nextDueDate: Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd') });
  });
  if (created.length) logActivity_(user.email, 'processRecurring', created.length + ' generated');
  return { ok: true, created: created };
}

function advanceDate_(date, freq) {
  const d = new Date(date);
  if (freq === 'Weekly') d.setDate(d.getDate() + 7);
  else if (freq === 'Yearly') d.setFullYear(d.getFullYear() + 1);
  else d.setMonth(d.getMonth() + 1);
  return d;
}

// ----------------------------------------------------------------------
// RECEIPTS (Google Drive)
// ----------------------------------------------------------------------

function uploadReceipt_(body, user) {
  const folder = getOrCreateReceiptsFolder_();
  const bytes = Utilities.base64Decode(body.base64);
  const blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', body.filename || (Utilities.getUuid() + '.jpg'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  logActivity_(user.email, 'uploadReceipt', file.getId());
  return { ok: true, url: 'https://drive.google.com/uc?id=' + file.getId(), fileId: file.getId() };
}
function getOrCreateReceiptsFolder_() {
  const name = 'Ledger Receipts';
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

// ----------------------------------------------------------------------
// EXCHANGE RATES (free public API, no key) — global, not workspace-scoped
// ----------------------------------------------------------------------

function fetchLiveRates_(body, user) {
  const base = body.base || 'USD';
  try {
    const resp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/' + encodeURIComponent(base), { muteHttpExceptions: true });
    const data = JSON.parse(resp.getContentText());
    if (!data.rates) return { ok: false, error: 'Rate provider returned no data.' };
    const targets = body.targets || Object.keys(data.rates).slice(0, 30);
    const now = new Date().toISOString();
    const results = [];
    targets.forEach(function(cur) {
      if (!data.rates[cur]) return;
      const existing = readAll_('ExchangeRates').find(function(r) { return r.fromCurrency === base && r.toCurrency === cur; });
      const row = { id: existing ? existing.id : Utilities.getUuid(), fromCurrency: base, toCurrency: cur, rate: data.rates[cur], updatedAt: now };
      if (existing) updateRowById_('ExchangeRates', existing.id, row); else appendRow_('ExchangeRates', row);
      results.push(row);
    });
    logActivity_(user.email, 'fetchRates', base);
    return { ok: true, rates: results };
  } catch (err) {
    return { ok: false, error: 'Could not reach rate provider: ' + err };
  }
}
