// USAGE (3 STEPS):
// 1. Add BOT_TOKEN in Project Settings > Script Properties.
// 2. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
// 3. Run setupWebhook() function, then message "start" to your bot on Zalo.
//
// ---------------------------------------------------------
// TROUBLESHOOTING (If setupWebhook fails):
// The app requires the following Script Properties (Project Settings) to run:
// - BOT_TOKEN : Your Zalo bot API token (Get from Zalo Developer portal). Used to call Zalo APIs.
// - SHEET_ID  : ID of this Google Sheet (Found in the URL between /d/ and /edit). Used to read/write data.
// - WEBAPP_URL: The /exec link of your deployment (Get it after Step 2). Used to register the webhook with Zalo.
// - URL_KEY   : A random secret string (You can make one up). Used to verify requests from Zalo.
//
// SCRIPT VARIABLES:
// - DEFAULT_TYPE: Required default type. Format: [code, name, target, group].
// - TYPES: Optional array of other predefined types. Format: [[code, name, target, group], ...].

const DEFAULT_TYPE = ['needs', 'Needs', 0, 'expense'];

const TYPES = [
  ['income', 'Income', 0, 'income']
];

const API_BASE = 'https://bot-api.zaloplatforms.com/bot';

const SHEETS = {
  Expenses: ['date', 'note', 'type', 'amount'],
  Types: ['code', 'name', 'monthly_target', 'group'],
};

function prop_(name) {
  return PropertiesService.getScriptProperties().getProperty(name);
}

function book_() {
  return SpreadsheetApp.openById(String(prop_('SHEET_ID') || '').trim());
}

function sheet_(name) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (SHEETS[name]) {
      sh.appendRow(SHEETS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold');
    }
  }
  return sh;
}

function dict_() {
  return Object.create(null);
}

function clean_(s) {
  return String(s === null || s === undefined ? '' : s).trim().toLowerCase();
}

function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase();
}

function money_(n) {
  const a = Math.abs(n);
  if (a === 0) return '0';
  const units = [['d', 0.001], ['k', 1], ['m', 1000], ['b', 1000000]];
  let i = a >= 1000000 ? 3 : a >= 1000 ? 2 : a >= 1 ? 1 : 0;
  let v = Math.round(a / units[i][1] * 100) / 100;
  if (v >= 1000 && i < 3) {
    i++;
    v = Math.round(a / units[i][1] * 100) / 100;
  }
  return (n < 0 ? '-' : '') + String(v).replace('.', ',') + units[i][0];
}

function parseAmount_(s, allowZero) {
  s = String(s).trim();
  if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  else s = s.replace(',', '.');
  if (!/^\d+(\.\d+)?$/.test(s)) return null;
  const n = parseFloat(s);
  return (n > 0 || (allowZero && n === 0)) ? n : null;
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseAmount_(v, true);
  return n === null ? 0 : n;
}

function callZaloRaw_(method, payload) {
  const res = UrlFetchApp.fetch(API_BASE + prop_('BOT_TOKEN') + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true,
  });
  return { code: res.getResponseCode(), text: res.getContentText() };
}

function sendText_(chatId, text) {
  const r = callZaloRaw_('sendMessage', { chat_id: chatId, text: text });
  if (r.code !== 200) console.error('sendMessage ' + r.code + ' ' + r.text);
}

function getChatIds_() {
  try { return JSON.parse(prop_('CHAT_IDS') || '[]'); } catch (e) { return []; }
}

function rememberChat_(chatId) {
  const ids = getChatIds_();
  const id = String(chatId);
  if (ids.indexOf(id) > -1) return false;
  ids.push(id);
  PropertiesService.getScriptProperties().setProperty('CHAT_IDS', JSON.stringify(ids));
  return true;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e.parameter || e.parameter.key !== prop_('URL_KEY')) return out_('forbidden');

    const body = JSON.parse(e.postData.contents);
    const msg = (body.result && body.result.message) || body.message;
    if (!msg || !msg.text) return out_('ignored');

    const chatId = msg.chat.id;
    const msgId = msg.message_id;

    const cache = CacheService.getScriptCache();
    if (msgId) {
      if (cache.get('m_' + msgId)) return out_('duplicate');
      cache.put('m_' + msgId, '1', 600);
    }

    lock.waitLock(15000);
    let reply;
    try {
      const firstWord = norm_(String(msg.text).trim().split(/\s+/)[0]).replace(/^\//, '');
      const isNew = rememberChat_(chatId);
      if (isNew && ['start', 'help'].indexOf(firstWord) === -1) {
        sendText_(chatId, 'Hello! I am your expense tracker bot.\n\n' + helpText_());
      }
      reply = handle_(msg.text);
    } catch (err) {
      console.error('handle: ' + err);
      reply = 'Error processing request. Type "help" for instructions.';
    }
    sendText_(chatId, reply);
  } catch (err) {
    console.error('doPost: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
  return out_('ok');
}

function out_(text) {
  return ContentService.createTextOutput(text);
}

function doGet(e) {
  const u = ScriptApp.getService().getUrl();
  if (u && /\/exec$/.test(u) && prop_('WEBAPP_URL') !== u) {
    PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', u);
  }
  return ContentService.createTextOutput('Bot is running. Web app URL recorded, go back to Apps Script and run setupWebhook.');
}

function handle_(text) {
  const parts = String(text).trim().split(/\s+/);
  const cmd = norm_(parts[0]).replace(/^\//, '');
  switch (cmd) {
    case 'start':
    case 'help':
      return helpText_();
    case 'report':
      return report_(parts[1]);
    case 'types':
      return listTypes_();
    case 'undo':
      return undoExpense_();
  }
  return addExpense_(parts);
}

function loadTypes_() {
  const map = dict_();
  const fromScript = [{ def: DEFAULT_TYPE, isDefault: true }]
    .concat(TYPES.map(function (t) { return { def: t, isDefault: false }; }));
  fromScript.forEach(function (x) {
    const code = clean_(x.def && x.def[0]);
    if (!code || map[code]) return;
    map[code] = {
      code: code,
      name: String(x.def[1] || '').trim() || code,
      target: num_(x.def[2]),
      group: String(x.def[3] || 'expense').trim().toLowerCase() === 'income' ? 'income' : 'expense',
      source: 'script',
      isDefault: x.isDefault,
    };
  });

  const sh = sheet_('Types');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    sh.getRange(2, 1, n, 4).getValues().forEach(function (r) {
      const code = clean_(r[0]);
      if (!code) return;
      const isDefault = map[code] ? map[code].isDefault : false;
      const group = String(r[3] || 'expense').trim().toLowerCase() === 'income' ? 'income' : 'expense';
      map[code] = { code: code, name: String(r[1]).trim() || code, target: num_(r[2]), group: group, source: 'sheet', isDefault: isDefault };
    });
  }
  
  return map;
}

function defaultType_(types) {
  return types[clean_(DEFAULT_TYPE && DEFAULT_TYPE[0])] || null;
}

function resolveType_(text, types) {
  const k = clean_(text);
  if (!k) return null;
  if (types[k]) return types[k];
  const codes = Object.keys(types);
  for (let i = 0; i < codes.length; i++) {
    if (clean_(types[codes[i]].name) === k) return types[codes[i]];
  }
  return null;
}

function findTypeRow_(code) {
  const sh = sheet_('Types');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const codes = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (clean_(codes[i][0]) === code) return i + 2;
    }
  }
  return -1;
}

function sumByType_(from, to, types) {
  const sh = sheet_('Expenses');
  const lastRow = sh.getLastRow();
  const out = dict_();
  if (lastRow < 2) return out;

  const BATCH_SIZE = 500;
  let startRow = lastRow;
  let keepGoing = true;

  while (startRow >= 2 && keepGoing) {
    const fetchStart = Math.max(2, startRow - BATCH_SIZE + 1);
    const fetchCount = startRow - fetchStart + 1;
    const data = sh.getRange(fetchStart, 1, fetchCount, 4).getValues();
    
    let allDatesBeforeFrom = true;
    for (let i = fetchCount - 1; i >= 0; i--) {
      const d = data[i][0];
      if (!(d instanceof Date)) continue;
      
      if (d >= from) {
        allDatesBeforeFrom = false;
        if (d < to) {
          const raw = String(data[i][2]).trim();
          const t = resolveType_(raw, types);
          const key = t ? t.code : '?' + raw;
          out[key] = (out[key] || 0) + num_(data[i][3]);
        }
      }
    }
    if (allDatesBeforeFrom) keepGoing = false;
    startRow = fetchStart - 1;
  }
  return out;
}

function helpText_() {
  return [
    'COMMANDS & EXAMPLES',
    'Add record: <note> <amount> [code]',
    '  Ex 1: food 50 (auto saved to Needs)',
    '  Ex 2: salary 5000 income (saved to Income)',
    'Report: report [month/year]',
    '  Ex: report 9/2026',
    'List types: types',
    'Undo: undo',
    '  Deletes the last added record.',
    '',
    'NOTES',
    '- Amount unit is thousands (50 = 50k, 1500 = 1.5m).',
    '- If no code is provided, the record defaults to Needs.',
    '- Manage types (add/edit) directly in the Types sheet.'
  ].join('\n');
}

function unknownTypeMsg_(token, types) {
  return 'Type "' + token + '" not found.\nAvailable codes: ' + Object.keys(types).join(', ') + '\nType "help" for instructions.';
}

function addExpense_(parts) {
  const n = parts.length;
  if (n < 2) return 'Invalid format. Ex: food 10. Type "help" for instructions.';

  const types = loadTypes_();
  const last = parts[n - 1];
  const prev = parseAmount_(parts[n - 2]);
  let t = null, amount = null, noteParts = [], hint = '';

  if (n >= 3 && prev !== null) {
    t = resolveType_(last, types);
    if (t) { amount = prev; noteParts = parts.slice(0, -2); }
  }

  if (!t) {
    amount = parseAmount_(last);
    if (amount === null) {
      if (prev !== null) {
        if (resolveType_(last, types)) return 'Missing note. Ex: food ' + parts[0] + ' ' + last + '\nType "help" for instructions.';
        return unknownTypeMsg_(last, types);
      }
      return 'Invalid amount. Format: <note> <amount> [code]\nEx: coffee 35\nType "help" for instructions.';
    }
    noteParts = parts.slice(0, -1);
    t = defaultType_(types);
    if (!t) return 'Missing DEFAULT_TYPE in script.';
    const stray = n >= 3 ? types[clean_(parts[n - 2])] : null;
    if (stray) {
      hint = '\nTip: Want to save as "' + stray.name + '"? Put code at the end: ' + parts.slice(0, -2).join(' ') + ' ' + last + ' ' + stray.code;
    }
  }

  const note = noteParts.join(' ');
  sheet_('Expenses').appendRow([new Date(), note, t.code, amount]);

  const now = new Date();
  const sums = sumByType_(new Date(now.getFullYear(), now.getMonth(), 1),
                          new Date(now.getFullYear(), now.getMonth() + 1, 1), types);
  
  const targetTxt = t.target > 0 ? ' / ' + money_(t.target) : '';
  const currentSum = sums[t.code] || 0;
  
  return 'note: ' + note + '\n' +
         'total: ' + money_(amount) + '\n' +
         'type: ' + t.name + ' - ' + money_(currentSum) + targetTxt + hint;
}

function undoExpense_() {
  const sh = sheet_('Expenses');
  const n = sh.getLastRow();
  if (n > 1) {
    const row = sh.getRange(n, 1, 1, 4).getValues()[0];
    sh.deleteRow(n);
    return 'Deleted last record:\n' + row[1] + ' - ' + money_(row[3]) + ' (' + row[2] + ')';
  }
  return 'No records to delete.';
}

function parsePeriod_(arg) {
  const now = new Date();
  let y = now.getFullYear(), m = now.getMonth() + 1;
  if (arg) {
    if (arg.indexOf('/') > -1) {
      const p = arg.split('/');
      m = parseInt(p[0], 10);
      y = parseInt(p[1], 10);
    } else {
      m = parseInt(arg, 10);
    }
  }
  if (isNaN(m) || isNaN(y) || m < 1 || m > 12 || y < 2000 || y > 2100) return null;
  const monthNames = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  return { from: new Date(y, m - 1, 1), to: new Date(y, m, 1), label: monthNames[m - 1] + ' ' + y };
}

function pct_(s, t) {
  if (!(t > 0)) return '';
  return ' (' + Math.round(s / t * 100) + '%' + (s > t ? ', OVER' : '') + ')';
}

function report_(arg) {
  const p = parsePeriod_(arg);
  if (!p) return 'Invalid period. Ex: report 9/2026\nType "help" for instructions.';

  const types = loadTypes_();
  const sums = sumByType_(p.from, p.to, types);
  const lines = ['REPORT ' + p.label.toUpperCase()];
  
  let totalExp = 0, totalInc = 0;
  const expLines = [], incLines = [];

  Object.keys(types).forEach(function (code) {
    const t = types[code];
    const s = sums[code] || 0;
    if (t.group === 'income') {
      totalInc += s;
      if (s > 0 || t.target > 0) {
         const tgt = t.target > 0 ? ' / ' + money_(t.target) + pct_(s, t.target) : '';
         incLines.push('+ ' + t.name + ': ' + money_(s) + tgt);
      }
    } else {
      totalExp += s;
      if (s > 0 || t.target > 0) {
         const tgt = t.target > 0 ? ' / ' + money_(t.target) + pct_(s, t.target) : '';
         expLines.push('- ' + t.name + ': ' + money_(s) + tgt);
      }
    }
  });

  Object.keys(sums).forEach(function (key) {
    if (key.charAt(0) === '?') {
      totalExp += sums[key];
      expLines.push('- (Unknown) ' + (key.slice(1) || 'Empty') + ': ' + money_(sums[key]));
    }
  });

  if (expLines.length > 0) {
    lines.push.apply(lines, expLines);
    if (expLines.length > 1) {
      lines.push('= Total Expenses: ' + money_(totalExp));
    }
  }

  if (incLines.length > 0) {
    lines.push.apply(lines, incLines);
    if (incLines.length > 1) {
      lines.push('= Total Income: ' + money_(totalInc));
    }
  }
  
  const net = totalInc - totalExp;
  lines.push('Total: ' + (net >= 0 ? '+' : '') + money_(net));
  
  return lines.join('\n');
}

function listTypes_() {
  const types = loadTypes_();
  return 'TYPES:\n' + Object.keys(types).map(function (c) {
    return c + ': ' + types[c].name;
  }).join('\n');
}

function buildGuide_(sh) {
  const rows = [
    ['section', 'command', 'description', 'example'],
    ['Log expense', '<note> <amount>',
      'Quickest way. The last word is the amount, everything before is the note. Auto-saved to DEFAULT_TYPE.',
      'food 10'],
    ['Log expense', '<note> <amount> <type>',
      'Save to a specific type: add type code at the end. Code is only parsed if there are 3+ words and the previous word is a number.',
      'coffee 35 fun'],
    ['Amount', 'unit: k',
      'Amounts are in thousands (remove 3 zeros). 20 = 20,000. 2000 = 2,000,000.',
      'rent 2000'],
    ['Amount', 'display',
      'Bot displays k (thousands), m (millions), or b (billions).',
      '20 -> 20k | 2000 -> 2m'],
    ['Report', 'report', 'Current month report.', 'report'],
    ['Report', 'report <month>', 'Report for a specific month this year (1-12).', 'report 9'],
    ['Report', 'report <month>/<year>', 'Report for a specific month and year.', 'report 9/2026'],
    ['Types', 'types', 'View all types: code, name, target, group.', 'types'],
    ['Undo', 'undo', 'Deletes the last record in the Expenses sheet.', 'undo'],
    ['Help', 'help', 'Show usage guide in Zalo.', 'help'],
    ['Sheet Types', 'code, name, monthly_target, group',
      'Manage types here. Add rows to create types. code: used in messages. name: display name. group: "Income" or "Expense".',
      'fun | Fun | 1000 | Expense'],
    ['Setup', 'setupWebhook', 'Run in Apps Script after deploying to apply changes.', '']
  ];
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  [150, 270, 560, 290].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.getRange(1, 1, rows.length, 4).setWrap(true).setVerticalAlignment('top');
}

function syncScriptTypes_() {
  const sh = sheet_('Types');
  const seen = dict_();
  const result = { added: 0 };
  [DEFAULT_TYPE].concat(TYPES).forEach(function (def) {
    const code = clean_(def && def[0]);
    if (!code || seen[code]) return;
    seen[code] = true;
    const groupStr = String(def[3] || 'expense').trim().toLowerCase() === 'income' ? 'Income' : 'Expense';
    const values = [code, String(def[1] || '').trim() || code, num_(def[2]), groupStr];
    let row = findTypeRow_(code);
    if (row < 0) {
      sh.appendRow(values);
      row = sh.getLastRow();
      result.added++;
      sh.getRange(row, 1).setNote('Script declared (add-only). Sheet data takes priority.');
    }
  });
  return result;
}

function ensureSheets_() {
  const ss = book_();
  const own = ['Expenses', 'Types', 'Guide'];
  const result = { created: [], removed: [] };

  own.forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      if (name === 'Guide') buildGuide_(ss.insertSheet('Guide'));
      else sheet_(name);
      result.created.push(name);
    } else if (name === 'Types') {
      const h = sh.getRange(1, 4).getValue();
      if (!h) sh.getRange(1, 4).setValue('group').setFontWeight('bold');
    }
  });

  const legacy = ss.getSheetByName('Log');
  if (legacy) {
    result.removed.push('Log');
    ss.deleteSheet(legacy);
  }
  ss.getSheets().forEach(function (sh) {
    const name = sh.getName();
    if (own.indexOf(name) > -1) return;
    if (sh.getLastRow() === 0 && sh.getLastColumn() === 0) {
      result.removed.push(name);
      ss.deleteSheet(sh);
    }
  });
  return result;
}

function probeUrl_(url) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const text = res.getContentText();
    if (text.indexOf('Bot is running') > -1) return 'ok';
    if (/accounts\.google\.com|ServiceLogin/i.test(text)) return 'login';
    if (res.getResponseCode() === 404) return 'notfound';
    return 'other';
  } catch (err) {
    return 'other';
  }
}

function findWebAppUrl_() {
  const list = [];
  const add = function (u) {
    if (u && /\/exec$/.test(u) && list.indexOf(u) < 0) list.push(u);
  };
  add(ScriptApp.getService().getUrl());
  add(prop_('WEBAPP_URL'));

  let problem = list.length ? 'other' : 'nolink';
  for (let i = 0; i < list.length; i++) {
    const r = probeUrl_(list[i]);
    if (r === 'ok') {
      PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', list[i]);
      return { url: list[i], problem: '' };
    }
    if (r === 'login' || problem === 'other') problem = r;
  }
  return { url: '', problem: problem };
}

function webAppProblemMsg_(problem) {
  switch (problem) {
    case 'login':
      return 'Web app URL requires Google login. Go to Deploy > Manage deployments > edit > Who has access = Anyone > New version > Deploy, then run setupWebhook again.';
    case 'notfound':
      return 'Web app URL not found (404). Copy the /exec link from Deploy, open it in browser once, then run setupWebhook again.';
    case 'nolink':
      return 'Web app URL not found. Deploy > New deployment > Web app, copy /exec link, open in browser once, then run setupWebhook again.';
    default:
      return 'Web app URL does not run new code. Deploy > Manage deployments > edit > New version > Deploy, then run setupWebhook again.';
  }
}

function setupWebhook() {
  const dcode = clean_(Array.isArray(DEFAULT_TYPE) ? DEFAULT_TYPE[0] : '');
  if (!/^[a-z0-9_]+$/.test(dcode)) {
    Logger.log('MISSING DEFAULT_TYPE: Script requires a default type. Example: const DEFAULT_TYPE = ["needs", "Needs", 0, "expense"];');
    return;
  }
  if (!String(prop_('BOT_TOKEN') || '').trim()) {
    Logger.log('MISSING BOT_TOKEN: Add BOT_TOKEN in Project Settings > Script Properties, then run again.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  if (String(prop_('URL_KEY') || '').length < 16) {
    props.setProperty('URL_KEY', Utilities.getUuid().replace(/-/g, ''));
    Logger.log('Generated random URL_KEY.');
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) props.setProperty('SHEET_ID', active.getId());
  if (!prop_('SHEET_ID')) {
    Logger.log('Script is not bound to a Google Sheet. Open Apps Script from Extensions menu in a Sheet.');
    return;
  }

  const sheets = ensureSheets_();
  Logger.log(sheets.created.length ? 'Created sheets: ' + sheets.created.join(', ') : 'Sheets are ready.');
  if (sheets.removed.length) Logger.log('Removed old/empty sheets: ' + sheets.removed.join(', '));

  const synced = syncScriptTypes_();
  Logger.log('Script types synced to Types sheet: added ' + synced.added + ' types.');

  const found = findWebAppUrl_();
  if (!found.url) {
    Logger.log('FAILED TO SET WEBHOOK. ' + webAppProblemMsg_(found.problem));
    return;
  }
  Logger.log('Using Web app URL: ' + found.url);

  const url = found.url + '?key=' + prop_('URL_KEY');
  const r = callZaloRaw_('setWebhook', { url: url, secret_token: prop_('URL_KEY') });
  Logger.log(r.code + ' ' + r.text);

  let data = {};
  try { data = JSON.parse(r.text); } catch (e) {}

  if (!(r.code === 200 && data.ok === true)) {
    Logger.log('FAILED: Zalo rejected webhook. Check your BOT_TOKEN.');
    return;
  }
  const v = data.result && data.result.verification;
  if (v && v.ok !== true) {
    Logger.log('PENDING: Zalo verification failed. Check "Who has access = Anyone", Deploy New version, then run this again.');
    return;
  }

  const ids = getChatIds_();
  if (!ids.length) {
    Logger.log('SUCCESS. Send "start" to the bot in Zalo to receive the guide.');
    return;
  }
  ids.forEach(function (id) {
    sendText_(id, 'Setup successful! Bot is ready.\n\n' + helpText_());
  });
  Logger.log('SUCCESS. Guide sent to ' + ids.length + ' chats.');
}