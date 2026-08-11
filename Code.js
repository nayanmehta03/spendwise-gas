// ============================================================
// SPENDWISE — Code.gs  v1.2.0
// Runtime backend only. All setup/admin logic lives in AdminOps.gs.
//
// FILE MAP:
//   Code.gs             — this file: all server-side functions
//   AdminOps.gs         — SETUP, REPAIR, STATUS + import/diagnostic tools
//   ChatHandler.gs      — Google Chat App: /spend, /summary, daily summary
//   EmailIngestion.gs   — Gmail receipt import pipeline + parsers
//   index.html          — SPA shell, router, nav, FAB
//   shared-styles.html  — all CSS
//   shared-nav.html     — overlay nav, toast, edit modal, shared JS utilities
//   page-add.html       — Add Expense page
//   page-dashboard.html — Dashboard page
//   page-expenses.html  — Expenses list page
//   page-analytics.html — Analytics page
//   page-income.html    — Income page
//   page-standing.html  — Recurring / Standing Instructions page
//   page-settings.html  — Settings page
//
// SETUP: run SETUP() from AdminOps.gs — creates all sheets, stores IDs in
// Script Properties. No IDs are hardcoded anywhere in this file.
// ============================================================

const SPENDWISE_VERSION = '1.2.0';

const SHEET_NAME = 'Expenses';
const CATEGORIES_TAB = 'Categories';
const SHARD_REGISTRY = 'ShardRegistry';
const SETTINGS_TAB = 'Settings';
const INCOME_TAB = 'Income';     // stored in Config sheet; not sharded
const SI_TAB = 'StandingInstructions'; // stored in Config sheet; not sharded
const KEYWORD_MAP_TAB = 'KeywordMap';       // keyword → category mapping
const EMAIL_SOURCES_TAB = 'EmailSources';   // email sender/subject rules
const LOGS_TAB = 'Logs';                   // import/chat audit trail

const TTL_CATEGORIES = 3600;
const TTL_ANALYTICS = 300;
const TTL_SHARD_REG = 7200;
const TTL_EXPENSES = 300;
const TTL_SETTINGS = 3600;
const TTL_SI = 3600; // standing instructions — changes infrequently
const TTL_KEYWORD_MAP = 3600; // keyword mappings — changes infrequently

const SCRIPT_CACHE = CacheService.getScriptCache();

const DEFAULT_CATEGORIES = [
  { name: 'Food & Drink', icon: 'fork-knife', budget: 8000 },
  { name: 'Groceries', icon: 'shopping-cart', budget: 6000 },
  { name: 'Transport', icon: 'car', budget: 4000 },
  { name: 'Shopping', icon: 'tote', budget: 5000 },
  { name: 'Health & Wellbeing', icon: 'heartbeat', budget: 3000 },
  { name: 'Entertainment', icon: 'film-strip', budget: 2000 },
  { name: 'Subscriptions', icon: 'device-mobile', budget: 1500 },
  { name: 'Bills', icon: 'lightbulb', budget: 3000 },
  { name: 'Rent', icon: 'house', budget: 20000 },
  { name: 'Travel', icon: 'airplane', budget: 10000 },
  { name: 'Gifts', icon: 'gift', budget: 2000 },
  { name: 'Investment', icon: 'trend-up', budget: 10000 },
  { name: 'Business', icon: 'briefcase', budget: 5000 },
  { name: 'Other', icon: 'package', budget: 2000 },
];

const DEFAULT_SETTINGS = {
  currency: '₹',
  currencyCode: 'INR',
  defaultPayment: 'UPI',
  weekStartDay: 'Monday',
  weeklyReportEnabled: 'false',
  weeklyReportDay: 'Monday',
  weeklyReportTime: '8',
  weeklyReportEmail: '',
  // Chat integration (opt-in)
  chatEnabled: 'false',
  chatSpaceId: '',
  // Daily summary (opt-in)
  dailySummaryEnabled: 'false',
  dailySummaryTime: '21',
  // Email ingestion (opt-in)
  emailIngestionEnabled: 'false',
  emailIngestionIntervalMinutes: '15',
  // Recurring reminder email (opt-in) — sent N days before each due date
  recurringReminderEnabled: 'false',
  recurringReminderTime: '9',
  recurringReminderEmail: '',
  recurringReminderDaysBefore: '1',
};

// How many days back processStandingInstructions() will look for occurrences
// it missed (trigger failure, auth lapse). Only occurrences AFTER the SI's
// LastLoggedDate are caught up, so this can never re-log history.
const SI_CATCHUP_DAYS = 3;

const _ssCache = {};
function _openSS(id) {
  if (!_ssCache[id]) _ssCache[id] = SpreadsheetApp.openById(id);
  return _ssCache[id];
}

// Reads CONFIG_SHEET_ID from Script Properties (set by SETUP() in AdminOps.gs)
function getConfigSS() {
  const id = PropertiesService.getScriptProperties().getProperty('CONFIG_SHEET_ID');
  if (!id) throw new Error('Spendwise is not set up. Run SETUP() from AdminOps.gs first.');
  return _openSS(id);
}

function getActiveShardId() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('ACTIVE_SHARD_ID');
  if (stored) return stored;
  // Fallback: derive from registry — newest month = active shard
  try {
    const records = _getAllShardRecords();
    if (records && records.length > 0) {
      const newest = records[0].id;
      props.setProperty('ACTIVE_SHARD_ID', newest);
      return newest;
    }
  } catch (e) { }
  throw new Error('No active shard found. Run SETUP() or RESET() from AdminOps.gs.');
}

function getActiveShardSS() { return _openSS(getActiveShardId()); }

// ── SPA entry point ──────────────────────────────────────────
function doGet(e) {
  const activeEmail = Session.getActiveUser().getEmail();
  const ownerEmail = Session.getEffectiveUser().getEmail();
  if (!activeEmail || activeEmail !== ownerEmail) {
    return HtmlService.createHtmlOutput('<h1>🔒 Unauthorized</h1><p>You do not have permission to access this application.</p>');
  }

  // Remember our own URL so reminder emails can link back to us
  _rememberWebAppUrl();

  // One-click actions from reminder emails — handled before page routing
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action === 'logSI' || action === 'undoSI') return _handleSIEmailAction(action, e.parameter);

  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'add';
  const validPages = ['add', 'dashboard', 'expenses', 'analytics', 'settings', 'income', 'standing'];
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.initPage = validPages.includes(page) ? page : 'add';
  return tmpl.evaluate()
    .setTitle('Spendwise')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setFaviconUrl('https://raw.githubusercontent.com/nayanmehta03/spendwise-gas/refs/heads/main/screenshots/svgviewer-png-output.png');
}

// ── Webhook / HTTP endpoint entry point for Google Chat ──────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ text: 'Error: No payload' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const event = JSON.parse(e.postData.contents);
    
    // Security check: Only allow the owner of the script to interact with the bot
    const ownerEmail = Session.getEffectiveUser().getEmail();
    const callerEmail = event.user && event.user.email;
    if (!callerEmail || callerEmail !== ownerEmail) {
      return ContentService.createTextOutput(JSON.stringify({ text: '🔒 Unauthorized: This bot is private to its owner.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    let response;
    if (event.type === 'APP_COMMAND') {
      response = onAppCommand(event);
    } else {
      response = onMessage(event);
    }
    
    return ContentService.createTextOutput(JSON.stringify(response || {}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ text: '❌ Error: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Initialization ─────────────────────────────────────────────
// Called by AdminOps.gs SETUP() only. Not called at runtime.
// Idempotent — safe to call multiple times.
function initializeSheets() {
  const props = PropertiesService.getScriptProperties();
  const configId = props.getProperty('CONFIG_SHEET_ID');
  const activeShardId = props.getProperty('ACTIVE_SHARD_ID');
  if (!configId) throw new Error('CONFIG_SHEET_ID not set. Run SETUP() first.');
  if (!activeShardId) throw new Error('ACTIVE_SHARD_ID not set. Run SETUP() first.');
  _initConfigSheet(activeShardId);
  _ensureShardSheet(_openSS(activeShardId));
  return { success: true, message: 'All sheets initialized.' };
}

function _initConfigSheet(firstShardId) {
  const ss = getConfigSS();
  const month = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  if (!ss.getSheetByName(CATEGORIES_TAB)) {
    const sheet = ss.insertSheet(CATEGORIES_TAB);
    const rows = [['Category', 'Icon', 'Budget'], ...DEFAULT_CATEGORIES.map(c => [c.name, c.icon, c.budget])];
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    sheet.getRange(1, 1, 1, 3).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SHARD_REGISTRY)) {
    const sheet = ss.insertSheet(SHARD_REGISTRY);
    sheet.getRange(1, 1, 1, 4).setValues([['ShardID', 'Month', 'IsActive', 'Label']]);
    if (firstShardId) sheet.getRange(2, 1, 1, 4).setValues([[firstShardId, month, true, 'Shard 01 — ' + month]]);
    sheet.getRange(1, 1, 1, 4).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SETTINGS_TAB)) {
    const sheet = ss.insertSheet(SETTINGS_TAB);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    Object.entries(DEFAULT_SETTINGS).forEach(([k, v], i) => sheet.getRange(i + 2, 1, 1, 2).setValues([[k, v]]));
    sheet.getRange(1, 1, 1, 2).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    const sheet = ss.getSheetByName(SETTINGS_TAB);
    if (sheet.getLastRow() > 1) {
      const existingKeys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
      const missingEntries = Object.entries(DEFAULT_SETTINGS).filter(([k, v]) => !existingKeys.includes(k));
      if (missingEntries.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, missingEntries.length, 2).setValues(missingEntries);
      }
    }
  }
  if (!ss.getSheetByName(INCOME_TAB)) {
    const sheet = ss.insertSheet(INCOME_TAB);
    const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Notes', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // _getSISheet() owns the schema and migrates existing sheets to it
  _getSISheet();
  // ── KeywordMap tab ───────────────────────────────────────────
  if (!ss.getSheetByName(KEYWORD_MAP_TAB)) {
    const sheet = ss.insertSheet(KEYWORD_MAP_TAB);
    const headers = ['Keyword', 'Category'];
    const defaultMappings = [
      ['lunch', 'Food & Drink'], ['dinner', 'Food & Drink'], ['breakfast', 'Food & Drink'],
      ['tea', 'Food & Drink'], ['coffee', 'Food & Drink'], ['snack', 'Food & Drink'],
      ['uber', 'Transport'], ['ola', 'Transport'], ['fuel', 'Transport'],
      ['metro', 'Transport'], ['auto', 'Transport'],
      ['grocery', 'Groceries'], ['dmart', 'Groceries'], ['bigbasket', 'Groceries'],
      ['movie', 'Entertainment'], ['netflix', 'Subscriptions'], ['spotify', 'Subscriptions'],
      ['rent', 'Rent'], ['electricity', 'Bills'], ['wifi', 'Bills'], ['mobile', 'Bills'],
      ['medicine', 'Health & Wellbeing'], ['doctor', 'Health & Wellbeing'],
      ['gym', 'Health & Wellbeing'], ['amazon', 'Shopping'], ['flipkart', 'Shopping'],
    ];
    const rows = [headers, ...defaultMappings];
    sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    sheet.getRange(1, 1, 1, 2).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // ── EmailSources tab ────────────────────────────────────────
  if (!ss.getSheetByName(EMAIL_SOURCES_TAB)) {
    const sheet = ss.insertSheet(EMAIL_SOURCES_TAB);
    const headers = ['Sender', 'SubjectPattern', 'ParserType', 'Enabled'];
    const exampleRows = [
      ['noreply@swiggy.in', 'Order Delivered', 'swiggy', true],
      ['noreply@zomato.com', 'Order Summary', 'zomato', true],
      ['auto-confirm@amazon.in', 'Your order', 'amazon', true],
    ];
    const rows = [headers, ...exampleRows];
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);
    sheet.getRange(1, 1, 1, 4).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // ── Logs tab ────────────────────────────────────────────────
  if (!ss.getSheetByName(LOGS_TAB)) {
    const sheet = ss.insertSheet(LOGS_TAB);
    const headers = ['Timestamp', 'Module', 'Status', 'Message', 'ExternalId'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function _ensureShardSheet(ss) {
  if (!ss.getSheetByName(SHEET_NAME)) {
    const sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'PaymentMethod', 'Notes', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return ss.getSheetByName(SHEET_NAME);
}

// ── Shard rotation ───────────────────────────────────────────
function rotateShardForNewMonth() {
  const tz = Session.getScriptTimeZone();
  const month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const existing = _getAllShardRecords();
  if (existing.find(s => s.month === month)) return { skipped: true };
  const newSS = SpreadsheetApp.create('Expenses_Shard_' + month);
  const newId = newSS.getId();
  _ensureShardSheet(newSS);
  getConfigSS().getSheetByName(SHARD_REGISTRY).appendRow([newId, month, true, 'Shard ' + (existing.length + 2)]);
  PropertiesService.getScriptProperties().setProperty('ACTIVE_SHARD_ID', newId);
  SCRIPT_CACHE.remove('shard_registry');
  return { success: true, shardId: newId, month, url: newSS.getUrl() };
}

// ── Shard registry ───────────────────────────────────────────
function _getAllShardRecords() {
  const cached = SCRIPT_CACHE.get('shard_registry');
  if (cached) return JSON.parse(cached);
  const sheet = getConfigSS().getSheetByName(SHARD_REGISTRY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone();
  const records = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(r => r[0]).map(r => ({
      id: String(r[0]),
      month: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM') : String(r[1] || ''),
      active: !!r[2],
      label: String(r[3] || '')
    }))
    .sort((a, b) => (b.month || '').localeCompare(a.month || '')); // newest first
  SCRIPT_CACHE.put('shard_registry', JSON.stringify(records), TTL_SHARD_REG);
  return records;
}

function _getShardsForRange(startDate, endDate) {
  if (!startDate && !endDate) return [getActiveShardId()];
  const startMonth = startDate ? startDate.substring(0, 7) : null;
  const endMonth = endDate ? endDate.substring(0, 7) : null;
  const ids = _getAllShardRecords()
    .filter(s => {
      if (!s.month) return true;
      if (startMonth && s.month < startMonth) return false;
      if (endMonth && s.month > endMonth) return false;
      return true;
    }).map(s => s.id);
  const activeId = getActiveShardId();
  if (!ids.includes(activeId)) ids.push(activeId);
  return [...new Set(ids)];
}

function getShardInfo() {
  const records = _getAllShardRecords();
  return { shards: records, activeShardId: getActiveShardId(), totalShards: records.length };
}

// ── Shard batch read ─────────────────────────────────────────
function _readShardExpenses(shardId) {
  const tz = Session.getScriptTimeZone();
  try {
    const sheet = _openSS(shardId).getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
      .filter(r => r[0] !== '')
      .map(r => ({
        id: String(r[0]),
        date: r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') : '',
        category: String(r[2] || ''),
        description: String(r[3] || ''),
        amount: parseFloat(r[4]) || 0,
        paymentMethod: String(r[5] || 'Cash'),
        notes: String(r[6] || '')
      }));
  } catch (err) { Logger.log('Read error ' + shardId + ': ' + err.message); return []; }
}

// ── Categories ───────────────────────────────────────────────
function getCategories() {
  const cached = SCRIPT_CACHE.get('categories');
  if (cached) { try { const p = JSON.parse(cached); if (Array.isArray(p)) return p; } catch (e) { } }
  try {
    const sheet = getConfigSS().getSheetByName(CATEGORIES_TAB);
    if (!sheet || sheet.getLastRow() < 2) return DEFAULT_CATEGORIES;

    // Mapping from old emojis / old FA names to new Phosphor names
    const legacyIconMap = {
      '🍔': 'fork-knife', 'utensils': 'fork-knife',
      '🛒': 'shopping-cart', 'cart-shopping': 'shopping-cart',
      '🚗': 'car', '🛍️': 'tote', 'bag-shopping': 'tote', '🛍': 'tote',
      '💊': 'heartbeat', 'heart-pulse': 'heartbeat',
      '🎬': 'film-strip', 'film': 'film-strip',
      '📱': 'device-mobile', 'mobile-screen': 'device-mobile',
      '💡': 'lightbulb', '🏠': 'house',
      '✈️': 'airplane', 'plane': 'airplane', '✈': 'airplane',
      '🎁': 'gift', '📈': 'trend-up', 'chart-line': 'trend-up',
      '💼': 'briefcase', '📦': 'package', 'box-open': 'package',
      '💰': 'coins', '🎯': 'target', 'bullseye': 'target'
    };

    const cats = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
      .filter(r => r[0]).map(r => {
        let rawIcon = String(r[1]).trim();
        let mappedIcon = legacyIconMap[rawIcon] || rawIcon || 'package';
        return { name: String(r[0]), icon: mappedIcon, budget: parseFloat(r[2]) || 0 };
      });

    SCRIPT_CACHE.put('categories', JSON.stringify(cats), TTL_CATEGORIES);
    return cats;
  } catch (err) { Logger.log('Cat fallback: ' + err.message); return DEFAULT_CATEGORIES; }
}

// Replaces all category rows. Invalidates categories + analytics cache.
function saveCategories(categories) {
  const sheet = getConfigSS().getSheetByName(CATEGORIES_TAB);
  if (!sheet) return { success: false, message: 'Sheet not found.' };
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (categories && categories.length > 0) {
    sheet.getRange(2, 1, categories.length, 3).setValues(
      categories.map(c => [c.name || '', c.icon || 'package', parseFloat(c.budget) || 0])
    );
  }
  invalidateCache(); // clears categories + all analytics
  return { success: true, message: 'Categories saved.' };
}

// ── Settings ─────────────────────────────────────────────────
function getSettings() {
  const cached = SCRIPT_CACHE.get('settings');
  if (cached) return JSON.parse(cached);
  try {
    const sheet = getConfigSS().getSheetByName(SETTINGS_TAB);
    if (!sheet || sheet.getLastRow() < 2) return DEFAULT_SETTINGS;
    const settings = { ...DEFAULT_SETTINGS };
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
      .filter(r => r[0]).forEach(r => { settings[String(r[0])] = String(r[1] || ''); });
    SCRIPT_CACHE.put('settings', JSON.stringify(settings), TTL_SETTINGS);
    return settings;
  } catch (err) { return DEFAULT_SETTINGS; }
}

function saveSettings(settings) {
  // Always merge with existing settings — never wipe keys not included in this call
  const existing = getSettings();
  const merged = { ...existing, ...settings };

  let sheet = getConfigSS().getSheetByName(SETTINGS_TAB);
  if (!sheet) {
    sheet = getConfigSS().insertSheet(SETTINGS_TAB);
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  }
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  const rows = Object.entries(merged);
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  SCRIPT_CACHE.remove('settings');

  // The daily Chat summary has no Settings UI, so its keys are edited directly
  // in the sheet — reconcile the trigger whenever they change, or the new time
  // silently never takes effect. Runs after the cache clear so it reads fresh.
  const touchesDailySummary = settings && (
    Object.prototype.hasOwnProperty.call(settings, 'dailySummaryEnabled') ||
    Object.prototype.hasOwnProperty.call(settings, 'dailySummaryTime'));
  if (touchesDailySummary) {
    try { reconcileDailySummaryTrigger(); }
    catch (e) { Logger.log('reconcileDailySummaryTrigger error: ' + e.message); }
  }

  return { success: true };
}

// ── Expenses ─────────────────────────────────────────────────
function getExpenses(filters) {
  try {
    filters = filters || {};
    const cacheKey = 'exp_' + JSON.stringify(filters);
    try {
      const cached = SCRIPT_CACHE.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) { }
    const shardIds = _getShardsForRange(filters.startDate || null, filters.endDate || null);
    let expenses = [];
    shardIds.forEach(id => {
      try { expenses = expenses.concat(_readShardExpenses(id)); } catch (e) { }
    });
    if (filters.startDate) expenses = expenses.filter(e => e.date >= filters.startDate);
    if (filters.endDate) expenses = expenses.filter(e => e.date <= filters.endDate);
    if (filters.category && filters.category !== 'All') expenses = expenses.filter(e => e.category === filters.category);
    if (filters.minAmount) expenses = expenses.filter(e => e.amount >= parseFloat(filters.minAmount));
    if (filters.maxAmount) expenses = expenses.filter(e => e.amount <= parseFloat(filters.maxAmount));
    expenses.sort((a, b) => new Date(b.date) - new Date(a.date));
    try { SCRIPT_CACHE.put(cacheKey, JSON.stringify(expenses), TTL_EXPENSES); } catch (e) { }
    return expenses;
  } catch (e) {
    Logger.log('getExpenses error: ' + e.message);
    return [];
  }
}

// ── Analytics ────────────────────────────────────────────────

// Returns a Date set to the start of the given named period.
// Used by getAnalytics() and getMonthlyComparison() to avoid duplication.
function _periodStartDate(period, now) {
  const d = new Date(now);
  // Trailing windows are filtered inclusively at both ends, so subtracting the
  // full N produced N+1 days — 'week' was really 8 days and 'month' 31.
  if (period === 'week') d.setDate(now.getDate() - 6);
  else if (period === 'month') d.setDate(now.getDate() - 29);
  else if (period === 'current_month') { d.setDate(1); d.setHours(0, 0, 0, 0); }
  else if (period === 'quarter') d.setMonth(now.getMonth() - 3);
  else if (period === 'half') d.setMonth(now.getMonth() - 6);
  else if (period === 'current_year') { d.setMonth(0); d.setDate(1); d.setHours(0, 0, 0, 0); }
  else if (period === 'year') d.setFullYear(now.getFullYear() - 1);
  else d.setFullYear(now.getFullYear() - 1); // default: trailing year
  return d;
}

function _emptyAnalytics() {
  return {
    totalSpent: 0, totalTransactions: 0, avgPerDay: 0, avgPerTransaction: 0,
    topCategories: [], trends: [], shardCount: 0, periodDays: 0
  };
}

// Days elapsed in a period, counting both ends. avgPerDay used to divide by a
// hardcoded 30 whatever the period, so "This Week" was shown as a 30-day
// average and "This Month" was only ever right on the 30th of the month.
function _elapsedDays(startDateStr, endDateStr) {
  const start = _parseLocalDate(startDateStr);
  const end = _parseLocalDate(endDateStr);
  return Math.max(1, Math.round((end - start) / 86400000) + 1);
}

// Builds the trend series for the period being viewed. This was hardcoded to
// the last 30 days, so "Last 6 Months" and "This Year" plotted an identical
// 30-bar chart. Longer periods are bucketed so the chart stays readable, and
// each point carries its own axis label.
function _buildTrends(byDay, startDateStr, endDateStr, tz) {
  const days = _elapsedDays(startDateStr, endDateStr);
  const bucketDays = days <= 31 ? 1 : days <= 130 ? 7 : 0; // 0 = calendar month
  const end = _parseLocalDate(endDateStr);
  const out = [];

  if (bucketDays === 0) {
    const cursor = _parseLocalDate(startDateStr);
    cursor.setDate(1);
    while (cursor <= end) {
      const key = Utilities.formatDate(cursor, tz, 'yyyy-MM');
      let amount = 0;
      Object.keys(byDay).forEach(d => { if (d.substring(0, 7) === key) amount += byDay[d]; });
      out.push({
        date: key + '-01',
        amount: Math.round(amount * 100) / 100,
        label: Utilities.formatDate(cursor, tz, 'MMM')
      });
      cursor.setMonth(cursor.getMonth() + 1);
    }
    return out;
  }

  const cursor = _parseLocalDate(startDateStr);
  while (cursor <= end) {
    const bucketStart = new Date(cursor);
    let amount = 0;
    // Inner loop always advances at least once, so the outer loop terminates
    for (let i = 0; i < bucketDays && cursor <= end; i++) {
      amount += byDay[Utilities.formatDate(cursor, tz, 'yyyy-MM-dd')] || 0;
      cursor.setDate(cursor.getDate() + 1);
    }
    out.push({
      date: Utilities.formatDate(bucketStart, tz, 'yyyy-MM-dd'),
      amount: Math.round(amount * 100) / 100,
      label: Utilities.formatDate(bucketStart, tz, 'd/M')
    });
  }
  return out;
}

function getAnalytics(period) {
  period = period || 'current_month';
  try {
    const cacheKey = 'analytics_' + period;
    try {
      const cached = SCRIPT_CACHE.get(cacheKey);
      if (cached) { const p = JSON.parse(cached); if (p && typeof p === 'object') return p; }
    } catch (e) { }

    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const startDateStr = Utilities.formatDate(_periodStartDate(period, now), tz, 'yyyy-MM-dd');
    const shardIds = _getShardsForRange(startDateStr, null);
    const byCategory = {}, byDay = {};
    let totalSpent = 0, totalTransactions = 0;

    shardIds.forEach(id => {
      try {
        _readShardExpenses(id).forEach(e => {
          if (!e.date || e.date < startDateStr) return;
          const cat = e.category || 'Other';
          totalSpent += e.amount; totalTransactions++;
          byCategory[cat] = (byCategory[cat] || 0) + e.amount;
          byDay[e.date] = (byDay[e.date] || 0) + e.amount;
        });
      } catch (shardErr) { Logger.log('Analytics shard error ' + id + ': ' + shardErr.message); }
    });

    const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const periodDays = _elapsedDays(startDateStr, todayStr);
    const trends = _buildTrends(byDay, startDateStr, todayStr, tz);

    const result = {
      totalSpent: Math.round(totalSpent * 100) / 100, totalTransactions, periodDays,
      avgPerDay: Math.round((totalSpent / periodDays) * 100) / 100,
      avgPerTransaction: totalTransactions > 0 ? Math.round((totalSpent / totalTransactions) * 100) / 100 : 0,
      topCategories: Object.entries(byCategory).sort((a, b) => b[1] - a[1])
        .map(([cat, amt]) => ({
          category: cat, amount: Math.round(amt * 100) / 100,
          percentage: totalSpent > 0 ? Math.round((amt / totalSpent) * 100) : 0
        })),
      trends, shardCount: shardIds.length
    };

    try { SCRIPT_CACHE.put(cacheKey, JSON.stringify(result), TTL_ANALYTICS); } catch (e) { }
    return result;
  } catch (e) {
    Logger.log('getAnalytics error: ' + e.message);
    return _emptyAnalytics();
  }
}

// Monthly comparison table: matrix[month][category] = amount.
// Called by analytics page for the multi-month comparison table.
function getMonthlyComparison(period) {
  period = period || 'year';
  const cacheKey = 'monthly_comparison_' + period;
  try {
    const cached = SCRIPT_CACHE.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) { }

  try {
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const startStr = Utilities.formatDate(_periodStartDate(period, now), tz, 'yyyy-MM-dd');
    const shardIds = _getShardsForRange(startStr, null);
    const categories = getCategories();
    const catNames = categories.map(c => c.name);
    const budgetMap = {};
    categories.forEach(c => { budgetMap[c.name] = c.budget || 0; });

    // matrix[month][category] = total amount
    const matrix = {};

    shardIds.forEach(id => {
      try {
        _readShardExpenses(id).forEach(e => {
          if (!e.date || e.date < startStr) return;
          const mKey = e.date.substring(0, 7);
          const cat = e.category || 'Other';
          if (!matrix[mKey]) matrix[mKey] = {};
          matrix[mKey][cat] = Math.round(((matrix[mKey][cat] || 0) + e.amount) * 100) / 100;
        });
      } catch (err) { Logger.log('Comparison shard error ' + id + ': ' + err.message); }
    });

    // All months in range, sorted ascending
    const months = Object.keys(matrix).sort();

    // Income by month from Income tab
    const incomeByMonth = {};
    try {
      const incSheet = _getIncomeSheet();
      if (incSheet && incSheet.getLastRow() > 1) {
        incSheet.getRange(2, 1, incSheet.getLastRow() - 1, 5).getValues()
          .filter(r => r[0])
          .forEach(r => {
            if (!r[1]) return;
            const mKey = Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM');
            if (mKey >= startStr.substring(0, 7)) {
              incomeByMonth[mKey] = Math.round(((incomeByMonth[mKey] || 0) + (parseFloat(r[4]) || 0)) * 100) / 100;
            }
          });
      }
    } catch (e) { Logger.log('Income fetch for comparison: ' + e.message); }

    // Compute totals and averages per category
    const catTotals = {}, catAverages = {};
    const monthCount = months.length || 1;
    catNames.forEach(cat => {
      const total = months.reduce((s, m) => s + ((matrix[m] && matrix[m][cat]) || 0), 0);
      catTotals[cat] = Math.round(total * 100) / 100;
      catAverages[cat] = Math.round((total / monthCount) * 100) / 100;
    });

    // Row totals
    const rowTotals = {};
    months.forEach(m => {
      rowTotals[m] = Math.round(Object.values(matrix[m] || {}).reduce((s, v) => s + v, 0) * 100) / 100;
    });
    const grandTotal = Math.round(Object.values(rowTotals).reduce((s, v) => s + v, 0) * 100) / 100;
    const grandAverage = Math.round((grandTotal / monthCount) * 100) / 100;

    const result = {
      months, categories: catNames, matrix, rowTotals,
      catTotals, catAverages, budgetMap, incomeByMonth,
      grandTotal, grandAverage, monthCount
    };

    try { SCRIPT_CACHE.put(cacheKey, JSON.stringify(result), TTL_ANALYTICS); } catch (e) { }
    return result;
  } catch (e) {
    Logger.log('getMonthlyComparison error: ' + e.message);
    return { months: [], categories: [], matrix: {}, rowTotals: {}, catTotals: {}, catAverages: {}, budgetMap: {}, incomeByMonth: {}, grandTotal: 0, grandAverage: 0, monthCount: 0 };
  }
}
function getBudgetSummary() {
  const categories = getCategories();
  const analytics = getAnalytics('current_month') || _emptyAnalytics();
  const spentMap = {};
  (analytics.topCategories || []).forEach(c => { spentMap[c.category] = c.amount; });
  return categories.map(cat => {
    const spent = spentMap[cat.name] || 0;
    return {
      category: cat.name, icon: cat.icon, budget: cat.budget, spent,
      remaining: cat.budget - spent,
      // NOT capped at 100 — a category at 3x its budget must be able to say so.
      // Consumers clamp separately where a bar width needs it.
      percentage: cat.budget > 0 ? Math.round((spent / cat.budget) * 100) : 0
    };
  });
}

// Lean analytics splits — avoids 50KB GAS serialization limit
function getAnalyticsSummary(period) {
  try {
    const a = getAnalytics(period || 'current_month') || _emptyAnalytics();
    return {
      totalSpent: a.totalSpent, totalTransactions: a.totalTransactions,
      avgPerDay: a.avgPerDay, avgPerTransaction: a.avgPerTransaction,
      periodDays: a.periodDays || 0, shardCount: a.shardCount || 1,
      topCategories: (a.topCategories || []).filter(c => c.amount > 0)
    };
  } catch (e) { Logger.log('getAnalyticsSummary: ' + e.message); return _emptyAnalytics(); }
}

function getAnalyticsTrends(period) {
  try {
    const a = getAnalytics(period || 'current_month') || _emptyAnalytics();
    return { trends: a.trends || [] };
  } catch (e) { Logger.log('getAnalyticsTrends: ' + e.message); return { trends: [] }; }
}
function getAddPageData() {
  try {
    return { categories: getCategories() || [], settings: getSettings() || DEFAULT_SETTINGS };
  } catch (e) {
    Logger.log('getAddPageData error: ' + e.message);
    return { categories: DEFAULT_CATEGORIES, settings: DEFAULT_SETTINGS };
  }
}

// Split into 3 lean calls to stay well under GAS 50KB serialization limit.
// Dashboard fires all three in parallel via separate google.script.run calls.

function getDashboardStats() {
  try {
    const a = getAnalytics('current_month') || _emptyAnalytics();
    // Return only the scalar summary — not byCategory/byDay maps
    return {
      totalSpent: a.totalSpent,
      totalTransactions: a.totalTransactions,
      avgPerDay: a.avgPerDay,
      avgPerTransaction: a.avgPerTransaction,
      periodDays: a.periodDays || 0,
      topCategory: (a.topCategories && a.topCategories[0]) ? a.topCategories[0] : null,
      lastUpdated: new Date().toISOString()
    };
  } catch (e) {
    Logger.log('getDashboardStats error: ' + e.message);
    return { totalSpent: 0, totalTransactions: 0, avgPerDay: 0, avgPerTransaction: 0, periodDays: 0, topCategory: null, lastUpdated: new Date().toISOString() };
  }
}

function getDashboardBudget() {
  try {
    // getBudgetSummary internally calls getAnalytics('current_month')
    return { budgetSummary: getBudgetSummary() || [], categories: getCategories() || [] };
  } catch (e) {
    Logger.log('getDashboardBudget error: ' + e.message);
    return { budgetSummary: [], categories: [] };
  }
}

function getDashboardRecent() {
  try {
    const recent = _readShardExpenses(getActiveShardId()) || [];
    recent.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { recentExpenses: recent.slice(0, 5) };
  } catch (e) {
    Logger.log('getDashboardRecent error: ' + e.message);
    return { recentExpenses: [] };
  }
}

function getSettingsPageData() {
  try {
    const settings = getSettings() || DEFAULT_SETTINGS;
    // Pre-populate report email with account email if not yet set
    if (!settings.weeklyReportEmail) {
      try { settings.weeklyReportEmail = Session.getActiveUser().getEmail(); } catch (e) { }
    }
    return { categories: getCategories() || [], settings, shardInfo: getShardInfo() || { shards: [], totalShards: 0 } };
  } catch (e) {
    Logger.log('getSettingsPageData error: ' + e.message);
    return { categories: DEFAULT_CATEGORIES, settings: DEFAULT_SETTINGS, shardInfo: { shards: [], totalShards: 0 } };
  }
}

// Parse a yyyy-MM-dd string as LOCAL midnight (not UTC midnight)
// Without this, new Date('2026-03-21') = UTC midnight = March 20 18:30 IST
function _parseLocalDate(dateStr) {
  if (!dateStr) return new Date();
  const parts = String(dateStr).split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]), 12, 0, 0);
    // Use noon (12:00) so DST edge cases never shift the date
  }
  return new Date(dateStr);
}
function addExpense(expense) {
  const sheet = _ensureShardSheet(getActiveShardSS());
  const id = 'EXP-' + new Date().getTime();
  sheet.appendRow([id, _parseLocalDate(expense.date), expense.category, expense.description,
    parseFloat(expense.amount), expense.paymentMethod || 'Cash', expense.notes || '', new Date()]);
  invalidateCache();
  return { success: true, id };
}

function updateExpense(id, expense) {
  const shardId = _findShardForExpense(id);
  if (!shardId) return { success: false, message: 'Not found.' };
  const sheet = _openSS(shardId).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 2, 1, 6).setValues([[_parseLocalDate(expense.date), expense.category,
      expense.description, parseFloat(expense.amount), expense.paymentMethod || 'Cash', expense.notes || '']]);
      invalidateCache();
      return { success: true };
    }
  }
  return { success: false, message: 'Row not found.' };
}

function deleteExpense(id) {
  const shardId = _findShardForExpense(id);
  if (!shardId) return { success: false, message: 'Not found.' };
  const sheet = _openSS(shardId).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); invalidateCache(); return { success: true }; }
  }
  return { success: false };
}

function _findShardForExpense(expId) {
  const activeId = getActiveShardId();
  const allIds = [activeId, ..._getAllShardRecords().map(s => s.id).filter(id => id !== activeId)];
  for (const shardId of allIds) {
    try {
      const sheet = _openSS(shardId).getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
      if (ids.includes(expId)) return shardId;
    } catch (err) { }
  }
  return null;
}

// ============================================================
// INCOME — stored in single Income tab in Config sheet
// No sharding needed — data volume is small
// ============================================================

const INCOME_CATEGORIES = ['Salary', 'Bonus', 'Miscellaneous'];

function _getIncomeSheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(INCOME_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(INCOME_TAB);
    const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Notes', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getIncome(filters) {
  filters = filters || {};
  const cacheKey = 'income_' + JSON.stringify(filters);
  const cached = SCRIPT_CACHE.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (e) { } }

  try {
    const tz = Session.getScriptTimeZone();
    const sheet = _getIncomeSheet();
    if (sheet.getLastRow() <= 1) return [];

    let rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
      .filter(r => r[0])
      .map(r => ({
        id: String(r[0]),
        date: r[1] ? Utilities.formatDate(new Date(r[1]), tz, 'yyyy-MM-dd') : '',
        category: String(r[2] || ''),
        description: String(r[3] || ''),
        amount: parseFloat(r[4]) || 0,
        notes: String(r[5] || ''),
        timestamp: r[6]
      }));

    if (filters.startDate) rows = rows.filter(r => r.date >= filters.startDate);
    if (filters.endDate) rows = rows.filter(r => r.date <= filters.endDate);
    if (filters.category && filters.category !== 'All') rows = rows.filter(r => r.category === filters.category);

    rows.sort((a, b) => new Date(b.date) - new Date(a.date));
    SCRIPT_CACHE.put(cacheKey, JSON.stringify(rows), TTL_EXPENSES);
    return rows;
  } catch (e) {
    Logger.log('getIncome error: ' + e.message);
    return [];
  }
}

// Clears the two income cache keys written by getIncome().
// Called after any income mutation (add/update/delete).
function _invalidateIncomeCache() {
  SCRIPT_CACHE.removeAll(['income_{}', 'income_{"startDate":null,"endDate":null}']);
}

function addIncome(income) {
  try {
    const sheet = _getIncomeSheet();
    const id = 'INC-' + new Date().getTime();
    sheet.appendRow([
      id,
      _parseLocalDate(income.date),
      income.category || 'Miscellaneous',
      income.description || '',
      parseFloat(income.amount),
      income.notes || '',
      new Date()
    ]);
    _invalidateIncomeCache();
    return { success: true, id };
  } catch (e) {
    Logger.log('addIncome error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateIncome(id, income) {
  try {
    const sheet = _getIncomeSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[
          _parseLocalDate(income.date),
          income.category || 'Miscellaneous',
          income.description || '',
          parseFloat(income.amount),
          income.notes || ''
        ]]);
        _invalidateIncomeCache();
        return { success: true };
      }
    }
    return { success: false, message: 'Income entry not found.' };
  } catch (e) {
    Logger.log('updateIncome error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function deleteIncome(id) {
  try {
    const sheet = _getIncomeSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.deleteRow(i + 1);
        _invalidateIncomeCache();
        return { success: true };
      }
    }
    return { success: false, message: 'Income entry not found.' };
  } catch (e) {
    Logger.log('deleteIncome error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getAllIncome() {
  // Return lean array only — no wrapper object, no timestamp field
  // Avoids GAS 50KB serialization limit
  try {
    const rows = getIncome({});
    return rows.map(r => ({
      id: r.id,
      date: r.date,
      category: r.category,
      description: r.description,
      amount: r.amount,
      notes: r.notes
    }));
  } catch (e) {
    Logger.log('getAllIncome error: ' + e.message);
    return [];
  }
}

// ── STANDING INSTRUCTIONS ────────────────────────────────────
// Recurring commitments stored in the StandingInstructions tab
// of the Config sheet. Auto-logged daily by processStandingInstructions().

const SI_COLUMNS = ['ID', 'Name', 'Category', 'Amount', 'Frequency', 'DayOfMonth', 'DayOfWeek',
  'StartDate', 'EndDate', 'PaymentMethod', 'Notes', 'AutoLog', 'IsActive', 'LastLoggedDate', 'CustomIntervalDays',
  'LoggedOccurrences', 'LastRemindedOccurrence'];

// Column numbers (1-based) for the fields written outside of the full-row writes.
const SI_COL_LAST_LOGGED = 14;
const SI_COL_CUSTOM_INTERVAL = 15;
const SI_COL_LOGGED_OCCURRENCES = 16;
const SI_COL_LAST_REMINDED = 17;

// Keeps the LoggedOccurrences cell bounded — a weekly SI would otherwise grow
// this by 52 dates a year. Older occurrences fall out of the reminder/catch-up
// window long before this cap is reached.
const SI_MAX_OCCURRENCE_HISTORY = 30;

function _getSISheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(SI_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SI_TAB);
    sheet.getRange(1, 1, 1, SI_COLUMNS.length).setValues([SI_COLUMNS]);
    sheet.getRange(1, 1, 1, SI_COLUMNS.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Migration — sheets created before v1.3 are missing trailing columns.
  // Append any header the current schema expects but the sheet doesn't have.
  const width = sheet.getLastColumn();
  if (width < SI_COLUMNS.length) {
    const missing = SI_COLUMNS.slice(width);
    sheet.getRange(1, width + 1, 1, missing.length).setValues([missing])
      .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
  }
  return sheet;
}

function _rowToSI(r) {
  const tz = Session.getScriptTimeZone();
  return {
    id: String(r[0] || ''),
    name: String(r[1] || ''),
    category: String(r[2] || ''),
    amount: parseFloat(r[3]) || 0,
    frequency: String(r[4] || 'monthly'),
    dayOfMonth: parseInt(r[5]) || 1,
    dayOfWeek: String(r[6] || 'Monday'),
    startDate: r[7] ? Utilities.formatDate(new Date(r[7]), tz, 'yyyy-MM-dd') : '',
    endDate: r[8] ? Utilities.formatDate(new Date(r[8]), tz, 'yyyy-MM-dd') : '',
    paymentMethod: String(r[9] || 'Auto-debit'),
    notes: String(r[10] || ''),
    autoLog: r[11] === true || r[11] === 'true' || r[11] === 'TRUE',
    isActive: r[12] === true || r[12] === 'true' || r[12] === 'TRUE' || r[12] === '',
    lastLoggedDate: r[13] ? Utilities.formatDate(new Date(r[13]), tz, 'yyyy-MM-dd') : '',
    customIntervalDays: parseInt(r[14]) || 0,
    loggedOccurrences: _parseOccurrenceList(r[15]),
    lastRemindedOccurrence: r[16] ? String(r[16]).trim() : '',
  };
}

// LoggedOccurrences is stored as a comma-separated yyyy-MM-dd list so the cell
// stays human-readable in the sheet.
function _parseOccurrenceList(raw) {
  if (!raw) return [];
  return String(raw).split(',').map(s => s.trim()).filter(s => /^\d{4}-\d{2}-\d{2}$/.test(s));
}

function getStandingInstructions() {
  const cached = SCRIPT_CACHE.get('si_all');
  if (cached) { try { return JSON.parse(cached); } catch (e) { } }
  try {
    const sheet = _getSISheet();
    if (sheet.getLastRow() <= 1) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, SI_COLUMNS.length).getValues()
      .filter(r => r[0])
      .map(_rowToSI);
    SCRIPT_CACHE.put('si_all', JSON.stringify(rows), TTL_SI);
    return rows;
  } catch (e) {
    Logger.log('getStandingInstructions error: ' + e.message);
    return [];
  }
}

function addStandingInstruction(si) {
  try {
    const sheet = _getSISheet();
    const id = 'SI-' + new Date().getTime();
    const tz = Session.getScriptTimeZone();
    sheet.appendRow([
      id, si.name, si.category, parseFloat(si.amount) || 0,
      si.frequency || 'monthly', parseInt(si.dayOfMonth) || 1, si.dayOfWeek || 'Monday',
      si.startDate ? _parseLocalDate(si.startDate) : new Date(),
      si.endDate ? _parseLocalDate(si.endDate) : '',
      si.paymentMethod || 'Auto-debit', si.notes || '',
      si.autoLog === true || si.autoLog === 'true', true, '',
      parseInt(si.customIntervalDays) || 0
    ]);
    SCRIPT_CACHE.remove('si_all');
    return { success: true, id };
  } catch (e) {
    Logger.log('addStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateStandingInstruction(id, si) {
  try {
    const sheet = _getSISheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.getRange(i + 1, 2, 1, 11).setValues([[
          si.name, si.category, parseFloat(si.amount) || 0,
          si.frequency || 'monthly', parseInt(si.dayOfMonth) || 1, si.dayOfWeek || 'Monday',
          si.startDate ? _parseLocalDate(si.startDate) : new Date(),
          si.endDate ? _parseLocalDate(si.endDate) : '',
          si.paymentMethod || 'Auto-debit', si.notes || '',
          si.autoLog === true || si.autoLog === 'true'
        ]]);
        // Write CustomIntervalDays separately — it sits after LastLoggedDate,
        // which this update must not touch
        sheet.getRange(i + 1, SI_COL_CUSTOM_INTERVAL).setValue(parseInt(si.customIntervalDays) || 0);
        SCRIPT_CACHE.remove('si_all');
        return { success: true };
      }
    }
    return { success: false, message: 'Not found.' };
  } catch (e) {
    Logger.log('updateStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function deleteStandingInstruction(id) {
  try {
    const sheet = _getSISheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.deleteRow(i + 1);
        SCRIPT_CACHE.remove('si_all');
        return { success: true };
      }
    }
    return { success: false, message: 'Not found.' };
  } catch (e) {
    Logger.log('deleteStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function toggleStandingInstruction(id, isActive) {
  try {
    const sheet = _getSISheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) {
        sheet.getRange(i + 1, 13).setValue(isActive); // IsActive column
        SCRIPT_CACHE.remove('si_all');
        return { success: true };
      }
    }
    return { success: false, message: 'Not found.' };
  } catch (e) {
    Logger.log('toggleStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Log a single SI occurrence as an expense.
// occurrenceDate (yyyy-MM-dd) identifies WHICH due date is being settled and
// defaults to today — that is what makes the call idempotent, so the same
// occurrence can never be logged twice from the UI, the email button and the
// auto-log trigger all at once.
function logStandingInstruction(id, occurrenceDate) {
  try {
    const all = getStandingInstructions();
    const si = all.find(s => s.id === id);
    if (!si) return { success: false, message: 'Standing instruction not found.' };
    if (!si.isActive) return { success: false, message: 'This instruction is paused.' };

    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const occurrence = /^\d{4}-\d{2}-\d{2}$/.test(occurrenceDate || '') ? occurrenceDate : today;

    if (_isOccurrenceLogged(si, occurrence)) {
      return { success: false, alreadyLogged: true, message: si.name + ' is already logged for ' + occurrence + '.' };
    }

    const expenseDate = _expenseDateForOccurrence(occurrence, today);
    const result = addExpense({
      date: expenseDate,
      category: si.category,
      description: si.name,
      amount: si.amount,
      paymentMethod: si.paymentMethod,
      notes: si.notes + (si.notes ? ' · ' : '') + '[Standing: ' + si.id + ' · due ' + occurrence + ']'
    });

    if (result.success) _recordSIOccurrence(id, occurrence, expenseDate);
    return { ...result, occurrenceDate: occurrence, expenseDate, name: si.name, amount: si.amount };
  } catch (e) {
    Logger.log('logStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Expenses always land in the ACTIVE shard, and shards are pruned by month at
// read time — so an expense dated outside the active shard's month would go
// missing from later queries. Log-ahead (clicking "Log" from a reminder for
// next month's rent) therefore falls back to today's date; the true due date is
// still recorded in the notes and in LoggedOccurrences.
function _expenseDateForOccurrence(occurrence, today) {
  return occurrence.substring(0, 7) === _getActiveShardMonth(today) ? occurrence : today;
}

function _getActiveShardMonth(today) {
  try {
    const activeId = getActiveShardId();
    const rec = _getAllShardRecords().find(s => s.id === activeId);
    if (rec && rec.month) return rec.month;
  } catch (e) { }
  return today.substring(0, 7);
}

// Appends an occurrence to the ledger and refreshes LastLoggedDate.
function _recordSIOccurrence(id, occurrence, expenseDate) {
  const sheet = _getSISheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== id) continue;
    const list = _parseOccurrenceList(data[i][SI_COL_LOGGED_OCCURRENCES - 1]);
    if (list.indexOf(occurrence) === -1) list.push(occurrence);
    list.sort();
    sheet.getRange(i + 1, SI_COL_LAST_LOGGED).setValue(expenseDate || occurrence);
    sheet.getRange(i + 1, SI_COL_LOGGED_OCCURRENCES)
      .setValue(list.slice(-SI_MAX_OCCURRENCE_HISTORY).join(','));
    break;
  }
  SCRIPT_CACHE.remove('si_all');
}

function _forgetSIOccurrence(id, occurrence) {
  const sheet = _getSISheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== id) continue;
    const list = _parseOccurrenceList(data[i][SI_COL_LOGGED_OCCURRENCES - 1]).filter(d => d !== occurrence);
    sheet.getRange(i + 1, SI_COL_LOGGED_OCCURRENCES).setValue(list.join(','));
    sheet.getRange(i + 1, SI_COL_LAST_LOGGED).setValue(list.length ? list[list.length - 1] : '');
    break;
  }
  SCRIPT_CACHE.remove('si_all');
}

// Has this specific due date already been settled?
//
// Rows written before the LoggedOccurrences ledger existed have no per-occurrence
// history, so for those we fall back to the old month-granularity rule — but only
// for frequencies that genuinely occur at most once a month. Weekly and custom
// SIs get no fallback, which is exactly the bug the ledger fixes.
function _isOccurrenceLogged(si, dueDate) {
  const ledger = si.loggedOccurrences || [];
  if (ledger.indexOf(dueDate) !== -1) return true;
  if (ledger.length > 0) return false;
  const oncePerMonth = ['monthly', 'quarterly', 'yearly'].indexOf(si.frequency) !== -1;
  return oncePerMonth && !!si.lastLoggedDate && si.lastLoggedDate.substring(0, 7) === dueDate.substring(0, 7);
}

// Undo a log made from an email reminder — deletes the expense and releases the
// occurrence so it shows as due again.
function undoStandingLog(id, occurrenceDate, expenseId) {
  try {
    if (expenseId) deleteExpense(expenseId);
    _forgetSIOccurrence(id, occurrenceDate);
    return { success: true };
  } catch (e) {
    Logger.log('undoStandingLog error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Expands every active SI into individual occurrences between two dates,
// each tagged with its own logged status. This is the single source of truth
// for the Recurring page, the auto-log trigger, the reminder email and the
// weekly report — none of which are limited to one calendar month any more.
function getSIOccurrences(fromStr, toStr, opts) {
  opts = opts || {};
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const out = [];

  getStandingInstructions()
    .filter(si => si.isActive || opts.includePaused)
    .filter(si => opts.autoLogOnly ? si.autoLog : true)
    .forEach(si => {
      _getDueDatesInRange(si, fromStr, toStr, tz).forEach(dueDate => {
        const logged = _isOccurrenceLogged(si, dueDate);
        if (opts.unloggedOnly && logged) return;
        const daysDiff = Math.round((_parseLocalDate(dueDate) - _parseLocalDate(today)) / 86400000);
        const status = logged ? 'logged'
          : dueDate < today ? 'overdue'
            : daysDiff <= 3 ? 'due_soon'
              : 'upcoming';
        out.push({ ...si, dueDate, status, daysDiff, isLogged: logged });
      });
    });

  out.sort((a, b) => a.dueDate.localeCompare(b.dueDate) || a.name.localeCompare(b.name));
  return out;
}

// Returns SIs due in the current calendar month with their logged status.
// Used by the Recurring page Zone 1 (upcoming this month).
function getUpcomingSIs() {
  try {
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const monthStart = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), tz, 'yyyy-MM-dd');
    const monthEnd = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 0), tz, 'yyyy-MM-dd');
    return getSIOccurrences(monthStart, monthEnd);
  } catch (e) {
    Logger.log('getUpcomingSIs error: ' + e.message);
    return [];
  }
}

// Returns the committed monthly total (all active SIs normalised to per-month amount)
// and a per-category breakdown. Used by Zone 2 and the Dashboard card.
function getCommittedMonthlyTotal() {
  try {
    const allSIs = getStandingInstructions().filter(si => si.isActive);
    const monthly = _normaliseToMonthly(allSIs);
    const byCategory = {};
    allSIs.forEach(si => {
      const m = _monthlyAmount(si);
      byCategory[si.category] = (byCategory[si.category] || 0) + m;
    });
    // Get average monthly income from last 3 months for committed % calculation
    let monthlyIncome = 0;
    try {
      const tz = Session.getScriptTimeZone();
      const now = new Date();
      const rows = getAllIncome();
      const cutoff = new Date(now); cutoff.setMonth(now.getMonth() - 3);
      const cutStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');
      const recent = rows.filter(r => r.date >= cutStr);
      const months = [...new Set(recent.map(r => r.date.substring(0, 7)))].length || 1;
      const total = recent.reduce((s, r) => s + r.amount, 0);
      monthlyIncome = Math.round(total / months);
    } catch (e) { }

    return {
      total: Math.round(monthly * 100) / 100,
      count: allSIs.length,
      byCategory,
      monthlyIncome,
      committedPct: monthlyIncome > 0 ? Math.round((monthly / monthlyIncome) * 100) : 0,
      discretionary: Math.max(0, monthlyIncome - monthly)
    };
  } catch (e) {
    Logger.log('getCommittedMonthlyTotal error: ' + e.message);
    return { total: 0, count: 0, byCategory: {}, monthlyIncome: 0, committedPct: 0, discretionary: 0 };
  }
}

// Page bootstrap — returns everything the Recurring page needs in one call.
function getStandingPageData() {
  try {
    const categories = getCategories() || [];
    const settings = getSettings() || DEFAULT_SETTINGS;
    const upcoming = getUpcomingSIs();
    const committed = getCommittedMonthlyTotal();
    const all = getStandingInstructions();
    return { categories, settings, upcoming, committed, all };
  } catch (e) {
    Logger.log('getStandingPageData error: ' + e.message);
    return { categories: DEFAULT_CATEGORIES, settings: DEFAULT_SETTINGS, upcoming: [], committed: { total: 0, count: 0, byCategory: {}, monthlyIncome: 0, committedPct: 0, discretionary: 0 }, all: [] };
  }
}

// Auto-log trigger — called daily at 07:00 by a time-based trigger.
// Logs every unlogged AutoLog occurrence due today, plus any it missed in the
// last SI_CATCHUP_DAYS days (a skipped trigger run used to lose that occurrence
// permanently). Occurrences on or before the SI's LastLoggedDate are never
// caught up, so rows that predate the ledger can't be re-logged retroactively.
function processStandingInstructions() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const from = Utilities.formatDate(new Date(now.getTime() - SI_CATCHUP_DAYS * 86400000), tz, 'yyyy-MM-dd');

  const due = getSIOccurrences(from, today, { autoLogOnly: true, unloggedOnly: true })
    .filter(o => o.dueDate === today || !o.lastLoggedDate || o.dueDate > o.lastLoggedDate);

  let logged = 0, failed = 0;
  due.forEach(o => {
    try {
      const result = logStandingInstruction(o.id, o.dueDate);
      if (result.success) {
        logged++;
        Logger.log('Auto-logged: ' + o.name + ' (due ' + o.dueDate + ')' + (o.dueDate < today ? ' [catch-up]' : ''));
      } else if (!result.alreadyLogged) {
        failed++;
        Logger.log('Auto-log failed for ' + o.name + ': ' + result.message);
      }
    } catch (e) {
      failed++;
      Logger.log('processStandingInstructions error for ' + o.id + ': ' + e.message);
    }
  });

  Logger.log('processStandingInstructions: logged=' + logged + ' failed=' + failed + ' candidates=' + due.length);
  return { logged, failed, candidates: due.length };
}

// ── SI helpers ───────────────────────────────────────────────

// Returns due dates for an SI across an arbitrary date range by walking the
// calendar months it spans. Ranges that cross a month boundary are the whole
// point: a reminder sent on 31 Jan has to know about rent due 1 Feb.
function _getDueDatesInRange(si, fromStr, toStr, tz) {
  if (!fromStr || !toStr || toStr < fromStr) return [];
  const out = [];
  let year = parseInt(fromStr.substring(0, 4));
  let mIndex = parseInt(fromStr.substring(5, 7)) - 1;
  const endYear = parseInt(toStr.substring(0, 4));
  const endMIndex = parseInt(toStr.substring(5, 7)) - 1;

  let guard = 0;
  while ((year < endYear || (year === endYear && mIndex <= endMIndex)) && guard++ < 120) {
    const month = year + '-' + ('0' + (mIndex + 1)).slice(-2);
    _getDueDatesInMonth(si, month, tz).forEach(d => {
      if (d >= fromStr && d <= toStr) out.push(d);
    });
    mIndex++;
    if (mIndex > 11) { mIndex = 0; year++; }
  }
  return out;
}

// Returns due dates (yyyy-MM-dd strings) for a given SI within a month.
function _getDueDatesInMonth(si, month, tz) {
  const dates = [];
  const year = parseInt(month.substring(0, 4));
  const mIndex = parseInt(month.substring(5, 7)) - 1; // 0-indexed month

  if (si.frequency === 'monthly') {
    const day = si.dayOfMonth || 1;
    const maxDay = new Date(year, mIndex + 1, 0).getDate();
    const d = new Date(year, mIndex, Math.min(day, maxDay));
    dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));

  } else if (si.frequency === 'weekly') {
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const target = dayNames.indexOf(si.dayOfWeek || 'Monday');
    const d = new Date(year, mIndex, 1);
    while (d.getMonth() === mIndex) {
      if (d.getDay() === target) dates.push(Utilities.formatDate(new Date(d), tz, 'yyyy-MM-dd'));
      d.setDate(d.getDate() + 1);
    }

  } else if (si.frequency === 'quarterly') {
    // Anchored on the start month; without one, fall back to calendar quarters
    // (Jan/Apr/Jul/Oct) rather than never being due at all.
    const startMonth = si.startDate ? parseInt(si.startDate.substring(5, 7)) - 1 : 0;
    if ((mIndex - startMonth + 12) % 3 === 0) {
      const day = si.dayOfMonth || 1;
      const maxDay = new Date(year, mIndex + 1, 0).getDate();
      const d = new Date(year, mIndex, Math.min(day, maxDay));
      dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
    }

  } else if (si.frequency === 'yearly') {
    // Anchored on the start date; without one, use January + DayOfMonth.
    const startM = si.startDate ? parseInt(si.startDate.substring(5, 7)) - 1 : 0;
    const startD = si.startDate ? parseInt(si.startDate.substring(8, 10)) : (si.dayOfMonth || 1);
    if (mIndex === startM) {
      const maxDay = new Date(year, mIndex + 1, 0).getDate();
      const d = new Date(year, mIndex, Math.min(startD, maxDay));
      dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
    }

  } else if (si.frequency === 'custom' && si.customIntervalDays > 0 && si.startDate) {
    // Custom: every N days from start date
    const interval = si.customIntervalDays;
    const start = new Date(parseInt(si.startDate.substring(0, 4)),
      parseInt(si.startDate.substring(5, 7)) - 1,
      parseInt(si.startDate.substring(8, 10)));
    const monthStart = new Date(year, mIndex, 1);
    const monthEnd = new Date(year, mIndex + 1, 0);

    // Fast-forward from start to the first occurrence >= monthStart
    let cursor = new Date(start);
    if (cursor < monthStart) {
      const daysBetween = Math.floor((monthStart - cursor) / 86400000);
      const skip = Math.floor(daysBetween / interval) * interval;
      cursor.setDate(cursor.getDate() + skip);
    }
    while (cursor <= monthEnd) {
      if (cursor >= monthStart && cursor.getMonth() === mIndex) {
        dates.push(Utilities.formatDate(new Date(cursor), tz, 'yyyy-MM-dd'));
      }
      cursor.setDate(cursor.getDate() + interval);
    }
  }

  return dates.filter(d => {
    if (si.startDate && d < si.startDate) return false;
    if (si.endDate && d > si.endDate) return false;
    return true;
  });
}

// Converts a single SI's amount to a monthly equivalent.
function _monthlyAmount(si) {
  const a = si.amount || 0;
  if (si.frequency === 'weekly') return (a * 52) / 12;
  if (si.frequency === 'quarterly') return a / 3;
  if (si.frequency === 'yearly') return a / 12;
  if (si.frequency === 'custom' && si.customIntervalDays > 0) return (a * 365.25 / si.customIntervalDays) / 12;
  return a; // monthly default
}

// Returns the total normalised monthly cost of an array of SIs.
function _normaliseToMonthly(sis) {
  return sis.reduce((sum, si) => sum + _monthlyAmount(si), 0);
}

// ── CSV EXPORT ───────────────────────────────────────────────
function exportToCSV(filters) {
  filters = filters || {};
  // If no date range specified, explicitly read ALL shards rather than
  // defaulting to active shard only (_getShardsForRange behaviour with no args)
  if (!filters.startDate && !filters.endDate) {
    const allShardIds = _getAllShardRecords().map(s => s.id);
    const activeId = getActiveShardId();
    if (!allShardIds.includes(activeId)) allShardIds.push(activeId);
    let expenses = [];
    allShardIds.forEach(id => {
      try { expenses = expenses.concat(_readShardExpenses(id)); } catch (e) { }
    });
    if (filters.category && filters.category !== 'All') {
      expenses = expenses.filter(e => e.category === filters.category);
    }
    expenses.sort((a, b) => new Date(b.date) - new Date(a.date));
    const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Payment Method', 'Notes'];
    const rows = [headers, ...expenses.map(e => [e.id, e.date, e.category, e.description, e.amount, e.paymentMethod, e.notes])];
    return rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  }
  // Date range provided — use getExpenses which handles shard pruning
  const expenses = getExpenses(filters);
  const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Payment Method', 'Notes'];
  const rows = [headers, ...expenses.map(e => [e.id, e.date, e.category, e.description, e.amount, e.paymentMethod, e.notes])];
  return rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
}

// ── Weekly Email Report ───────────────────────────────────────

// Entry point called by the time-based trigger.
// Checks if still enabled in settings before sending.
function weeklyReportTrigger() {
  const settings = getSettings();
  if (settings.weeklyReportEnabled !== 'true') return;
  const email = settings.weeklyReportEmail ||
    PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') ||
    Session.getActiveUser().getEmail();
  if (!email) { Logger.log('weeklyReportTrigger: no email configured'); return; }
  sendWeeklyReport(email);
}

// Run this function from the Apps Script Editor dropdown to trigger the Google OAuth permission popup!
// MUST NOT use try/catch so Apps Script intercepts the scope requirement and prompts for consent.
function GRANT_EMAIL_PERMISSIONS() {
  const quota = MailApp.getRemainingDailyQuota();
  Logger.log('✓ MailApp permission granted successfully! Remaining daily quota: ' + quota);
  return 'Permission active! Quota: ' + quota;
}

// Checks if MailApp permission has been granted by trying a no-op
// call that requires the same scope but sends nothing.
// Returns { granted: true } or { granted: false, message, instructions }
function checkMailPermission() {
  let executingAs = '';
  try { executingAs = Session.getEffectiveUser().getEmail() || ''; } catch (e) { }

  try {
    MailApp.getRemainingDailyQuota(); // requires script.send_mail scope, sends nothing
    return { granted: true, email: executingAs };
  } catch (e) {
    var errMsg = (e && e.message) ? e.message : String(e);
    Logger.log('checkMailPermission error: ' + errMsg);

    let editorUrl = '';
    try { editorUrl = 'https://script.google.com/home/projects/' + ScriptApp.getScriptId() + '/edit'; } catch (e2) { }

    return {
      granted: false,
      message: errMsg,
      email: executingAs,
      editorUrl: editorUrl,
      instructions: [
        '1. Click "Open Apps Script Editor" below',
        '2. In the top dropdown, select "GRANT_EMAIL_PERMISSIONS"',
        '3. Click Run. Google will show the "Authorization Required" popup!',
        '4. Click "Review permissions" → choose ' + (executingAs || 'your account') + ' → "Advanced" → "Go to Spendwise (unsafe)" → "Allow"',
        '5. Return here and click Save or Send Test again!'
      ]
    };
  }
}

// Returns the most recent COMPLETE week [start, end], honouring the configured
// week start day. Reporting the *current* partial week meant the default
// schedule (send on Monday, week starts Monday) covered midnight → 08:00 and
// reported ₹0 against a full previous week.
function _lastCompleteWeek(now, weekStartDay) {
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const startIdx = Math.max(0, dayNames.indexOf(weekStartDay || 'Monday'));

  const start = new Date(now);
  start.setHours(12, 0, 0, 0); // noon — never let DST shift the date
  const daysIntoWeek = (start.getDay() - startIdx + 7) % 7;
  start.setDate(start.getDate() - daysIntoWeek - 7);

  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  return { start, end };
}

// Builds and sends the weekly report. Call from editor to test.
function sendWeeklyReport(toEmail) {
  try {
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const settings = getSettings();

    // The week being reported on, and the week before it for comparison.
    // Both windows are a full 7 days, so the percentage change is like-for-like.
    const week = _lastCompleteWeek(now, settings.weekStartDay);
    const prevStart = new Date(week.start); prevStart.setDate(week.start.getDate() - 7);
    const prevEnd = new Date(week.end); prevEnd.setDate(week.end.getDate() - 7);

    const fmt2 = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const thisWeekExpenses = getExpenses({ startDate: fmt2(week.start), endDate: fmt2(week.end) });
    const lastWeekExpenses = getExpenses({ startDate: fmt2(prevStart), endDate: fmt2(prevEnd) });
    const thisWeekIncome = getIncome({ startDate: fmt2(week.start), endDate: fmt2(week.end) });

    const thisTotal = thisWeekExpenses.reduce((s, e) => s + e.amount, 0);
    const lastTotal = lastWeekExpenses.reduce((s, e) => s + e.amount, 0);
    const thisInc = thisWeekIncome.reduce((s, r) => s + r.amount, 0);

    // Category breakdown this week
    const byCat = {};
    thisWeekExpenses.forEach(e => {
      byCat[e.category] = (byCat[e.category] || 0) + e.amount;
    });
    const topCats = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 5);

    // Budget alerts — categories over 80% of monthly budget
    const budgetSummary = getBudgetSummary().filter(b => b.budget > 0 && b.percentage >= 80);

    // Week-over-week change
    const pctChange = lastTotal > 0
      ? Math.round(((thisTotal - lastTotal) / lastTotal) * 100)
      : null;
    // Zero change is neither good nor bad — it used to render as a green decrease
    const changeStr = pctChange === null ? ''
      : pctChange > 0 ? `▲ ${pctChange}% vs last week`
        : pctChange < 0 ? `▼ ${Math.abs(pctChange)}% vs last week`
          : 'Flat vs last week';
    const changeColor = pctChange > 0 ? '#eb5757' : pctChange < 0 ? '#6fcf97' : '#8892a4';
    const changeBg = pctChange > 0 ? 'rgba(235,87,87,0.15)'
      : pctChange < 0 ? 'rgba(111,207,151,0.15)' : 'rgba(136,146,164,0.15)';

    const weekLabel = Utilities.formatDate(week.start, tz, 'MMM d') +
      ' – ' + Utilities.formatDate(week.end, tz, 'MMM d, yyyy');

    // ── Build HTML email ──────────────────────────────────────
    const fmtRs = amt => '&#8377;' + parseFloat(amt).toLocaleString('en-IN',
      { minimumFractionDigits: 0, maximumFractionDigits: 0 });

    // Top categories table rows — category name left, bar, amount right
    const topCatRows = topCats.map(([cat, amt]) => {
      const pct = thisTotal > 0 ? Math.round((amt / thisTotal) * 100) : 0;
      const barWidth = Math.max(4, pct);
      return `
        <tr>
          <td style="padding:10px 0 0;vertical-align:top;">
            <table style="width:100%;border-collapse:collapse;">
              <tr>
                <td style="color:#c8d0dc;font-size:13px;font-weight:500;padding-bottom:4px;">${cat}</td>
                <td style="text-align:right;color:#e8ecf0;font-size:13px;font-weight:700;padding-bottom:4px;">${fmtRs(amt)}</td>
              </tr>
              <tr>
                <td colspan="2" style="padding-bottom:2px;">
                  <table style="width:100%;border-collapse:collapse;">
                    <tr>
                      <td style="width:${barWidth}%;background:#4f8ef7;height:3px;border-radius:2px;"></td>
                      <td style="background:#1e2330;height:3px;border-radius:2px;"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td style="color:#5c6478;font-size:11px;">${pct}% of total</td>
                <td></td>
              </tr>
            </table>
          </td>
        </tr>`;
    }).join('');

    // Budget alerts — table layout so spacing is consistent
    const budgetAlertRows = budgetSummary.map(b => {
      const overBudget = b.percentage >= 100;
      const color = overBudget ? '#eb5757' : '#f2994a';
      const icon = overBudget ? '&#x1F6A8;' : '&#9888;';
      return `
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #1e2330;">
            <table style="width:100%;border-collapse:collapse;">
              <tr>
                <td style="color:${color};font-size:13px;font-weight:500;">${icon}&nbsp;&nbsp;${b.category}</td>
                <td style="text-align:right;white-space:nowrap;">
                  <span style="background:${overBudget ? 'rgba(235,87,87,0.15)' : 'rgba(242,153,74,0.15)'};color:${color};font-size:12px;font-weight:700;padding:2px 8px;border-radius:99px;">${b.percentage}%</span>
                </td>
              </tr>
              <tr>
                <td colspan="2" style="color:#5c6478;font-size:11px;padding-top:2px;">
                  Spent &#8377;${b.spent.toLocaleString('en-IN')} of &#8377;${b.budget.toLocaleString('en-IN')} budget
                </td>
              </tr>
            </table>
          </td>
        </tr>`;
    }).join('');

    const budgetAlerts = budgetSummary.length === 0 ? '' :
      `<div style="margin-top:24px;padding-top:20px;border-top:1px solid #272d3d;">
        <p style="font-size:11px;font-weight:700;letter-spacing:0.08em;color:#8892a4;text-transform:uppercase;margin:0 0 4px;">Budget Alerts</p>
        <p style="font-size:11px;color:#5c6478;margin:0 0 12px;">Categories at or near their monthly limit</p>
        <table style="width:100%;border-collapse:collapse;">${budgetAlertRows}</table>
      </div>`;

    // Income section
    const netAmt = thisInc - thisTotal;
    const netColor = netAmt >= 0 ? '#6fcf97' : '#eb5757';
    const incomeSection = thisInc > 0 ?
      `<div style="margin-top:16px;">
        <table style="width:100%;border-collapse:collapse;background:#0d1a0f;border-radius:8px;border:1px solid rgba(111,207,151,0.2);">
          <tr>
            <td style="padding:16px 18px;">
              <p style="font-size:11px;font-weight:700;letter-spacing:0.08em;color:#4a8c5c;text-transform:uppercase;margin:0 0 8px;">Income This Week</p>
              <table style="width:100%;border-collapse:collapse;">
                <tr>
                  <td style="font-size:26px;font-weight:700;color:#6fcf97;font-family:Georgia,serif;">+${fmtRs(thisInc)}</td>
                  <td style="text-align:right;vertical-align:bottom;">
                    <span style="font-size:13px;color:${netColor};font-weight:600;">Net ${netAmt >= 0 ? '+' : ''}${fmtRs(netAmt)}</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>` : '';

    // Upcoming standing instructions — due in the next 7 days, not yet logged.
    // Uses the range expander so items early next month still show up.
    const in7Days = new Date(now); in7Days.setDate(now.getDate() + 7);
    const upcomingSIs = getSIOccurrences(fmt2(now), fmt2(in7Days), { unloggedOnly: true });

    const upcomingSIRows = upcomingSIs.map(si => {
      const d = new Date(si.dueDate + 'T00:00:00');
      const lbl = d.toLocaleString('en-IN', { weekday: 'short', day: 'numeric', month: 'short' });
      return `
        <tr>
          <td style="padding:7px 0;color:#c8d0dc;font-size:13px;border-bottom:1px solid #1e2330;">${lbl}</td>
          <td style="padding:7px 0;color:#c8d0dc;font-size:13px;border-bottom:1px solid #1e2330;">${si.name}</td>
          <td style="padding:7px 0;text-align:right;font-weight:700;color:#e8ecf0;font-size:13px;border-bottom:1px solid #1e2330;white-space:nowrap;">${fmtRs(si.amount)}</td>
        </tr>`;
    }).join('');

    const upcomingSection = upcomingSIs.length === 0 ? '' :
      `<table style="width:100%;border-collapse:collapse;background:#1a1f2e;border-radius:14px;border:1px solid #272d3d;margin-bottom:14px;">
        <tr>
          <td style="padding:20px 22px 16px;">
            <p style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#5c6478;text-transform:uppercase;margin:0 0 12px;">Upcoming This Week</p>
            <table style="width:100%;border-collapse:collapse;">${upcomingSIRows}</table>
            <p style="font-size:11px;color:#5c6478;margin:10px 0 0;">
              ${fmtRs(upcomingSIs.reduce((s, si) => s + si.amount, 0))} in scheduled debits
            </p>
          </td>
        </tr>
      </table>`;

    // Week-over-week change badge
    const changeBadge = pctChange !== null ?
      `<span style="display:inline-block;margin-left:10px;font-size:12px;font-weight:700;padding:3px 10px;border-radius:99px;background:${changeBg};color:${changeColor};">${changeStr}</span>` : '';

    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#0a0d14;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;">
  <div style="max-width:540px;margin:0 auto;padding:32px 16px 48px;">

    <!-- Wordmark -->
    <div style="margin-bottom:28px;">
      <span style="font-size:18px;font-weight:800;color:#e8ecf0;letter-spacing:-0.02em;">Spendwise</span>
      <span style="font-size:12px;color:#5c6478;margin-left:8px;">Weekly Report</span>
    </div>

    <!-- Week label -->
    <p style="font-size:13px;color:#8892a4;margin:0 0 20px;">${weekLabel}</p>

    <!-- Total spend hero card -->
    <table style="width:100%;border-collapse:collapse;background:#1a1f2e;border-radius:14px;border:1px solid #272d3d;margin-bottom:14px;">
      <tr>
        <td style="padding:22px 22px 18px;">
          <p style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#5c6478;text-transform:uppercase;margin:0 0 10px;">Total Spent This Week</p>
          <table style="border-collapse:collapse;">
            <tr>
              <td style="font-size:38px;font-weight:800;color:#e8ecf0;font-family:Georgia,serif;line-height:1;">${fmtRs(thisTotal)}</td>
              <td style="vertical-align:bottom;padding-bottom:4px;padding-left:4px;">${changeBadge}</td>
            </tr>
          </table>
          ${pctChange !== null ? `<p style="font-size:12px;color:#5c6478;margin:6px 0 0;">vs ${fmtRs(lastTotal)} last week</p>` : ''}
        </td>
      </tr>
    </table>

    <!-- Top categories card -->
    <table style="width:100%;border-collapse:collapse;background:#1a1f2e;border-radius:14px;border:1px solid #272d3d;margin-bottom:14px;">
      <tr>
        <td style="padding:20px 22px 16px;">
          <p style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#5c6478;text-transform:uppercase;margin:0 0 6px;">Top Categories</p>
          ${topCats.length > 0
        ? `<table style="width:100%;border-collapse:collapse;">${topCatRows}</table>`
        : '<p style="color:#5c6478;font-size:13px;margin:12px 0 0;">No expenses recorded this week.</p>'}
          ${budgetAlerts}
          ${incomeSection}
        </td>
      </tr>
    </table>

    ${upcomingSection}

    <!-- Footer -->
    <table style="width:100%;border-collapse:collapse;">
      <tr>
        <td style="padding-top:20px;text-align:center;">
          <p style="font-size:11px;color:#3a4050;margin:0;">
            You're receiving this because weekly reports are enabled in your Spendwise settings.
          </p>
        </td>
      </tr>
    </table>

  </div>
</body>
</html>`;

    const subject = 'Spendwise: \u20B9' + parseFloat(thisTotal).toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + ' spent this week (' + weekLabel + ')';
    MailApp.sendEmail({ to: toEmail, subject, htmlBody: html });
    Logger.log('✓ Weekly report sent to ' + toEmail);
    return { success: true };
  } catch (e) {
    Logger.log('sendWeeklyReport error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Install or update the weekly trigger based on current settings.
// Called by saveSettings() whenever report settings change.
function scheduleWeeklyReport() {
  // Remove any existing weekly report trigger first
  cancelWeeklyReport();

  const settings = getSettings();
  if (settings.weeklyReportEnabled !== 'true') return { cancelled: true };

  // Must use ScriptApp.WeekDay enum — plain integers silently fail
  const dayEnumMap = {
    Sunday: ScriptApp.WeekDay.SUNDAY,
    Monday: ScriptApp.WeekDay.MONDAY,
    Tuesday: ScriptApp.WeekDay.TUESDAY,
    Wednesday: ScriptApp.WeekDay.WEDNESDAY,
    Thursday: ScriptApp.WeekDay.THURSDAY,
    Friday: ScriptApp.WeekDay.FRIDAY,
    Saturday: ScriptApp.WeekDay.SATURDAY
  };
  const day = dayEnumMap[settings.weeklyReportDay] || ScriptApp.WeekDay.MONDAY;
  const hour = parseInt(settings.weeklyReportTime) || 8;

  const trigger = ScriptApp.newTrigger('weeklyReportTrigger')
    .timeBased().onWeekDay(day).atHour(hour).nearMinute(0).create();

  // Store trigger ID and owner email in Script Properties
  // Trigger ID lets cancelWeeklyReport() delete the exact trigger
  const props = PropertiesService.getScriptProperties();
  props.setProperty('WEEKLY_REPORT_TRIGGER_ID', trigger.getUniqueId());
  try {
    props.setProperty('OWNER_EMAIL', settings.weeklyReportEmail || Session.getActiveUser().getEmail());
  } catch (e) { }

  Logger.log('✓ Weekly report trigger set: ' + settings.weeklyReportDay + ' at ' + hour + ':00 (id: ' + trigger.getUniqueId() + ')');
  return { success: true, day: settings.weeklyReportDay, hour, triggerId: trigger.getUniqueId() };
}

// Remove the weekly report trigger using stored ID for precision.
// Falls back to scanning by handler name if ID not found.
function cancelWeeklyReport() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('WEEKLY_REPORT_TRIGGER_ID');
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'weeklyReportTrigger') {
      if (!storedId || t.getUniqueId() === storedId) {
        ScriptApp.deleteTrigger(t);
      }
    }
  });

  props.deleteProperty('WEEKLY_REPORT_TRIGGER_ID');
}

// Save report settings and reschedule trigger.
// saveSettings() handles merging so general keys are preserved.
function saveSettingsWithReport(reportSettings) {
  const result = saveSettings(reportSettings);
  if (!result.success) return result;
  try { scheduleWeeklyReport(); } catch (e) {
    Logger.log('scheduleWeeklyReport error: ' + e.message);
  }
  return result;
}

// Send a test report immediately to the configured email.
// Called from the Settings page "Send Test" button.
function sendTestReport() {
  const settings = getSettings();
  const email = settings.weeklyReportEmail ||
    PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') ||
    Session.getActiveUser().getEmail();
  if (!email) return { success: false, message: 'No email address configured.' };
  return sendWeeklyReport(email);
}

// ── Recurring Reminder Email ──────────────────────────────────
// Sends a heads-up N days before each due date, with a one-click Log button
// per manual item. The buttons hit doGet(action=logSI) on the web app.

// The web app can't ask for its own URL from a trigger context reliably, so we
// capture it the first time the app is opened and reuse it for email links.
function _getWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('WEBAPP_URL');
  if (stored) return stored;
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) { props.setProperty('WEBAPP_URL', url); return url; }
  } catch (e) { }
  return '';
}

// Called on every page load, so it must stay cheap — bail out once stored.
function _rememberWebAppUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('WEBAPP_URL')) return;
    const url = ScriptApp.getService().getUrl();
    if (url && url.indexOf('/exec') !== -1) props.setProperty('WEBAPP_URL', url);
  } catch (e) { }
}

// Links in email are signed so a stray or forwarded URL can't log an expense.
// The web app is already owner-only; this is a second lock, not the only one.
function _siActionSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('SI_ACTION_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SI_ACTION_SECRET', secret);
  }
  return secret;
}

function _siActionToken(payload) {
  return Utilities.computeHmacSha256Signature(String(payload), _siActionSecret())
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').substring(0, 24);
}

function _siActionUrl(action, params) {
  const base = _getWebAppUrl();
  if (!base) return '';
  const qs = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  return base + (base.indexOf('?') === -1 ? '?' : '&') + 'action=' + action + '&' + qs;
}

// Entry point called by the time-based trigger.
function recurringReminderTrigger() {
  const settings = getSettings();
  if (settings.recurringReminderEnabled !== 'true') return;
  sendRecurringReminders();
}

// Builds and sends the reminder. Pass { force: true } from the test button to
// send even when nothing is due and to ignore the already-reminded guard.
function sendRecurringReminders(opts) {
  opts = opts || {};
  try {
    const settings = getSettings();
    const email = opts.toEmail || settings.recurringReminderEmail || settings.weeklyReportEmail ||
      PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') ||
      Session.getActiveUser().getEmail();
    if (!email) return { success: false, message: 'No email address configured.' };

    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const daysBefore = Math.max(0, Math.min(7, parseInt(settings.recurringReminderDaysBefore) || 1));
    const targetDate = Utilities.formatDate(new Date(now.getTime() + daysBefore * 86400000), tz, 'yyyy-MM-dd');
    const currency = settings.currency || '₹';

    // Upcoming — the occurrences that fall exactly on the target date
    let upcoming = getSIOccurrences(targetDate, targetDate, { unloggedOnly: true });
    if (!opts.force) {
      upcoming = upcoming.filter(o => !o.lastRemindedOccurrence || o.lastRemindedOccurrence < o.dueDate);
    }

    // Overdue — anything still unsettled from the last 30 days. Occurrences on
    // or before LastLoggedDate are excluded so rows that predate the ledger
    // don't arrive as a wall of false alarms on the first send.
    const overdueFrom = Utilities.formatDate(new Date(now.getTime() - 30 * 86400000), tz, 'yyyy-MM-dd');
    const yesterday = Utilities.formatDate(new Date(now.getTime() - 86400000), tz, 'yyyy-MM-dd');
    const overdue = getSIOccurrences(overdueFrom, yesterday, { unloggedOnly: true })
      .filter(o => !o.lastLoggedDate || o.dueDate > o.lastLoggedDate);

    if (!upcoming.length && !overdue.length && !opts.force) {
      Logger.log('sendRecurringReminders: nothing due on ' + targetDate + ', no email sent');
      return { success: true, skipped: true, reason: 'nothing due' };
    }

    const manual = upcoming.filter(o => !o.autoLog);
    const auto = upcoming.filter(o => o.autoLog);
    const html = _buildReminderEmail({ today, targetDate, daysBefore, manual, auto, overdue, currency, tz });

    const dueTotal = upcoming.reduce((s, o) => s + o.amount, 0);
    const count = upcoming.length + overdue.length;
    const whenLabel = daysBefore === 0 ? 'today' : daysBefore === 1 ? 'tomorrow'
      : 'in ' + daysBefore + ' days';
    const subject = upcoming.length
      ? 'Spendwise: ' + currency + Math.round(dueTotal).toLocaleString('en-IN') + ' due ' + whenLabel +
      ' (' + count + ' item' + (count !== 1 ? 's' : '') + ')'
      : 'Spendwise: ' + overdue.length + ' recurring item' + (overdue.length !== 1 ? 's' : '') + ' still unlogged';

    MailApp.sendEmail({ to: email, subject, htmlBody: html });

    // Mark the upcoming ones so a re-run today doesn't send twice.
    // Test sends are excluded — they must never suppress the real reminder.
    if (!opts.force) {
      upcoming.forEach(o => {
        try { _markSIReminded(o.id, o.dueDate); } catch (e) { }
      });
    }

    Logger.log('✓ Recurring reminder sent to ' + email + ' (' + upcoming.length + ' upcoming, ' + overdue.length + ' overdue)');
    return { success: true, upcoming: upcoming.length, overdue: overdue.length };
  } catch (e) {
    Logger.log('sendRecurringReminders error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function _markSIReminded(id, dueDate) {
  const sheet = _getSISheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.getRange(i + 1, SI_COL_LAST_REMINDED).setValue(dueDate);
      break;
    }
  }
  SCRIPT_CACHE.remove('si_all');
}

function _buildReminderEmail(ctx) {
  const fmtAmt = amt => ctx.currency + parseFloat(amt).toLocaleString('en-IN',
    { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const dateLabel = ds => Utilities.formatDate(_parseLocalDate(ds), ctx.tz, 'EEE, d MMM');

  // Gmail-safe button: a table-wrapped anchor, no flexbox, no external assets.
  const logButton = o => {
    const url = _siActionUrl('logSI', { id: o.id, d: o.dueDate, t: _siActionToken(o.id + '|' + o.dueDate) });
    if (!url) return '<span style="font-size:11px;color:#5c6478;">Open Spendwise to log</span>';
    return `<a href="${url}" style="display:inline-block;background:#4f8ef7;color:#0a0d14;font-size:12px;font-weight:700;text-decoration:none;padding:7px 14px;border-radius:8px;white-space:nowrap;">Log it</a>`;
  };

  const itemRow = (o, withButton) => `
    <tr>
      <td style="padding:12px 0;border-bottom:1px solid #1e2330;">
        <table style="width:100%;border-collapse:collapse;">
          <tr>
            <td style="vertical-align:top;">
              <div style="color:#e8ecf0;font-size:14px;font-weight:600;">${o.name}</div>
              <div style="color:#5c6478;font-size:11px;padding-top:3px;">${o.category} · ${dateLabel(o.dueDate)} · ${o.paymentMethod}</div>
            </td>
            <td style="text-align:right;vertical-align:top;white-space:nowrap;padding-left:12px;">
              <div style="color:#e8ecf0;font-size:15px;font-weight:700;font-family:Georgia,serif;">${fmtAmt(o.amount)}</div>
              ${withButton ? `<div style="padding-top:7px;">${logButton(o)}</div>` : ''}
            </td>
          </tr>
        </table>
      </td>
    </tr>`;

  const card = (title, subtitle, rows, accent) => `
    <table style="width:100%;border-collapse:collapse;background:#1a1f2e;border-radius:14px;border:1px solid ${accent || '#272d3d'};margin-bottom:14px;">
      <tr>
        <td style="padding:20px 22px 14px;">
          <p style="font-size:11px;font-weight:700;letter-spacing:0.1em;color:#5c6478;text-transform:uppercase;margin:0 0 2px;">${title}</p>
          <p style="font-size:11px;color:#5c6478;margin:0 0 8px;">${subtitle}</p>
          <table style="width:100%;border-collapse:collapse;">${rows}</table>
        </td>
      </tr>
    </table>`;

  const whenLabel = ctx.daysBefore === 0 ? 'today'
    : ctx.daysBefore === 1 ? 'tomorrow' : 'in ' + ctx.daysBefore + ' days';

  const overdueCard = ctx.overdue.length ? card(
    'Overdue &#9888;',
    'Past their due date and not logged yet',
    ctx.overdue.map(o => itemRow(o, true)).join(''),
    'rgba(235,87,87,0.35)') : '';

  const manualCard = ctx.manual.length ? card(
    'Due ' + whenLabel,
    dateLabel(ctx.targetDate) + ' · tap Log to record it now',
    ctx.manual.map(o => itemRow(o, true)).join('')) : '';

  const autoCard = ctx.auto.length ? card(
    'Auto-debit ' + whenLabel,
    'Logged automatically — no action needed',
    ctx.auto.map(o => itemRow(o, false)).join('')) : '';

  const grandTotal = [].concat(ctx.manual, ctx.auto, ctx.overdue).reduce((s, o) => s + o.amount, 0);
  const appUrl = _getWebAppUrl();
  const emptyNote = (!ctx.overdue.length && !ctx.manual.length && !ctx.auto.length)
    ? `<table style="width:100%;border-collapse:collapse;background:#1a1f2e;border-radius:14px;border:1px solid #272d3d;margin-bottom:14px;">
         <tr><td style="padding:22px;text-align:center;color:#5c6478;font-size:13px;">Nothing due ${whenLabel}. This is what a quiet day looks like.</td></tr>
       </table>` : '';

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#0a0d14;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;">
  <div style="max-width:540px;margin:0 auto;padding:32px 16px 48px;">

    <div style="margin-bottom:24px;">
      <span style="font-size:18px;font-weight:800;color:#e8ecf0;letter-spacing:-0.02em;">Spendwise</span>
      <span style="font-size:12px;color:#5c6478;margin-left:8px;">Recurring Reminder</span>
    </div>

    <p style="font-size:13px;color:#8892a4;margin:0 0 20px;">
      ${grandTotal > 0 ? fmtAmt(grandTotal) + ' across ' + (ctx.manual.length + ctx.auto.length + ctx.overdue.length) + ' item' + ((ctx.manual.length + ctx.auto.length + ctx.overdue.length) !== 1 ? 's' : '') : 'Your recurring commitments'}
    </p>

    ${overdueCard}
    ${manualCard}
    ${autoCard}
    ${emptyNote}

    <table style="width:100%;border-collapse:collapse;">
      <tr>
        <td style="padding-top:8px;text-align:center;">
          ${appUrl ? `<a href="${appUrl}?page=standing" style="display:inline-block;color:#4f8ef7;font-size:12px;font-weight:600;text-decoration:none;">Open Recurring in Spendwise &rarr;</a>` : ''}
          <p style="font-size:11px;color:#3a4050;margin:14px 0 0;">
            You're receiving this because recurring reminders are enabled in your Spendwise settings.
          </p>
        </td>
      </tr>
    </table>

  </div>
</body>
</html>`;
}

// Install or update the reminder trigger based on current settings.
function scheduleRecurringReminder() {
  cancelRecurringReminder();

  const settings = getSettings();
  if (settings.recurringReminderEnabled !== 'true') return { cancelled: true };

  const hour = parseInt(settings.recurringReminderTime);
  const trigger = ScriptApp.newTrigger('recurringReminderTrigger')
    .timeBased().everyDays(1).atHour(isNaN(hour) ? 9 : hour).nearMinute(0).create();

  const props = PropertiesService.getScriptProperties();
  props.setProperty('RECURRING_REMINDER_TRIGGER_ID', trigger.getUniqueId());
  try {
    props.setProperty('OWNER_EMAIL', settings.recurringReminderEmail ||
      props.getProperty('OWNER_EMAIL') || Session.getActiveUser().getEmail());
  } catch (e) { }

  Logger.log('✓ Recurring reminder trigger set: daily at ' + hour + ':00 (id: ' + trigger.getUniqueId() + ')');
  return { success: true, hour, triggerId: trigger.getUniqueId() };
}

function cancelRecurringReminder() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('RECURRING_REMINDER_TRIGGER_ID');
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'recurringReminderTrigger') {
      if (!storedId || t.getUniqueId() === storedId) ScriptApp.deleteTrigger(t);
    }
  });
  props.deleteProperty('RECURRING_REMINDER_TRIGGER_ID');
}

// Save reminder settings and reschedule the trigger.
function saveSettingsWithReminder(reminderSettings) {
  const result = saveSettings(reminderSettings);
  if (!result.success) return result;
  try { scheduleRecurringReminder(); } catch (e) {
    Logger.log('scheduleRecurringReminder error: ' + e.message);
    return { success: false, message: e.message };
  }
  return result;
}

// Send a reminder immediately, even if nothing is due — Settings "Send Test".
// toEmail lets the test go to the address currently typed in the form, before
// the settings have been saved.
function sendTestRecurringReminder(toEmail) {
  return sendRecurringReminders({ force: true, toEmail: toEmail || '' });
}

// ── Reminder email actions (doGet) ────────────────────────────
// Reached only by an authenticated owner (doGet rejects everyone else) AND with
// a valid signature, so a link that leaks out of the inbox is inert.
function _handleSIEmailAction(action, params) {
  const id = String(params.id || '');
  const occurrence = String(params.d || '');
  const token = String(params.t || '');
  const expenseId = String(params.x || '');

  if (action === 'undoSI') {
    if (token !== _siActionToken('undo|' + id + '|' + occurrence + '|' + expenseId)) {
      return _siActionPage('error', 'Link expired', 'This undo link is no longer valid.', '');
    }
    const undone = undoStandingLog(id, occurrence, expenseId);
    return undone.success
      ? _siActionPage('info', 'Undone', 'The expense was removed and the item is due again.', '')
      : _siActionPage('error', 'Could not undo', undone.message || 'Please remove it from the Expenses page.', '');
  }

  if (token !== _siActionToken(id + '|' + occurrence)) {
    return _siActionPage('error', 'Link expired', 'This log link is no longer valid. Open Spendwise to log it manually.', '');
  }

  const result = logStandingInstruction(id, occurrence);

  if (result.alreadyLogged) {
    return _siActionPage('info', 'Already logged', result.message, '');
  }
  if (!result.success) {
    return _siActionPage('error', 'Could not log it', result.message || 'Something went wrong.', '');
  }

  const settings = getSettings();
  const amount = (settings.currency || '₹') +
    parseFloat(result.amount).toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  const dateNote = result.expenseDate === result.occurrenceDate
    ? 'Dated ' + result.expenseDate
    : 'Dated ' + result.expenseDate + ' (due ' + result.occurrenceDate + ')';

  // result.id is the new expense's ID — needed so Undo can delete it
  const undoUrl = _siActionUrl('undoSI', {
    id: id, d: occurrence, x: result.id,
    t: _siActionToken('undo|' + id + '|' + occurrence + '|' + result.id)
  });

  return _siActionPage('success', 'Logged', result.name + ' · ' + amount + '<br>' + dateNote, undoUrl);
}

function _siActionPage(kind, title, detail, undoUrl) {
  const accent = kind === 'success' ? '#6fcf97' : kind === 'error' ? '#eb5757' : '#4f8ef7';
  const mark = kind === 'success' ? '&#10003;' : kind === 'error' ? '&#33;' : '&#8505;';
  const appUrl = _getWebAppUrl();

  return HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Spendwise</title>
</head>
<body style="margin:0;background:#0a0d14;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:#e8ecf0;">
  <div style="max-width:420px;margin:0 auto;padding:64px 20px;text-align:center;">
    <div style="font-size:18px;font-weight:800;letter-spacing:-0.02em;margin-bottom:32px;">Spendwise</div>
    <div style="width:56px;height:56px;line-height:56px;border-radius:50%;margin:0 auto 18px;background:${accent}22;color:${accent};font-size:26px;font-weight:700;">${mark}</div>
    <h1 style="font-size:22px;font-weight:700;margin:0 0 10px;">${title}</h1>
    <p style="font-size:14px;color:#8892a4;line-height:1.6;margin:0 0 28px;">${detail}</p>
    ${undoUrl ? `<a href="${undoUrl}" style="display:inline-block;font-size:13px;color:#8892a4;text-decoration:underline;margin-bottom:20px;">Undo this</a><br>` : ''}
    ${appUrl ? `<a href="${appUrl}?page=standing" style="display:inline-block;margin-top:8px;background:#4f8ef7;color:#0a0d14;font-size:13px;font-weight:700;text-decoration:none;padding:11px 22px;border-radius:10px;">Open Spendwise</a>` : ''}
  </div>
</body>
</html>`)
    .setTitle('Spendwise')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// KEYWORD-BASED CATEGORIZATION ENGINE
// Sheet-driven: reads KeywordMap tab from Config sheet.
// No hardcoded keyword→category mappings.
// ============================================================

function _getKeywordMapSheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(KEYWORD_MAP_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(KEYWORD_MAP_TAB);
    sheet.getRange(1, 1, 1, 2).setValues([['Keyword', 'Category']]);
    sheet.getRange(1, 1, 1, 2).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Returns { keyword: category } map (lowercased keys). Cached.
function _getKeywordMap() {
  const cached = SCRIPT_CACHE.get('keyword_map');
  if (cached) { try { return JSON.parse(cached); } catch (e) { } }
  try {
    const sheet = _getKeywordMapSheet();
    if (sheet.getLastRow() <= 1) return {};
    const map = {};
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
      .filter(r => r[0])
      .forEach(r => { map[String(r[0]).toLowerCase().trim()] = String(r[1] || 'Other'); });
    SCRIPT_CACHE.put('keyword_map', JSON.stringify(map), TTL_KEYWORD_MAP);
    return map;
  } catch (e) { Logger.log('_getKeywordMap error: ' + e.message); return {}; }
}

// Looks up keyword in the map. Case-insensitive, supports partial match.
// Returns category name or 'Other'.
function categorizeByKeyword(keyword) {
  if (!keyword) return 'Other';
  const map = _getKeywordMap();
  const key = String(keyword).toLowerCase().trim();
  // Exact match first
  if (map[key]) return map[key];
  // Partial match: check if keyword contains or is contained by a map key
  for (const mapKey of Object.keys(map)) {
    if (key.includes(mapKey) || mapKey.includes(key)) return map[mapKey];
  }
  return 'Other';
}

// Returns all keyword mappings as an array of { keyword, category } objects.
function getKeywordMappings() {
  const map = _getKeywordMap();
  return Object.entries(map).map(([keyword, category]) => ({ keyword, category }));
}

// Saves keyword mappings. Replaces all rows.
function saveKeywordMappings(mappings) {
  try {
    const sheet = _getKeywordMapSheet();
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    if (mappings && mappings.length > 0) {
      sheet.getRange(2, 1, mappings.length, 2).setValues(
        mappings.map(m => [String(m.keyword || '').toLowerCase().trim(), m.category || 'Other'])
      );
    }
    SCRIPT_CACHE.remove('keyword_map');
    return { success: true, message: 'Keyword mappings saved.' };
  } catch (e) {
    Logger.log('saveKeywordMappings error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ============================================================
// HISTORY-BASED CATEGORIZATION (auto keyword learning)
// Learns category↔description associations from past expenses so
// new descriptions are auto-categorized even without a KeywordMap row.
// ============================================================

const _STOPWORDS = {
  the: 1, and: 1, for: 1, with: 1, from: 1, paid: 1, payment: 1, pay: 1,
  expense: 1, bought: 1, buy: 1, order: 1, bill: 1, rs: 1, inr: 1, via: 1, upi: 1
};

// Splits a description into meaningful lowercase tokens (>=3 chars,
// no stopwords, no pure numbers).
function _tokenizeDesc(s) {
  return String(s || '').toLowerCase().split(/[^a-z0-9]+/)
    .filter(t => t.length >= 3 && !_STOPWORDS[t] && !/^\d+$/.test(t));
}

// Returns the highest-count key in a { key: count } object, or null.
function _topVote(votes) {
  let best = null, bestN = 0;
  for (const k in votes) { if (votes[k] > bestN) { bestN = votes[k]; best = k; } }
  return best;
}

// Builds (and caches) an index from all historical expenses:
//   { full: { "<description>": {cat: n} }, tokens: { "<token>": {cat: n} } }
// Expenses categorized as 'Other' are ignored so they don't pollute votes.
function _getHistoryCategoryIndex() {
  const cached = SCRIPT_CACHE.get('hist_cat_index');
  if (cached) { try { return JSON.parse(cached); } catch (e) { } }
  const full = {}, tokens = {};
  try {
    const shardIds = _getAllShardRecords().map(s => s.id);
    if (!shardIds.includes(getActiveShardId())) shardIds.push(getActiveShardId());
    [...new Set(shardIds)].forEach(id => {
      let exps = [];
      try { exps = _readShardExpenses(id); } catch (e) { }
      exps.forEach(e => {
        const cat = String(e.category || '').trim();
        const desc = String(e.description || '').toLowerCase().trim();
        if (!cat || cat === 'Other' || !desc) return;
        full[desc] = full[desc] || {};
        full[desc][cat] = (full[desc][cat] || 0) + 1;
        _tokenizeDesc(desc).forEach(tok => {
          tokens[tok] = tokens[tok] || {};
          tokens[tok][cat] = (tokens[tok][cat] || 0) + 1;
        });
      });
    });
  } catch (e) { Logger.log('_getHistoryCategoryIndex error: ' + e.message); }
  const index = { full, tokens };
  try { SCRIPT_CACHE.put('hist_cat_index', JSON.stringify(index), 1800); } catch (e) { }
  return index;
}

// Categorizes a description by matching it against expense history.
// Exact full-description match wins; otherwise tokens vote. Returns a
// category name, or null when history offers no match.
function categorizeByHistory(description) {
  if (!description) return null;
  const index = _getHistoryCategoryIndex();
  const desc = String(description).toLowerCase().trim();
  if (index.full[desc]) return _topVote(index.full[desc]);
  const votes = {};
  _tokenizeDesc(desc).forEach(tok => {
    const m = index.tokens[tok];
    if (m) for (const cat in m) votes[cat] = (votes[cat] || 0) + m[cat];
  });
  return _topVote(votes);
}

// Appends a learned keyword→category row to the KeywordMap sheet so it
// becomes visible and editable in the Settings UI. No-ops if a row for
// this keyword already exists, or for empty/'Other' values.
function _learnKeyword(keyword, category) {
  try {
    const key = String(keyword || '').toLowerCase().trim();
    if (!key || !category || category === 'Other') return;
    if (_getKeywordMap()[key]) return; // already mapped exactly
    _getKeywordMapSheet().appendRow([key, category]);
    SCRIPT_CACHE.remove('keyword_map');
    logEvent('CATEGORIZE', 'LEARNED', key + ' → ' + category, '');
  } catch (e) { Logger.log('_learnKeyword error: ' + e.message); }
}

// ── One-time backfill ────────────────────────────────────────
// Run manually from the editor to seed the KeywordMap from ALL existing
// history in one pass (instead of learning incrementally per expense).
// For every distinct word in past descriptions, assigns the category it
// was used with most often — keeping only confident, unambiguous words.
//   minCount       min times a word must appear      (default 2)
//   minConfidence  min share for the winning category (default 0.6)
// Existing KeywordMap rows are never overwritten. Returns a report.
// Tip: backfillKeywordMap(1) is aggressive and also captures one-offs.
function backfillKeywordMap(minCount, minConfidence) {
  minCount = minCount || 2;
  minConfidence = minConfidence || 0.6;

  SCRIPT_CACHE.remove('hist_cat_index'); // force a fresh scan of history
  const index = _getHistoryCategoryIndex();
  const existing = _getKeywordMap();

  const toAdd = [];
  Object.keys(index.tokens).forEach(tok => {
    if (existing[tok]) return; // don't overwrite a manual/learned rule
    const votes = index.tokens[tok];
    let total = 0, top = null, topN = 0;
    for (const cat in votes) {
      total += votes[cat];
      if (votes[cat] > topN) { topN = votes[cat]; top = cat; }
    }
    if (total < minCount) return;
    if ((topN / total) < minConfidence) return;
    toAdd.push([tok, top]);
  });

  if (toAdd.length) {
    const sheet = _getKeywordMapSheet();
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
    SCRIPT_CACHE.remove('keyword_map');
    logEvent('CATEGORIZE', 'BACKFILL', 'Added ' + toAdd.length + ' keywords', '');
  }

  const report = 'Backfill complete — added ' + toAdd.length + ' keyword(s) ' +
    '(minCount=' + minCount + ', minConfidence=' + minConfidence + '). ' +
    'Existing rows and low-confidence words were skipped.';
  Logger.log(report);
  toAdd.forEach(([k, c]) => Logger.log('  + ' + k + ' → ' + c));
  return { added: toAdd.length, keywords: toAdd.map(([k, c]) => ({ keyword: k, category: c })), message: report };
}

// Best-effort categorization: explicit KeywordMap → learned history → 'Other'.
// When history resolves a category, the mapping is persisted to KeywordMap.
function smartCategorize(description) {
  if (!description) return 'Other';
  const mapped = categorizeByKeyword(description);
  if (mapped && mapped !== 'Other') return mapped;
  const learned = categorizeByHistory(description);
  if (learned) {
    _learnKeyword(description, learned);
    return learned;
  }
  return 'Other';
}

// ============================================================
// EMAIL SOURCES — CRUD for the EmailSources config tab
// ============================================================

function _getEmailSourcesSheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(EMAIL_SOURCES_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(EMAIL_SOURCES_TAB);
    sheet.getRange(1, 1, 1, 4).setValues([['Sender', 'SubjectPattern', 'ParserType', 'Enabled']]);
    sheet.getRange(1, 1, 1, 4).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getEmailSources() {
  try {
    const sheet = _getEmailSourcesSheet();
    if (sheet.getLastRow() <= 1) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
      .filter(r => r[0])
      .map(r => ({
        sender: String(r[0] || ''),
        subjectPattern: String(r[1] || ''),
        parserType: String(r[2] || 'generic'),
        enabled: r[3] === true || r[3] === 'TRUE' || r[3] === 'true'
      }));
  } catch (e) { Logger.log('getEmailSources error: ' + e.message); return []; }
}

function saveEmailSources(sources) {
  try {
    const sheet = _getEmailSourcesSheet();
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    if (sources && sources.length > 0) {
      sheet.getRange(2, 1, sources.length, 4).setValues(
        sources.map(s => [s.sender || '', s.subjectPattern || '', s.parserType || 'generic', s.enabled !== false])
      );
    }
    return { success: true, message: 'Email sources saved.' };
  } catch (e) {
    Logger.log('saveEmailSources error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ============================================================
// STRUCTURED LOGGING — writes to Logs tab in Config sheet
// ============================================================

function _getLogsSheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(LOGS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(LOGS_TAB);
    sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Module', 'Status', 'Message', 'ExternalId']]);
    sheet.getRange(1, 1, 1, 5).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Log an event. Module: 'EMAIL', 'CHAT', 'SUMMARY'. Status: 'SUCCESS', 'ERROR', 'SKIP', 'DUPLICATE'.
function logEvent(module, status, message, externalId) {
  try {
    const sheet = _getLogsSheet();
    sheet.appendRow([new Date(), module, status, message, externalId || '']);
    // Trim logs to last 1000 rows to prevent unbounded growth
    if (sheet.getLastRow() > 1001) {
      sheet.deleteRows(2, sheet.getLastRow() - 1001);
    }
  } catch (e) { Logger.log('logEvent error: ' + e.message); }
}

// Returns recent log entries for the UI.
function getRecentLogs(limit) {
  try {
    const sheet = _getLogsSheet();
    const count = Math.max(0, sheet.getLastRow() - 1);
    if (count === 0) return [];
    const n = Math.min(limit || 50, count);
    const startRow = sheet.getLastRow() - n + 1;
    const tz = Session.getScriptTimeZone();
    return sheet.getRange(startRow, 1, n, 5).getValues()
      .filter(r => r[0])
      .map(r => ({
        timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd HH:mm') : '',
        module: String(r[1] || ''),
        status: String(r[2] || ''),
        message: String(r[3] || ''),
        externalId: String(r[4] || '')
      }))
      .reverse(); // newest first
  } catch (e) { Logger.log('getRecentLogs error: ' + e.message); return []; }
}

// Check if an external ID already exists in any shard (for duplicate prevention).
// Scans Notes field for [EXT:externalId] tag.
function _isDuplicateExternal(externalId) {
  if (!externalId) return false;
  const tag = '[EXT:' + externalId + ']';
  const shardIds = _getAllShardRecords().map(s => s.id);
  const activeId = getActiveShardId();
  if (!shardIds.includes(activeId)) shardIds.unshift(activeId);
  for (const shardId of shardIds) {
    try {
      const sheet = _openSS(shardId).getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      // Notes column is column 7 (index 6 in 0-based)
      const notes = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues().flat();
      if (notes.some(n => String(n).includes(tag))) return true;
    } catch (e) { }
  }
  return false;
}

// Add an expense with source tagging. Used by Chat and Email handlers.
// Encodes source and externalId in the Notes field for backward compatibility.
function addExpenseWithSource(expense, source, externalId) {
  // Build notes with source tags
  let notes = expense.notes || '';
  if (source) notes += (notes ? ' · ' : '') + '[SOURCE:' + source + ']';
  if (externalId) notes += ' [EXT:' + externalId + ']';

  return addExpense({
    date: expense.date,
    category: expense.category,
    description: expense.description,
    amount: expense.amount,
    paymentMethod: expense.paymentMethod || 'Auto',
    notes: notes
  });
}


// ── Cache ────────────────────────────────────────────────────
// Called after any expense or category write.
// Income cache is managed separately by _invalidateIncomeCache().
function invalidateCache() {
  SCRIPT_CACHE.removeAll([
    'categories', 'shard_registry', 'settings', 'si_all',
    'income_{}', 'keyword_map', 'hist_cat_index',
    'analytics_week', 'analytics_month', 'analytics_current_month',
    'analytics_quarter', 'analytics_half', 'analytics_current_year', 'analytics_year',
    'monthly_comparison_quarter', 'monthly_comparison_half',
    'monthly_comparison_year', 'monthly_comparison_current_month', 'monthly_comparison_current_year'
  ]);
}

// Full cache wipe + Script Properties reset. Called from REPAIR in AdminOps / Settings UI.
// Expense cache uses dynamic keys (exp_*) — only the two common patterns are known statically.
// After this, ACTIVE_SHARD_ID re-derives from ShardRegistry on next request.
function purgeAllCache() {
  invalidateCache();
  SCRIPT_CACHE.removeAll(['exp_{}', 'exp_{"startDate":null,"endDate":null}']);
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('ACTIVE_SHARD_ID');
    props.deleteProperty('OLDEST_SHARD_ID');
  } catch (e) { }
  Logger.log('All caches purged.');
  return { success: true };
}

// ── System admin (UI-callable) ───────────────────────────────
// These wrappers are called from the Settings page via google.script.run.
// They return { ok/success, lines[] } for display in the output panel.
// Full implementations live in AdminOps.gs for editor use.

function getSystemStatus() {
  const props = PropertiesService.getScriptProperties();
  const configId = props.getProperty('CONFIG_SHEET_ID');
  const activeId = props.getProperty('ACTIVE_SHARD_ID');
  const lines = [];
  let allOk = true;

  // Version
  lines.push({ type: 'heading', text: 'Spendwise v' + SPENDWISE_VERSION });
  lines.push({ type: 'heading', text: 'Script Properties' });
  lines.push({
    type: configId ? 'ok' : 'error',
    text: 'CONFIG_SHEET_ID: ' + (configId ? configId.substring(0, 20) + '...' : 'NOT SET — run SETUP')
  });
  lines.push({
    type: activeId ? 'ok' : 'warn',
    text: 'ACTIVE_SHARD_ID: ' + (activeId ? activeId.substring(0, 20) + '...' : 'Not set — run RESET')
  });
  if (!configId) allOk = false;

  // Config sheet tabs
  lines.push({ type: 'heading', text: 'Config Sheet' });
  if (!configId) {
    lines.push({ type: 'error', text: 'Cannot check — CONFIG_SHEET_ID not set' });
    allOk = false;
  } else {
    try {
      const ss = SpreadsheetApp.openById(configId);
      const tabs = ss.getSheets().map(s => s.getName());
      lines.push({ type: 'ok', text: 'Accessible: ' + ss.getName() });
      ['Categories', 'ShardRegistry', 'Settings', 'Income', 'StandingInstructions',
       'KeywordMap', 'EmailSources', 'Logs'].forEach(t => {
        const ok = tabs.includes(t);
        if (!ok) allOk = false;
        lines.push({ type: ok ? 'ok' : 'error', text: t + ' tab: ' + (ok ? 'present' : 'MISSING') });
      });
    } catch (e) {
      lines.push({ type: 'error', text: 'Cannot open: ' + e.message });
      allOk = false;
    }
  }

  // Shard registry
  lines.push({ type: 'heading', text: 'Shards' });
  try {
    const records = _getAllShardRecords();
    if (records.length === 0) {
      lines.push({ type: 'warn', text: 'No shards in registry' });
    } else {
      records.forEach(r => {
        let ok = false;
        try { SpreadsheetApp.openById(r.id); ok = true; } catch (e) { }
        if (!ok) allOk = false;
        lines.push({
          type: ok ? 'ok' : 'error',
          text: r.label + ': ' + (ok ? 'accessible' : 'INACCESSIBLE')
        });
      });
    }
  } catch (e) {
    lines.push({ type: 'error', text: 'Registry read failed: ' + e.message });
    allOk = false;
  }

  // Triggers
  lines.push({ type: 'heading', text: 'Triggers' });
  const triggers = ScriptApp.getProjectTriggers();
  const hasRotation = triggers.some(t => t.getHandlerFunction() === 'rotateShardForNewMonth');
  const hasWeeklyRpt = triggers.some(t => t.getHandlerFunction() === 'weeklyReportTrigger');
  const hasSI = triggers.some(t => t.getHandlerFunction() === 'processStandingInstructions');
  lines.push({
    type: hasRotation ? 'ok' : 'warn',
    text: 'Monthly shard rotation: ' + (hasRotation ? 'active' : 'NOT installed')
  });
  lines.push({
    type: hasSI ? 'ok' : 'warn',
    text: 'Daily SI auto-log: ' + (hasSI ? 'active (07:00)' : 'NOT installed — run REPAIR')
  });
  const rptSettings = getSettings();
  if (rptSettings.weeklyReportEnabled === 'true') {
    lines.push({
      type: hasWeeklyRpt ? 'ok' : 'warn',
      text: 'Weekly email report: ' + (hasWeeklyRpt ? 'active (' + rptSettings.weeklyReportDay + ' ' + rptSettings.weeklyReportTime + ':00)' : 'ENABLED but no trigger — re-save report settings')
    });
  } else {
    lines.push({ type: 'info', text: 'Weekly email report: not enabled' });
  }

  // Recurring reminder
  const hasReminder = triggers.some(t => t.getHandlerFunction() === 'recurringReminderTrigger');
  if (rptSettings.recurringReminderEnabled === 'true') {
    lines.push({
      type: hasReminder ? 'ok' : 'warn',
      text: 'Recurring reminder: ' + (hasReminder
        ? 'active (' + (rptSettings.recurringReminderDaysBefore || '1') + 'd before, ' + (rptSettings.recurringReminderTime || '9') + ':00)'
        : 'ENABLED but no trigger — re-save reminder settings')
    });
    if (!_getWebAppUrl()) {
      lines.push({ type: 'warn', text: 'Reminder Log buttons inactive — deploy as a web app, then open Spendwise once' });
    }
  } else {
    lines.push({ type: 'info', text: 'Recurring reminder: not enabled' });
  }

  // Email ingestion
  const hasEmailIngestion = triggers.some(t => t.getHandlerFunction() === 'processEmailReceipts');
  if (rptSettings.emailIngestionEnabled === 'true') {
    lines.push({
      type: hasEmailIngestion ? 'ok' : 'warn',
      text: 'Email ingestion: ' + (hasEmailIngestion ? 'active' : 'ENABLED but no trigger — run REPAIR')
    });
  } else {
    lines.push({ type: 'info', text: 'Email ingestion: not enabled' });
  }

  // Daily summary
  const hasDailySummary = triggers.some(t => t.getHandlerFunction() === 'sendDailySummary');
  if (rptSettings.dailySummaryEnabled === 'true') {
    lines.push({
      type: hasDailySummary ? 'ok' : 'warn',
      text: 'Daily Chat summary: ' + (hasDailySummary ? 'active (' + (rptSettings.dailySummaryTime || '21') + ':00)' : 'ENABLED but no trigger — run REPAIR')
    });
  } else {
    lines.push({ type: 'info', text: 'Daily Chat summary: not enabled' });
  }

  // Chat integration
  if (rptSettings.chatEnabled === 'true') {
    lines.push({
      type: 'ok',
      text: 'Chat integration: enabled' + (rptSettings.chatSpaceId ? ' (space linked)' : ' (no space — message the bot)')
    });
  } else {
    lines.push({ type: 'info', text: 'Chat integration: not enabled' });
  }

  // Cache warmth
  lines.push({ type: 'heading', text: 'Cache' });
  const cacheKeys = ['categories', 'shard_registry', 'settings', 'analytics_current_month'];
  cacheKeys.forEach(k => {
    const warm = !!CacheService.getScriptCache().get(k);
    lines.push({ type: 'info', text: k + ': ' + (warm ? '● warm' : '○ cold') });
  });

  return { ok: allOk, isSetUp: !!configId, lines };
}

// Runs SETUP() — only works if not already configured
function runSetup() {
  const configId = PropertiesService.getScriptProperties().getProperty('CONFIG_SHEET_ID');
  if (configId) {
    return {
      success: false, alreadySetUp: true,
      lines: [{ type: 'warn', text: 'Already configured. CONFIG_SHEET_ID exists: ' + configId }]
    };
  }
  try {
    const result = SETUP();
    const lines = [];
    if (result.success) {
      lines.push({ type: 'ok', text: 'Config sheet created: ' + result.configUrl });
      lines.push({ type: 'ok', text: 'First shard created (' + result.month + '): ' + result.shardUrl });
      lines.push({ type: 'ok', text: 'All tabs seeded, trigger installed' });
      lines.push({ type: 'info', text: 'Next: Deploy → New Deployment → Web App' });
    } else {
      lines.push({ type: 'error', text: 'Setup failed: ' + (result.error || 'Unknown error') });
    }
    return { success: !!result.success, lines };
  } catch (e) {
    return { success: false, lines: [{ type: 'error', text: e.message }] };
  }
}

// runRepair — merges cache clearing + recovery into one safe operation.
// Runs all recovery steps: cache clear, shard re-derive, tab
// verification, trigger reinstall.
// No data is modified or deleted. Safe to run anytime.
function runRepair() {
  try {
    const lines = [];
    const props = PropertiesService.getScriptProperties();

    // 1. Purge all caches
    purgeAllCache();
    lines.push({ type: 'ok', text: 'All caches cleared' });

    // 2. Re-derive ACTIVE_SHARD_ID from registry
    try {
      const records = _getAllShardRecords();
      if (records.length > 0) {
        const newest = records[0];
        props.setProperty('ACTIVE_SHARD_ID', newest.id);
        lines.push({ type: 'ok', text: 'ACTIVE_SHARD_ID re-derived: ' + newest.label });
      } else {
        lines.push({ type: 'warn', text: 'No shard records found — run fixShardRegistry() from editor' });
      }
    } catch (e) { lines.push({ type: 'error', text: 'Registry read failed: ' + e.message }); }

    // 3. Verify + create missing Config tabs
    try {
      const activeId = props.getProperty('ACTIVE_SHARD_ID');
      _initConfigSheet(activeId);
      lines.push({ type: 'ok', text: 'Config sheet tabs verified' });
    } catch (e) { lines.push({ type: 'warn', text: 'Config tab check: ' + e.message }); }

    // 4. Verify + create missing Expenses tab on all shards
    try {
      const records = _getAllShardRecords();
      records.forEach(r => {
        try { _ensureShardSheet(SpreadsheetApp.openById(r.id)); }
        catch (e) { lines.push({ type: 'warn', text: 'Could not verify ' + r.label + ': ' + e.message }); }
      });
      lines.push({ type: 'ok', text: 'All ' + records.length + ' shard(s) Expenses tab verified' });
    } catch (e) { lines.push({ type: 'warn', text: 'Shard tab check: ' + e.message }); }

    // 5. Reinstall triggers if missing
    try {
      const hasTrigger = ScriptApp.getProjectTriggers()
        .some(t => t.getHandlerFunction() === 'rotateShardForNewMonth');
      if (!hasTrigger) {
        ScriptApp.newTrigger('rotateShardForNewMonth').timeBased().onMonthDay(1).atHour(0).create();
        lines.push({ type: 'ok', text: 'Monthly rotation trigger reinstalled' });
      } else {
        lines.push({ type: 'ok', text: 'Monthly rotation trigger: already present' });
      }
    } catch (e) { lines.push({ type: 'warn', text: 'Trigger check: ' + e.message }); }

    try {
      const hasSI = ScriptApp.getProjectTriggers()
        .some(t => t.getHandlerFunction() === 'processStandingInstructions');
      if (!hasSI) {
        ScriptApp.newTrigger('processStandingInstructions').timeBased().everyDays(1).atHour(7).create();
        lines.push({ type: 'ok', text: 'Daily SI auto-log trigger reinstalled' });
      } else {
        lines.push({ type: 'ok', text: 'Daily SI auto-log trigger: present' });
      }
    } catch (e) { lines.push({ type: 'warn', text: 'SI trigger check: ' + e.message }); }

    try {
      const settings = getSettings();
      const hasReminder = ScriptApp.getProjectTriggers()
        .some(t => t.getHandlerFunction() === 'recurringReminderTrigger');
      if (settings.recurringReminderEnabled === 'true' && !hasReminder) {
        scheduleRecurringReminder();
        lines.push({ type: 'ok', text: 'Recurring reminder trigger reinstalled' });
      } else {
        lines.push({
          type: 'ok',
          text: 'Recurring reminder trigger: ' + (hasReminder ? 'present' : 'not enabled')
        });
      }
    } catch (e) { lines.push({ type: 'warn', text: 'Reminder trigger check: ' + e.message }); }

    lines.push({ type: 'info', text: 'No data was modified or deleted.' });
    lines.push({ type: 'info', text: 'Run STATUS to verify everything looks healthy.' });
    return { success: true, lines };
  } catch (e) {
    return { success: false, lines: [{ type: 'error', text: e.message }] };
  }
}