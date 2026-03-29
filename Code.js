// ============================================================
// SPENDWISE — Code.gs  v1.1.0
// Runtime backend only. All setup/admin logic lives in AdminOps.gs.
//
// FILE MAP:
//   Code.gs             — this file: all server-side functions
//   AdminOps.gs         — SETUP, REPAIR, STATUS + import/diagnostic tools
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

const SPENDWISE_VERSION = '1.1.0';

const SHEET_NAME = 'Expenses';
const CATEGORIES_TAB = 'Categories';
const SHARD_REGISTRY = 'ShardRegistry';
const SETTINGS_TAB = 'Settings';
const INCOME_TAB = 'Income';     // stored in Config sheet; not sharded
const SI_TAB = 'StandingInstructions'; // stored in Config sheet; not sharded

const TTL_CATEGORIES = 3600;
const TTL_ANALYTICS = 300;
const TTL_SHARD_REG = 7200;
const TTL_EXPENSES = 300;
const TTL_SETTINGS = 3600;
const TTL_SI = 3600; // standing instructions — changes infrequently

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
};

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
  }
  if (!ss.getSheetByName(INCOME_TAB)) {
    const sheet = ss.insertSheet(INCOME_TAB);
    const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Notes', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName(SI_TAB)) {
    const sheet = ss.insertSheet(SI_TAB);
    const headers = ['ID', 'Name', 'Category', 'Amount', 'Frequency', 'DayOfMonth', 'DayOfWeek',
      'StartDate', 'EndDate', 'PaymentMethod', 'Notes', 'AutoLog', 'IsActive', 'LastLoggedDate'];
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
  if (period === 'week') d.setDate(now.getDate() - 7);
  else if (period === 'month') d.setDate(now.getDate() - 30);
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
    topCategories: [], trends: [], shardCount: 0
  };
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

    const trends = [];
    for (let i = 29; i >= 0; i--) {
      const d = new Date(); d.setDate(d.getDate() - i);
      const k = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      trends.push({ date: k, amount: byDay[k] || 0 });
    }

    const result = {
      totalSpent: Math.round(totalSpent * 100) / 100, totalTransactions,
      avgPerDay: Math.round((totalSpent / 30) * 100) / 100,
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
      percentage: cat.budget > 0 ? Math.min(100, Math.round((spent / cat.budget) * 100)) : 0
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
      shardCount: a.shardCount || 1,
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
      topCategory: (a.topCategories && a.topCategories[0]) ? a.topCategories[0] : null,
      lastUpdated: new Date().toISOString()
    };
  } catch (e) {
    Logger.log('getDashboardStats error: ' + e.message);
    return { totalSpent: 0, totalTransactions: 0, avgPerDay: 0, avgPerTransaction: 0, topCategory: null, lastUpdated: new Date().toISOString() };
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
  'StartDate', 'EndDate', 'PaymentMethod', 'Notes', 'AutoLog', 'IsActive', 'LastLoggedDate', 'CustomIntervalDays'];

function _getSISheet() {
  const ss = getConfigSS();
  let sheet = ss.getSheetByName(SI_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(SI_TAB);
    sheet.getRange(1, 1, 1, SI_COLUMNS.length).setValues([SI_COLUMNS]);
    sheet.getRange(1, 1, 1, SI_COLUMNS.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    sheet.setFrozenRows(1);
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
  };
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
        // Write CustomIntervalDays (column 15) separately
        sheet.getRange(i + 1, 15).setValue(parseInt(si.customIntervalDays) || 0);
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

// Manually log a single SI as an expense entry now.
// Called from the UI quick-log button.
function logStandingInstruction(id) {
  try {
    const all = getStandingInstructions();
    const si = all.find(s => s.id === id);
    if (!si) return { success: false, message: 'Standing instruction not found.' };
    if (!si.isActive) return { success: false, message: 'This instruction is paused.' };

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const result = addExpense({
      date: today,
      category: si.category,
      description: si.name,
      amount: si.amount,
      paymentMethod: si.paymentMethod,
      notes: si.notes + (si.notes ? ' · ' : '') + '[Standing: ' + si.id + ']'
    });

    if (result.success) {
      // Update LastLoggedDate
      const sheet = _getSISheet();
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === id) {
          sheet.getRange(i + 1, 14).setValue(today); // LastLoggedDate column
          break;
        }
      }
      SCRIPT_CACHE.remove('si_all');
    }
    return result;
  } catch (e) {
    Logger.log('logStandingInstruction error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Returns SIs due in the current calendar month with their logged status.
// Used by the Recurring page Zone 1 (upcoming this month).
function getUpcomingSIs() {
  try {
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const month = Utilities.formatDate(now, tz, 'yyyy-MM');
    const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const allSIs = getStandingInstructions().filter(si => si.isActive);
    const upcoming = [];

    allSIs.forEach(si => {
      const dueDates = _getDueDatesInMonth(si, month, tz);
      dueDates.forEach(dueDate => {
        const isLogged = si.lastLoggedDate && si.lastLoggedDate.substring(0, 7) === month;
        const daysDiff = Math.ceil((new Date(dueDate) - now) / 86400000);
        const status = isLogged ? 'logged'
          : dueDate < today ? 'overdue'
            : daysDiff <= 3 ? 'due_soon'
              : 'upcoming';
        upcoming.push({ ...si, dueDate, status, daysDiff });
      });
    });

    upcoming.sort((a, b) => a.dueDate.localeCompare(b.dueDate));
    return upcoming;
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
// Processes all active SIs and logs those due today with AutoLog=true.
function processStandingInstructions() {
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const month = today.substring(0, 7);
  const allSIs = getStandingInstructions().filter(si => si.isActive && si.autoLog);
  let logged = 0, skipped = 0;

  allSIs.forEach(si => {
    try {
      if (si.lastLoggedDate && si.lastLoggedDate.substring(0, 7) === month) { skipped++; return; }
      if (si.startDate && today < si.startDate) { skipped++; return; }
      if (si.endDate && today > si.endDate) { skipped++; return; }

      const dueDates = _getDueDatesInMonth(si, month, tz);
      const isDueToday = dueDates.some(d => d === today);
      if (!isDueToday) return;

      const result = logStandingInstruction(si.id);
      if (result.success) { logged++; Logger.log('Auto-logged: ' + si.name); }
    } catch (e) {
      Logger.log('processStandingInstructions error for ' + si.id + ': ' + e.message);
    }
  });

  Logger.log('processStandingInstructions: logged=' + logged + ' skipped=' + skipped);
  return { logged, skipped };
}

// ── SI helpers ───────────────────────────────────────────────

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
    // Due only if current month is in the correct quarter cycle from start date
    if (si.startDate) {
      const startMonth = parseInt(si.startDate.substring(5, 7)) - 1;
      if ((mIndex - startMonth + 12) % 3 === 0) {
        const day = si.dayOfMonth || 1;
        const maxDay = new Date(year, mIndex + 1, 0).getDate();
        const d = new Date(year, mIndex, Math.min(day, maxDay));
        dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
      }
    }

  } else if (si.frequency === 'yearly') {
    if (si.startDate) {
      const startM = parseInt(si.startDate.substring(5, 7)) - 1;
      const startD = parseInt(si.startDate.substring(8, 10));
      if (mIndex === startM) {
        const d = new Date(year, mIndex, startD);
        dates.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
      }
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

// Checks if MailApp permission has been granted by trying a no-op
// call that requires the same scope but sends nothing.
// Returns { granted: true } or { granted: false, message, instructions }
function checkMailPermission() {
  try {
    MailApp.getRemainingDailyQuota(); // requires gmail.send scope, sends nothing
    return { granted: true };
  } catch (e) {
    return {
      granted: false,
      message: e.message,
      instructions: [
        '1. Open the Apps Script editor for this project',
        '2. Select any function in the dropdown (e.g. STATUS)',
        '3. Click Run — Google will show an authorization dialog',
        '4. Click "Review permissions" → choose your account → Allow',
        '5. Return here and save report settings again'
      ]
    };
  }
}

// Builds and sends the weekly report. Call from editor to test.
function sendWeeklyReport(toEmail) {
  try {
    const tz = Session.getScriptTimeZone();
    const now = new Date();

    // This week: Mon–today
    const thisWeekStart = new Date(now);
    thisWeekStart.setDate(now.getDate() - now.getDay() + (now.getDay() === 0 ? -6 : 1));
    thisWeekStart.setHours(0, 0, 0, 0);

    // Last week: same 7-day window shifted back
    const lastWeekStart = new Date(thisWeekStart);
    lastWeekStart.setDate(thisWeekStart.getDate() - 7);
    const lastWeekEnd = new Date(thisWeekStart);
    lastWeekEnd.setDate(thisWeekStart.getDate() - 1);

    const fmt2 = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const thisWeekExpenses = getExpenses({ startDate: fmt2(thisWeekStart), endDate: fmt2(now) });
    const lastWeekExpenses = getExpenses({ startDate: fmt2(lastWeekStart), endDate: fmt2(lastWeekEnd) });
    const thisWeekIncome = getIncome({ startDate: fmt2(thisWeekStart), endDate: fmt2(now) });

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
    const changeStr = pctChange === null ? '' :
      pctChange > 0 ? `▲ ${pctChange}% vs last week` : `▼ ${Math.abs(pctChange)}% vs last week`;
    const changeColor = pctChange > 0 ? '#eb5757' : '#6fcf97';

    const weekLabel = Utilities.formatDate(thisWeekStart, tz, 'MMM d') +
      ' – ' + Utilities.formatDate(now, tz, 'MMM d, yyyy');

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

    // Upcoming standing instructions — due in the next 7 days, not yet logged
    const tz7 = Session.getScriptTimeZone();
    const in7Days = new Date(now); in7Days.setDate(now.getDate() + 7);
    const in7Str = Utilities.formatDate(in7Days, tz7, 'yyyy-MM-dd');
    const todayStr2 = Utilities.formatDate(now, tz7, 'yyyy-MM-dd');
    const upcomingSIs = getUpcomingSIs().filter(si =>
      si.dueDate >= todayStr2 && si.dueDate <= in7Str && si.status !== 'logged'
    );

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
      `<span style="display:inline-block;margin-left:10px;font-size:12px;font-weight:700;padding:3px 10px;border-radius:99px;background:${pctChange > 0 ? 'rgba(235,87,87,0.15)' : 'rgba(111,207,151,0.15)'};color:${changeColor};">${changeStr}</span>` : '';

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

// ── Cache ────────────────────────────────────────────────────
// Called after any expense or category write.
// Income cache is managed separately by _invalidateIncomeCache().
function invalidateCache() {
  SCRIPT_CACHE.removeAll([
    'categories', 'shard_registry', 'settings', 'si_all',
    'income_{}',
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
      ['Categories', 'ShardRegistry', 'Settings', 'Income', 'StandingInstructions'].forEach(t => {
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

    lines.push({ type: 'info', text: 'No data was modified or deleted.' });
    lines.push({ type: 'info', text: 'Run STATUS to verify everything looks healthy.' });
    return { success: true, lines };
  } catch (e) {
    return { success: false, lines: [{ type: 'error', text: e.message }] };
  }
}