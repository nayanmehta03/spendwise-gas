// ============================================================
// SPENDWISE — AdminOps.gs  v1.0.0
// Editor-only tools: setup, diagnostics, recovery, import.
// None of these functions are called by the web app at runtime.
//
// QUICK REFERENCE:
//   SETUP()                    — first-time only, creates all sheets
//   STATUS()                   — health check for all components
//   RESET()                    — cache + registry recovery
//   UPDATE()                   — schema migrations after a code update
//   importFromCSV()            — bulk import historical expense data
//   forceReimportFromCSV()     — wipe existing rows and reimport
//   fixShardRegistry()         — reorder + validate shard registry
//   resetCategoriesToDefaults() — overwrite categories with defaults
//   runShardExperiment()       — benchmark shard fan-out performance
// ============================================================


// ── SETUP ────────────────────────────────────────────────────
// Run once after copying all files into a new Apps Script project.
// Creates Config Sheet + first Shard Sheet, seeds all tabs,
// installs rotation trigger, stores all IDs in Script Properties.
function SETUP() {
  const log = [];
  const props = PropertiesService.getScriptProperties();

  log.push('=== SPENDWISE SETUP ===');
  log.push('Starting at ' + new Date().toLocaleString());
  log.push('');

  const existingConfigId = props.getProperty('CONFIG_SHEET_ID');
  if (existingConfigId) {
    log.push('⚠ CONFIG_SHEET_ID already exists: ' + existingConfigId);
    log.push('  SETUP has already been run.');
    log.push('  To repair without data loss: run STATUS() then RESET().');
    log.push('  To start over: delete Script Properties + both sheets, then re-run SETUP().');
    Logger.log(log.join('\n'));
    return { alreadySetUp: true, configId: existingConfigId };
  }

  // Create Config Sheet
  let configSS;
  try {
    configSS = SpreadsheetApp.create('Spendwise — Config');
    props.setProperty('CONFIG_SHEET_ID', configSS.getId());
    log.push('✓ Config sheet: ' + configSS.getUrl());
  } catch (e) {
    log.push('✗ Failed to create Config sheet: ' + e.message);
    Logger.log(log.join('\n'));
    return { success: false, error: e.message };
  }

  // Create first Shard Sheet
  let shardSS;
  const tz = Session.getScriptTimeZone();
  const month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  try {
    shardSS = SpreadsheetApp.create('Spendwise — Expenses_Shard_' + month);
    const shardId = shardSS.getId();
    props.setProperty('ACTIVE_SHARD_ID', shardId);
    log.push('✓ First shard (' + month + '): ' + shardSS.getUrl());
    _initConfigSheet(shardId);
    _ensureShardSheet(shardSS);
    log.push('✓ All tabs seeded (Categories, ShardRegistry, Settings, Income, Expenses)');
  } catch (e) {
    log.push('✗ Failed to create shard sheet: ' + e.message);
    Logger.log(log.join('\n'));
    return { success: false, error: e.message };
  }

  // Install monthly rotation trigger
  try {
    _installRotationTrigger();
    log.push('✓ Monthly rotation trigger installed');
  } catch (e) {
    log.push('⚠ Trigger install failed: ' + e.message + ' — run setupMonthlyRotationTrigger() from editor');
  }

  // Install daily SI auto-log trigger
  try {
    _installSITrigger();
    log.push('✓ Daily standing instructions trigger installed (07:00)');
  } catch (e) {
    log.push('⚠ SI trigger install failed: ' + e.message);
  }

  // Smoke test
  try {
    const cats = getCategories();
    log.push('✓ Categories readable: ' + cats.length + ' categories');
  } catch (e) {
    log.push('⚠ Category read test failed: ' + e.message);
  }
  try {
    const settings = getSettings();
    log.push('✓ Settings readable: defaultPayment=' + settings.defaultPayment);
  } catch (e) {
    log.push('⚠ Settings read test failed: ' + e.message);
  }

  log.push('');
  log.push('=== SETUP COMPLETE ===');
  log.push('Config Sheet:  ' + configSS.getUrl());
  log.push('First Shard:   ' + shardSS.getUrl());
  log.push('');
  log.push('Next step: Deploy → New Deployment → Web App');
  log.push('  Execute as: Me');
  log.push('  Who has access: Anyone');

  Logger.log(log.join('\n'));
  return {
    success: true,
    configId: configSS.getId(),
    shardId: shardSS.getId(),
    month,
    configUrl: configSS.getUrl(),
    shardUrl: shardSS.getUrl()
  };
}


// ── STATUS ───────────────────────────────────────────────────
// Full health check: Script Properties, Config Sheet tabs,
// shard accessibility, triggers, and cache warmth.
function STATUS() {
  const log = [];
  const props = PropertiesService.getScriptProperties();

  log.push('=== SPENDWISE STATUS (v' + SPENDWISE_VERSION + ') ===');
  log.push('Checked at ' + new Date().toLocaleString());
  log.push('');

  // ── Script Properties ──────────────────────────────────────
  const configId = props.getProperty('CONFIG_SHEET_ID');
  const activeId = props.getProperty('ACTIVE_SHARD_ID');
  const oldestId = props.getProperty('OLDEST_SHARD_ID');
  log.push('SCRIPT PROPERTIES');
  log.push('  CONFIG_SHEET_ID  : ' + (configId || '✗ NOT SET — run SETUP()'));
  log.push('  ACTIVE_SHARD_ID  : ' + (activeId || '✗ NOT SET — run RESET()'));
  log.push('  OLDEST_SHARD_ID  : ' + (oldestId || '(not set — optional)'));

  if (!configId) {
    log.push('');
    log.push('✗ Spendwise is not set up. Run SETUP() first.');
    Logger.log(log.join('\n'));
    return { setupRequired: true };
  }

  // ── Config Sheet ───────────────────────────────────────────
  log.push('');
  log.push('CONFIG SHEET (' + configId + ')');
  try {
    const ss = SpreadsheetApp.openById(configId);
    const tabs = ss.getSheets().map(s => s.getName());
    log.push('  ✓ Accessible: ' + ss.getName());
    log.push('  Tabs found   : ' + tabs.join(', '));
    ['Categories', 'ShardRegistry', 'Settings', 'Income'].forEach(t => {
      log.push('  ' + (tabs.includes(t) ? '✓' : '✗ MISSING') + ' ' + t + ' tab');
    });
  } catch (e) {
    log.push('  ✗ CANNOT OPEN: ' + e.message);
  }

  // ── Categories ─────────────────────────────────────────────
  log.push('');
  log.push('CATEGORIES');
  try {
    const cats = getCategories();
    log.push('  ✓ ' + cats.length + ' categories loaded');
  } catch (e) {
    log.push('  ✗ ' + e.message);
  }

  // ── Settings ───────────────────────────────────────────────
  log.push('');
  log.push('SETTINGS');
  try {
    const s = getSettings();
    log.push('  ✓ currency=' + s.currency + ' defaultPayment=' + s.defaultPayment + ' weekStart=' + s.weekStartDay);
  } catch (e) {
    log.push('  ✗ ' + e.message);
  }

  // ── Shard Registry ─────────────────────────────────────────
  log.push('');
  log.push('SHARD REGISTRY');
  try {
    const records = _getAllShardRecords();
    log.push('  ✓ ' + records.length + ' shard(s) registered');
    records.forEach((r, i) => {
      let accessible = false;
      try { SpreadsheetApp.openById(r.id); accessible = true; } catch (e) { }
      log.push('  ' + (i + 1) + '. ' + r.label + ' (' + r.month + ') — ' + (accessible ? '✓ accessible' : '✗ INACCESSIBLE'));
    });
  } catch (e) {
    log.push('  ✗ ' + e.message);
  }

  // ── Active Shard ───────────────────────────────────────────
  log.push('');
  log.push('ACTIVE SHARD (' + (activeId || 'NOT SET') + ')');
  if (activeId) {
    try {
      const ss = SpreadsheetApp.openById(activeId);
      const sheet = ss.getSheetByName('Expenses');
      const rowCount = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
      log.push('  ✓ ' + ss.getName() + ' — ' + rowCount + ' expense row(s)');
    } catch (e) {
      log.push('  ✗ CANNOT OPEN: ' + e.message + ' — run RESET()');
    }
  }

  // ── Income Tab ─────────────────────────────────────────────
  log.push('');
  log.push('INCOME TAB');
  try {
    const sheet = _getIncomeSheet();
    const count = Math.max(0, sheet.getLastRow() - 1);
    log.push('  ✓ ' + count + ' income row(s)');
  } catch (e) {
    log.push('  ✗ ' + e.message);
  }

  // ── Triggers ───────────────────────────────────────────────
  log.push('');
  log.push('TRIGGERS');
  const triggers = ScriptApp.getProjectTriggers();
  const storedTrigId = PropertiesService.getScriptProperties().getProperty('WEEKLY_REPORT_TRIGGER_ID');
  const hasRotation = triggers.some(t => t.getHandlerFunction() === 'rotateShardForNewMonth');
  const hasWeeklyRpt = triggers.some(t => t.getHandlerFunction() === 'weeklyReportTrigger');
  const hasSI = triggers.some(t => t.getHandlerFunction() === 'processStandingInstructions');
  if (triggers.length === 0) {
    log.push('  ⚠ No triggers found — run SETUP() or setupMonthlyRotationTrigger()');
  } else {
    log.push('  ' + (hasRotation ? '✓' : '✗') + ' Monthly shard rotation');
    log.push('  ' + (hasSI ? '✓' : '⚠') + ' Daily standing instructions auto-log' + (hasSI ? '' : ' — run UPDATE() to install'));
    if (hasWeeklyRpt) {
      const wt = triggers.find(t => t.getHandlerFunction() === 'weeklyReportTrigger');
      log.push('  ✓ Weekly email report (trigger id: ' + wt.getUniqueId().substring(0, 12) + '...)');
      if (storedTrigId && wt.getUniqueId() !== storedTrigId) {
        log.push('  ⚠ Stored trigger ID mismatch — re-save report settings to resync');
      }
    } else {
      const settings = getSettings();
      if (settings.weeklyReportEnabled === 'true') {
        log.push('  ⚠ Weekly report is enabled in settings but no trigger found — re-save report settings');
      } else {
        log.push('  ○ Weekly email report (not enabled)');
      }
    }
  }

  // ── Cache ──────────────────────────────────────────────────
  log.push('');
  log.push('CACHE');
  const cacheKeys = ['categories', 'shard_registry', 'settings',
    'analytics_current_month', 'analytics_week', 'analytics_quarter'];
  cacheKeys.forEach(k => {
    const hit = CacheService.getScriptCache().get(k);
    log.push('  ' + k + ': ' + (hit ? '● warm' : '○ cold'));
  });

  log.push('');
  log.push('=== END STATUS ===');
  Logger.log(log.join('\n'));
  return { checked: true };
}


// ── RESET ────────────────────────────────────────────────────
// Clears all caches, re-derives ACTIVE_SHARD_ID from registry,
// verifies Config and shard tabs exist, reinstalls trigger if missing.
// Does not delete or modify any expense or income data.
function RESET() {
  const log = [];
  const props = PropertiesService.getScriptProperties();

  log.push('=== SPENDWISE RESET ===');

  purgeAllCache();
  log.push('✓ All caches cleared');

  try {
    const records = _getAllShardRecords();
    if (records.length > 0) {
      const newest = records[0]; // sorted newest-first by _getAllShardRecords
      props.setProperty('ACTIVE_SHARD_ID', newest.id);
      log.push('✓ ACTIVE_SHARD_ID re-derived: ' + newest.label + ' (' + newest.month + ')');
    } else {
      log.push('⚠ No shard records found — Config sheet may be inaccessible');
      log.push('  Run STATUS() to diagnose, or SETUP() to start fresh');
    }
  } catch (e) {
    log.push('✗ Could not read shard registry: ' + e.message);
  }

  try {
    _initConfigSheet(props.getProperty('ACTIVE_SHARD_ID'));
    log.push('✓ Config sheet tabs verified');
  } catch (e) {
    log.push('⚠ Config sheet check failed: ' + e.message);
  }

  try {
    const activeId = props.getProperty('ACTIVE_SHARD_ID');
    if (activeId) {
      _ensureShardSheet(SpreadsheetApp.openById(activeId));
      log.push('✓ Active shard Expenses tab verified');
    }
  } catch (e) {
    log.push('⚠ Active shard check failed: ' + e.message);
  }

  try {
    const existing = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'rotateShardForNewMonth');
    if (existing.length === 0) {
      _installRotationTrigger();
      log.push('✓ Monthly rotation trigger reinstalled');
    } else {
      log.push('✓ Monthly rotation trigger: present');
    }
  } catch (e) {
    log.push('⚠ Trigger check failed: ' + e.message);
  }

  try {
    const hasSI = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === 'processStandingInstructions');
    if (!hasSI) {
      _installSITrigger();
      log.push('✓ Daily SI trigger reinstalled');
    } else {
      log.push('✓ Daily SI trigger: present');
    }
  } catch (e) {
    log.push('⚠ SI trigger check failed: ' + e.message);
  }

  log.push('');
  log.push('=== RESET COMPLETE ===');
  log.push('Run STATUS() to verify.');

  Logger.log(log.join('\n'));
  return { success: true };
}


// ── UPDATE ───────────────────────────────────────────────────
// Run after pulling new source code. Applies schema migrations
// (idempotent — safe to re-run), verifies tabs, clears cache.
function UPDATE() {
  const log = [];
  log.push('=== SPENDWISE UPDATE ===');
  log.push('Running at ' + new Date().toLocaleString());
  log.push('');

  // Ensure Income tab exists (added in v1.0.0)
  try {
    const ss = getConfigSS();
    if (!ss.getSheetByName('Income')) {
      const sheet = ss.insertSheet('Income');
      const headers = ['ID', 'Date', 'Category', 'Description', 'Amount', 'Notes', 'Timestamp'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      log.push('✓ Income tab created');
    } else {
      log.push('✓ Income tab: present');
    }
  } catch (e) {
    log.push('✗ Income tab check failed: ' + e.message);
  }

  // Ensure all shard sheets have Expenses tab
  try {
    const records = _getAllShardRecords();
    records.forEach(r => {
      try { _ensureShardSheet(SpreadsheetApp.openById(r.id)); }
      catch (e) { log.push('⚠ Could not verify shard ' + r.label + ': ' + e.message); }
    });
    log.push('✓ All ' + records.length + ' shard Expenses tabs verified');
  } catch (e) {
    log.push('⚠ Shard check failed: ' + e.message);
  }

  // Ensure rotation trigger installed
  try {
    const existing = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'rotateShardForNewMonth');
    if (existing.length === 0) {
      _installRotationTrigger();
      log.push('✓ Rotation trigger installed');
    } else {
      log.push('✓ Rotation trigger: present');
    }
  } catch (e) {
    log.push('⚠ Trigger check failed: ' + e.message);
  }

  // Ensure daily SI trigger installed
  try {
    const hasSI = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === 'processStandingInstructions');
    if (!hasSI) {
      _installSITrigger();
      log.push('✓ Daily SI trigger installed');
    } else {
      log.push('✓ Daily SI trigger: present');
    }
  } catch (e) {
    log.push('⚠ SI trigger check failed: ' + e.message);
  }

  purgeAllCache();
  log.push('✓ Cache cleared');
  log.push('');
  log.push('=== UPDATE COMPLETE ===');
  log.push('Redeploy the web app to apply changes.');
  Logger.log(log.join('\n'));
  return { success: true };
}


// Deletes and recreates the monthly rotation trigger.
function _installRotationTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'rotateShardForNewMonth')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('rotateShardForNewMonth').timeBased().onMonthDay(1).atHour(0).create();
}

// Deletes and recreates the daily standing instructions trigger (07:00).
function _installSITrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processStandingInstructions')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processStandingInstructions').timeBased().everyDays(1).atHour(7).create();
}

// Public wrapper — callable directly from editor
function setupMonthlyRotationTrigger() {
  _installRotationTrigger();
  Logger.log('✓ Monthly rotation trigger set for rotateShardForNewMonth on day 1 of each month.');
  return 'Monthly rotation trigger installed.';
}


// ── fixShardRegistry ─────────────────────────────────────────
// Validates every registered shard, removes inaccessible ones,
// rewrites the registry in chronological order, and updates
// ACTIVE_SHARD_ID / OLDEST_SHARD_ID in Script Properties.
// Run after bulk import, shard deletion, or shard migration.
function fixShardRegistry() {
  const log = [];
  const tz = Session.getScriptTimeZone();

  log.push('=== FIX SHARD REGISTRY ===');

  const configSS = getConfigSS();
  const regSheet = configSS.getSheetByName(SHARD_REGISTRY);
  if (!regSheet || regSheet.getLastRow() < 2) {
    log.push('✗ ShardRegistry tab not found or empty.');
    Logger.log(log.join('\n'));
    return;
  }

  const rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 4).getValues();
  log.push('Found ' + rows.length + ' registry row(s)');

  const validated = [];
  rows.forEach(r => {
    const shardId = String(r[0] || '').trim();
    const month = r[1] instanceof Date
      ? Utilities.formatDate(r[1], tz, 'yyyy-MM')
      : String(r[1] || '').trim();
    if (!shardId) { log.push('  ↷ Skipping empty row'); return; }

    try {
      const ss = SpreadsheetApp.openById(shardId);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const dataRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
      validated.push({ shardId, month, dataRows });
      log.push('  ✓ Kept: ' + month + ' — ' + dataRows + ' rows (' + shardId.substring(0, 12) + '...)');
    } catch (e) {
      log.push('  ✗ Removed inaccessible shard ' + shardId.substring(0, 12) + '... (' + month + ')');
    }
  });

  // Sort chronologically oldest → newest
  validated.sort((a, b) => (a.month || '').localeCompare(b.month || ''));

  // Rewrite registry
  if (regSheet.getLastRow() > 1) regSheet.deleteRows(2, regSheet.getLastRow() - 1);
  if (validated.length > 0) {
    const newRows = validated.map((s, i) => [
      s.shardId,
      s.month,
      true,
      'Shard ' + String(i + 1).padStart(2, '0') + ' — ' + s.month
    ]);
    regSheet.getRange(2, 1, newRows.length, 4).setValues(newRows);
  }

  // Update Script Properties
  const props = PropertiesService.getScriptProperties();
  if (validated.length > 0) {
    const newest = validated[validated.length - 1];
    const oldest = validated[0];
    props.setProperty('ACTIVE_SHARD_ID', newest.shardId);
    props.setProperty('OLDEST_SHARD_ID', oldest.shardId);
    log.push('');
    log.push('✓ ACTIVE_SHARD_ID → ' + newest.month + ' (' + newest.shardId.substring(0, 12) + '...)');
    log.push('✓ OLDEST_SHARD_ID → ' + oldest.month + ' (' + oldest.shardId.substring(0, 12) + '...)');
  }

  // Clear caches
  SCRIPT_CACHE.removeAll(['shard_registry', 'analytics_week', 'analytics_month',
    'analytics_current_month', 'analytics_quarter', 'analytics_half',
    'analytics_current_year', 'analytics_year']);

  log.push('✓ Registry rewritten with ' + validated.length + ' shard(s) in chronological order');
  log.push('=== DONE ===');

  Logger.log(log.join('\n'));
  return { shards: validated.length, active: validated.length > 0 ? validated[validated.length - 1].month : null };
}


// Overwrites the Categories tab with DEFAULT_CATEGORIES from Code.gs.
function resetCategoriesToDefaults() {
  const result = saveCategories(DEFAULT_CATEGORIES);
  if (result.success) {
    Logger.log('✓ Categories tab reset with ' + DEFAULT_CATEGORIES.length + ' default categories:');
    DEFAULT_CATEGORIES.forEach(c => Logger.log('  ' + c.icon + '  ' + c.name + '  (budget: ₹' + c.budget + ')'));
  } else {
    Logger.log('✗ Failed: ' + result.message);
  }
  return result;
}


// Overwrites Settings tab with DEFAULT_SETTINGS from Code.gs.
function resetSettingsToDefaults() {
  const result = saveSettings(DEFAULT_SETTINGS);
  if (result.success) {
    Logger.log('✓ Settings reset to defaults:');
    Object.entries(DEFAULT_SETTINGS).forEach(([k, v]) => Logger.log('  ' + k + ': ' + v));
  } else {
    Logger.log('✗ Failed to reset settings');
  }
  return result;
}


// ── importFromCSV / forceReimportFromCSV ─────────────────────
// Bulk-imports historical expense data from CSV.
// HOW TO USE:
//   1. Copy the full contents of spendwise_import.csv
//   2. Paste into _getImportRows() below, between the backticks
//   3. Run importFromCSV()        — skips rows with duplicate IDs
//      or forceReimportFromCSV() — clears existing rows first

// Clears existing rows in each affected shard before importing.
function forceReimportFromCSV() {
  return importFromCSV(true);
}

function importFromCSV(forceReimport) {
  forceReimport = forceReimport === true;
  const startTime = Date.now();
  const log = [];

  log.push('=== IMPORT FROM CSV ===');
  log.push(forceReimport ? '⚠ FORCE MODE — existing rows will be cleared' : 'Normal mode — skips duplicates');

  const rows = _getImportRows();
  if (!rows || rows.length === 0) {
    log.push('✗ No rows found. Paste CSV data into _getImportRows() first.');
    Logger.log(log.join('\n'));
    return { success: false };
  }
  log.push('Parsed ' + rows.length + ' rows from CSV');

  // Group by month
  const byMonth = {};
  rows.forEach(row => {
    const month = row.Date ? row.Date.substring(0, 7) : 'unknown';
    if (!byMonth[month]) byMonth[month] = [];
    byMonth[month].push(row);
  });
  log.push('Data spans ' + Object.keys(byMonth).length + ' month(s): ' + Object.keys(byMonth).sort().join(', '));

  const allShards = _getAllShardRecords();
  let totalImported = 0, totalSkipped = 0;

  Object.keys(byMonth).sort().forEach(month => {
    const monthRows = byMonth[month];
    const shardRecord = allShards.find(s => s.month === month);
    let shardId, shardSS;

    if (shardRecord) {
      // Validate registered shard is accessible
      try {
        shardSS = SpreadsheetApp.openById(shardRecord.id);
        shardId = shardRecord.id;
        _ensureShardSheet(shardSS);
      } catch (e) {
        log.push('  ⚠ Registered shard for ' + month + ' inaccessible — recreating...');
        shardId = null;
      }
    }

    if (!shardId) {
      // Create a new shard for this month
      try {
        const newSS = SpreadsheetApp.create('Spendwise — Expenses_Shard_' + month);
        shardId = newSS.getId();
        shardSS = newSS;
        _ensureShardSheet(newSS);

        const regSheet = getConfigSS().getSheetByName(SHARD_REGISTRY);
        if (shardRecord) {
          // Update stale row in registry
          const regData = regSheet.getDataRange().getValues();
          for (let i = 1; i < regData.length; i++) {
            const rowMonth = regData[i][1] instanceof Date
              ? Utilities.formatDate(regData[i][1], Session.getScriptTimeZone(), 'yyyy-MM')
              : String(regData[i][1] || '');
            if (rowMonth === month) { regSheet.getRange(i + 1, 1).setValue(shardId); break; }
          }
        } else {
          regSheet.appendRow([shardId, month, true, 'Shard ' + month]);
        }
        SCRIPT_CACHE.remove('shard_registry');
        log.push('  ✓ Created shard for ' + month + ': ' + shardId.substring(0, 12) + '...');
      } catch (e) {
        log.push('  ✗ Failed to create shard for ' + month + ': ' + e.message);
        totalSkipped += monthRows.length;
        return;
      }
    }

    try {
      const sheet = shardSS.getSheetByName(SHEET_NAME);

      // Force mode: clear existing rows before writing
      if (forceReimport && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
        log.push('  🗑 ' + month + ': cleared existing rows');
      }

      // Duplicate check
      const existingIds = (!forceReimport && sheet.getLastRow() > 1)
        ? new Set(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat())
        : new Set();

      const newRows = monthRows
        .filter(r => !existingIds.has(r.ID))
        .map(r => [
          r.ID,
          _parseLocalDate(r.Date),
          r.Category || 'Other',
          r.Description || '',
          parseFloat(r.Amount) || 0,
          r.PaymentMethod || 'UPI',
          r.Notes || '',
          new Date(r.Timestamp || r.Date)
        ]);

      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
        totalImported += newRows.length;
        log.push('  ✓ ' + month + ': imported ' + newRows.length + ' rows' +
          (monthRows.length - newRows.length > 0 ? ' (skipped ' + (monthRows.length - newRows.length) + ' duplicates)' : ''));
      } else {
        log.push('  ↷ ' + month + ': all ' + monthRows.length + ' rows already exist, skipped');
        totalSkipped += monthRows.length;
      }
    } catch (e) {
      log.push('  ✗ Error writing to shard for ' + month + ': ' + e.message);
      totalSkipped += monthRows.length;
    }
  });

  invalidateCache();

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  log.push('');
  log.push('=== IMPORT COMPLETE ===');
  log.push('Imported : ' + totalImported);
  log.push('Skipped  : ' + totalSkipped);
  log.push('Time     : ' + elapsed + 's');
  log.push('');
  log.push('Run fixShardRegistry() to reorder the registry, then STATUS() to verify.');

  Logger.log(log.join('\n'));
  return { success: true, imported: totalImported, skipped: totalSkipped };
}

// ── Paste your CSV content between the backticks below ───────
function _getImportRows() {
  const csv = ``;  // ← PASTE spendwise_import.csv CONTENTS HERE

  if (!csv || !csv.trim()) return [];

  const lines = csv.trim().split('\n');
  const header = lines[0].split(',').map(h => h.replace(/^"|"$/g, '').trim());

  return lines.slice(1).map(line => {
    // Handle quoted fields containing commas
    const fields = [];
    let cur = '', inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') { inQuote = !inQuote; }
      else if (ch === ',' && !inQuote) { fields.push(cur); cur = ''; }
      else { cur += ch; }
    }
    fields.push(cur);
    const row = {};
    header.forEach((h, i) => { row[h] = (fields[i] || '').trim(); });
    return row;
  }).filter(r => r.ID && r.Date && r.Amount);
}


// ── runShardExperiment ───────────────────────────────────────
// Benchmarks shard read performance: compares 1-month vs 6-month
// fan-out, measures per-shard overhead, and prints an interpretation.
// Run after import to verify sharding overhead is acceptable.
function runShardExperiment() {
  const log = [];
  const tz = Session.getScriptTimeZone();
  const now = new Date();

  log.push('=== SHARD EXPERIMENT ===');

  // Clear analytics cache so reads are cold
  SCRIPT_CACHE.removeAll(['analytics_current_month', 'analytics_quarter', 'shard_registry']);

  // ── Test 1: 1-month read (active shard only) ───────────────
  const t1Start = Date.now();
  const activeId = getActiveShardId();
  const currentMonth = Utilities.formatDate(now, tz, 'yyyy-MM');
  const oneMonthRows = _readShardExpenses(activeId)
    .filter(e => e.date && e.date.substring(0, 7) === currentMonth);
  const t1 = Date.now() - t1Start;

  log.push('');
  log.push('[1-MONTH READ]');
  log.push('  Shards opened : 1');
  log.push('  Rows returned : ' + oneMonthRows.length);
  log.push('  Time          : ' + t1 + 'ms');

  // ── Test 2: 6-month fan-out ────────────────────────────────
  const sixMonthsAgo = new Date(now);
  sixMonthsAgo.setMonth(now.getMonth() - 6);
  const startStr = Utilities.formatDate(sixMonthsAgo, tz, 'yyyy-MM-dd');
  const shardIds = _getShardsForRange(startStr, null);

  const t2Start = Date.now();
  let sixMonthRows = 0;
  const shardTimings = [];
  shardIds.forEach(id => {
    const ts = Date.now();
    const rows = _readShardExpenses(id).filter(e => e.date && e.date >= startStr);
    sixMonthRows += rows.length;
    shardTimings.push({
      id: id.substring(0, 12), rows: rows.length, ms: Date.now() - ts,
      month: _getAllShardRecords().find(s => s.id === id)?.month || '?'
    });
  });
  const t2 = Date.now() - t2Start;

  log.push('');
  log.push('[6-MONTH READ]');
  log.push('  Shards opened : ' + shardIds.length);
  log.push('  Rows returned : ' + sixMonthRows);
  log.push('  Time          : ' + t2 + 'ms');

  // ── Test 3: Full analytics ─────────────────────────────────
  const t3Start = Date.now();
  const a = getAnalytics('current_month');
  const t3 = Date.now() - t3Start;

  log.push('');
  log.push('[6-MONTH ANALYTICS (getAnalytics)]');
  log.push('  Total spent   : ₹' + a.totalSpent);
  log.push('  Transactions  : ' + a.totalTransactions);
  log.push('  Shards queried: ' + a.shardCount);
  log.push('  Time          : ' + t3 + 'ms');

  // ── Per-shard breakdown ────────────────────────────────────
  log.push('');
  log.push('[PER-SHARD BREAKDOWN]');
  shardTimings.forEach(s => {
    log.push('  ' + s.month + ': ' + s.rows + ' rows in ' + s.ms + 'ms');
  });

  // ── Summary ────────────────────────────────────────────────
  const overhead = t2 - t1;
  const extraShards = Math.max(1, shardIds.length - 1);
  const perShard = Math.round(overhead / extraShards);

  log.push('');
  log.push('[SUMMARY]');
  log.push('  1-month  : ' + t1 + 'ms for ' + oneMonthRows.length + ' rows across 1 shard');
  log.push('  6-months : ' + t2 + 'ms for ' + sixMonthRows + ' rows across ' + shardIds.length + ' shard(s)');
  log.push('  Overhead : +' + overhead + 'ms for ' + extraShards + ' extra shard(s)');
  log.push('  Per extra shard: ~' + perShard + 'ms');
  log.push('');
  if (perShard < 500) {
    log.push('INTERPRETATION:');
    log.push('  ✓ Shard overhead is LOW — monthly sharding is efficient for your data volume.');
  } else if (perShard < 1500) {
    log.push('INTERPRETATION:');
    log.push('  ⚠ Shard overhead is MODERATE — consider quarterly sharding if you add more months.');
  } else {
    log.push('INTERPRETATION:');
    log.push('  ✗ Shard overhead is HIGH — switch to quarterly sharding (3 months per shard).');
  }
  log.push('=== END EXPERIMENT ===');

  Logger.log(log.join('\n'));
  return { t1, t2, t3, overhead, perShard };
}