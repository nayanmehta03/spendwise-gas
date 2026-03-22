// ============================================================
// fixShardRegistry()
// Run ONCE from the Apps Script editor after bulk import.
// Does three things:
//   1. Reads all ShardRegistry rows
//   2. Removes any rows where the shard sheet has 0 data rows
//      (cleans up the original blank shard if unused)
//   3. Rewrites registry sorted chronologically (oldest → newest)
//   4. Updates ACTIVE_SHARD_ID and OLDEST_SHARD_ID in Script Properties
// ============================================================
function fixShardRegistry() {
  const log = [];
  const tz  = Session.getScriptTimeZone();

  const configSS  = getConfigSS();
  const regSheet  = configSS.getSheetByName(SHARD_REGISTRY);
  if (!regSheet || regSheet.getLastRow() < 2) {
    Logger.log('ShardRegistry not found or empty.');
    return;
  }

  // Read all existing registry rows
  const rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 4).getValues();
  log.push(`Found ${rows.length} registry rows`);

  // Inspect each shard — check if it has actual data
  const validated = [];
  rows.forEach(r => {
    const shardId = String(r[0] || '').trim();
    const month   = r[1] instanceof Date
      ? Utilities.formatDate(r[1], tz, 'yyyy-MM')
      : String(r[1] || '').trim();
    const label   = String(r[3] || '').trim();

    if (!shardId) { log.push(`  ↷ Skipping empty row`); return; }

    try {
      const ss    = SpreadsheetApp.openById(shardId);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const dataRows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;

      if (dataRows === 0 && shardId === FIRST_SHARD_ID) {
        log.push(`  ✗ Removed: ${label || shardId} (${month}) — original blank shard, 0 data rows`);
        return; // skip — don't include in rewritten registry
      }

      validated.push({ shardId, month, dataRows, label });
      log.push(`  ✓ Kept: ${month} — ${dataRows} rows (${shardId.substring(0,12)}...)`);
    } catch(e) {
      log.push(`  ✗ Could not open shard ${shardId}: ${e.message} — removing from registry`);
    }
  });

  // Sort chronologically: oldest (2025-03) → newest (2026-03)
  validated.sort((a, b) => (a.month || '').localeCompare(b.month || ''));

  // Rewrite registry sheet cleanly
  if (regSheet.getLastRow() > 1) regSheet.deleteRows(2, regSheet.getLastRow() - 1);

  if (validated.length > 0) {
    const newRows = validated.map((s, i) => [
      s.shardId,
      s.month,
      true,
      `Shard ${String(i + 1).padStart(2, '0')} — ${s.month}`
    ]);
    regSheet.getRange(2, 1, newRows.length, 4).setValues(newRows);
  }

  // Update Script Properties
  const props      = PropertiesService.getScriptProperties();
  const newestShard = validated[validated.length - 1]; // last = newest month
  const oldestShard = validated[0];                     // first = oldest month

  if (newestShard) {
    props.setProperty('ACTIVE_SHARD_ID', newestShard.shardId);
    log.push(`\n✓ ACTIVE_SHARD_ID  → ${newestShard.month} (${newestShard.shardId.substring(0,12)}...)`);
  }
  if (oldestShard) {
    props.setProperty('OLDEST_SHARD_ID', oldestShard.shardId);
    log.push(`✓ OLDEST_SHARD_ID  → ${oldestShard.month} (${oldestShard.shardId.substring(0,12)}...)`);
  }

  // Clear all caches so next request re-reads clean registry
  SCRIPT_CACHE.removeAll(['shard_registry','analytics_week','analytics_month',
    'analytics_current_month','analytics_quarter','analytics_year']);

  log.push(`\nRegistry rewritten with ${validated.length} shards in chronological order.`);
  log.push('Run runShardExperiment() again to verify.');

  const result = '\n=== FIX SHARD REGISTRY ===\n' + log.join('\n') + '\n==========================';
  Logger.log(result);
  return { shards: validated.length, active: newestShard?.month, oldest: oldestShard?.month };
}

// ============================================================
// PART 2 — SHARD EXPERIMENT: 1 month vs 6 months
// Measures actual GAS execution time reading different shard
// configurations for the same analytics query.
// ============================================================

function runShardExperiment() {
  const log = [];
  log.push('=== SHARD EXPERIMENT: 1-month vs 6-month read ===\n');

  const tz  = Session.getScriptTimeZone();
  const now = new Date();

  // ── Test 1: Current month only (1 shard) ─────────────────
  const t1Start   = Date.now();
  const monthStart = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), tz, 'yyyy-MM-dd');
  const monthEnd   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  const shard1Ids  = _getShardsForRange(monthStart, monthEnd);
  let   rows1      = [];
  shard1Ids.forEach(id => { rows1 = rows1.concat(_readShardExpenses(id)); });
  rows1 = rows1.filter(e => e.date >= monthStart && e.date <= monthEnd);
  const t1Ms = Date.now() - t1Start;

  log.push(`[1-MONTH READ]`);
  log.push(`  Shards opened : ${shard1Ids.length}`);
  log.push(`  Rows returned : ${rows1.length}`);
  log.push(`  Time          : ${t1Ms}ms`);
  log.push('');

  // ── Test 2: Last 6 months (multiple shards) ───────────────
  const t2Start   = Date.now();
  const sixMoStart = new Date(now);
  sixMoStart.setMonth(sixMoStart.getMonth() - 5);
  sixMoStart.setDate(1);
  const sixMoStartStr = Utilities.formatDate(sixMoStart, tz, 'yyyy-MM-dd');

  const shard6Ids  = _getShardsForRange(sixMoStartStr, monthEnd);
  let   rows6      = [];
  shard6Ids.forEach(id => { rows6 = rows6.concat(_readShardExpenses(id)); });
  rows6 = rows6.filter(e => e.date >= sixMoStartStr && e.date <= monthEnd);
  const t2Ms = Date.now() - t2Start;

  log.push(`[6-MONTH READ]`);
  log.push(`  Shards opened : ${shard6Ids.length}`);
  log.push(`  Rows returned : ${rows6.length}`);
  log.push(`  Time          : ${t2Ms}ms`);
  log.push('');

  // ── Test 3: Full analytics over 6 months ─────────────────
  const t3Start = Date.now();
  // Invalidate cache to force a real read
  SCRIPT_CACHE.remove('analytics_month');
  const analytics6 = getAnalytics('month');
  const t3Ms = Date.now() - t3Start;

  log.push(`[6-MONTH ANALYTICS (getAnalytics)]`);
  log.push(`  Total spent   : ₹${analytics6.totalSpent}`);
  log.push(`  Transactions  : ${analytics6.totalTransactions}`);
  log.push(`  Shards queried: ${analytics6.shardCount}`);
  log.push(`  Time          : ${t3Ms}ms`);
  log.push('');

  // ── Per-shard breakdown ───────────────────────────────────
  log.push('[PER-SHARD BREAKDOWN]');
  shard6Ids.forEach(id => {
    const t = Date.now();
    const shardRows = _readShardExpenses(id);
    const elapsed   = Date.now() - t;
    const record    = _getAllShardRecords().find(s => s.id === id);
    log.push(`  ${record ? record.month : id}: ${shardRows.length} rows in ${elapsed}ms`);
  });
  log.push('');

  // ── Summary ───────────────────────────────────────────────
  const overhead = t2Ms - t1Ms;
  log.push('[SUMMARY]');
  log.push(`  1-month  : ${t1Ms}ms for ${rows1.length} rows across ${shard1Ids.length} shard(s)`);
  log.push(`  6-months : ${t2Ms}ms for ${rows6.length} rows across ${shard6Ids.length} shard(s)`);
  log.push(`  Overhead : +${overhead}ms for ${shard6Ids.length - shard1Ids.length} extra shard(s)`);
  log.push(`  Per extra shard: ~${shard6Ids.length > 1 ? Math.round(overhead / (shard6Ids.length - 1)) : 'N/A'}ms`);
  log.push('');
  log.push('INTERPRETATION:');
  if (overhead < 500) {
    log.push('  ✓ Shard overhead is LOW — each additional shard adds minimal latency.');
    log.push('    Monthly sharding is efficient for your data volume.');
  } else if (overhead < 2000) {
    log.push('  ⚠ Shard overhead is MODERATE — consider quarterly shards if analytics');
    log.push('    over long periods feel slow.');
  } else {
    log.push('  ✗ Shard overhead is HIGH — GAS sheet open() calls are expensive at this');
    log.push('    volume. Consider quarterly shards (3 months per sheet) instead of monthly.');
  }

  const result = log.join('\n');
  Logger.log('\n' + result);
  return { t1Ms, t2Ms, t3Ms, rows1: rows1.length, rows6: rows6.length, shards1: shard1Ids.length, shards6: shard6Ids.length };
}