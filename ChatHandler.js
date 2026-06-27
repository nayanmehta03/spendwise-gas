// ============================================================
// SPENDWISE — ChatHandler.gs  v1.0.0
// Google Chat App integration: slash commands, DM expense logging,
// daily spending summary.
//
// ENTRY POINTS:
//   onMessage(event)         — Chat App message handler (auto-wired by GAS)
//   sendDailySummary()       — Trigger entry: sends daily summary to Chat
//
// COMMANDS:
//   /spend <amount> <keyword>   — Log an expense
//   /summary                    — Today's spending summary
//   /today                      — Alias for /summary
//   /help                       — Usage instructions
//
// SETUP:
//   1. Enable chatEnabled='true' in Settings sheet
//   2. Create a Google Chat App in GCP Console
//   3. Link to this Apps Script project
//   4. Add the bot to a Chat space — chatSpaceId auto-captured
// ============================================================


// ── Diagnostics ──────────────────────────────────────────────
// Run this manually from the Apps Script editor to debug a
// "Spendwise not responding" error in Google Chat.
//   • First run triggers the OAuth consent screen for ALL declared
//     scopes (chat.bot, gmail, etc.) — granting these often fixes
//     "not responding" by itself.
//   • Then it checks setup + settings and simulates the slash
//     commands locally so you can confirm the handler works.
// Read the output in View → Logs (or the Executions page).
function runChatDiagnostics() {
  const out = [];
  const ok = (b) => (b ? '✓' : '✗');

  // 1. Force scope authorization + identity
  let effEmail = '', actEmail = '';
  try { effEmail = Session.getEffectiveUser().getEmail(); } catch (e) { effEmail = 'ERROR: ' + e.message; }
  try { actEmail = Session.getActiveUser().getEmail(); } catch (e) { actEmail = '(not available)'; }
  out.push('IDENTITY');
  out.push('  effectiveUser (bot runs as): ' + effEmail);
  out.push('  activeUser (caller):         ' + actEmail);

  // 2. Setup
  out.push('');
  out.push('SETUP');
  let setupOk = false;
  try {
    const configId = PropertiesService.getScriptProperties().getProperty('CONFIG_SHEET_ID');
    setupOk = !!configId;
    out.push('  ' + ok(setupOk) + ' CONFIG_SHEET_ID: ' + (configId || 'MISSING — run SETUP() in AdminOps'));
    out.push('  ' + ok(true) + ' activeShardId: ' + getActiveShardId());
  } catch (e) {
    out.push('  ✗ Setup error: ' + e.message + ' — run SETUP() in AdminOps');
  }

  // 3. Settings
  out.push('');
  out.push('SETTINGS');
  try {
    const s = getSettings();
    out.push('  ' + ok(s.chatEnabled === 'true') + ' chatEnabled: ' + s.chatEnabled +
      (s.chatEnabled === 'true' ? '' : '  ← MUST be "true" or the bot replies "not enabled"'));
    out.push('  chatSpaceId: ' + (s.chatSpaceId || '(empty — auto-captured on first message)'));
  } catch (e) {
    out.push('  ✗ getSettings error: ' + e.message);
  }

  // Self-heal: clear the fake spaceId an earlier version of this
  // function may have written, so a real Chat event can be captured.
  try {
    const s2 = getSettings();
    if (s2.chatSpaceId === 'spaces/DIAGNOSTIC') {
      saveSettings({ chatSpaceId: '' });
      out.push('  (cleared stale chatSpaceId=spaces/DIAGNOSTIC)');
    }
  } catch (e) {}

  // 4. Simulate slash commands (read-only ones). No `space` is passed,
  // so this never overwrites the real chatSpaceId.
  out.push('');
  out.push('SIMULATED COMMANDS');
  ['/help', '/summary'].forEach(cmd => {
    try {
      // Simulate the real add-on event shape.
      const resp = onMessage({ chat: { messagePayload: { message: { text: cmd } } } });
      const m = resp && resp.hostAppDataAction && resp.hostAppDataAction.chatDataAction &&
        resp.hostAppDataAction.chatDataAction.createMessageAction &&
        resp.hostAppDataAction.chatDataAction.createMessageAction.message;
      const kind = m && m.cardsV2 ? 'card' : (m && m.text ? 'text' : 'EMPTY (bad envelope!)');
      out.push('  ' + ok(!!m) + ' ' + cmd + ' → ' + kind +
        (m && m.text ? ': ' + m.text.split('\n')[0] : ''));
    } catch (e) {
      out.push('  ✗ ' + cmd + ' THREW: ' + e.message);
    }
  });

  const report = out.join('\n');
  Logger.log(report);
  return report;
}


// ── Chat App entry points (Google Workspace add-on framework) ─
// The add-on framework invokes these by name. Replies MUST be wrapped
// in the hostAppDataAction envelope via _chatReply(). Event data lives
// under event.chat.<payload>.*; we also accept the classic shape
// (event.message / event.space) so local tests still work.

function onMessage(event) {
  try {
    const settings = getSettings();
    if (settings.chatEnabled !== 'true') {
      return _chatReply(_chatTextResponse('⚙️ Spendwise Chat is not enabled. Set chatEnabled=true in the Settings sheet.'));
    }
    const ctx = _extractChatContext(event);
    _maybeCaptureSpace(ctx, settings);

    const text = (ctx.message && ctx.message.text) ? ctx.message.text.trim() : '';
    if (!text) return _chatReply(_chatTextResponse('Send /help for usage instructions.'));

    return _chatReply(_routeCommand(text, ctx));
  } catch (e) {
    Logger.log('onMessage error: ' + e.message);
    logEvent('CHAT', 'ERROR', 'onMessage: ' + e.message, '');
    return _chatReply(_chatTextResponse('❌ Something went wrong: ' + e.message));
  }
}

// Slash / app commands are routed by the framework to onAppCommand.
function onAppCommand(event) {
  try {
    const settings = getSettings();
    if (settings.chatEnabled !== 'true') {
      return _chatReply(_chatTextResponse('⚙️ Spendwise Chat is not enabled. Set chatEnabled=true in the Settings sheet.'));
    }
    const ctx = _extractChatContext(event);
    _maybeCaptureSpace(ctx, settings);

    // Map configured slash-command IDs → command text.
    const cmdId = ctx.appCommand ? String(ctx.appCommand.appCommandId) : '';
    let cmdText = '';
    if (cmdId === '1') cmdText = 'spend';
    else if (cmdId === '2') cmdText = 'summary';
    else if (cmdId === '3') cmdText = 'help';

    const argsText = (ctx.message && ctx.message.argumentText) ? ctx.message.argumentText.trim() : '';
    let fullText = cmdText ? (cmdText + (argsText ? ' ' + argsText : '')) : '';
    // Fallback: some payloads carry literal text instead of a command ID.
    if (!fullText && ctx.message && ctx.message.text) fullText = ctx.message.text.trim();
    if (!fullText) return _chatReply(_handleHelpCommand());

    return _chatReply(_routeCommand(fullText, ctx));
  } catch (e) {
    Logger.log('onAppCommand error: ' + e.message);
    logEvent('CHAT', 'ERROR', 'onAppCommand: ' + e.message, '');
    return _chatReply(_chatTextResponse('❌ Something went wrong: ' + e.message));
  }
}

// Framework lifecycle events. Without these defined, the add-on logs
// "Script function not found" when added to / removed from a space.
function onAddedToSpace(event) {
  try {
    _maybeCaptureSpace(_extractChatContext(event), getSettings());
  } catch (e) { Logger.log('onAddedToSpace: ' + e.message); }
  return _chatReply(_handleHelpCommand());
}

function onRemovedFromSpace(event) {
  try { logEvent('CHAT', 'INFO', 'Removed from space', ''); } catch (e) {}
  return {};
}

// Back-compat alias used by some Chat API versions.
function onChatMessage(event) {
  return onMessage(event);
}


// ── Routing & event helpers ──────────────────────────────────

// Normalizes the add-on event shape (event.chat.*Payload) and the
// classic shape (event.message / event.space) into one object.
function _extractChatContext(event) {
  event = event || {};
  const chat = event.chat || {};
  const payload = chat.messagePayload || chat.appCommandPayload ||
                  chat.addedToSpacePayload || chat.removedFromSpacePayload ||
                  chat.buttonClickedPayload || {};
  return {
    message: payload.message || event.message || {},
    space: payload.space || event.space || chat.space || {},
    user: chat.user || event.user || {},
    appCommand: (chat.appCommandPayload && chat.appCommandPayload.appCommandMetadata) || null
  };
}

function _maybeCaptureSpace(ctx, settings) {
  if (ctx.space && ctx.space.name && (!settings || !settings.chatSpaceId)) {
    try { saveSettings({ chatSpaceId: ctx.space.name }); }
    catch (e) { Logger.log('Auto-capture spaceId error: ' + e.message); }
  }
}

// Routes a text command to a handler, returning a PLAIN message object
// ({text} or {cardsV2}). The caller wraps it with _chatReply().
function _routeCommand(text, ctx) {
  const parts = text.split(/\s+/);
  const command = parts[0].toLowerCase().replace(/^\//, '');
  switch (command) {
    case 'spend':
      return _handleSpendCommand(parts.slice(1), ctx);
    case 'summary':
    case 'today':
      return _handleSummaryCommand();
    case 'help':
      return _handleHelpCommand();
    default:
      if (/^\d/.test(parts[0])) return _handleSpendCommand(parts, ctx);
      return _chatTextResponse(
        '❓ Unknown command: *' + command + '*\n\n' +
        'Try:\n• `/spend 250 lunch`\n• `/summary`\n• `/help`'
      );
  }
}

// Wraps a plain Chat message ({text}/{cardsV2}) in the Workspace
// add-on response envelope the Chat framework requires.
function _chatReply(message) {
  if (!message) return {};
  return { hostAppDataAction: { chatDataAction: { createMessageAction: { message: message } } } };
}


// ── /spend <amount> <keyword> ────────────────────────────────
function _handleSpendCommand(args, ctx) {
  // Validate arguments
  if (!args || args.length < 2) {
    return _chatTextResponse(
      '⚠️ *Usage:* `/spend <amount> <keyword>`\n\n' +
      'Examples:\n• `/spend 250 lunch`\n• `/spend 120 uber`\n• `/spend 1800 groceries`'
    );
  }

  const amountStr = args[0];
  const keyword = args.slice(1).join(' ');
  const amount = parseFloat(amountStr);

  // Validate amount
  if (isNaN(amount) || amount <= 0) {
    logEvent('CHAT', 'ERROR', 'Invalid amount: ' + amountStr, '');
    return _chatTextResponse('❌ *Invalid amount:* `' + amountStr + '`\n\nAmount must be a number greater than zero.');
  }

  // Validate keyword
  if (!keyword || keyword.trim().length === 0) {
    return _chatTextResponse('❌ *Keyword is required.*\n\nExample: `/spend 250 lunch`');
  }

  // Categorize: explicit keyword map → learned from history → 'Other'
  const category = smartCategorize(keyword);

  // Duplicate check using Chat message ID
  const externalId = (ctx && ctx.message && ctx.message.name) ? ctx.message.name : '';
  if (externalId && _isDuplicateExternal(externalId)) {
    logEvent('CHAT', 'DUPLICATE', 'Duplicate chat expense: ' + keyword + ' ₹' + amount, externalId);
    return _chatTextResponse('⚠️ This expense was already recorded.');
  }

  // Get today's date
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Save expense
  const result = addExpenseWithSource({
    date: today,
    category: category,
    description: keyword,
    amount: amount,
    paymentMethod: 'UPI',
    notes: ''
  }, 'CHAT', externalId);

  if (result.success) {
    logEvent('CHAT', 'SUCCESS', 'Logged: ' + keyword + ' ₹' + amount + ' → ' + category, externalId);

    // Build confirmation card
    const settings = getSettings();
    const currency = settings.currency || '₹';
    return _buildSpendCard(currency, amount, keyword, category, result.id);
  } else {
    logEvent('CHAT', 'ERROR', 'Failed to save expense: ' + (result.message || ''), externalId);
    return _chatTextResponse('❌ Failed to save expense: ' + (result.message || 'Unknown error'));
  }
}


// ── /summary — Today's spending ──────────────────────────────
function _handleSummaryCommand() {
  try {
    const settings = getSettings();
    const currency = settings.currency || '₹';
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Get today's expenses from active shard
    const allExpenses = _readShardExpenses(getActiveShardId());
    const todayExpenses = allExpenses.filter(e => e.date === today);

    if (todayExpenses.length === 0) {
      return _chatTextResponse('📊 *No expenses recorded today.*\n\nUse `/spend <amount> <keyword>` to log one.');
    }

    // Aggregate by category
    const byCategory = {};
    let total = 0;
    todayExpenses.forEach(e => {
      byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
      total += e.amount;
    });

    // Sort categories by amount descending
    const sortedCats = Object.entries(byCategory)
      .sort((a, b) => b[1] - a[1]);

    // Top expenses
    const topExpenses = todayExpenses
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 5);

    return _buildSummaryCard(currency, total, sortedCats, topExpenses, today);
  } catch (e) {
    Logger.log('_handleSummaryCommand error: ' + e.message);
    return _chatTextResponse('❌ Failed to generate summary: ' + e.message);
  }
}


// ── /help ────────────────────────────────────────────────────
function _handleHelpCommand() {
  return _chatTextResponse(
    '💰 *Spendwise Chat Commands*\n\n' +
    '`/spend <amount> <keyword>`\n' +
    '  Log an expense. Category is auto-detected from keyword.\n' +
    '  Examples: `/spend 250 lunch`, `/spend 120 uber`\n\n' +
    '`/summary`\n' +
    '  View today\'s spending breakdown.\n\n' +
    '`/help`\n' +
    '  Show this message.\n\n' +
    '📋 *Keyword Mappings*\n' +
    'Keywords like "lunch", "uber", "grocery" are mapped to categories.\n' +
    'Manage mappings in the KeywordMap sheet in your Config spreadsheet.'
  );
}


// ── Card builders ────────────────────────────────────────────

function _buildSpendCard(currency, amount, keyword, category, expenseId) {
  return {
    cardsV2: [{
      cardId: 'spend_' + expenseId,
      card: {
        header: {
          title: '✅ Expense Logged',
          subtitle: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE, MMM d')
        },
        sections: [{
          widgets: [
            {
              decoratedText: {
                topLabel: 'Amount',
                text: '<b>' + currency + parseFloat(amount).toLocaleString('en-IN') + '</b>',
                startIcon: { knownIcon: 'DOLLAR' }
              }
            },
            {
              decoratedText: {
                topLabel: 'Description',
                text: keyword,
                startIcon: { knownIcon: 'DESCRIPTION' }
              }
            },
            {
              decoratedText: {
                topLabel: 'Category',
                text: category,
                startIcon: { knownIcon: 'BOOKMARK' }
              }
            }
          ]
        }]
      }
    }]
  };
}

function _buildSummaryCard(currency, total, sortedCats, topExpenses, dateStr) {
  const catLines = sortedCats.map(([cat, amt]) =>
    cat + '  ' + currency + parseFloat(amt).toLocaleString('en-IN')
  ).join('\n');

  const topLines = topExpenses.map(e =>
    '• ' + e.description + '  ' + currency + parseFloat(e.amount).toLocaleString('en-IN')
  ).join('\n');

  return {
    cardsV2: [{
      cardId: 'summary_' + dateStr,
      card: {
        header: {
          title: '📊 Today\'s Spending',
          subtitle: dateStr
        },
        sections: [
          {
            header: 'Total',
            widgets: [{
              decoratedText: {
                text: '<b>' + currency + parseFloat(total).toLocaleString('en-IN') + '</b>',
                startIcon: { knownIcon: 'DOLLAR' }
              }
            }]
          },
          {
            header: 'By Category',
            widgets: [{
              textParagraph: { text: catLines || 'No categories' }
            }]
          },
          {
            header: 'Top Expenses',
            widgets: [{
              textParagraph: { text: topLines || 'No expenses' }
            }]
          }
        ]
      }
    }]
  };
}

function _chatTextResponse(text) {
  return { text: text };
}


// ============================================================
// DAILY SUMMARY — Proactive Chat DM
// Called by time-based trigger at configured hour.
// ============================================================

function sendDailySummary() {
  try {
    const settings = getSettings();
    if (settings.dailySummaryEnabled !== 'true') {
      Logger.log('sendDailySummary: disabled in settings');
      return { skipped: true };
    }

    const spaceId = settings.chatSpaceId;
    if (!spaceId) {
      Logger.log('sendDailySummary: no chatSpaceId configured. Send a message to the bot first.');
      logEvent('SUMMARY', 'ERROR', 'No chatSpaceId configured', '');
      return { success: false, message: 'No chatSpaceId' };
    }

    const currency = settings.currency || '₹';
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Get today's expenses
    const allExpenses = _readShardExpenses(getActiveShardId());
    const todayExpenses = allExpenses.filter(e => e.date === today);

    // Aggregate by category
    const byCategory = {};
    let total = 0;
    todayExpenses.forEach(e => {
      byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
      total += e.amount;
    });

    const sortedCats = Object.entries(byCategory).sort((a, b) => b[1] - a[1]);
    const topExpenses = todayExpenses.sort((a, b) => b.amount - a.amount).slice(0, 5);

    // Build message text
    let text = '📊 *Today\'s Spending*\n\n';

    if (todayExpenses.length === 0) {
      text += 'No expenses recorded today. 🎉';
    } else {
      sortedCats.forEach(([cat, amt]) => {
        text += cat + '  ' + currency + parseFloat(amt).toLocaleString('en-IN') + '\n';
      });
      text += '\n*Total  ' + currency + parseFloat(total).toLocaleString('en-IN') + '*\n';

      if (topExpenses.length > 0) {
        text += '\n*Top Expenses:*\n';
        topExpenses.forEach(e => {
          text += '• ' + e.description + '  ' + currency + parseFloat(e.amount).toLocaleString('en-IN') + '\n';
        });
      }
    }

    // Send via Chat API
    const message = { text: text };
    Chat.Spaces.Messages.create(message, spaceId);

    logEvent('SUMMARY', 'SUCCESS', 'Daily summary sent. Total: ' + currency + total, '');
    Logger.log('sendDailySummary: sent. Total: ' + currency + total);
    return { success: true, total };
  } catch (e) {
    Logger.log('sendDailySummary error: ' + e.message);
    logEvent('SUMMARY', 'ERROR', 'sendDailySummary: ' + e.message, '');
    return { success: false, message: e.message };
  }
}
