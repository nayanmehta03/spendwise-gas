// ============================================================
// SPENDWISE — EmailIngestion.gs  v1.0.0
// Automated email receipt import pipeline.
//
// ENTRY POINTS:
//   processEmailReceipts()      — Main trigger entry point
//   runEmailImportNow()         — Manual one-shot import (UI button)
//
// ARCHITECTURE:
//   Parser Registry pattern — each merchant has a registered parser.
//   EmailSources tab in Config sheet controls which senders to scan.
//   Adding a new merchant = add row to EmailSources + optionally
//   register a parser function.
//
// PARSERS:
//   generic  — regex-based, works for most receipts
//   swiggy   — Swiggy order delivery emails
//   zomato   — Zomato order summary emails
//   amazon   — Amazon order confirmation emails
//
// DUPLICATE PREVENTION:
//   Gmail message ID stored as [EXT:msgId] in Notes field.
//   Checked before every insert via _isDuplicateExternal().
// ============================================================


// ── Parser Registry ──────────────────────────────────────────
const PARSER_REGISTRY = {};

function registerParser(parserType, parserFn) {
  PARSER_REGISTRY[parserType.toLowerCase()] = parserFn;
}


// ── Generic Parser ───────────────────────────────────────────
// Extracts amount from email body/subject using common patterns.
// Works for most transactional emails with ₹ or Rs. amounts.
registerParser('generic', function (message, source) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const content = subject + '\n' + body;

  // Try to extract amount using common Indian currency patterns
  const amountPatterns = [
    /(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:total|amount|paid|charged|debited)\s*:?\s*(?:₹|Rs\.?|INR)?\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)\s*(?:only|paid|charged|debited)/gi,
  ];

  let amount = 0;
  for (const pattern of amountPatterns) {
    const matches = [...content.matchAll(pattern)];
    if (matches.length > 0) {
      // Take the largest amount found (likely the total)
      const amounts = matches.map(m => parseFloat(m[1].replace(/,/g, ''))).filter(a => a > 0);
      if (amounts.length > 0) {
        amount = Math.max(...amounts);
        break;
      }
    }
  }

  if (amount <= 0) return null;

  // Extract merchant from sender or subject
  const from = message.getFrom() || '';
  const merchant = _extractMerchantFromSender(from) || _extractFirstWord(subject);

  return {
    amount: amount,
    merchant: merchant,
    description: merchant + ' order',
    category: smartCategorize(merchant),
    date: message.getDate()
  };
});


// ── Swiggy Parser ────────────────────────────────────────────
registerParser('swiggy', function (message, source) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const content = subject + '\n' + body;

  // Swiggy emails typically contain "Order Delivered" and total amount
  let amount = 0;

  // Pattern: "Total ₹XXX" or "Grand Total: ₹XXX" or "Paid ₹XXX"
  const patterns = [
    /(?:grand\s*total|total\s*(?:bill|paid|amount))\s*:?\s*(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)\s*(?:paid|total)/gi,
    /(?:you\s*paid)\s*(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/gi,
  ];

  for (const pattern of patterns) {
    const matches = [...content.matchAll(pattern)];
    if (matches.length > 0) {
      const amounts = matches.map(m => parseFloat(m[1].replace(/,/g, ''))).filter(a => a > 0);
      if (amounts.length > 0) {
        amount = Math.max(...amounts);
        break;
      }
    }
  }

  if (amount <= 0) return null;

  // Try to extract restaurant name
  let restaurant = 'Swiggy';
  const restMatch = body.match(/(?:from|order\s*from|restaurant)\s*:?\s*([A-Za-z][A-Za-z &'-]{2,30})/i);
  if (restMatch) restaurant = restMatch[1].trim();

  return {
    amount: amount,
    merchant: 'Swiggy',
    description: restaurant + ' via Swiggy',
    category: categorizeByKeyword('swiggy') || 'Food & Drink',
    date: message.getDate()
  };
});


// ── Zomato Parser ────────────────────────────────────────────
registerParser('zomato', function (message, source) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const content = subject + '\n' + body;

  let amount = 0;

  const patterns = [
    /(?:grand\s*total|total\s*(?:bill|paid|amount)|you\s*paid)\s*:?\s*(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)\s*(?:paid|total)/gi,
    /(?:₹|Rs\.?)\s*([\d,]+(?:\.\d{1,2})?)/gi,
  ];

  for (const pattern of patterns) {
    const matches = [...content.matchAll(pattern)];
    if (matches.length > 0) {
      const amounts = matches.map(m => parseFloat(m[1].replace(/,/g, ''))).filter(a => a > 0);
      if (amounts.length > 0) {
        amount = Math.max(...amounts);
        break;
      }
    }
  }

  if (amount <= 0) return null;

  // Try to extract restaurant name
  let restaurant = 'Zomato';
  const restMatch = body.match(/(?:from|order\s*from|restaurant)\s*:?\s*([A-Za-z][A-Za-z &'-]{2,30})/i);
  if (restMatch) restaurant = restMatch[1].trim();

  return {
    amount: amount,
    merchant: 'Zomato',
    description: restaurant + ' via Zomato',
    category: categorizeByKeyword('zomato') || 'Food & Drink',
    date: message.getDate()
  };
});


// ── Amazon Parser ────────────────────────────────────────────
registerParser('amazon', function (message, source) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const content = subject + '\n' + body;

  let amount = 0;

  const patterns = [
    /(?:order\s*total|grand\s*total|total)\s*:?\s*(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)/gi,
    /(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)\s*(?:total|paid)/gi,
    /(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)/gi,
  ];

  for (const pattern of patterns) {
    const matches = [...content.matchAll(pattern)];
    if (matches.length > 0) {
      const amounts = matches.map(m => parseFloat(m[1].replace(/,/g, ''))).filter(a => a > 0);
      if (amounts.length > 0) {
        amount = Math.max(...amounts);
        break;
      }
    }
  }

  if (amount <= 0) return null;

  // Try to extract item description from subject
  let itemDesc = 'Amazon order';
  const itemMatch = subject.match(/(?:your\s*(?:order|delivery))\s*(?:of|for)\s*(.+?)(?:\s*has|\s*-|$)/i);
  if (itemMatch) itemDesc = itemMatch[1].trim().substring(0, 50);

  return {
    amount: amount,
    merchant: 'Amazon',
    description: itemDesc,
    category: categorizeByKeyword('amazon') || 'Shopping',
    date: message.getDate()
  };
});


// ── IDFC FIRST Bank card alert parser ────────────────────────
// Handles "Debit Alert" emails, e.g.:
//   "Transaction Successful! INR 2380.00 spent on your IDFC FIRST BANK
//    Credit Card ending XX2391 at <Merchant> on 22 JUN 2026."
// Targets the SPENT amount specifically so it never mistakes the
// "Available Limit" figure for the transaction value.
registerParser('idfc', function (message, source) {
  const subject = message.getSubject() || '';
  const body = message.getPlainBody() || '';
  const content = subject + '\n' + body;

  // Spent amount only (anchored on the word "spent").
  const amtMatch = content.match(/INR\s*([\d,]+(?:\.\d{1,2})?)\s*spent/i);
  const amount = amtMatch ? parseFloat(amtMatch[1].replace(/,/g, '')) : 0;
  if (amount <= 0) return null;

  // Merchant: the text between "at" and "on <date>".
  let merchant = 'IDFC Card';
  const merMatch = content.match(/\bat\s+(.+?)\s+on\s+\d/i);
  if (merMatch && merMatch[1].trim()) merchant = merMatch[1].trim();

  return {
    amount: amount,
    merchant: merchant,
    description: merchant,
    category: smartCategorize(merchant),
    date: message.getDate()
  };
});


// ── Helper functions ─────────────────────────────────────────

// One-time setup: registers the IDFC FIRST Bank debit-alert source and
// enables email ingestion. Run manually from the editor, then either run
// runEmailImportNow() to import now, or RESET() to install the trigger.
function addIdfcEmailSource() {
  const sources = getEmailSources();
  if (sources.some(s => (s.sender || '').toLowerCase().indexOf('idfcfirstbank.com') !== -1)) {
    const msg = 'IDFC email source already configured.';
    Logger.log(msg);
    return { added: false, message: msg };
  }
  sources.push({
    sender: 'noreply@idfcfirstbank.com',
    subjectPattern: 'Debit Alert',
    parserType: 'idfc',
    enabled: true
  });
  saveEmailSources(sources);
  if (getSettings().emailIngestionEnabled !== 'true') {
    saveSettings({ emailIngestionEnabled: 'true' });
  }
  const msg = 'Added IDFC email source and enabled email ingestion. ' +
    'Run runEmailImportNow() to import now, or RESET() to install the recurring trigger.';
  Logger.log(msg);
  logEvent('EMAIL', 'SUCCESS', 'IDFC email source added', '');
  return { added: true, message: msg };
}

function _extractMerchantFromSender(from) {
  // Extract domain or name from "Name <email@domain.com>" format
  const nameMatch = from.match(/^([^<]+)</);
  if (nameMatch) {
    const name = nameMatch[1].trim();
    if (name && name.length > 1 && name.length < 40) return name;
  }
  // Fallback: extract domain
  const domainMatch = from.match(/@([^.]+)\./);
  if (domainMatch) return domainMatch[1];
  return '';
}

function _extractFirstWord(text) {
  if (!text) return '';
  const words = text.trim().split(/\s+/);
  return words[0] || '';
}


// ============================================================
// MAIN PROCESSING PIPELINE
// ============================================================

// Entry point for the scheduled trigger.
// Reads enabled email sources, searches Gmail, parses, dedupes, saves.
function processEmailReceipts() {
  const startTime = Date.now();

  try {
    const settings = getSettings();
    if (settings.emailIngestionEnabled !== 'true') {
      Logger.log('processEmailReceipts: disabled in settings');
      return { skipped: true };
    }

    const sources = getEmailSources().filter(s => s.enabled);
    if (sources.length === 0) {
      Logger.log('processEmailReceipts: no enabled email sources');
      logEvent('EMAIL', 'SKIP', 'No enabled email sources configured', '');
      return { skipped: true, message: 'No enabled sources' };
    }

    let totalImported = 0;
    let totalSkipped = 0;
    let totalErrors = 0;

    sources.forEach(source => {
      try {
        const result = _processEmailSource(source);
        totalImported += result.imported;
        totalSkipped += result.skipped;
        totalErrors += result.errors;
      } catch (e) {
        Logger.log('processEmailReceipts source error (' + source.sender + '): ' + e.message);
        logEvent('EMAIL', 'ERROR', 'Source processing failed: ' + source.sender + ' — ' + e.message, '');
        totalErrors++;
      }
    });

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    Logger.log('processEmailReceipts: imported=' + totalImported + ' skipped=' + totalSkipped +
      ' errors=' + totalErrors + ' time=' + elapsed + 's');

    return { success: true, imported: totalImported, skipped: totalSkipped, errors: totalErrors, elapsed };
  } catch (e) {
    Logger.log('processEmailReceipts error: ' + e.message);
    logEvent('EMAIL', 'ERROR', 'processEmailReceipts: ' + e.message, '');
    return { success: false, message: e.message };
  }
}

// Process a single email source: search → parse → dedupe → save.
function _processEmailSource(source) {
  let imported = 0, skipped = 0, errors = 0;

  // Build Gmail search query
  // Search last 2 days to handle timezone edge cases and trigger gaps
  const query = 'from:(' + source.sender + ') subject:(' + source.subjectPattern + ') newer_than:2d';

  let threads;
  try {
    threads = GmailApp.search(query, 0, 20); // cap at 20 threads per source per cycle
  } catch (e) {
    Logger.log('Gmail search error for ' + source.sender + ': ' + e.message);
    logEvent('EMAIL', 'ERROR', 'Gmail search failed: ' + source.sender + ' — ' + e.message, '');
    return { imported: 0, skipped: 0, errors: 1 };
  }

  if (!threads || threads.length === 0) return { imported: 0, skipped: 0, errors: 0 };

  // Get the parser for this source type
  const parser = PARSER_REGISTRY[source.parserType.toLowerCase()] || PARSER_REGISTRY['generic'];
  if (!parser) {
    logEvent('EMAIL', 'ERROR', 'No parser found for type: ' + source.parserType, '');
    return { imported: 0, skipped: 0, errors: 1 };
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      try {
        const msgId = message.getId();

        // Duplicate check
        if (_isDuplicateExternal(msgId)) {
          skipped++;
          return; // skip this message
        }

        // Parse the email
        const transaction = parser(message, source);
        if (!transaction || !transaction.amount || transaction.amount <= 0) {
          logEvent('EMAIL', 'SKIP', 'Could not parse amount from: ' + source.sender + ' — ' + (message.getSubject() || '').substring(0, 50), msgId);
          skipped++;
          return;
        }

        // Save the transaction
        const tz = Session.getScriptTimeZone();
        const txDate = transaction.date
          ? Utilities.formatDate(new Date(transaction.date), tz, 'yyyy-MM-dd')
          : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

        const result = addExpenseWithSource({
          date: txDate,
          category: transaction.category || 'Other',
          description: transaction.description || transaction.merchant || source.sender,
          amount: transaction.amount,
          paymentMethod: 'Auto',
          notes: 'Email: ' + (message.getSubject() || '').substring(0, 60)
        }, 'EMAIL', msgId);

        if (result.success) {
          imported++;
          logEvent('EMAIL', 'SUCCESS',
            'Imported: ' + transaction.merchant + ' ₹' + transaction.amount + ' → ' + transaction.category,
            msgId);
        } else {
          errors++;
          logEvent('EMAIL', 'ERROR', 'Save failed: ' + (result.message || ''), msgId);
        }
      } catch (e) {
        errors++;
        logEvent('EMAIL', 'ERROR', 'Message processing error: ' + e.message, '');
      }
    });
  });

  return { imported, skipped, errors };
}


// ── Manual trigger — callable from Settings UI ───────────────
function runEmailImportNow() {
  return processEmailReceipts();
}
