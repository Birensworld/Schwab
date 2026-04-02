/**
 * NetLiquidity.gs — Per-account Net Liquidity data management.
 *
 * Each account gets its own sheet tab: "Net Liq 418", "Net Liq 973", etc.
 * Sheet layout: Date | Net Liquidity ($) | Source
 *
 * Three ways to populate historical data:
 *   1. Daily Trigger  — most accurate; logs live balance each market close.
 *   2. Transaction Reconstruction — estimates from HISTORY_START by replaying
 *      every transaction backwards from today's live balance.
 *      Captures external cash flows (deposits/withdrawals) only;
 *      unrealized P&L on open positions is linearly interpolated.
 *   3. CSV Import — import from schwab.com → Accounts → Performance → Export.
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateNetLiqSheet_(suffix) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheetNames_(suffix).netLiq;
  let sheet       = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const hdr = sheet.getRange(1, 1, 1, 3);
    hdr.setValues([['Date', 'Net Liquidity ($)', 'Source']]);
    hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 110);
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
// Public: capture today's value
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches live Net Liquidity from Schwab for one account and logs it.
 * @param {string} suffix  e.g. '418'
 */
function fetchTodayNetLiqForAccount(suffix) {
  try {
    const value   = fetchNetLiqForSuffix_(suffix);
    const dateStr = todayStr_();
    upsertNetLiqRow_(suffix, dateStr, value, 'API');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Account …' + suffix + '  Net Liquidity: $' +
        value.toLocaleString('en-US', { minimumFractionDigits: 2 }),
      '✅ Captured', 5
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error – Account ' + suffix, e.message, SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: transaction-based reconstruction
// ─────────────────────────────────────────────────────────────────

/**
 * Reconstructs approximate daily portfolio values for one account
 * by replaying all transactions backwards from today's live balance.
 * @param {string} suffix  e.g. '418'
 */
function reconstructNetLiqHistoryForAccount(suffix) {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Reconstruct Net Liquidity — Account …' + suffix,
    'This will:\n' +
    '  • Fetch all transactions from ' + HISTORY_START + ' to today\n' +
    '  • Estimate daily values working backwards from your current live balance\n\n' +
    '⚠️  LIMITATION: Only external cash flows (deposits/withdrawals) are\n' +
    'removed. Unrealized P&L on open positions is linearly interpolated.\n\n' +
    'For higher accuracy use "Import Net Liq from CSV" instead.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Fetching current balance for account …' + suffix + '…', 'Working', -1);
    const currentNetLiq = fetchNetLiqForSuffix_(suffix);
    const hash          = getHashForSuffix_(suffix);

    ss.toast('Downloading transaction history (this may take a minute)…', 'Working', -1);
    const allTx = fetchAllTransactionsInChunks_(hash);

    ss.toast('Processing ' + allTx.length + ' transactions…', 'Working', -1);
    const externalFlows = buildExternalCashFlowMap_(allTx);
    const dailyValues   = reconstructBackwards_(currentNetLiq, externalFlows);

    writeNetLiqData_(suffix, dailyValues, 'Reconstructed');

    ss.toast(
      'Reconstructed ' + Object.keys(dailyValues).length +
        ' daily values for account …' + suffix + '.',
      '✅ Complete', 10
    );
  } catch (e) {
    ui.alert('Reconstruction Error – Account ' + suffix, e.message, ui.ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: CSV import
// ─────────────────────────────────────────────────────────────────

function showCsvImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CsvImport')
    .setWidth(480)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Net Liquidity from CSV');
}

/**
 * Called from CsvImport.html.
 * @param {Array<{date: string, value: number}>} rows
 * @param {string} suffix  account suffix, e.g. '418'
 */
function importNetLiqFromCsv(rows, suffix) {
  suffix = suffix || '418';
  if (!rows || rows.length === 0) throw new Error('No data received.');

  const sheet = getOrCreateNetLiqSheet_(suffix);
  const tz    = Session.getScriptTimeZone();

  const existing = {};
  sheet.getDataRange().getValues().slice(1).forEach(function(r, i) {
    if (r[0]) existing[fmtDate_(new Date(r[0]), tz)] = i + 2;
  });

  rows.forEach(function(row) {
    const d = new Date(row.date);
    if (isNaN(d.getTime())) return;
    const ds = fmtDate_(d, tz);
    if (existing[ds]) {
      sheet.getRange(existing[ds], 2, 1, 2).setValues([[row.value, 'CSV Import']]);
    } else {
      sheet.appendRow([d, row.value, 'CSV Import']);
    }
  });

  sortNetLiqSheet_(sheet);
  return 'Imported ' + rows.length + ' rows into account …' + suffix + '.';
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

function fetchAllTransactionsInChunks_(accountHash) {
  const allTx = [];
  const start = new Date(HISTORY_START);
  const end   = new Date();

  for (var cur = new Date(start); cur < end; ) {
    var chunkEnd = new Date(cur);
    chunkEnd.setMonth(chunkEnd.getMonth() + 3);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());

    try {
      var txs = getTransactions(accountHash, cur.toISOString(), chunkEnd.toISOString());
      if (Array.isArray(txs)) Array.prototype.push.apply(allTx, txs);
    } catch (e) {
      console.warn('Skipping chunk ' + cur.toISOString() + ': ' + e.message);
    }

    cur = new Date(chunkEnd);
    cur.setDate(cur.getDate() + 1);
  }

  return allTx;
}

function buildExternalCashFlowMap_(transactions) {
  const flows        = {};
  const externalTypes = { ELECTRONIC_FUND: true, WIRE_IN: true, WIRE_OUT: true, JOURNAL: true };

  transactions.forEach(function(tx) {
    if (!externalTypes[tx.type]) return;
    if (!tx.tradeDate) return;
    var d = tx.tradeDate.substring(0, 10);
    flows[d] = (flows[d] || 0) + (tx.netAmount || 0);
  });

  return flows;
}

function reconstructBackwards_(currentValue, externalFlowsByDate) {
  const tz      = Session.getScriptTimeZone();
  const result  = {};
  const startDt = new Date(HISTORY_START);

  var running = currentValue;
  for (var d = new Date(); d >= startDt; d.setDate(d.getDate() - 1)) {
    var ds = fmtDate_(d, tz);
    if (externalFlowsByDate[ds]) running -= externalFlowsByDate[ds];
    result[ds] = running;
  }

  return result;
}

function writeNetLiqData_(suffix, dailyValues, source) {
  const sheet   = getOrCreateNetLiqSheet_(suffix);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const rows = Object.entries(dailyValues)
    .filter(function(e) { return e[1] > 0; })
    .sort(function(a, b) { return a[0].localeCompare(b[0]); })
    .map(function(e) { return [new Date(e[0]), e[1], source]; });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('"$"#,##0.00');
  }
}

function upsertNetLiqRow_(suffix, dateStr, value, source) {
  const sheet = getOrCreateNetLiqSheet_(suffix);
  const tz    = Session.getScriptTimeZone();
  const data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (fmtDate_(new Date(data[i][0]), tz) === dateStr) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[value, source]]);
      return;
    }
  }
  sheet.appendRow([new Date(dateStr), value, source]);
  sortNetLiqSheet_(sheet);
}

function sortNetLiqSheet_(sheet) {
  var last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, 3).sort(1);
}

function fmtDate_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayStr_() {
  return fmtDate_(new Date());
}

/**
 * Returns a date-keyed map of Net Liquidity values for use in chart building.
 * @param {string} suffix  e.g. '418'
 * @returns {Object}  { 'YYYY-MM-DD': number }
 */
function getNetLiqMap_(suffix) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetNames_(suffix).netLiq);
  if (!sheet) return {};

  const tz  = Session.getScriptTimeZone();
  const map = {};
  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (!row[0] || !row[1]) return;
    var v = parseFloat(row[1]);
    if (isNaN(v) || v <= 0) return;
    map[fmtDate_(new Date(row[0]), tz)] = v;
  });
  return map;
}
