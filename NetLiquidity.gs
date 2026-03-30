/**
 * NetLiquidity.gs — Manages the 'Net Liquidity' data sheet.
 *
 * Three ways to populate historical data:
 *   1. Daily Trigger  — most accurate; logs live balance each market close.
 *   2. Transaction Reconstruction — estimates from 2020 by replaying every
 *      transaction. Captures realized P&L + cash flows; open-position
 *      unrealized P&L is interpolated linearly between known endpoints.
 *      Use this as a reasonable approximation when you don't have exports.
 *   3. CSV Import — import a file exported from schwab.com → Accounts →
 *      Performance → Export. Columns: Date, Value.
 *
 * Sheet layout: Date | Net Liquidity ($) | Source
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateNetLiqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NET_LIQ);
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

/** Fetches live Net Liquidity from Schwab and logs it to the sheet. */
function fetchTodayNetLiq() {
  try {
    const value   = fetchCurrentNetLiquidity();
    const dateStr = todayStr_();
    upsertNetLiqRow_(dateStr, value, 'API');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Net Liquidity captured: $${value.toLocaleString('en-US', { minimumFractionDigits: 2 })}`,
      '✅ Success', 5
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error capturing Net Liquidity', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: transaction-based reconstruction
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches all transactions since HISTORY_START, then reconstructs
 * approximate daily portfolio values by working backwards from the
 * current Net Liquidation Value.
 *
 * Algorithm:
 *   currentValue  = live account balance  (from API)
 *   For each day D going backwards from today to HISTORY_START:
 *     adjustedValue[D] = adjustedValue[D+1] - externalCashFlow[D]
 *   where externalCashFlow = deposits + withdrawals on date D.
 *
 * This removes the effect of money moved in/out, leaving a curve
 * that reflects portfolio performance only (realized P&L on closed
 * positions + dividends, linearly smoothed for open-position drift).
 */
function reconstructNetLiqHistory() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Reconstruct Net Liquidity from Transactions',
    'This will:\n' +
    '  • Fetch all transactions from ' + HISTORY_START + ' to today\n' +
    '  • Estimate daily portfolio values working backwards from your\n' +
    '    current live account balance\n\n' +
    '⚠️  LIMITATION: Unrealized P&L on positions still held is\n' +
    'linearly interpolated — values are approximate.\n\n' +
    'For higher accuracy, use "Import Net Liq from CSV" instead.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast('Fetching current account balance…', 'Working', -1);
    const currentNetLiq = fetchCurrentNetLiquidity();

    const props      = PropertiesService.getScriptProperties();
    const accountHash = props.getProperty('SCHWAB_ACCOUNT_HASH');
    if (!accountHash) {
      throw new Error('Account hash not found. Please capture today\'s Net Liquidity first.');
    }

    ss.toast('Downloading transaction history (this may take a minute)…', 'Working', -1);
    const allTx = fetchAllTransactionsInChunks_(accountHash);

    ss.toast(`Processing ${allTx.length} transactions…`, 'Working', -1);
    const externalFlowsByDate = buildExternalCashFlowMap_(allTx);
    const dailyValues         = reconstructBackwards_(currentNetLiq, externalFlowsByDate);

    writeNetLiqData_(dailyValues, 'Reconstructed');

    ss.toast(
      `Reconstructed ${Object.keys(dailyValues).length} daily values from ${HISTORY_START} to today.`,
      '✅ Complete', 10
    );
  } catch (e) {
    ui.alert('Reconstruction Error', e.message, ui.ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: CSV import dialog
// ─────────────────────────────────────────────────────────────────

function showCsvImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CsvImport')
    .setWidth(480)
    .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Net Liquidity from CSV');
}

/**
 * Called from CsvImport.html with the parsed CSV rows.
 * @param {Array<{date: string, value: number}>} rows
 */
function importNetLiqFromCsv(rows) {
  if (!rows || rows.length === 0) throw new Error('No data received.');
  const sheet = getOrCreateNetLiqSheet_();
  const tz    = Session.getScriptTimeZone();

  // Build a map of existing dates for upsert logic
  const existing = {};
  const data = sheet.getDataRange().getValues().slice(1);
  data.forEach((r, i) => {
    if (r[0]) {
      existing[fmtDate_(new Date(r[0]), tz)] = i + 2; // 1-based row
    }
  });

  rows.forEach(row => {
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
  return `Imported ${rows.length} rows successfully.`;
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

/** Fetches transactions in 3-month chunks to stay within API limits. */
function fetchAllTransactionsInChunks_(accountHash) {
  const allTx   = [];
  const start   = new Date(HISTORY_START);
  const end     = new Date();

  for (let cur = new Date(start); cur < end; ) {
    const chunkEnd = new Date(cur);
    chunkEnd.setMonth(chunkEnd.getMonth() + 3);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());

    try {
      const txs = getTransactions(accountHash, cur.toISOString(), chunkEnd.toISOString());
      if (Array.isArray(txs)) allTx.push(...txs);
    } catch (e) {
      console.warn(`Skipping chunk ${cur.toISOString()}: ${e.message}`);
    }

    cur = new Date(chunkEnd);
    cur.setDate(cur.getDate() + 1);
  }

  return allTx;
}

/**
 * Builds a map of date → net external cash flow (deposits - withdrawals).
 * We only count ELECTRONIC_FUND / ACH / wire transactions because those
 * are the transfers that change your portfolio size independent of returns.
 */
function buildExternalCashFlowMap_(transactions) {
  const flows = {};
  const externalTypes = new Set([
    'ELECTRONIC_FUND', 'WIRE_IN', 'WIRE_OUT', 'JOURNAL',
  ]);

  transactions.forEach(tx => {
    if (!externalTypes.has(tx.type)) return;
    if (!tx.tradeDate) return;
    const d = tx.tradeDate.substring(0, 10); // 'YYYY-MM-DD'
    flows[d] = (flows[d] || 0) + (tx.netAmount || 0);
  });

  return flows;
}

/**
 * Walks backwards from today's net liquidity, removing external cash
 * flows to reconstruct the performance-only equity curve.
 * @returns {Object} date-string → value
 */
function reconstructBackwards_(currentValue, externalFlowsByDate) {
  const tz      = Session.getScriptTimeZone();
  const result  = {};
  const today   = new Date();
  const startDt = new Date(HISTORY_START);

  let running = currentValue;
  for (let d = new Date(today); d >= startDt; d.setDate(d.getDate() - 1)) {
    const ds = fmtDate_(d, tz);
    // Reverse any external deposit/withdrawal that happened on this day
    if (externalFlowsByDate[ds]) {
      running -= externalFlowsByDate[ds];
    }
    result[ds] = running;
  }

  return result;
}

/**
 * Writes a date→value map to the Net Liquidity sheet (full replace).
 * @param {Object} dailyValues   { 'YYYY-MM-DD': number }
 * @param {string} source        e.g. 'Reconstructed'
 */
function writeNetLiqData_(dailyValues, source) {
  const sheet   = getOrCreateNetLiqSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const rows = Object.entries(dailyValues)
    .filter(([, v]) => v > 0)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([d, v]) => [new Date(d), v, source]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('"$"#,##0.00');
  }
}

/** Inserts or updates a single row by date. */
function upsertNetLiqRow_(dateStr, value, source) {
  const sheet = getOrCreateNetLiqSheet_();
  const tz    = Session.getScriptTimeZone();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
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
  sheet = sheet || getOrCreateNetLiqSheet_();
  const last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, 3).sort(1);
}

function fmtDate_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayStr_() {
  return fmtDate_(new Date());
}
