/**
 * NetLiquidity.gs — Per-account Net Liquidity data management.
 *
 * Each account gets its own sheet tab: "Net Liq 418", "Net Liq 973", etc.
 * Sheet layout: Date | Net Liquidity ($) | Source
 *
 * Reconstruction approach (accurate):
 *   1. Fetches all transactions from HISTORY_START to today
 *   2. Walks backwards from today, reversing each trade to reconstruct
 *      historical position quantities at every date
 *   3. Fetches historical daily prices for every symbol ever held
 *   4. Daily Net Liq = reconstructed cash + Σ(qty × historical_close)
 *
 * Limitations:
 *   - Options positions are treated as pure cash flows (no qty tracking)
 *   - Stock splits / mergers may cause inaccuracies for affected symbols
 *   - Schwab transactions API may not return data before ~2 years ago;
 *     the earliest available date is shown in the log after running
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
    SpreadsheetApp.getUi().alert('Error – Account ' + suffix, e.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Public: accurate transaction-based reconstruction
// ─────────────────────────────────────────────────────────────────

function reconstructNetLiqHistoryForAccount(suffix) {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Reconstruct Net Liquidity — Account …' + suffix,
    'This will:\n' +
    '  1. Download all transactions from ' + HISTORY_START + ' to today\n' +
    '  2. Reconstruct historical position quantities by replaying trades\n' +
    '  3. Fetch historical daily prices for every symbol ever held\n' +
    '  4. Calculate daily Net Liq = cash + Σ(qty × price)\n\n' +
    'This takes 2–5 minutes depending on how many symbols you have traded.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // ── Step 1: current account state ─────────────────────────────
    ss.toast('Fetching current account state…', 'Working', -1);
    const hash          = getHashForSuffix_(suffix);
    const accountData   = getAccountDetails(hash);
    const sec           = accountData.securitiesAccount || accountData;
    const bal           = sec.currentBalances || {};
    const currentCash   = bal.cashBalance || 0;

    // Current equity positions: { symbol → longQuantity }
    const currentPositions = {};
    if (sec.positions) {
      sec.positions.forEach(function(pos) {
        const sym = pos.instrument && pos.instrument.symbol;
        const qty = pos.longQuantity || 0;
        if (sym && qty > 0 && pos.instrument.assetType === 'EQUITY') {
          currentPositions[sym] = qty;
        }
      });
    }
    console.log('Current cash: ' + currentCash);
    console.log('Current equity positions: ' + JSON.stringify(currentPositions));

    // ── Step 2: fetch all transactions ────────────────────────────
    ss.toast('Downloading transaction history…', 'Working', -1);
    const allTx = fetchAllTransactionsInChunks_(hash);
    console.log('Total transactions fetched: ' + allTx.length);

    if (allTx.length === 0) {
      throw new Error(
        'No transactions returned by the API.\n\n' +
        'Schwab may only provide 1–2 years of transaction history. ' +
        'Try setting HISTORY_START in Code.gs to a more recent date.'
      );
    }

    // ── Step 3: collect all equity symbols ever traded ────────────
    const tradedSymbols = new Set(Object.keys(currentPositions));
    allTx.forEach(function(tx) {
      var item = tx.transactionItem;
      if (!item || !item.instrument) return;
      if (item.instrument.assetType !== 'EQUITY') return;
      if (item.instrument.symbol) tradedSymbols.add(item.instrument.symbol);
    });
    console.log('Symbols to price: ' + JSON.stringify(Array.from(tradedSymbols)));

    // ── Step 4: fetch price history for every symbol ──────────────
    ss.toast('Fetching price history for ' + tradedSymbols.size + ' symbol(s)…', 'Working', -1);
    const priceMap = fetchPriceHistoriesForSymbols_(tradedSymbols);

    // ── Step 5: group transactions by date ────────────────────────
    const txByDate = groupTransactionsByDate_(allTx);

    // ── Step 6: reconstruct daily values backwards ────────────────
    ss.toast('Reconstructing daily portfolio values…', 'Working', -1);
    const dailyValues = reconstructDailyValues_(
      currentCash, currentPositions, txByDate, priceMap
    );

    const count = Object.keys(dailyValues).length;
    if (count === 0) {
      throw new Error('Reconstruction produced no data. Check the Apps Script logs for details.');
    }

    writeNetLiqData_(suffix, dailyValues, 'Reconstructed');

    const dates = Object.keys(dailyValues).sort();
    ss.toast(
      'Reconstructed ' + count + ' days  (' + dates[0] + ' → ' + dates[dates.length - 1] + ')',
      '✅ Account …' + suffix + ' Complete', 15
    );
    console.log('Reconstruction complete. Days: ' + count +
      '  Range: ' + dates[0] + ' → ' + dates[dates.length - 1]);

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
    .setWidth(480).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Net Liquidity from CSV');
}

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
// Core reconstruction engine
// ─────────────────────────────────────────────────────────────────

/**
 * Walks backwards day by day from today, reversing each trade to get
 * historical position quantities and cash balance, then values each day.
 *
 * @param {number}  currentCash        Today's cash balance from API
 * @param {Object}  currentPositions   { symbol: qty } from API
 * @param {Object}  txByDate           { 'YYYY-MM-DD': [tx, …] }
 * @param {Object}  priceMap           { symbol: { 'YYYY-MM-DD': closePrice } }
 * @returns {Object}  { 'YYYY-MM-DD': portfolioValue }
 */
function reconstructDailyValues_(currentCash, currentPositions, txByDate, priceMap) {
  const tz        = Session.getScriptTimeZone();
  const result    = {};
  const startDt   = new Date(HISTORY_START);
  const positions = copyObj_(currentPositions);
  var   cash      = currentCash;

  for (var d = new Date(); d >= startDt; d.setDate(d.getDate() - 1)) {
    var ds      = fmtDate_(d, tz);
    var dayTxs  = txByDate[ds] || [];

    // Reverse every transaction that happened on this day
    dayTxs.forEach(function(tx) {
      // Reverse the cash effect for ALL transaction types
      cash -= (tx.netAmount || 0);

      // For equity trades, reverse the position quantity change
      if (tx.type === 'TRADE' || tx.type === 'RECEIVE_AND_DELIVER') {
        var item = tx.transactionItem;
        if (!item || !item.instrument) return;
        if (item.instrument.assetType !== 'EQUITY') return;

        var sym = item.instrument.symbol;
        var qty = item.quantity || 0;

        if (item.instruction === 'BUY' || item.positionEffect === 'OPENING') {
          // Going forward we bought; reverse = remove those shares
          positions[sym] = (positions[sym] || 0) - qty;
        } else if (item.instruction === 'SELL' || item.positionEffect === 'CLOSING') {
          // Going forward we sold; reverse = add those shares back
          positions[sym] = (positions[sym] || 0) + qty;
        }
      }
    });

    // Value positions using historical close prices
    var positionValue = 0;
    Object.keys(positions).forEach(function(sym) {
      var qty = positions[sym];
      if (qty <= 0) return;
      var symPrices = priceMap[sym];
      if (!symPrices) return;

      // Use exact date price, or walk back up to 5 days for holidays/weekends
      var price = symPrices[ds];
      if (!price) {
        for (var back = 1; back <= 5; back++) {
          var prev = new Date(d);
          prev.setDate(prev.getDate() - back);
          price = symPrices[fmtDate_(prev, tz)];
          if (price) break;
        }
      }
      if (price) positionValue += qty * price;
    });

    var totalValue = cash + positionValue;
    if (totalValue > 0) result[ds] = Math.round(totalValue * 100) / 100;
  }

  return result;
}

// ─────────────────────────────────────────────────────────────────
// Data fetch helpers
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

/**
 * Fetches daily close prices for a set of symbols from HISTORY_START to today.
 * Returns { symbol: { 'YYYY-MM-DD': closePrice } }.
 * Symbols that fail (delisted, etc.) are logged and skipped.
 */
function fetchPriceHistoriesForSymbols_(symbolSet) {
  const result  = {};
  const tz      = Session.getScriptTimeZone();
  const today   = new Date().toISOString().slice(0, 10);
  const symbols = Array.from(symbolSet);

  symbols.forEach(function(sym) {
    try {
      var resp = getPriceHistory(sym, HISTORY_START, today);
      if (!resp.candles || resp.candles.length === 0) {
        console.warn('No price history returned for ' + sym);
        return;
      }
      result[sym] = {};
      resp.candles.forEach(function(c) {
        var ds = fmtDate_(new Date(c.datetime), tz);
        result[sym][ds] = c.close;
      });
      console.log(sym + ': ' + resp.candles.length + ' price points');
    } catch (e) {
      console.warn('Could not fetch prices for ' + sym + ': ' + e.message);
    }
  });

  return result;
}

function groupTransactionsByDate_(transactions) {
  const byDate = {};
  transactions.forEach(function(tx) {
    if (!tx.tradeDate) return;
    var ds = tx.tradeDate.substring(0, 10);
    if (!byDate[ds]) byDate[ds] = [];
    byDate[ds].push(tx);
  });
  return byDate;
}

// ─────────────────────────────────────────────────────────────────
// Sheet write / read helpers
// ─────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────
// Utilities
// ─────────────────────────────────────────────────────────────────

function fmtDate_(date, tz) {
  return Utilities.formatDate(date, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayStr_() {
  return fmtDate_(new Date());
}

function copyObj_(obj) {
  var copy = {};
  Object.keys(obj).forEach(function(k) { copy[k] = obj[k]; });
  return copy;
}
