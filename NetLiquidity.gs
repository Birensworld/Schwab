/**
 * NetLiquidity.gs — Manages the shared "Net Liquidity" sheet.
 *
 * Sheet layout:
 *   Date | Net Liq 418 ($) | Net Liq 973 ($) | Source
 *
 * Column positions are determined by ACCOUNT_ORDER in Code.gs.
 * All accounts share one sheet; each has its own column.
 *
 * Daily trigger writes only if no value exists yet for that date
 * (preserves manually entered values).
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateNetLiqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NET_LIQ);

    // Build header row: Date | Net Liq 418 ($) | Net Liq 973 ($) | Source
    var headers = ['Date'];
    ACCOUNT_ORDER.forEach(function(s) { headers.push('Net Liq ' + s + ' ($)'); });
    headers.push('Source');

    var hdrRange = sheet.getRange(1, 1, 1, headers.length);
    hdrRange.setValues([headers]);
    hdrRange.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    ACCOUNT_ORDER.forEach(function(_, i) { sheet.setColumnWidth(i + 2, 180); });
    sheet.setColumnWidth(ACCOUNT_ORDER.length + 2, 110); // Source col
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
// Public: capture today's value
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches live Net Liquidity from Schwab for one account and logs it.
 * @param {string}  suffix        e.g. '418'
 * @param {boolean} skipIfExists  If true, don't overwrite an existing value for today
 */
function fetchTodayNetLiqForAccount(suffix, skipIfExists) {
  try {
    var value   = fetchNetLiqForSuffix_(suffix);
    var dateStr = todayStr_();
    upsertNetLiqRow_(suffix, dateStr, value, 'API', skipIfExists);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Account …' + suffix + '  Net Liq: $' +
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
// Public: CSV import
// ─────────────────────────────────────────────────────────────────

function showCsvImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('CsvImport')
    .setWidth(480).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Net Liquidity from CSV');
}

function importNetLiqFromCsv(rows, suffix) {
  suffix = suffix || '418';
  if (!rows || rows.length === 0) throw new Error('No data received.');
  var tz = Session.getScriptTimeZone();
  rows.forEach(function(row) {
    var d = new Date(row.date);
    if (isNaN(d.getTime())) return;
    upsertNetLiqRow_(suffix, fmtDate_(d, tz), row.value, 'CSV Import', false);
  });
  return 'Imported ' + rows.length + ' rows into account …' + suffix + '.';
}

// ─────────────────────────────────────────────────────────────────
// Diagnostics
// ─────────────────────────────────────────────────────────────────

/**
 * Run from the Apps Script editor to validate reconstruction accuracy.
 * Check View → Logs after running.
 */
function debugReconstruction() {
  var suffix = '418';
  var hash   = getHashForSuffix_(suffix);

  var accountData  = getAccountDetails(hash);
  var sec          = accountData.securitiesAccount || accountData;
  var bal          = sec.currentBalances || {};
  var cash         = bal.cashBalance || 0;
  var actualNetLiq = extractNetLiquidityFromAccount_(accountData);

  var positions = {};
  if (sec.positions) {
    sec.positions.forEach(function(pos) {
      var sym = pos.instrument && pos.instrument.symbol;
      var qty = pos.longQuantity || 0;
      if (sym && qty > 0) positions[sym] = qty;
    });
  }

  console.log('=== Starting state ===');
  console.log('cashBalance: ' + cash);
  console.log('Actual Net Liq (API): ' + actualNetLiq);
  console.log('Positions: ' + JSON.stringify(positions));

  var end   = new Date();
  var start = new Date();
  start.setDate(start.getDate() - 30);
  var txs = getTransactions(hash, start.toISOString(), end.toISOString());

  console.log('\n=== Last 30 days: ' + txs.length + ' transactions ===');
  txs.slice(0, 5).forEach(function(tx, i) {
    console.log('\nTransaction ' + (i + 1) + ':');
    console.log('  type: '        + tx.type);
    console.log('  tradeDate: '   + tx.tradeDate);
    console.log('  netAmount: '   + tx.netAmount);
    var item = tx.transactionItem;
    if (item) {
      console.log('  instruction: '   + item.instruction);
      console.log('  positionEffect: '+ item.positionEffect);
      console.log('  quantity: '      + item.quantity);
      var inst = item.instrument;
      if (inst) {
        console.log('  symbol: '    + inst.symbol);
        console.log('  assetType: ' + inst.assetType);
      }
    }
  });
}

// ─────────────────────────────────────────────────────────────────
// Sheet read / write helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Writes or updates a single date row for one account column.
 * @param {string}  suffix        Account suffix, e.g. '418'
 * @param {string}  dateStr       'YYYY-MM-DD'
 * @param {number}  value         Net Liquidity value
 * @param {string}  source        'API', 'Manual', 'CSV Import', etc.
 * @param {boolean} skipIfExists  If true and a non-empty value already exists, do nothing
 */
function upsertNetLiqRow_(suffix, dateStr, value, source, skipIfExists) {
  var sheet  = getOrCreateNetLiqSheet_();
  var tz     = Session.getScriptTimeZone();
  var col    = netLiqCol_(suffix);                   // column for this account
  var srcCol = ACCOUNT_ORDER.length + 2;             // Source is last column
  var data   = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (fmtDate_(new Date(data[i][0]), tz) !== dateStr) continue;

    // Row found for this date
    var existing = data[i][col - 1];
    if (skipIfExists && existing !== '' && existing !== 0 && existing !== null) return;
    sheet.getRange(i + 1, col).setValue(value);
    sheet.getRange(i + 1, srcCol).setValue(source);
    return;
  }

  // No row for this date — append a new one
  var newRow = [new Date(dateStr)];
  ACCOUNT_ORDER.forEach(function(s) {
    newRow.push(s === suffix ? value : '');
  });
  newRow.push(source);
  sheet.appendRow(newRow);
  sortNetLiqSheet_(sheet);
}

function sortNetLiqSheet_(sheet) {
  var last = sheet.getLastRow();
  if (last > 2) sheet.getRange(2, 1, last - 1, ACCOUNT_ORDER.length + 2).sort(1);
}

/**
 * Returns a date-keyed map of Net Liquidity values for one account.
 * Used by the chart builder.
 * @param {string} suffix  e.g. '418'
 * @returns {Object}  { 'YYYY-MM-DD': number }
 */
function getNetLiqMap_(suffix) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) return {};

  var tz  = Session.getScriptTimeZone();
  var col = netLiqCol_(suffix) - 1;  // 0-based for array indexing
  var map = {};

  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (!row[0]) return;
    var v = parseFloat(row[col]);
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

function todayStr_() { return fmtDate_(new Date()); }
