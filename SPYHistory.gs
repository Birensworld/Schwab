/**
 * SPYHistory.gs — Fetches and stores daily SPY and QQQ closing prices.
 *
 * Sheet layout (SHEET_SPY):
 *   Date | SPY Close ($) | SPY Indexed (Base=100) | QQQ Close ($) | QQQ Indexed (Base=100)
 *
 * Both series are indexed to 100 at the first date with data for both symbols.
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateSPYSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SPY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SPY);
    const hdr = sheet.getRange(1, 1, 1, 5);
    hdr.setValues([[
      'Date',
      'SPY Close ($)', 'SPY Indexed (Base=100)',
      'QQQ Close ($)', 'QQQ Indexed (Base=100)',
    ]]);
    hdr.setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    [2, 3, 4, 5].forEach(function(c) { sheet.setColumnWidth(c, 160); });
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
// Public
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches daily closes for SPY and QQQ from HISTORY_START to today
 * and writes them to the SPY History sheet.
 */
function fetchSPYHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Fetching SPY and QQQ price history…', 'Working', -1);

    const today   = new Date().toISOString().slice(0, 10);
    const spyData = fetchCandlesForSymbol_('SPY', HISTORY_START, today);
    const qqqData = fetchCandlesForSymbol_('QQQ', HISTORY_START, today);

    if (spyData.length === 0) {
      SpreadsheetApp.getUi().alert('No SPY data returned. Check authorization and try again.');
      return;
    }

    // Build date-keyed maps
    const tz = Session.getScriptTimeZone();
    const spyMap = {}, qqqMap = {};
    spyData.forEach(function(c) { spyMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });
    qqqData.forEach(function(c) { qqqMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });

    // All dates that appear in SPY (QQQ fills in where available)
    const allDates = Object.keys(spyMap).sort();

    // Base = first trading day on or after HISTORY_START (Jan 1 2026 is a holiday;
    // the market opens Jan 2 2026, which will be the first candle in the data)
    const baseDate = allDates.find(function(d) { return d >= HISTORY_START; }) || allDates[0];
    const baseSPY  = spyMap[baseDate];
    const baseQQQ  = qqqMap[baseDate] || null;

    const rows = allDates.map(function(d) {
      const spyClose = spyMap[d];
      const qqqClose = qqqMap[d] || '';
      return [
        new Date(d),
        spyClose,
        roundTo2_(spyClose / baseSPY * 100),
        qqqClose,
        qqqClose && baseQQQ ? roundTo2_(qqqClose / baseQQQ * 100) : '',
      ];
    });

    const sheet = getOrCreateSPYSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('"$"#,##0.00');
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0.00');
    sheet.getRange(2, 4, rows.length, 1).setNumberFormat('"$"#,##0.00');
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('0.00');

    backupSPYHistory_();
    ss.toast(
      '✅ SPY: ' + spyData.length + ' days  |  QQQ: ' + qqqData.length + ' days',
      'Benchmarks Updated', 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Fetch Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

/** Fetches deduplicated, sorted daily candles for one symbol. */
function fetchCandlesForSymbol_(symbol, startDate, endDate) {
  const allCandles = [];
  const start = new Date(startDate);
  const end   = new Date(endDate);

  for (var cur = new Date(start); cur <= end; ) {
    var chunkEnd = new Date(cur);
    chunkEnd.setFullYear(chunkEnd.getFullYear() + 5);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());

    try {
      var resp = getPriceHistory(symbol, cur.toISOString().slice(0, 10),
                                          chunkEnd.toISOString().slice(0, 10));
      if (resp && Array.isArray(resp.candles)) {
        Array.prototype.push.apply(allCandles, resp.candles);
      }
    } catch (e) {
      console.warn(symbol + ' chunk error: ' + e.message);
    }

    cur = new Date(chunkEnd);
    cur.setDate(cur.getDate() + 1);
  }

  // Deduplicate and sort
  var seen = {};
  return allCandles
    .filter(function(c) {
      if (seen[c.datetime]) return false;
      seen[c.datetime] = true;
      return true;
    })
    .sort(function(a, b) { return a.datetime - b.datetime; });
}

/**
 * Returns date-keyed close price maps for SPY and QQQ from the sheet.
 * The chart builder uses these raw closes; it re-normalizes to 100 at
 * the first date where all series (portfolio + both benchmarks) overlap.
 * @returns {{ spy: {date: price}, qqq: {date: price} }}
 */
function getBenchmarkCloseMaps_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SPY);
  if (!sheet) return { spy: {}, qqq: {} };

  const tz  = Session.getScriptTimeZone();
  const spy = {}, qqq = {};

  sheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if (!row[0] || !row[1]) return;
    var d = fmtDate_(new Date(row[0]), tz);
    spy[d] = parseFloat(row[1]) || 0;
    if (row[3]) qqq[d] = parseFloat(row[3]) || 0;
  });

  return { spy: spy, qqq: qqq };
}

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}
