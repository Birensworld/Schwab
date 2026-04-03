/**
 * SPYHistory.gs — Fetches and stores daily SPY and QQQ closing prices.
 * Version: 1.6 (2026-04-03) — Fix: parse lastDateStr as local midnight for correct nextDay calculation.
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
 * Fetches daily closes for SPY and QQQ and writes them to the SPY History sheet.
 *
 * Incremental mode (sheet already has data):
 *   - Reads the last date from the sheet
 *   - Fetches only dates after that → appends new rows
 *   - Base prices are read from the first data row (where indexed = 100)
 *
 * Full mode (sheet is empty):
 *   - Fetches from HISTORY_START to today and writes all rows
 */
function fetchSPYHistory() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date().toISOString().slice(0, 10);
  const tz    = Session.getScriptTimeZone();

  try {
    ss.toast('Fetching SPY and QQQ price history…', 'Working', -1);

    const sheet   = getOrCreateSPYSheet_();
    const lastRow = sheet.getLastRow();

    // ── Determine fetch window and base prices ────────────────────
    var fetchStart, baseSPY, baseQQQ, fullRewrite;
    var existingDates = {};   // populated in incremental mode for dedup guard

    if (lastRow > 1) {
      // Incremental: only fetch dates after the last row.
      const dataRows = lastRow - 1;

      // Read numeric values for base prices (row 2 = base date, indexed = 100)
      const firstRow = sheet.getRange(2, 1, 1, 4).getValues()[0];
      baseSPY = firstRow[1];          // SPY Close ($)
      baseQQQ = firstRow[3] || null;  // QQQ Close ($)

      // Use getDisplayValues() to read the last date as the string shown in the
      // cell (e.g. '2026-04-02'). This avoids timezone shift issues that occur
      // when converting a stored Date serial back through fmtDate_().
      const lastDateStr = sheet.getRange(lastRow, 1).getDisplayValue(); // 'yyyy-mm-dd'

      // Build a set of all dates already in the sheet for the dedup guard below
      const existingDates = {};
      sheet.getRange(2, 1, dataRows, 1).getDisplayValues().forEach(function(r) {
        if (r[0]) existingDates[r[0]] = true;
      });

      // Parse lastDateStr ('yyyy-mm-dd') as LOCAL midnight to avoid UTC
      // timezone shift in setDate()+1 calculation.
      var lp = lastDateStr.split('-');
      var nextDay = new Date(+lp[0], +lp[1] - 1, +lp[2] + 1);
      fetchStart  = fmtDate_(nextDay, tz);
      fullRewrite = false;

      if (fetchStart > today) {
        ss.toast('SPY/QQQ history is already up to date.', 'No Update Needed', 5);
        return;
      }
    } else {
      // Full fetch from scratch
      fetchStart  = HISTORY_START;
      fullRewrite = true;
    }

    // ── Fetch from Schwab ─────────────────────────────────────────
    const spyData = fetchCandlesForSymbol_('SPY', fetchStart, today);
    const qqqData = fetchCandlesForSymbol_('QQQ', fetchStart, today);

    if (spyData.length === 0) {
      ss.toast('No new SPY data since ' + fetchStart + '.', 'Already Up To Date', 5);
      return;
    }

    const spyMap = {}, qqqMap = {};
    spyData.forEach(function(c) { spyMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });
    qqqData.forEach(function(c) { qqqMap[fmtDate_(new Date(c.datetime), tz)] = c.close; });

    const newDates = Object.keys(spyMap).sort();

    if (fullRewrite) {
      // Establish base prices from the first trading day
      const baseDate = newDates.find(function(d) { return d >= HISTORY_START; }) || newDates[0];
      baseSPY = spyMap[baseDate];
      baseQQQ = qqqMap[baseDate] || null;
      // Clear all existing data rows
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    }

    // ── Build and write new rows ──────────────────────────────────
    // Dedup guard: skip any date already present in the sheet
    const datesToWrite = fullRewrite
      ? newDates
      : newDates.filter(function(d) { return !existingDates[d]; });

    if (datesToWrite.length === 0) {
      ss.toast('SPY/QQQ history is already up to date.', 'No Update Needed', 5);
      return;
    }

    const newRows = datesToWrite.map(function(d) {
      const spyClose = spyMap[d];
      const qqqClose = qqqMap[d] || '';
      // Parse as LOCAL midnight so reading back with fmtDate_() always
      // returns the same YYYY-MM-DD regardless of the script timezone.
      // new Date('YYYY-MM-DD') is UTC midnight and shifts to previous day in US timezones.
      var p = d.split('-');
      var localDate = new Date(+p[0], +p[1] - 1, +p[2]);
      return [
        localDate,
        spyClose,
        roundTo2_(spyClose / baseSPY * 100),
        qqqClose,
        qqqClose && baseQQQ ? roundTo2_(qqqClose / baseQQQ * 100) : '',
      ];
    });

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    sheet.getRange(startRow, 1, newRows.length, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(startRow, 2, newRows.length, 1).setNumberFormat('"$"#,##0.00');
    sheet.getRange(startRow, 3, newRows.length, 1).setNumberFormat('0.00');
    sheet.getRange(startRow, 4, newRows.length, 1).setNumberFormat('"$"#,##0.00');
    sheet.getRange(startRow, 5, newRows.length, 1).setNumberFormat('0.00');

    backupSPYHistory_();
    ss.toast(
      '✅ Added ' + newRows.length + ' new day(s) — SPY & QQQ',
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
