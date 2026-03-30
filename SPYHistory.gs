/**
 * SPYHistory.gs — Fetches and stores daily SPY closing prices via the
 * Schwab Market Data (price history) API.
 *
 * Sheet layout: Date | Close ($) | Indexed (Base = 100)
 *
 * The "Indexed" column is normalized so that the earliest date = 100,
 * making it directly comparable to the indexed Portfolio series in the
 * Equity Curve chart.
 */

// ─────────────────────────────────────────────────────────────────
// Sheet bootstrap
// ─────────────────────────────────────────────────────────────────

function getOrCreateSPYSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SPY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SPY);
    const hdr = sheet.getRange(1, 1, 1, 3);
    hdr.setValues([['Date', 'Close ($)', 'Indexed (Base = 100)']]);
    hdr.setFontWeight('bold').setBackground('#cc0000').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}

// ─────────────────────────────────────────────────────────────────
// Public
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches all SPY daily closes from HISTORY_START to today using the
 * Schwab Market Data API and writes them to the SPY History sheet.
 *
 * The Schwab API caps a single pricehistory request at ~1825 days
 * (5 years), so we split into ≤5-year chunks when the range exceeds that.
 */
function fetchSPYHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Fetching SPY price history from Schwab…', 'Working', -1);

    const startDate = new Date(HISTORY_START);
    const endDate   = new Date();

    // Fetch in ≤5-year chunks
    const allCandles = [];
    for (let cur = new Date(startDate); cur < endDate; ) {
      const chunkEnd = new Date(cur);
      chunkEnd.setFullYear(chunkEnd.getFullYear() + 5);
      if (chunkEnd > endDate) chunkEnd.setTime(endDate.getTime());

      const resp = getPriceHistory(
        'SPY',
        cur.toISOString().slice(0, 10),
        chunkEnd.toISOString().slice(0, 10)
      );

      if (resp && Array.isArray(resp.candles)) {
        allCandles.push(...resp.candles);
      }

      cur = new Date(chunkEnd);
      cur.setDate(cur.getDate() + 1);
    }

    if (allCandles.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No SPY Data',
        'The Schwab API returned no candles. Check your authorization and try again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Deduplicate and sort by timestamp ascending
    const seen = new Set();
    const unique = allCandles
      .filter(c => {
        if (seen.has(c.datetime)) return false;
        seen.add(c.datetime);
        return true;
      })
      .sort((a, b) => a.datetime - b.datetime);

    writeSPYData_(unique);

    ss.toast(`✅ Fetched ${unique.length} SPY trading days (${HISTORY_START} → today).`, 'Complete', 10);
  } catch (e) {
    SpreadsheetApp.getUi().alert('SPY Fetch Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Writes sorted candle objects to the SPY History sheet.
 * @param {Array<{datetime: number, close: number}>} candles  Sorted ascending.
 */
function writeSPYData_(candles) {
  const sheet   = getOrCreateSPYSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const baseClose = candles[0].close;  // index to 100 at first trading day
  const rows = candles.map(c => [
    new Date(c.datetime),
    c.close,
    (c.close / baseClose) * 100,
  ]);

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('"$"#,##0.00');
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0.00');
}

/**
 * Returns a date-keyed map of SPY close prices for use in chart building.
 * @returns {Object}  { 'YYYY-MM-DD': closePrice }
 */
function getSPYCloseMap_() {
  const sheet = getOrCreateSPYSheet_();
  const data  = sheet.getDataRange().getValues().slice(1);
  const tz    = Session.getScriptTimeZone();
  const map   = {};
  data.forEach(row => {
    if (!row[0] || !row[1]) return;
    map[fmtDate_(new Date(row[0]), tz)] = row[1];
  });
  return map;
}
