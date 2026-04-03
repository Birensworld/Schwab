/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 *
 * Each chart plots three series indexed to 100 at the first date where
 * all three have data:
 *   • Portfolio (Net Liq for that account)
 *   • SPY
 *   • QQQ
 *
 * Chart sheet: "Equity Curve 418", "Equity Curve 973", etc.
 * Data layout: row 1 = header, rows 2+ = Date | Portfolio | SPY | QQQ (all indexed)
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point
// ─────────────────────────────────────────────────────────────────

function buildEquityCurveChartForAccount(suffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building equity curve for account …' + suffix + '…', 'Working', -1);

    // ── 1. Load data ───────────────────────────────────────────────
    var netLiqMap   = getNetLiqMap_(suffix);
    var benchmarks  = getBenchmarkCloseMaps_();
    var spyMap      = benchmarks.spy;
    var qqqMap      = benchmarks.qqq;

    if (Object.keys(netLiqMap).length === 0) {
      throw new Error(
        'No Net Liquidity data found for account …' + suffix + '.\n\n' +
        'Enter values in the "' + SHEET_NET_LIQ + '" sheet first.'
      );
    }
    if (Object.keys(spyMap).length === 0) {
      throw new Error('No SPY/QQQ data found. Run "Fetch SPY + QQQ History" first.');
    }

    // ── 2. Find dates where portfolio AND SPY both have values ─────
    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d]; })
      .sort();

    if (commonDates.length < 2) {
      throw new Error(
        'Not enough overlapping dates between Net Liq …' + suffix + ' and SPY.\n' +
        'Make sure both datasets share common trading days.'
      );
    }

    // ── 3. Normalize to 100 at first common date ──────────────────
    var baseDate   = commonDates[0];
    var baseNetLiq = netLiqMap[baseDate];
    var baseSPY    = spyMap[baseDate];
    var baseQQQ    = qqqMap[baseDate] || null;

    var rows = commonDates.map(function(d) {
      // Write date as a plain string so the chart treats col A as text
      // categories (avoids Date serial number domain detection issues).
      var qqqIndexed = (baseQQQ && qqqMap[d]) ? roundTo2_(qqqMap[d] / baseQQQ * 100) : 0;
      return [
        d,                                             // 'YYYY-MM-DD' string
        roundTo2_(netLiqMap[d] / baseNetLiq * 100),
        roundTo2_(spyMap[d]    / baseSPY    * 100),
        qqqIndexed,
      ];
    });

    console.log('Common dates: ' + commonDates.length +
      '  First: ' + commonDates[0] + '  Last: ' + commonDates[commonDates.length - 1]);
    console.log('Sample row[0]: ' + JSON.stringify(rows[0]));

    // ── 4. Write to chart sheet ───────────────────────────────────
    var chartSheet = getOrCreateChartSheet_(ss, suffix);
    writeChartData_(chartSheet, rows, suffix, baseDate, baseNetLiq, baseSPY, baseQQQ);

    // Flush ensures all cell writes are committed before the chart
    // builder reads the range — without this the chart sees empty cells.
    SpreadsheetApp.flush();

    // ── 5. Draw chart ─────────────────────────────────────────────
    chartSheet.getCharts().forEach(function(c) { chartSheet.removeChart(c); });
    insertLineChart_(chartSheet, rows.length, suffix);

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ ' + rows.length + ' days  (' + baseDate + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Equity Curve …' + suffix, 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error – Account ' + suffix, e.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Sheet setup
// ─────────────────────────────────────────────────────────────────

function getOrCreateChartSheet_(ss, suffix) {
  var name  = chartSheetName_(suffix);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function writeChartData_(sheet, rows, suffix, baseDate, baseNetLiq, baseSPY, baseQQQ) {
  // ── Header row (row 1) ────────────────────────────────────────
  var hdr = sheet.getRange(1, 1, 1, 4);
  hdr.setValues([['Date', 'Portfolio …' + suffix + ' (Indexed)', 'SPY (Indexed)', 'QQQ (Indexed)']]);
  hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);

  // ── Data rows (row 2+) ────────────────────────────────────────
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat('0.00');

  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0) sheet.getRange(2 + i, 1, 1, 4).setBackground('#f8f9fa');
  }

  // ── Meta (below data) ─────────────────────────────────────────
  var metaRow = rows.length + 3;
  sheet.getRange(metaRow, 1, 3, 2).setValues([
    ['Account',     '…' + suffix],
    ['Base Date',   baseDate],
    ['Base Net Liq','$' + baseNetLiq.toLocaleString('en-US', { minimumFractionDigits: 2 })],
  ]);
  sheet.getRange(metaRow, 1, 3, 2).setFontColor('#888888').setFontStyle('italic');
}

// ─────────────────────────────────────────────────────────────────
// Chart rendering
// ─────────────────────────────────────────────────────────────────

function insertLineChart_(sheet, dataRows, suffix) {
  // Bare-minimum chart — no styling options — to confirm data renders.
  // If this shows lines, options will be added back one by one.
  var n = dataRows + 1; // header row + data rows
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 4))
    .setNumHeaders(1)
    .setPosition(dataRows + 5, 1, 0, 0)
    .build();

  sheet.insertChart(chart);
}

// ─────────────────────────────────────────────────────────────────
// Debug helper — run this manually from Apps Script editor if chart
// is still blank. Logs the first 5 rows of the chart sheet so you
// can confirm what's actually written to the cells.
// ─────────────────────────────────────────────────────────────────

function debugChartSheet(suffix) {
  suffix = suffix || '418';
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(chartSheetName_(suffix));
  if (!sheet) { console.log('Sheet not found: ' + chartSheetName_(suffix)); return; }

  var vals  = sheet.getRange(1, 1, 6, 4).getValues();
  var types = sheet.getRange(1, 1, 6, 4).getValues().map(function(row) {
    return row.map(function(v) { return typeof v + ' | ' + Object.prototype.toString.call(v); });
  });

  console.log('=== Chart sheet values (rows 1-6) ===');
  vals.forEach(function(row, i) { console.log('Row ' + (i+1) + ': ' + JSON.stringify(row)); });
  console.log('=== Cell types ===');
  types.forEach(function(row, i) { console.log('Row ' + (i+1) + ': ' + row.join(' || ')); });

  var charts = sheet.getCharts();
  console.log('Charts on sheet: ' + charts.length);
  if (charts.length > 0) {
    var ranges = charts[0].getRanges();
    console.log('Chart ranges: ' + ranges.length);
    ranges.forEach(function(r, i) {
      console.log('  Range ' + i + ': ' + r.getA1Notation() + '  numRows=' + r.getNumRows() + '  numCols=' + r.getNumColumns());
    });
  }
}

// ─────────────────────────────────────────────────────────────────
// Standalone chart test — run directly from Apps Script editor.
// Creates a fresh sheet "Test Chart" with 5 hardcoded rows and
// builds the simplest possible LINE chart. No dependency on any
// other function. If this is ALSO blank the issue is in the
// spreadsheet/Apps Script environment itself, not our code.
// ─────────────────────────────────────────────────────────────────

function testMinimalChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Fresh sheet every time
  var existing = ss.getSheetByName('Test Chart');
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet('Test Chart');

  // Hardcoded 5-row dataset — no Date objects, just strings + numbers
  sheet.getRange(1, 1, 6, 3).setValues([
    ['Date',       'Series A', 'Series B'],
    ['2026-01-01', 100,        100       ],
    ['2026-01-02', 101,        99        ],
    ['2026-01-03', 103,        98        ],
    ['2026-01-04', 102,        101       ],
    ['2026-01-05', 105,        103       ],
  ]);

  SpreadsheetApp.flush();

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, 6, 3))
    .setNumHeaders(1)
    .setPosition(8, 1, 0, 0)
    .build();

  sheet.insertChart(chart);
  ss.setActiveSheet(sheet);
  ss.toast('Test chart created — do you see lines?', 'Chart Test', 10);
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}
