/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 * Version: 1.6 (2026-04-03) — Tight Y-axis viewWindow + restored styling for better line separation.
 *           charts that caused blank rendering; bare-minimum chart config.
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
    insertLineChart_(chartSheet, rows, suffix);

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
  var name     = chartSheetName_(suffix);
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
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

function insertLineChart_(sheet, rows, suffix) {
  var dataRows = rows.length;
  var n        = dataRows + 1; // header + data rows

  // ── Compute tight Y-axis bounds from actual data ──────────────
  // Columns 1-3 (0-indexed) hold the three indexed series.
  var allVals = [];
  rows.forEach(function(r) {
    [r[1], r[2], r[3]].forEach(function(v) {
      if (typeof v === 'number' && v > 0) allVals.push(v);
    });
  });
  var dataMin = Math.min.apply(null, allVals);
  var dataMax = Math.max.apply(null, allVals);
  // Add 10% padding so lines don't hug the edges
  var padding  = Math.max((dataMax - dataMin) * 0.10, 1);
  var yMin     = Math.floor(dataMin - padding);
  var yMax     = Math.ceil(dataMax  + padding);

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 4))
    .setNumHeaders(1)
    .setPosition(dataRows + 5, 1, 0, 0)
    .setOption('title', 'Equity Curve — Account …' + suffix + ' vs. SPY & QQQ')
    .setOption('titleTextStyle', { fontSize: 16, bold: true, color: '#202124' })
    .setOption('hAxis', {
      title: 'Date',
      titleTextStyle: { bold: true, color: '#444' },
      slantedText: true, slantedTextAngle: 30,
      gridlines: { color: '#e0e0e0' },
    })
    .setOption('vAxis', {
      title: 'Indexed Value (Base = 100)',
      titleTextStyle: { bold: true, color: '#444' },
      viewWindow: { min: yMin, max: yMax },
      gridlines: { count: 8, color: '#e0e0e0' },
      format: '0.0',
    })
    .setOption('series', {
      0: { color: '#1a73e8', lineWidth: 2, pointSize: 0 },  // Portfolio — blue
      1: { color: '#ea4335', lineWidth: 2, pointSize: 0 },  // SPY — red
      2: { color: '#fbbc04', lineWidth: 2, pointSize: 0 },  // QQQ — amber
    })
    .setOption('legend',      { position: 'top', textStyle: { fontSize: 12 } })
    .setOption('width',        1200)
    .setOption('height',       550)
    .setOption('backgroundColor', { fill: '#ffffff' })
    .setOption('chartArea',   { left: 80, top: 60, width: '85%', height: '75%' })
    .setOption('interpolateNulls', true)
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
