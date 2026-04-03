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
      // Parse date as LOCAL midnight (not UTC) to avoid timezone shift on sheet
      var p = d.split('-');
      var localDate = new Date(+p[0], +p[1] - 1, +p[2]);
      // Use 0 for missing QQQ so the chart range stays clean (no nulls/empty)
      var qqqIndexed = (baseQQQ && qqqMap[d]) ? roundTo2_(qqqMap[d] / baseQQQ * 100) : 0;
      return [
        localDate,
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
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
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
  // Use separate ranges so the chart builder unambiguously knows:
  //   col 1 = X-axis domain (dates)
  //   cols 2-4 = three series
  // setNumHeaders(1) tells it row 1 is labels, not data points.
  var n = dataRows + 1; // header + data rows
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 1))   // Date (domain)
    .addRange(sheet.getRange(1, 2, n, 1))   // Portfolio
    .addRange(sheet.getRange(1, 3, n, 1))   // SPY
    .addRange(sheet.getRange(1, 4, n, 1))   // QQQ
    .setNumHeaders(1)
    .setPosition(dataRows + 5, 1, 0, 0)
    .setOption('title', 'Equity Curve — Account …' + suffix + ' vs. SPY & QQQ')
    .setOption('titleTextStyle', { fontSize: 17, bold: true, color: '#202124' })
    .setOption('hAxis', {
      title: 'Date',
      titleTextStyle: { bold: true, color: '#444' },
      slantedText: true, slantedTextAngle: 30,
      format: 'MMM yyyy',
      gridlines: { color: '#e0e0e0' },
    })
    .setOption('vAxis', {
      title: 'Indexed Value (Base = 100)',
      titleTextStyle: { bold: true, color: '#444' },
      gridlines: { count: 10, color: '#e0e0e0' },
      format: '0',
    })
    .setOption('series', {
      0: { color: '#1a73e8', lineWidth: 2, pointSize: 0 },  // Portfolio — blue
      1: { color: '#ea4335', lineWidth: 2, pointSize: 0 },  // SPY — red
      2: { color: '#fbbc04', lineWidth: 2, pointSize: 0 },  // QQQ — amber
    })
    .setOption('legend',      { position: 'top', textStyle: { fontSize: 13 } })
    .setOption('width',        1200)
    .setOption('height',       550)
    .setOption('backgroundColor', { fill: '#ffffff' })
    .setOption('chartArea',   { left: 80, top: 80, width: '87%', height: '72%' })
    .setOption('interpolateNulls', true)
    .setOption('focusTarget', 'category')
    .build();

  sheet.insertChart(chart);
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}
