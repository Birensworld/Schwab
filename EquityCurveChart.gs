/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 *
 * Merges Portfolio (Net Liq ###) and SPY data, normalizes both to 100
 * at the first date where both series overlap, and renders a line chart
 * on the "Equity Curve ###" tab.
 *
 * SPY History is shared; each account has its own chart tab.
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point
// ─────────────────────────────────────────────────────────────────

/**
 * Builds or refreshes the equity curve chart for one account.
 * @param {string} suffix  e.g. '418'
 */
function buildEquityCurveChartForAccount(suffix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building equity curve for account …' + suffix + '…', 'Working', -1);

    // ── 1. Load data ───────────────────────────────────────────────
    const netLiqMap = getNetLiqMap_(suffix);
    const spyMap    = getSPYCloseMap_();

    if (Object.keys(netLiqMap).length === 0) {
      throw new Error(
        'No Net Liquidity data found for account …' + suffix + '.\n\n' +
        'Run "Capture Today\'s Net Liquidity" or "Reconstruct History" first.'
      );
    }
    if (Object.keys(spyMap).length === 0) {
      throw new Error('No SPY data found. Run "Fetch SPY History" first.');
    }

    // ── 2. Find overlapping dates ─────────────────────────────────
    const commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d] !== undefined; })
      .sort();

    if (commonDates.length < 2) {
      throw new Error(
        'Not enough overlapping dates between Net Liq …' + suffix + ' and SPY.\n\n' +
        'Make sure both datasets share at least a few common trading days.'
      );
    }

    // ── 3. Normalize to 100 at first common date ──────────────────
    const baseDate   = commonDates[0];
    const baseNetLiq = netLiqMap[baseDate];
    const baseSPY    = spyMap[baseDate];

    const rows = commonDates.map(function(d) {
      return [
        new Date(d),
        roundTo2_(netLiqMap[d] / baseNetLiq * 100),
        roundTo2_(spyMap[d]    / baseSPY    * 100),
      ];
    });

    // ── 4. Write data to chart sheet ─────────────────────────────
    const chartSheet = getOrCreateChartSheet_(ss, suffix);
    writeChartData_(chartSheet, rows, suffix, baseDate, baseNetLiq, baseSPY);

    // ── 5. Draw chart ─────────────────────────────────────────────
    chartSheet.getCharts().forEach(function(c) { chartSheet.removeChart(c); });
    insertLineChart_(chartSheet, rows.length, suffix);

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ Chart ready — ' + rows.length + ' trading days (' +
        baseDate + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Equity Curve …' + suffix, 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Chart Error – Account ' + suffix, e.message, SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Sheet setup
// ─────────────────────────────────────────────────────────────────

function getOrCreateChartSheet_(ss, suffix) {
  const name  = sheetNames_(suffix).chart;
  let sheet   = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function writeChartData_(sheet, rows, suffix, baseDate, baseNetLiq, baseSPY) {
  // ── Meta info (rows 1–3) ──────────────────────────────────────
  sheet.getRange(1, 1, 3, 2).setValues([
    ['Account',       '…' + suffix],
    ['Base Date',     baseDate],
    ['Base Net Liq',  '$' + baseNetLiq.toLocaleString('en-US', { minimumFractionDigits: 2 })],
  ]);
  sheet.getRange(1, 1, 3, 2).setFontColor('#888888').setFontStyle('italic');

  // ── Header row (row 5) ────────────────────────────────────────
  const hdr = sheet.getRange(5, 1, 1, 3);
  hdr.setValues([['Date', 'Portfolio …' + suffix + ' (Indexed)', 'SPY (Indexed)']]);
  hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(5);
  sheet.setColumnWidth(1, 125);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 160);

  // ── Data rows (row 6+) ────────────────────────────────────────
  sheet.getRange(6, 1, rows.length, 3).setValues(rows);
  sheet.getRange(6, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(6, 2, rows.length, 2).setNumberFormat('0.00');

  // Alternating row shading
  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0) sheet.getRange(6 + i, 1, 1, 3).setBackground('#f8f9fa');
  }
}

// ─────────────────────────────────────────────────────────────────
// Chart rendering
// ─────────────────────────────────────────────────────────────────

function insertLineChart_(sheet, dataRows, suffix) {
  const dataRange = sheet.getRange(5, 1, dataRows + 1, 3);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setPosition(dataRows + 8, 1, 0, 0)
    .setOption('title', 'Equity Curve — Account …' + suffix + ' vs. SPY')
    .setOption('titleTextStyle', { fontSize: 17, bold: true, color: '#202124' })
    .setOption('hAxis', {
      title: 'Date',
      titleTextStyle: { bold: true, color: '#444' },
      slantedText: true,
      slantedTextAngle: 30,
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
