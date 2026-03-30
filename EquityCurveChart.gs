/**
 * EquityCurveChart.gs — Merges Portfolio and SPY data, normalizes both
 * series to a common base of 100, and renders an interactive line chart
 * on the 'Equity Curve' sheet.
 *
 * Normalization: both series are divided by their value on the earliest
 * date where BOTH datasets have a value, then multiplied by 100.
 * This makes the chart answer: "If I started with $100 in each on the
 * same day, where would I be today?"
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point
// ─────────────────────────────────────────────────────────────────

function buildEquityCurveChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building equity curve…', 'Working', -1);

    // ── 1. Load raw data ───────────────────────────────────────────
    const netLiqMap = getNetLiqMap_();
    const spyMap    = getSPYCloseMap_();

    if (Object.keys(netLiqMap).length === 0) {
      throw new Error(
        'No Net Liquidity data found.\n\n' +
        'Please fetch today\'s value or run "Reconstruct from Transactions" first.'
      );
    }
    if (Object.keys(spyMap).length === 0) {
      throw new Error(
        'No SPY data found. Please run "Fetch SPY History" first.'
      );
    }

    // ── 2. Find overlapping dates (market days with both series) ───
    const commonDates = Object.keys(netLiqMap)
      .filter(d => spyMap[d] !== undefined)
      .sort();

    if (commonDates.length < 2) {
      throw new Error(
        'Not enough overlapping dates between Net Liquidity and SPY data.\n\n' +
        'Make sure both datasets cover at least some common trading days.'
      );
    }

    // ── 3. Normalize to 100 at first common date ───────────────────
    const baseDate   = commonDates[0];
    const baseNetLiq = netLiqMap[baseDate];
    const baseSPY    = spyMap[baseDate];

    const rows = commonDates.map(d => [
      new Date(d),
      roundTo2_(netLiqMap[d] / baseNetLiq * 100),
      roundTo2_(spyMap[d]    / baseSPY    * 100),
    ]);

    // ── 4. Write to Equity Curve sheet ─────────────────────────────
    const chartSheet = getOrCreateChartSheet_(ss);
    writeChartData_(chartSheet, rows, baseDate, baseNetLiq, baseSPY);

    // ── 5. Draw the chart ──────────────────────────────────────────
    // Remove old charts first
    chartSheet.getCharts().forEach(c => chartSheet.removeChart(c));
    insertLineChart_(chartSheet, rows.length);

    ss.setActiveSheet(chartSheet);
    ss.toast(
      `✅ Chart built — ${rows.length} trading days from ${baseDate} to ${commonDates.at(-1)}.`,
      'Equity Curve Ready', 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Sheet setup
// ─────────────────────────────────────────────────────────────────

function getOrCreateChartSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_CHART);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CHART);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return sheet;
}

function writeChartData_(sheet, rows, baseDate, baseNetLiq, baseSPY) {
  // ── Meta info block (rows 1-3) ─────────────────────────────────
  sheet.getRange(1, 1, 1, 2).setValues([['Base Date', baseDate]]);
  sheet.getRange(2, 1, 1, 2).setValues([
    ['Base Net Liq', '$' + baseNetLiq.toLocaleString('en-US', { minimumFractionDigits: 2 })],
  ]);
  sheet.getRange(3, 1, 1, 2).setValues([
    ['Base SPY Close', '$' + baseSPY.toFixed(2)],
  ]);
  sheet.getRange(1, 1, 3, 2).setFontColor('#888888').setFontStyle('italic');

  // ── Header row (row 5) ─────────────────────────────────────────
  const headerRange = sheet.getRange(5, 1, 1, 3);
  headerRange.setValues([['Date', 'Portfolio (Indexed)', 'SPY (Indexed)']]);
  headerRange
    .setFontWeight('bold')
    .setBackground('#0b5394')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(5);
  sheet.setColumnWidth(1, 125);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 160);

  // ── Data rows (starting row 6) ────────────────────────────────
  const dataRange = sheet.getRange(6, 1, rows.length, 3);
  dataRange.setValues(rows);
  sheet.getRange(6, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(6, 2, rows.length, 2).setNumberFormat('0.00');

  // Conditional formatting: portfolio cells green/red vs SPY
  // (simple alternating row shading for readability)
  for (let i = 0; i < rows.length; i++) {
    if (i % 2 === 0) {
      sheet.getRange(6 + i, 1, 1, 3).setBackground('#f8f9fa');
    }
  }
}

// ─────────────────────────────────────────────────────────────────
// Chart rendering
// ─────────────────────────────────────────────────────────────────

function insertLineChart_(sheet, dataRows) {
  // Data is in rows 5 (header) through 5+dataRows
  const dataRange = sheet.getRange(5, 1, dataRows + 1, 3);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    // Position the chart below the data table with some padding
    .setPosition(dataRows + 8, 1, 0, 0)
    // ── Titles ────────────────────────────────────────────────────
    .setOption('title', 'Portfolio Equity Curve vs. SPY')
    .setOption('titleTextStyle', { fontSize: 18, bold: true, color: '#202124' })
    .setOption('subtitle', 'Both series indexed to 100 at first overlapping trading day')
    // ── Axes ──────────────────────────────────────────────────────
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
      minorGridlines: { count: 1 },
      format: '0',
    })
    // ── Series colors ─────────────────────────────────────────────
    .setOption('series', {
      0: { color: '#1a73e8', lineWidth: 2, pointSize: 0 },  // Portfolio — blue
      1: { color: '#ea4335', lineWidth: 2, pointSize: 0 },  // SPY — red
    })
    // ── Legend ────────────────────────────────────────────────────
    .setOption('legend', {
      position: 'top',
      textStyle: { fontSize: 13 },
    })
    // ── Layout ────────────────────────────────────────────────────
    .setOption('width', 1200)
    .setOption('height', 550)
    .setOption('backgroundColor', { fill: '#ffffff' })
    .setOption('chartArea', {
      left: 80, top: 80, width: '87%', height: '72%',
    })
    .setOption('interpolateNulls', true)
    .setOption('focusTarget', 'category')   // tooltip shows both series on hover
    .setOption('curveType', 'none')
    .build();

  sheet.insertChart(chart);
}

// ─────────────────────────────────────────────────────────────────
// Data loading helpers
// ─────────────────────────────────────────────────────────────────

/** Returns date → net liquidity value map from the Net Liq sheet. */
function getNetLiqMap_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NET_LIQ);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues().slice(1);
  const tz   = Session.getScriptTimeZone();
  const map  = {};
  data.forEach(row => {
    if (!row[0] || !row[1]) return;
    const v = parseFloat(row[1]);
    if (isNaN(v) || v <= 0) return;
    map[fmtDate_(new Date(row[0]), tz)] = v;
  });
  return map;
}

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}
