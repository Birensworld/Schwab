/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 * Version: 2.2 (2026-04-05) — QQQ red; 15th-of-month markers; yearly chart with year prompt.
 *
 * Chart sheet layout (7 columns):
 *   Date | NL% | SPY% | QQQ% | NL Marker | SPY Marker | QQQ Marker
 *
 *   Series 0–2 : solid lines, no individual point markers
 *   Series 3–5 : month-end + 15th-of-month dots (lineWidth=0), value labels, color-matched
 *
 * Colors: NL = black (#000000), SPY = green (#34a853), QQQ = red (#ea4335)
 *
 * IMPORTANT — flat dot-notation setOption() ONLY.
 * Never pass nested objects to setOption(); it silently breaks the chart.
 */

// ─────────────────────────────────────────────────────────────────
// Public entry points
// ─────────────────────────────────────────────────────────────────

/** Full-data chart (all available dates). Called from menu wrapper. */
function buildEquityCurveChartForAccount(suffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building equity curve for account …' + suffix + '…', 'Working', -1);

    // ── 1. Load raw data ──────────────────────────────────────────
    var netLiqMap  = getNetLiqMap_(suffix);
    var benchmarks = getBenchmarkCloseMaps_();
    var spyMap     = benchmarks.spy;
    var qqqMap     = benchmarks.qqq;

    if (Object.keys(netLiqMap).length === 0) {
      throw new Error(
        'No Net Liquidity data found for account …' + suffix + '.\n\n' +
        'Enter values in the "' + SHEET_NET_LIQ + '" sheet first.'
      );
    }
    if (Object.keys(spyMap).length === 0) {
      throw new Error('No SPY/QQQ data found. Run "Fetch SPY + QQQ History" first.');
    }

    // ── 2. Common dates (NL + SPY both present) ───────────────────
    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d]; })
      .sort();

    if (commonDates.length < 2) {
      throw new Error(
        'Not enough overlapping dates between Net Liq …' + suffix + ' and SPY.\n' +
        'Make sure both datasets share common trading days.'
      );
    }

    // ── 3. Base values anchored to HISTORY_START ──────────────────
    // Each series uses the first available date on or after HISTORY_START
    // as its own base. This handles holidays (e.g. Jan 1) where NetLiq
    // may have a manual entry but SPY has no trading data.
    var allNLDates  = Object.keys(netLiqMap).sort();
    var allSPYDates = Object.keys(spyMap).sort();

    var baseNLDate  = allNLDates.find(function(d)  { return d >= HISTORY_START; }) || allNLDates[0];
    var baseSPYDate = allSPYDates.find(function(d) { return d >= HISTORY_START; }) || allSPYDates[0];

    var baseDate   = baseNLDate;
    var baseNetLiq = netLiqMap[baseNLDate];
    var baseSPY    = spyMap[baseSPYDate];
    var baseQQQ    = qqqMap[baseSPYDate] || null;

    // ── 4. Build rows + markers ───────────────────────────────────
    var markers = getMarkerDates_(commonDates);
    var rows    = buildRows_(commonDates, netLiqMap, spyMap, qqqMap,
                             baseNetLiq, baseSPY, baseQQQ, markers);

    // ── 5. Write + chart ──────────────────────────────────────────
    var sheetName  = chartSheetName_(suffix);
    var chartSheet = getOrCreateChartSheet_(ss, sheetName);
    writeChartData_(chartSheet, rows, suffix, baseDate);
    SpreadsheetApp.flush();
    insertLineChart_(chartSheet, rows,
      '% Return — Portfolio …' + suffix + ' vs. SPY & QQQ');

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

/**
 * Prompts for a 4-digit year then builds a year-scoped equity curve.
 * Sheet: "Equity Curve 418 (2026)" — separate from the full-data chart.
 */
function promptAndBuildYearlyEquityCurve_(suffix) {
  var ui     = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Equity Curve – Yearly View – Account …' + suffix,
    'Enter a 4-digit year (e.g. ' + new Date().getFullYear() + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var year = result.getResponseText().trim();
  if (!/^\d{4}$/.test(year)) {
    ui.alert('Invalid year', 'Please enter a 4-digit year such as 2026.',
      ui.ButtonSet.OK);
    return;
  }
  buildEquityCurveChartForYear_(suffix, year);
}

function buildEquityCurveChartForYear_(suffix, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building ' + year + ' equity curve for account …' + suffix + '…', 'Working', -1);

    // ── 1. Load raw data ──────────────────────────────────────────
    var netLiqMap  = getNetLiqMap_(suffix);
    var benchmarks = getBenchmarkCloseMaps_();
    var spyMap     = benchmarks.spy;
    var qqqMap     = benchmarks.qqq;

    // ── 2. Filter to requested year ───────────────────────────────
    var yStart = year + '-01-01';
    var yEnd   = year + '-12-31';

    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d] && d >= yStart && d <= yEnd; })
      .sort();

    if (commonDates.length === 0) {
      ss.toast('', '', 1);
      ui.alert(
        'No Data for ' + year,
        'No overlapping Net Liquidity and SPY/QQQ data found for ' + year + '.\n\n' +
        'Make sure you have:\n' +
        '  • Net Liquidity entries in "' + SHEET_NET_LIQ + '" for ' + year + '\n' +
        '  • SPY history fetched for ' + year,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    if (commonDates.length < 2) {
      throw new Error('Only 1 data point found for ' + year + ' — need at least 2 to draw a chart.');
    }

    // ── 3. Base = first common date in the requested year ─────────
    var baseDate   = commonDates[0];
    var baseNetLiq = netLiqMap[baseDate];
    var baseSPY    = spyMap[baseDate];
    var baseQQQ    = qqqMap[baseDate] || null;

    // ── 4. Build rows + markers ───────────────────────────────────
    var markers = getMarkerDates_(commonDates);
    var rows    = buildRows_(commonDates, netLiqMap, spyMap, qqqMap,
                             baseNetLiq, baseSPY, baseQQQ, markers);

    // ── 5. Write + chart ──────────────────────────────────────────
    var sheetName  = 'Equity Curve ' + suffix + ' (' + year + ')';
    var chartSheet = getOrCreateChartSheet_(ss, sheetName);
    writeChartData_(chartSheet, rows, suffix, baseDate);
    SpreadsheetApp.flush();
    insertLineChart_(chartSheet, rows,
      '% Return ' + year + ' — Portfolio …' + suffix + ' vs. SPY & QQQ');

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ ' + rows.length + ' days  (' + baseDate + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Equity Curve …' + suffix + ' (' + year + ')', 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error – Account ' + suffix + ' (' + year + ')',
      e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Row builder (shared by full and yearly charts)
// ─────────────────────────────────────────────────────────────────

function buildRows_(dates, netLiqMap, spyMap, qqqMap,
                    baseNetLiq, baseSPY, baseQQQ, markers) {
  return dates.map(function(d) {
    var nlPct  = roundTo2_((netLiqMap[d] / baseNetLiq - 1) * 100);
    var spyPct = roundTo2_((spyMap[d]    / baseSPY    - 1) * 100);
    var qqqPct = (baseQQQ && qqqMap[d])
      ? roundTo2_((qqqMap[d] / baseQQQ - 1) * 100)
      : '';

    var isMark = markers[d];
    return [
      d,
      nlPct,
      spyPct,
      qqqPct,
      isMark             ? nlPct  : '',
      isMark             ? spyPct : '',
      (isMark && qqqPct !== '') ? qqqPct : '',
    ];
  });
}

// ─────────────────────────────────────────────────────────────────
// Sheet setup
// ─────────────────────────────────────────────────────────────────

function getOrCreateChartSheet_(ss, sheetName) {
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(sheetName);
}

function writeChartData_(sheet, rows, suffix, baseDate) {
  // ── Header ────────────────────────────────────────────────────
  var hdr = sheet.getRange(1, 1, 1, 7);
  hdr.setValues([[
    'Date',
    'Portfolio …' + suffix + ' (% Return)',
    'SPY (% Return)',
    'QQQ (% Return)',
    'NL Marker',
    'SPY Marker',
    'QQQ Marker',
  ]]);
  hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  [2, 3, 4].forEach(function(c) { sheet.setColumnWidth(c, 185); });
  [5, 6, 7].forEach(function(c) { sheet.setColumnWidth(c, 120); });

  // ── Data rows ─────────────────────────────────────────────────
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);

  // Line series: plain % format
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat('0.00"%"');

  // Marker series: signed % format (e.g. +4.80% / -2.10%)
  sheet.getRange(2, 5, rows.length, 3).setNumberFormat('+0.00"%";-0.00"%";"0%"');

  // Alternating row shading
  for (var i = 0; i < rows.length; i++) {
    if (i % 2 === 0) sheet.getRange(2 + i, 1, 1, 7).setBackground('#f8f9fa');
  }

  // ── Meta block below data ─────────────────────────────────────
  var metaRow = rows.length + 3;
  sheet.getRange(metaRow, 1, 2, 2).setValues([
    ['Account',   '…' + suffix],
    ['Base Date', baseDate],
  ]);
  sheet.getRange(metaRow, 1, 2, 2).setFontColor('#888888').setFontStyle('italic');
}

// ─────────────────────────────────────────────────────────────────
// Chart rendering
// ─────────────────────────────────────────────────────────────────

function insertLineChart_(sheet, rows, title) {
  var n = rows.length + 1;  // header + data rows

  // ── Y-axis bounds from actual % change values ─────────────────
  var allPcts = [];
  rows.forEach(function(r) {
    [r[1], r[2], r[3]].forEach(function(v) {
      if (typeof v === 'number') allPcts.push(v);
    });
  });
  var dataMin = allPcts.length ? Math.min.apply(null, allPcts) : -5;
  var dataMax = allPcts.length ? Math.max.apply(null, allPcts) : 20;
  var padding = Math.max((dataMax - dataMin) * 0.15, 2);
  // Round to nearest 0.5 so gridlines land exactly on 0.5 increments
  var yMin          = Math.floor((dataMin - padding) * 2) / 2;
  var yMax          = Math.ceil((dataMax  + padding) * 2) / 2;
  var gridlineCount = Math.round((yMax - yMin) / 0.5) + 1;

  // ── Build chart ───────────────────────────────────────────────
  // IMPORTANT: flat dot-notation only — nested setOption() objects
  // silently break Google Sheets EmbeddedChartBuilder rendering.
  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 7))
    .setNumHeaders(1)
    .setPosition(rows.length + 5, 1, 0, 0)
    .setOption('title',  title)
    .setOption('width',  1200)
    .setOption('height', 550)
    .setOption('legend.position', 'top')
    .setOption('hAxis.title', 'Date')
    .setOption('hAxis.slantedText', true)
    .setOption('hAxis.slantedTextAngle', 30)
    .setOption('vAxis.title', '% Return')
    .setOption('vAxis.viewWindowMode',       'explicit')
    .setOption('vAxis.viewWindow.min',       yMin)
    .setOption('vAxis.viewWindow.max',       yMax)
    .setOption('vAxis.format',               '0.0')
    .setOption('vAxis.gridlines.count',      gridlineCount)
    .setOption('vAxis.titleTextStyle.bold',  true)
    .setOption('hAxis.titleTextStyle.bold',  true)
    // ── Series 0–2: solid lines ─────────────────────────────────
    // NL = black, SPY = green, QQQ = red
    .setOption('series.0.color',         '#000000')
    .setOption('series.0.lineWidth',     2)
    .setOption('series.0.pointsVisible', false)
    .setOption('series.1.color',         '#34a853')
    .setOption('series.1.lineWidth',     2)
    .setOption('series.1.pointsVisible', false)
    .setOption('series.2.color',         '#ea4335')
    .setOption('series.2.lineWidth',     2)
    .setOption('series.2.pointsVisible', false)
    // ── Series 3–5: marker dots + value labels ───────────────────
    .setOption('series.3.color',           '#000000')
    .setOption('series.3.lineWidth',       0)
    .setOption('series.3.pointsVisible',   true)
    .setOption('series.3.pointSize',       7)
    .setOption('series.3.dataLabel',       'value')
    .setOption('series.3.visibleInLegend', false)
    .setOption('series.4.color',           '#34a853')
    .setOption('series.4.lineWidth',       0)
    .setOption('series.4.pointsVisible',   true)
    .setOption('series.4.pointSize',       7)
    .setOption('series.4.dataLabel',       'value')
    .setOption('series.4.visibleInLegend', false)
    .setOption('series.5.color',           '#ea4335')
    .setOption('series.5.lineWidth',       0)
    .setOption('series.5.pointsVisible',   true)
    .setOption('series.5.pointSize',       7)
    .setOption('series.5.dataLabel',       'value')
    .setOption('series.5.visibleInLegend', false);

  sheet.insertChart(builder.build());
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Returns a set of marker dates: the last trading day of each month
 * AND the trading day closest to the 15th of each month.
 * Handles weekends/holidays naturally by working from actual data dates.
 * @param {string[]} sortedDates  Ascending 'YYYY-MM-DD' strings
 * @returns {Object}  { 'YYYY-MM-DD': true }
 */
function getMarkerDates_(sortedDates) {
  var markers = {};

  // Group dates by YYYY-MM
  var byMonth = {};
  sortedDates.forEach(function(d) {
    var ym = d.substring(0, 7);
    if (!byMonth[ym]) byMonth[ym] = [];
    byMonth[ym].push(d);
  });

  Object.keys(byMonth).forEach(function(ym) {
    var datesInMonth = byMonth[ym].sort();

    // Month-end: last trading day of the month
    markers[datesInMonth[datesInMonth.length - 1]] = true;

    // Mid-month: trading day whose day-of-month is closest to 15
    var midMonth = datesInMonth.reduce(function(best, d) {
      if (!best) return d;
      var dDay    = parseInt(d.substring(8),    10);
      var bestDay = parseInt(best.substring(8), 10);
      return Math.abs(dDay - 15) < Math.abs(bestDay - 15) ? d : best;
    }, null);
    if (midMonth) markers[midMonth] = true;
  });

  return markers;
}

function roundTo2_(n) {
  return Math.round(n * 100) / 100;
}

// ─────────────────────────────────────────────────────────────────
// Debug helpers
// ─────────────────────────────────────────────────────────────────

function debugChartSheet(suffix) {
  suffix = suffix || '418';
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(chartSheetName_(suffix));
  if (!sheet) { console.log('Sheet not found: ' + chartSheetName_(suffix)); return; }

  var vals = sheet.getRange(1, 1, 6, 7).getValues();
  console.log('=== Chart sheet values (rows 1-6) ===');
  vals.forEach(function(row, i) { console.log('Row ' + (i + 1) + ': ' + JSON.stringify(row)); });

  var charts = sheet.getCharts();
  console.log('Charts on sheet: ' + charts.length);
  if (charts.length > 0) {
    charts[0].getRanges().forEach(function(r, i) {
      console.log('  Range ' + i + ': ' + r.getA1Notation() +
        '  rows=' + r.getNumRows() + '  cols=' + r.getNumColumns());
    });
  }
}

function testMinimalChart() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var existing = ss.getSheetByName('Test Chart');
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet('Test Chart');

  sheet.getRange(1, 1, 6, 3).setValues([
    ['Date',       'Series A', 'Series B'],
    ['2026-01-01', 100,        100       ],
    ['2026-01-02', 101,        99        ],
    ['2026-01-03', 103,        98        ],
    ['2026-01-04', 102,        101       ],
    ['2026-01-05', 105,        103       ],
  ]);

  SpreadsheetApp.flush();

  sheet.insertChart(
    sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(1, 1, 6, 3))
      .setNumHeaders(1)
      .setPosition(8, 1, 0, 0)
      .build()
  );

  ss.setActiveSheet(sheet);
  ss.toast('Test chart created — do you see lines?', 'Chart Test', 10);
}
