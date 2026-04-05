/**
 * EquityCurveChart.gs — Builds a per-account equity curve chart.
 * Version: 2.0 (2026-04-05) — % change from day 1; month-end labeled dots.
 *
 * Chart sheet layout (7 columns):
 *   Date | NL% | SPY% | QQQ% | NL dot | SPY dot | QQQ dot
 *
 *   Series 0–2 : solid lines, no individual point markers
 *   Series 3–5 : month-end dots only (lineWidth=0), value labels, color-matched
 *
 * IMPORTANT — flat dot-notation setOption() ONLY.
 * Never pass nested objects to setOption(); it silently breaks the chart.
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point
// ─────────────────────────────────────────────────────────────────

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

    // ── 3. Compute % change from first common date ────────────────
    var baseDate   = commonDates[0];
    var baseNetLiq = netLiqMap[baseDate];
    var baseSPY    = spyMap[baseDate];
    var baseQQQ    = qqqMap[baseDate] || null;

    var monthEnds = getMonthEndDates_(commonDates);

    var rows = commonDates.map(function(d) {
      var nlPct  = roundTo2_((netLiqMap[d] / baseNetLiq - 1) * 100);
      var spyPct = roundTo2_((spyMap[d]    / baseSPY    - 1) * 100);
      var qqqPct = (baseQQQ && qqqMap[d])
        ? roundTo2_((qqqMap[d] / baseQQQ - 1) * 100)
        : '';

      var isME = monthEnds[d];
      return [
        d,
        nlPct,
        spyPct,
        qqqPct,
        isME             ? nlPct  : '',
        isME             ? spyPct : '',
        (isME && qqqPct !== '') ? qqqPct : '',
      ];
    });

    // ── 4. Write staging data ─────────────────────────────────────
    var chartSheet = getOrCreateChartSheet_(ss, suffix);
    writeChartData_(chartSheet, rows, suffix, baseDate);
    SpreadsheetApp.flush();

    // ── 5. Build chart ────────────────────────────────────────────
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

function writeChartData_(sheet, rows, suffix, baseDate) {
  // ── Header ────────────────────────────────────────────────────
  var hdr = sheet.getRange(1, 1, 1, 7);
  hdr.setValues([[
    'Date',
    'Portfolio …' + suffix + ' (% Return)',
    'SPY (% Return)',
    'QQQ (% Return)',
    'NL Month-End',
    'SPY Month-End',
    'QQQ Month-End',
  ]]);
  hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
     .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120);
  [2, 3, 4].forEach(function(c) { sheet.setColumnWidth(c, 185); });
  [5, 6, 7].forEach(function(c) { sheet.setColumnWidth(c, 120); });

  // ── Data rows ─────────────────────────────────────────────────
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);

  // Line series: plain % format (e.g. 4.80%)
  sheet.getRange(2, 2, rows.length, 3).setNumberFormat('0.00"%"');

  // Dot series: signed % format (e.g. +4.80% / -2.10%)
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

function insertLineChart_(sheet, rows, suffix) {
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
  var yMin    = Math.floor(dataMin - padding);
  var yMax    = Math.ceil(dataMax  + padding);

  // ── Build chart ───────────────────────────────────────────────
  // IMPORTANT: flat dot-notation only — nested setOption() objects
  // silently break Google Sheets EmbeddedChartBuilder rendering.
  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 7))
    .setNumHeaders(1)
    .setPosition(rows.length + 5, 1, 0, 0)
    .setOption('title',  '% Return — Portfolio …' + suffix + ' vs. SPY & QQQ')
    .setOption('width',  1200)
    .setOption('height', 550)
    .setOption('legend.position', 'top')
    .setOption('hAxis.title', 'Date')
    .setOption('hAxis.slantedText', true)
    .setOption('hAxis.slantedTextAngle', 30)
    .setOption('vAxis.title', '% Return')
    .setOption('vAxis.viewWindowMode', 'explicit')
    .setOption('vAxis.viewWindow.min', yMin)
    .setOption('vAxis.viewWindow.max', yMax)
    .setOption('vAxis.format', '0.0')
    // ── Series 0–2: solid lines ───────────────────────────────
    .setOption('series.0.color',         '#1a73e8')
    .setOption('series.0.lineWidth',     2)
    .setOption('series.0.pointsVisible', false)
    .setOption('series.1.color',         '#ea4335')
    .setOption('series.1.lineWidth',     2)
    .setOption('series.1.pointsVisible', false)
    .setOption('series.2.color',         '#fbbc04')
    .setOption('series.2.lineWidth',     2)
    .setOption('series.2.pointsVisible', false)
    // ── Series 3–5: month-end dots + value labels ─────────────
    .setOption('series.3.color',            '#1a73e8')
    .setOption('series.3.lineWidth',        0)
    .setOption('series.3.pointsVisible',    true)
    .setOption('series.3.pointSize',        7)
    .setOption('series.3.dataLabel',        'value')
    .setOption('series.3.visibleInLegend',  false)
    .setOption('series.4.color',            '#ea4335')
    .setOption('series.4.lineWidth',        0)
    .setOption('series.4.pointsVisible',    true)
    .setOption('series.4.pointSize',        7)
    .setOption('series.4.dataLabel',        'value')
    .setOption('series.4.visibleInLegend',  false)
    .setOption('series.5.color',            '#fbbc04')
    .setOption('series.5.lineWidth',        0)
    .setOption('series.5.pointsVisible',    true)
    .setOption('series.5.pointSize',        7)
    .setOption('series.5.dataLabel',        'value')
    .setOption('series.5.visibleInLegend',  false);

  sheet.insertChart(builder.build());
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Returns a set of dates that are month-end trading days.
 * A date is a month-end if the next date in the dataset falls in a
 * different calendar month (handles weekends/holidays naturally).
 * The last date in the dataset is always included.
 * @param {string[]} sortedDates  Ascending 'YYYY-MM-DD' strings
 * @returns {Object}  { 'YYYY-MM-DD': true }
 */
function getMonthEndDates_(sortedDates) {
  var monthEnds = {};
  for (var i = 0; i < sortedDates.length; i++) {
    var d    = sortedDates[i];
    var next = sortedDates[i + 1];
    // Different month prefix OR last date in set
    if (!next || next.substring(0, 7) !== d.substring(0, 7)) {
      monthEnds[d] = true;
    }
  }
  return monthEnds;
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
