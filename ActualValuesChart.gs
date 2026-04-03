/**
 * ActualValuesChart.gs — Dual Y-axis chart of actual (non-normalized) values.
 * Version: 1.0 (2026-04-03)
 *
 * Per-account chart sheet: "Actual Values 418", "Actual Values 973", etc.
 * Data layout: row 1 = header, rows 2+ = Date | Net Liq ($) | SPY Close ($)
 *
 * Left Y-axis  → Net Liquidity in dollars
 * Right Y-axis → SPY Close price in dollars
 * Same date range as the Net Liquidity sheet.
 * Hover over any point to see the exact value for that day.
 *
 * Chart rules (Apps Script quirks):
 *   - Flat dot-notation setOption() only — nested objects break the chart
 *   - Delete+recreate sheet (never clearContents) to avoid stale embedded charts
 *   - Single contiguous addRange() with setNumHeaders(1)
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point
// ─────────────────────────────────────────────────────────────────

function buildActualValuesChartForAccount(suffix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Building actual values chart for account …' + suffix + '…', 'Working', -1);

    // ── 1. Load Net Liq data for this account ─────────────────────
    var netLiqMap = getNetLiqMap_(suffix);
    if (Object.keys(netLiqMap).length === 0) {
      throw new Error('No Net Liquidity data for account …' + suffix + '. Enter values in the "' + SHEET_NET_LIQ + '" sheet first.');
    }

    // ── 2. Load SPY actual closes ──────────────────────────────────
    var benchmarks = getBenchmarkCloseMaps_();
    var spyMap     = benchmarks.spy;
    if (Object.keys(spyMap).length === 0) {
      throw new Error('No SPY data found. Run "Fetch SPY + QQQ History" first.');
    }

    // ── 3. Build rows for dates where both have data ───────────────
    var commonDates = Object.keys(netLiqMap)
      .filter(function(d) { return spyMap[d]; })
      .sort();

    if (commonDates.length < 2) {
      throw new Error('Not enough overlapping dates between Net Liq …' + suffix + ' and SPY.');
    }

    var rows = commonDates.map(function(d) {
      return [d, netLiqMap[d], spyMap[d]];
    });

    // ── 4. Write data to chart sheet ──────────────────────────────
    var chartSheet = getOrCreateActualSheet_(ss, suffix);

    var hdr = chartSheet.getRange(1, 1, 1, 3);
    hdr.setValues([['Date', 'Net Liquidity …' + suffix + ' ($)', 'SPY Close ($)']]);
    hdr.setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff')
       .setHorizontalAlignment('center');
    chartSheet.setFrozenRows(1);
    chartSheet.setColumnWidth(1, 120);
    chartSheet.setColumnWidth(2, 200);
    chartSheet.setColumnWidth(3, 150);

    chartSheet.getRange(2, 1, rows.length, 3).setValues(rows);
    chartSheet.getRange(2, 2, rows.length, 1).setNumberFormat('"$"#,##0.00');
    chartSheet.getRange(2, 3, rows.length, 1).setNumberFormat('"$"#,##0.00');

    SpreadsheetApp.flush();

    // ── 5. Draw dual Y-axis chart ─────────────────────────────────
    insertActualValuesChart_(chartSheet, rows.length, suffix);

    ss.setActiveSheet(chartSheet);
    ss.toast(
      '✅ ' + rows.length + ' days  (' + commonDates[0] + ' → ' + commonDates[commonDates.length - 1] + ')',
      'Actual Values …' + suffix, 10
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Chart Error – Account ' + suffix, e.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(e);
  }
}

// ─────────────────────────────────────────────────────────────────
// Sheet setup — always delete+recreate to avoid stale embedded charts
// ─────────────────────────────────────────────────────────────────

function getOrCreateActualSheet_(ss, suffix) {
  var name     = 'Actual Values ' + suffix;
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(name);
}

// ─────────────────────────────────────────────────────────────────
// Chart rendering
// ─────────────────────────────────────────────────────────────────

function insertActualValuesChart_(sheet, dataRows, suffix) {
  var n = dataRows + 1; // header + data rows

  // Net Liq Y-axis range
  var netLiqVals = sheet.getRange(2, 2, dataRows, 1).getValues()
    .map(function(r) { return r[0]; }).filter(function(v) { return v > 0; });
  var nlMin = netLiqVals.length ? Math.min.apply(null, netLiqVals) : 0;
  var nlMax = netLiqVals.length ? Math.max.apply(null, netLiqVals) : 1;
  var nlPad = Math.max((nlMax - nlMin) * 0.10, 1);

  // SPY Y-axis range
  var spyVals = sheet.getRange(2, 3, dataRows, 1).getValues()
    .map(function(r) { return r[0]; }).filter(function(v) { return v > 0; });
  var spyMin = spyVals.length ? Math.min.apply(null, spyVals) : 0;
  var spyMax = spyVals.length ? Math.max.apply(null, spyVals) : 1;
  var spyPad = Math.max((spyMax - spyMin) * 0.10, 1);

  var builder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(1, 1, n, 3))
    .setNumHeaders(1)
    .setPosition(dataRows + 5, 1, 0, 0)
    .setOption('title',  'Actual Values — Account …' + suffix + ' vs. SPY')
    .setOption('width',  1200)
    .setOption('height', 550)
    .setOption('legend.position', 'top')
    // Net Liq series — left axis (0)
    .setOption('series.0.color', '#1a73e8')
    .setOption('series.0.lineWidth', 2)
    .setOption('series.0.targetAxisIndex', 0)
    // SPY series — right axis (1)
    .setOption('series.1.color', '#ea4335')
    .setOption('series.1.lineWidth', 2)
    .setOption('series.1.targetAxisIndex', 1)
    // Left axis (Net Liq)
    .setOption('vAxes.0.title', 'Net Liquidity ($)')
    .setOption('vAxes.0.viewWindowMode', 'explicit')
    .setOption('vAxes.0.viewWindow.min', Math.floor(nlMin - nlPad))
    .setOption('vAxes.0.viewWindow.max', Math.ceil(nlMax  + nlPad))
    .setOption('vAxes.0.format', '"$"#,##0')
    // Right axis (SPY)
    .setOption('vAxes.1.title', 'SPY Close ($)')
    .setOption('vAxes.1.viewWindowMode', 'explicit')
    .setOption('vAxes.1.viewWindow.min', Math.floor(spyMin - spyPad))
    .setOption('vAxes.1.viewWindow.max', Math.ceil(spyMax  + spyPad))
    .setOption('vAxes.1.format', '"$"#,##0.00')
    // X-axis
    .setOption('hAxis.title', 'Date')
    .setOption('hAxis.slantedText', true)
    .setOption('hAxis.slantedTextAngle', 30);

  sheet.insertChart(builder.build());
}
