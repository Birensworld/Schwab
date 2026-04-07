/**
 * Code.gs — Schwab Portfolio + Equity Curve — Google Apps Script
 * Version: 2.5 (2026-04-07)
 *
 * Entry point: onOpen() builds all menus.
 * Functionality is split across separate files:
 *   Authentication.gs  — OAuth / token management
 *   AccountRefresh.gs  — Portfolio refresh (UpdateSheet)
 *   Liquidation.gs     — Position liquidation
 *   NetLiquidity.gs    — Net Liquidity sheet management
 *   SPYHistory.gs      — SPY + QQQ price history
 *   EquityCurveChart.gs — % change equity curve with month-end labels
 *   SchwabAPI.gs       — Low-level Schwab API wrappers
 *   Backup.gs          — Backup to SPY_NetLiq_Backup spreadsheet
 *   Triggers.gs        — Daily trigger setup / teardown
 *
 * ─── Account registry ────────────────────────────────────────────
 * Keys are the last-3-digit suffix; values are full account numbers.
 * ACCOUNT_ORDER determines column order in the Net Liquidity sheet.
 */

const ACCOUNT_CONFIGS = {
  '418': '52172418',
  '973': '55262973',
  '317': '30050317',
};
const ACCOUNT_ORDER = ['418', '973', '317'];

const SHEET_NET_LIQ = 'Net Liquidity';
const SHEET_SPY     = 'SPY History';
const HISTORY_START = '2026-01-01';

function netLiqCol_(suffix)      { return ACCOUNT_ORDER.indexOf(suffix) + 2; }
function totalNetLiqCol_()       { return ACCOUNT_ORDER.length + 2; }
function chartSheetName_(suffix) { return 'Equity Curve ' + suffix; }

// ─────────────────────────────────────────────────────────────────
// Menu — references functions defined in the files listed above
// ─────────────────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Schwab menu — portfolio refresh + auth (Authentication.gs, AccountRefresh.gs)
  var authMenu = ui.createMenu('Setup / Re-Authorize')
    .addItem('Get Authorization Link',   'ShowAuthUrl')
    .addItem('Exchange Code for Tokens', 'ExchangeAuthCode');

  ui.createMenu('Schwab')
    .addItem('Refresh Portfolio', 'UpdateSheet')
    .addSeparator()
    .addSubMenu(authMenu)
    .addToUi();

  // Equity Curve menu — SPYHistory.gs, NetLiquidity.gs, EquityCurveChart.gs, ActualValuesChart.gs
  ui.createMenu('📈 Equity Curve')
    .addItem('📊 Fetch SPY + QQQ History',                      'fetchSPYHistory')
    .addItem("📥 Capture Today's Net Liq – All Accounts",       'fetchTodayNetLiq')
    .addSeparator()
    .addSubMenu(ui.createMenu('💼 Account …418')
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_418')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_418'))
    .addSubMenu(ui.createMenu('💼 Account …973')
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_973')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_973'))
    .addSubMenu(ui.createMenu('💼 Account …317')
      .addItem('📈 Build / Refresh Equity Curve Chart', 'buildEquityCurveChart_317')
      .addItem('📅 Build Chart – By Year',              'buildEquityCurveChartYearly_317'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger'))
    .addToUi();
}

// ─── Per-account menu wrappers ────────────────────────────────────
function buildEquityCurveChart_418()         { buildEquityCurveChartForAccount('418'); }
function buildEquityCurveChart_973()         { buildEquityCurveChartForAccount('973'); }
function buildEquityCurveChart_317()         { buildEquityCurveChartForAccount('317'); }
function buildEquityCurveChartYearly_418()   { promptAndBuildYearlyEquityCurve_('418'); }
function buildEquityCurveChartYearly_973()   { promptAndBuildYearlyEquityCurve_('973'); }
function buildEquityCurveChartYearly_317()   { promptAndBuildYearlyEquityCurve_('317'); }

// ─── Batch helpers ────────────────────────────────────────────────

/** Daily trigger target — captures Net Liq for all accounts (skipIfExists). */
function fetchTodayNetLiq() {
  ACCOUNT_ORDER.forEach(function(suffix) {
    try { fetchTodayNetLiqForAccount(suffix, true); }
    catch (e) { console.error('Daily snapshot failed for account ' + suffix + ': ' + e.message); }
  });
}

/** Refresh benchmarks + all snapshots + rebuild all charts. */
function refreshAllData() {
  fetchSPYHistory();
  ACCOUNT_ORDER.forEach(function(suffix) {
    fetchTodayNetLiqForAccount(suffix, true);
    buildEquityCurveChartForAccount(suffix);
  });
}
