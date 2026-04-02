/**
 * Schwab Equity Curve — Google Apps Script
 *
 * Plots Portfolio Net Liquidity vs SPY, indexed to 100 at the first
 * common trading day, for one or more Schwab accounts.
 *
 * Each account gets its own "Net Liq ###" and "Equity Curve ###" tabs.
 * SPY History is shared across all accounts.
 *
 * ─── Account registry ────────────────────────────────────────────
 * Add accounts here as { 'last3digits': 'fullAccountNumber' }.
 * Account 418 is active; 973 is defined but not yet in the menu.
 */
const ACCOUNT_CONFIGS = {
  '418': '52172418',
  '973': '55262973',
};

// Shared constants
const SHEET_SPY     = 'SPY History';
const HISTORY_START = '2024-05-14';

/** Returns the sheet tab names for a given account suffix. */
function sheetNames_(suffix) {
  return {
    netLiq: 'Net Liq ' + suffix,
    chart:  'Equity Curve ' + suffix,
  };
}

// ─────────────────────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var authMenu = ui.createMenu('Setup / Re-Authorize')
    .addItem('Get Authorization Link', 'ShowAuthUrl')
    .addItem('Exchange Code for Tokens', 'ExchangeAuthCode');

  ui.createMenu('Schwab')
    .addItem('Refresh Portfolio', 'UpdateSheet')
    .addSeparator()
    .addSubMenu(authMenu)
    .addToUi();

  ui.createMenu('📈 Equity Curve')
    // ── SPY (shared) ──────────────────────────────────────────────
    .addItem('📊 Fetch SPY History (2020 → Today)', 'fetchSPYHistory')
    .addSeparator()
    // ── Account 418 ───────────────────────────────────────────────
    .addSubMenu(ui.createMenu('💼 Account …418')
      .addItem("Capture Today's Net Liquidity",            'fetchTodayNetLiq_418')
      .addItem('Reconstruct History from Transactions ⚠️', 'reconstructNetLiqHistory_418')
      .addItem('Import Net Liq from CSV…',                 'showCsvImportDialog')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',       'buildEquityCurveChart_418'))
    .addSeparator()
    // ── Automation ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger'))
    .addToUi();
}

// ─── Account 418 menu handlers ───────────────────────────────────
function fetchTodayNetLiq_418()         { fetchTodayNetLiqForAccount('418'); }
function reconstructNetLiqHistory_418() { reconstructNetLiqHistoryForAccount('418'); }
function buildEquityCurveChart_418()    { buildEquityCurveChartForAccount('418'); }

/**
 * Daily trigger target — captures Net Liq for every configured account.
 * Also callable manually from the Apps Script editor.
 */
function fetchTodayNetLiq() {
  Object.keys(ACCOUNT_CONFIGS).forEach(function(suffix) {
    try {
      fetchTodayNetLiqForAccount(suffix);
    } catch (e) {
      console.error('Daily snapshot failed for account ' + suffix + ': ' + e.message);
    }
  });
}

/** Refresh SPY + all accounts + rebuild all charts. */
function refreshAllData() {
  fetchSPYHistory();
  Object.keys(ACCOUNT_CONFIGS).forEach(function(suffix) {
    fetchTodayNetLiqForAccount(suffix);
    buildEquityCurveChartForAccount(suffix);
  });
}
