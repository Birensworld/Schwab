/**
 * Schwab Equity Curve — Google Apps Script
 *
 * Sheet layout:
 *   "Net Liquidity"  — Date | Net Liq 418 ($) | Net Liq 973 ($) | Source
 *   "SPY History"    — Date | SPY Close | SPY Indexed | QQQ Close | QQQ Indexed
 *   "Equity Curve 418" / "Equity Curve 973" — per-account charts vs SPY + QQQ
 *
 * ─── Account registry ────────────────────────────────────────────
 * Keys are the last-3-digit suffix used throughout; values are full account numbers.
 * ACCOUNT_ORDER determines column order in the Net Liquidity sheet.
 */
const ACCOUNT_CONFIGS = {
  '418': '52172418',
  '973': '55262973',
};
const ACCOUNT_ORDER = ['418', '973'];   // col 2 = 418, col 3 = 973

// Sheet name constants
const SHEET_NET_LIQ  = 'Net Liquidity'; // single shared sheet, one col per account
const SHEET_SPY      = 'SPY History';
const HISTORY_START  = '2026-01-01';

/** Returns the column index (1-based) for a given account suffix in the Net Liq sheet. */
function netLiqCol_(suffix) {
  return ACCOUNT_ORDER.indexOf(suffix) + 2;  // col 1 = Date, col 2+ = accounts
}

/** Returns the Equity Curve chart sheet name for an account. */
function chartSheetName_(suffix) {
  return 'Equity Curve ' + suffix;
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
    .addItem('📊 Fetch SPY + QQQ History',              'fetchSPYHistory')
    .addSeparator()
    .addSubMenu(ui.createMenu('💼 Account …418')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_418')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',     'buildEquityCurveChart_418'))
    .addSubMenu(ui.createMenu('💼 Account …973')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_973')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',     'buildEquityCurveChart_973'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger'))
    .addToUi();
}

// ─── Per-account menu wrappers ────────────────────────────────────
function fetchTodayNetLiq_418()      { fetchTodayNetLiqForAccount('418'); }
function fetchTodayNetLiq_973()      { fetchTodayNetLiqForAccount('973'); }
function buildEquityCurveChart_418() { buildEquityCurveChartForAccount('418'); }
function buildEquityCurveChart_973() { buildEquityCurveChartForAccount('973'); }

/**
 * Daily trigger target — captures Net Liq for every account,
 * skipping any date that already has a value (preserves manual entries).
 */
function fetchTodayNetLiq() {
  ACCOUNT_ORDER.forEach(function(suffix) {
    try {
      fetchTodayNetLiqForAccount(suffix, true); // true = skip if already exists
    } catch (e) {
      console.error('Daily snapshot failed for account ' + suffix + ': ' + e.message);
    }
  });
}

/** Refresh benchmarks + all account snapshots + rebuild all charts. */
function refreshAllData() {
  fetchSPYHistory();
  ACCOUNT_ORDER.forEach(function(suffix) {
    fetchTodayNetLiqForAccount(suffix, true);
    buildEquityCurveChartForAccount(suffix);
  });
}
