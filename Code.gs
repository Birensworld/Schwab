/**
 * Schwab Equity Curve — Google Apps Script
 *
 * Plots your Portfolio Net Liquidity against SPY daily closes, both indexed
 * to 100 at a common start date, from 2020 to present.
 *
 * ─────────────────────────────────────────────────────────────────
 * SETUP (one-time)
 * ─────────────────────────────────────────────────────────────────
 * 1. In Apps Script editor: Resources > Libraries
 *    Add OAuth2 library ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 * 2. Register an app at https://developer.schwab.com to get your
 *    Client ID (App Key) and Client Secret.
 *    Set the callback URL to:
 *    https://script.google.com/macros/d/<YOUR_SCRIPT_ID>/usercallback
 * 3. Open this spreadsheet, use the "📈 Equity Curve" menu →
 *    "Setup & Authorize Schwab", enter your credentials, then click
 *    the authorization link.
 *
 * ─────────────────────────────────────────────────────────────────
 * WORKFLOW
 * ─────────────────────────────────────────────────────────────────
 * Historical data (2020 → today):
 *   A) "Fetch SPY History"         → pulls daily SPY closes via the
 *      Schwab Market Data API (accurate, complete).
 *   B) "Reconstruct from Transactions" → estimates daily portfolio
 *      values by replaying all transactions since 2020.
 *      ⚠ Captures realized P&L + external cash flows only.
 *         Unrealized P&L on open positions is approximated linearly.
 *         For higher accuracy, use "Import Net Liq from CSV" instead
 *         (export from schwab.com → Accounts → Performance → Export).
 *
 * Ongoing (recommended):
 *   "Enable Daily Auto-Capture" installs a 4:30 PM ET trigger that
 *   logs your live Net Liquidity every trading day.
 *
 * Build chart:
 *   "Build / Refresh Equity Curve Chart" generates a line chart on
 *   the 'Equity Curve' sheet comparing both series.
 */

// ─── Sheet & date constants ───────────────────────────────────────
const SHEET_NET_LIQ   = 'Net Liquidity';
const SHEET_SPY       = 'SPY History';
const SHEET_CHART     = 'Equity Curve';
const HISTORY_START   = '2020-01-01';      // earliest date to pull

// ─────────────────────────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📈 Equity Curve')
    .addItem('⚙️  Setup & Authorize Schwab',         'showSetupDialog')
    .addItem('✅  Check Authorization Status',        'checkAuthStatus')
    .addSeparator()
    .addSubMenu(ui.createMenu('📥  Fetch / Import Data')
      .addItem('Fetch SPY History (2020 → Today)',    'fetchSPYHistory')
      .addItem("Capture Today's Net Liquidity",       'fetchTodayNetLiq')
      .addItem('Reconstruct History from Transactions ⚠️', 'reconstructNetLiqHistory')
      .addItem('Import Net Liq from CSV…',            'showCsvImportDialog'))
    .addSeparator()
    .addItem('📈  Build / Refresh Equity Curve Chart', 'buildEquityCurveChart')
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰  Automation')
      .addItem('Enable Daily Snapshot (4:30 PM ET)',  'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',              'removeDailyTrigger'))
    .addToUi();
}

/**
 * One-shot helper: fetch everything and rebuild the chart.
 * Useful after initial setup.
 */
function refreshAllData() {
  fetchSPYHistory();
  fetchTodayNetLiq();
  buildEquityCurveChart();
}
