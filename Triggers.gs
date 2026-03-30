/**
 * Triggers.gs — Time-driven trigger management.
 *
 * Installs a daily trigger that fires at 4:30 PM Eastern Time each day
 * to automatically log your Net Liquidity after the market closes.
 *
 * Only one trigger is maintained per handler function.
 */

const DAILY_HANDLER = 'fetchTodayNetLiq';

// ─────────────────────────────────────────────────────────────────

/** Installs (or replaces) the daily 4:30 PM ET Net Liquidity trigger. */
function setupDailyTrigger() {
  removeDailyTrigger();   // ensure no duplicate

  ScriptApp.newTrigger(DAILY_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(16)           // 4 PM — Apps Script rounds to nearest 15 min
    .nearMinute(30)
    .inTimezone('America/New_York')
    .create();

  SpreadsheetApp.getUi().alert(
    'Daily Auto-Capture Enabled',
    '✅  Net Liquidity will be automatically captured at ~4:30 PM ET each day ' +
    'and logged to the "' + SHEET_NET_LIQ + '" sheet.\n\n' +
    'The chart will NOT auto-rebuild — run "Build / Refresh Equity Curve Chart" ' +
    'whenever you want to update it.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Removes the daily Net Liquidity trigger if it exists. */
function removeDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === DAILY_HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** Convenience wrapper called by the menu item. */
function removeDailyTrigger_menu() {
  removeDailyTrigger();
  SpreadsheetApp.getUi().alert(
    'Daily Auto-Capture Disabled',
    'The daily Net Liquidity snapshot trigger has been removed.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
