/**
 * Code.gs — Schwab Portfolio + Equity Curve — Google Apps Script
 * Version: 2.0 (2026-04-03)
 *
 * Sheet layout:
 *   "Net Liquidity"     — Date | Net Liq 418 ($) | Net Liq 973 ($) | Net Liq 317 ($) | Total
 *   "SPY History"       — Date | SPY Close | SPY Indexed | QQQ Close | QQQ Indexed
 *   "Equity Curve 418/973/317" — per-account indexed charts vs SPY + QQQ
 *   "Actual Values 418/973/317" — per-account actual dollar charts vs SPY
 *
 * ─── Account registry ────────────────────────────────────────────
 * Keys are the last-3-digit suffix; values are full account numbers.
 * ACCOUNT_ORDER determines column order in the Net Liquidity sheet.
 *
 * ⚠️  CREDENTIALS: Store SW_CLIENT_ID, SW_CLIENT_SECRET, and SW_REDIRECT_URI
 *     in Script Properties (Project Settings → Script Properties) rather than
 *     hard-coding them here. Replace the placeholders below before running.
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

function netLiqCol_(suffix)  { return ACCOUNT_ORDER.indexOf(suffix) + 2; }
function totalNetLiqCol_()   { return ACCOUNT_ORDER.length + 2; }
function chartSheetName_(suffix) { return 'Equity Curve ' + suffix; }

// ─────────────────────────────────────────────────────────────────
// ⚠️  Replace these placeholders with your actual Schwab app credentials
// ─────────────────────────────────────────────────────────────────
const SW_CLIENT_ID     = "YOUR_CLIENT_ID_HERE";
const SW_CLIENT_SECRET = "YOUR_CLIENT_SECRET_HERE";
const SW_REDIRECT_URI  = "YOUR_REDIRECT_URI_HERE";

const SW_TOKEN_URL        = "https://api.schwabapi.com/v1/oauth/token";
const SW_ACCESS_BUFFER_MS = 90 * 1000;

// ─────────────────────────────────────────────────────────────────
// Menu
// ─────────────────────────────────────────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var authMenu = ui.createMenu('Setup / Re-Authorize')
    .addItem('Get Authorization Link',    'ShowAuthUrl')
    .addItem('Exchange Code for Tokens',  'ExchangeAuthCode');

  ui.createMenu('Schwab')
    .addItem('Refresh Portfolio', 'UpdateSheet')
    .addSeparator()
    .addSubMenu(authMenu)
    .addToUi();

  ui.createMenu('📈 Equity Curve')
    .addItem('📊 Fetch SPY + QQQ History', 'fetchSPYHistory')
    .addSeparator()
    .addSubMenu(ui.createMenu('💼 Account …418')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_418')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',  'buildEquityCurveChart_418')
      .addItem('Build / Refresh Actual Values Chart', 'buildActualValuesChart_418'))
    .addSubMenu(ui.createMenu('💼 Account …973')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_973')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',  'buildEquityCurveChart_973')
      .addItem('Build / Refresh Actual Values Chart', 'buildActualValuesChart_973'))
    .addSubMenu(ui.createMenu('💼 Account …317')
      .addItem("Capture Today's Net Liquidity (skip if exists)", 'fetchTodayNetLiq_317')
      .addSeparator()
      .addItem('Build / Refresh Equity Curve Chart',  'buildEquityCurveChart_317')
      .addItem('Build / Refresh Actual Values Chart', 'buildActualValuesChart_317'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Automation')
      .addItem('Enable Daily Snapshot – All Accounts (4:30 PM ET)', 'setupDailyTrigger')
      .addItem('Disable Daily Snapshot',                            'removeDailyTrigger'))
    .addToUi();
}

// ─── Per-account menu wrappers ────────────────────────────────────
function fetchTodayNetLiq_418()       { fetchTodayNetLiqForAccount('418'); }
function fetchTodayNetLiq_973()       { fetchTodayNetLiqForAccount('973'); }
function fetchTodayNetLiq_317()       { fetchTodayNetLiqForAccount('317'); }
function buildEquityCurveChart_418()  { buildEquityCurveChartForAccount('418'); }
function buildEquityCurveChart_973()  { buildEquityCurveChartForAccount('973'); }
function buildEquityCurveChart_317()  { buildEquityCurveChartForAccount('317'); }
function buildActualValuesChart_418() { buildActualValuesChartForAccount('418'); }
function buildActualValuesChart_973() { buildActualValuesChartForAccount('973'); }
function buildActualValuesChart_317() { buildActualValuesChartForAccount('317'); }

function fetchTodayNetLiq() {
  ACCOUNT_ORDER.forEach(function(suffix) {
    try { fetchTodayNetLiqForAccount(suffix, true); }
    catch (e) { console.error('Daily snapshot failed for account ' + suffix + ': ' + e.message); }
  });
}

function refreshAllData() {
  fetchSPYHistory();
  ACCOUNT_ORDER.forEach(function(suffix) {
    fetchTodayNetLiqForAccount(suffix, true);
    buildEquityCurveChartForAccount(suffix);
  });
}

// ─────────────────────────────────────────────────────────────────
// OAuth / Token management
// ─────────────────────────────────────────────────────────────────

function doGet(e) {
  const code = e.parameter.code;
  if (code) {
    PropertiesService.getScriptProperties().setProperty("SW_AUTH_CODE", code);
    return ContentService.createTextOutput("Schwab authorization successful. Code saved.");
  }
  return ContentService.createTextOutput("No code received from Schwab.");
}

function saveTokenResponse_(tokens) {
  const props = PropertiesService.getScriptProperties();
  const now   = Date.now();
  if (tokens.access_token) {
    props.setProperty("SW_ACCESS_TOKEN", tokens.access_token);
    props.setProperty("SW_ACCESS_TOKEN_EXPIRES_AT",
      String(now + ((tokens.expires_in || 1800) * 1000)));
  }
  if (tokens.refresh_token) {
    props.setProperty("SW_REFRESH_TOKEN", tokens.refresh_token);
    props.setProperty("SW_REFRESH_TOKEN_EXPIRES_AT",
      String(now + (7 * 24 * 60 * 60 * 1000)));
  }
}

function hasUsableAccessToken_() {
  const props     = PropertiesService.getScriptProperties();
  const token     = props.getProperty("SW_ACCESS_TOKEN");
  const expiresAt = Number(props.getProperty("SW_ACCESS_TOKEN_EXPIRES_AT") || 0);
  return !!token && Date.now() < (expiresAt - SW_ACCESS_BUFFER_MS);
}

function hasUsableRefreshToken_() {
  const props     = PropertiesService.getScriptProperties();
  const token     = props.getProperty("SW_REFRESH_TOKEN");
  const expiresAt = Number(props.getProperty("SW_REFRESH_TOKEN_EXPIRES_AT") || 0);
  return !!token && Date.now() < expiresAt;
}

function getValidAccessToken() {
  const props = PropertiesService.getScriptProperties();
  if (hasUsableAccessToken_()) return props.getProperty("SW_ACCESS_TOKEN");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (hasUsableAccessToken_()) return props.getProperty("SW_ACCESS_TOKEN");
    if (!hasUsableRefreshToken_())
      throw new Error("Refresh token missing or expired. Run Step 1 and Step 2 again.");
    return getAccessToken();
  } finally {
    lock.releaseLock();
  }
}

function ShowAuthUrl() {
  const url =
    "https://api.schwabapi.com/v1/oauth/authorize" +
    "?response_type=code" +
    "&client_id="     + encodeURIComponent(SW_CLIENT_ID) +
    "&redirect_uri="  + encodeURIComponent(SW_REDIRECT_URI) +
    "&scope="         + encodeURIComponent("read trade:write");
  SpreadsheetApp.getUi().alert(
    "Open this URL in your browser, log in, and approve access:\n\n" + url);
}

function ExchangeAuthCode() {
  const code = PropertiesService.getScriptProperties().getProperty("SW_AUTH_CODE");
  if (!code) {
    SpreadsheetApp.getUi().alert("No authorization code found. Please run Step 1 first.");
    return;
  }
  const response = UrlFetchApp.fetch(SW_TOKEN_URL, {
    method: "post",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(SW_CLIENT_ID + ":" + SW_CLIENT_SECRET),
      "Content-Type":  "application/x-www-form-urlencoded"
    },
    payload: { grant_type: "authorization_code", code: code, redirect_uri: SW_REDIRECT_URI },
    muteHttpExceptions: true
  });
  const tokens = JSON.parse(response.getContentText());
  if (tokens.access_token || tokens.refresh_token) {
    saveTokenResponse_(tokens);
    PropertiesService.getScriptProperties().deleteProperty("SW_AUTH_CODE");
  } else {
    SpreadsheetApp.getUi().alert("Error exchanging code:\n" + response.getContentText());
  }
}

function getAccessToken() {
  const props        = PropertiesService.getScriptProperties();
  const refreshToken = props.getProperty("SW_REFRESH_TOKEN");
  if (!refreshToken) {
    SpreadsheetApp.getUi().alert("No Schwab refresh token saved. Please run Step 2 first.");
    return null;
  }
  const response = UrlFetchApp.fetch(SW_TOKEN_URL, {
    method: "post",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(SW_CLIENT_ID + ":" + SW_CLIENT_SECRET),
      "Content-Type":  "application/x-www-form-urlencoded"
    },
    payload: { grant_type: "refresh_token", refresh_token: refreshToken },
    muteHttpExceptions: true
  });
  const tokens = JSON.parse(response.getContentText());
  if (tokens.access_token) {
    saveTokenResponse_(tokens);
    return tokens.access_token;
  }
  SpreadsheetApp.getUi().alert("Error refreshing token:\n" + response.getContentText());
  return null;
}

// ─────────────────────────────────────────────────────────────────
// Portfolio refresh (UpdateSheet)
// ─────────────────────────────────────────────────────────────────

function setSheetFont() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schwab");
  if (!sheet) return;
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily("Nunito");
}

function getAccountLabel(acctId) {
  switch (acctId) {
    case "30050317": return "SW Equity Sub";
    case "52172418": return "SW Equity";
    default:         return "SW IRA";
  }
}

function UpdateSheet() {
  let token;
  try { token = getValidAccessToken(); }
  catch (e) { SpreadsheetApp.getUi().alert("No valid token: " + e.message); return; }
  if (!token) { SpreadsheetApp.getUi().alert("No valid token. Please run Step 2."); return; }

  const url      = "https://api.schwabapi.com/trader/v1/accounts?fields=positions";
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });
  let data;
  try { data = JSON.parse(response.getContentText()); }
  catch (e) { SpreadsheetApp.getUi().alert("Error parsing Schwab response."); return; }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("Schwab");
  if (!sheet) sheet = ss.insertSheet("Schwab");
  sheet.clear();
  let rowIndex = 1;
  if (!Array.isArray(data)) { console.error("Unexpected JSON:", data); return; }
  setSheetFont();

  const accountOrderMap = { "SW Equity": 1, "SW IRA": 2, "SW Equity Sub": 3 };
  const allAccounts = [];

  data.forEach(wrapper => {
    const account  = wrapper.securitiesAccount;
    if (!account) return;
    const balances = account.currentBalances || {};
    const acctId   = account.accountNumber || "";
    const acctLabel = getAccountLabel(acctId);
    const acctValue = balances.liquidationValue || 0;
    const cash      = balances.cashBalance || 0;
    allAccounts.push({
      accountId:   acctId,
      label:       acctLabel,
      value:       acctValue,
      cash:        cash,
      cashPercent: acctValue ? (cash / acctValue) : 0,
      order:       accountOrderMap[acctLabel] || 999,
      positions:   account.positions || []
    });
  });

  allAccounts.sort((a, b) => a.order - b.order);

  allAccounts.forEach(acct => {
    const accountRange = sheet.getRange(rowIndex, 1, 1, 4);
    accountRange.setValues([[acct.label, acct.value, acct.cash, acct.cashPercent]]);
    accountRange.setFontWeight("bold");
    sheet.getRange(rowIndex, 1).setBackground("#d9d9d9").setFontSize(12);
    sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat("#,##0.00").setBackground("#d0f0c0");
    sheet.getRange(rowIndex, 4).setNumberFormat("0.00%").setBackground("#d0f0c0");
    rowIndex++;

    const posHeader = ["Symbol","Qty","Chg %","Avg Price","MV","P/L","P/L %","Weight"];
    const posHeaderRange = sheet.getRange(rowIndex, 1, 1, posHeader.length);
    posHeaderRange.setValues([posHeader]).setFontWeight("bold").setBackground("#cfe2f3");
    rowIndex++;

    const positions = acct.positions.filter(pos =>
      ["EQUITY","ETF","MUTUAL_FUND","COLLECTIVE_INVESTMENT"].includes(pos.instrument?.assetType));
    positions.sort((a, b) => (b.marketValue || 0) - (a.marketValue || 0));

    if (positions.length > 0) {
      const posRows = positions.map(pos => {
        const qty      = (pos.longQuantity || 0) - (pos.shortQuantity || 0);
        const avgPrice = pos.averagePrice || 0;
        const mv       = pos.marketValue || 0;
        const openPL   = pos.longOpenProfitLoss || 0;
        return [
          pos.instrument?.symbol || "",
          qty, "",
          avgPrice, mv, openPL,
          (avgPrice && qty) ? (openPL / (avgPrice * qty)) : 0,
          acct.value ? (mv / acct.value) : 0
        ];
      });

      const posRange = sheet.getRange(rowIndex, 1, posRows.length, posHeader.length);
      posRange.setValues(posRows);
      sheet.getRange(rowIndex, 1, posRows.length, 1).setFontWeight("bold");
      sheet.getRange(rowIndex, 2, posRows.length, 1).setNumberFormat("#,##0").setFontWeight("bold");
      sheet.getRange(rowIndex, 4, posRows.length, 3).setNumberFormat("#,##0.00");
      sheet.getRange(rowIndex, 5, posRows.length, 1).setFontWeight("bold");
      sheet.getRange(rowIndex, 7, posRows.length, 1).setNumberFormat("0.00%").setFontWeight("bold");
      sheet.getRange(rowIndex, 8, posRows.length, 1).setNumberFormat("0.00%").setFontWeight("bold");
      sheet.getRange(rowIndex, 9, posRows.length, 1).setFontColor("#666666").setFontSize(9).setBackground("#f5f5f5");

      const formulas = positions.map(pos =>
        [`=IFERROR(GOOGLEFINANCE("${pos.instrument?.symbol}", "changepct")/100, 0)`]);
      sheet.getRange(rowIndex, 3, formulas.length, 1).setFormulas(formulas);
      sheet.getRange(rowIndex, 9, posRows.length, 1)
        .setValues(new Array(posRows.length).fill([acct.accountId]));

      const rules = sheet.getConditionalFormatRules();
      [[3, 0], [6, 0], [7, -7]].forEach(([col, threshold]) => {
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0).setFontColor("green").setBold(true)
          .setRanges([sheet.getRange(rowIndex, col, posRows.length, 1)]).build());
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0).setFontColor("red").setBold(true)
          .setRanges([sheet.getRange(rowIndex, col, posRows.length, 1)]).build());
        rules.push(SpreadsheetApp.newConditionalFormatRule()
          .whenNumberEqualTo(0).setFontColor("black").setBold(true)
          .setRanges([sheet.getRange(rowIndex, col, posRows.length, 1)]).build());
      });
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(-7).setBackground("#FFCCCB").setFontColor("black").setBold(true)
        .setRanges([sheet.getRange(rowIndex, 7, posRows.length, 1)]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=AND($G" + rowIndex + "<0,$G" + rowIndex + ">-7)")
        .setFontColor("red").setBold(true)
        .setRanges([sheet.getRange(rowIndex, 7, posRows.length, 1)]).build());
      sheet.setConditionalFormatRules(rules);

      rowIndex += posRows.length;
    } else {
      sheet.getRange(rowIndex, 1).setValue("No positions for this account");
      rowIndex++;
    }
    rowIndex++;
  });

  const totalValue = allAccounts.reduce((s, a) => s + a.value, 0);
  const totalCash  = allAccounts.reduce((s, a) => s + a.cash,  0);
  const totalsRange = sheet.getRange(rowIndex, 1, 1, 4);
  totalsRange.setValues([["TOTALS", totalValue, totalCash, totalValue ? (totalCash / totalValue) : 0]]);
  totalsRange.setFontWeight("bold").setFontSize(11).setBackground("#d0f0c0");
  sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat("#,##0.00");
  sheet.getRange(rowIndex, 4).setNumberFormat("0.00%");
}

// ─────────────────────────────────────────────────────────────────
// Liquidation
// ─────────────────────────────────────────────────────────────────

function close_SW_Equity() { closePositionForAccount("52172418", "SW Equity"); }
function close_SW_IRA()    { closePositionForAccount("55262973", "SW IRA");    }

function getSchwabAccountMap() {
  let token = getValidAccessToken();
  const url  = "https://api.schwabapi.com/trader/v1/accounts/accountNumbers";
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "Authorization": "Bearer " + token, "Accept": "application/json" },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code === 401) { token = getAccessToken(); if (!token) throw new Error("Unable to refresh token."); return getSchwabAccountMap(); }
  if (code !== 200 || !text) throw new Error("Failed to fetch account map.");
  const json = JSON.parse(text);
  const arr  = Array.isArray(json.accounts) ? json.accounts : Array.isArray(json) ? json : [];
  const map  = {};
  arr.forEach(a => { if (a.accountNumber && (a.hashValue || a.accountId)) map[a.accountNumber] = a.hashValue || a.accountId; });
  PropertiesService.getScriptProperties().setProperty("SCHWAB_ACCOUNT_MAP", JSON.stringify(map));
  return map;
}

function placeSchwabOrder(acctId, token, instruction, symbol, qty) {
  const order = {
    orderType: "MARKET", orderStrategyType: "SINGLE", session: "NORMAL", duration: "DAY",
    orderLegCollection: [{ instruction, quantity: qty, instrument: { symbol, assetType: "EQUITY" } }]
  };
  const resp = UrlFetchApp.fetch(`https://api.schwabapi.com/trader/v1/accounts/${acctId}/orders`, {
    method: "post",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json", "Accept": "application/json" },
    payload: JSON.stringify(order),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode(), body = resp.getContentText();
  if (code === 401) throw new Error("Unauthorized. Refresh token and retry.");
  if (code >= 400)  throw new Error(`Order error [${code}]: ${body}`);
  if (!body || body.trim() === "") return { orderId: "", status: "ACCEPTED" };
  try { return JSON.parse(body); } catch (e) { return { orderId: "", status: "ACCEPTED" }; }
}

function logTradeResultSchwab(accountNumber, accountName, symbol, qty, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log  = ss.getSheetByName("Trade Log") || ss.insertSheet("Trade Log");
  if (log.getLastRow() === 0)
    log.appendRow(["Timestamp","Broker","Account #","Account Name","Symbol","Qty","Filled Price","Order ID","Status"]);
  log.appendRow([new Date(), "Schwab", accountNumber, accountName, symbol, qty,
    result.price || "", result.orderId || "", result.status || ""]);
}

function closePositionForAccount(accountNumber, accountName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Schwab");
  const ui    = SpreadsheetApp.getUi();
  if (!sheet) return ui.alert("⚠️ Sheet 'Schwab' not found.");

  const data     = sheet.getDataRange().getValues().slice(1);
  const acctRows = data.filter(r => String(r[8]).trim() === String(accountNumber).trim());
  if (!acctRows.length) return ui.alert(`No positions found for ${accountName} (#${accountNumber}).`);

  const confirm = ui.alert(
    `Confirm liquidation for ${accountName} (#${accountNumber})?\n\n${acctRows.length} positions will be SOLD at MARKET.`,
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  const mapJson = PropertiesService.getScriptProperties().getProperty("SCHWAB_ACCOUNT_MAP");
  if (!mapJson) return ui.alert("❌ Account map not found. Run getSchwabAccountMap() first.");
  const acctId  = JSON.parse(mapJson)[accountNumber];
  if (!acctId)   return ui.alert(`❌ Could not resolve trading ID for account ${accountNumber}.`);

  const token = getValidAccessToken();
  let successCount = 0, failCount = 0;

  acctRows.forEach(r => {
    const symbol = String(r[0]).trim();
    const qty    = parseFloat(String(r[1]).replace(/,/g, ""));
    if (!symbol || isNaN(qty) || qty <= 0) return;
    try {
      const result = placeSchwabOrder(acctId, token, "SELL", symbol, qty);
      logTradeResultSchwab(accountNumber, accountName, symbol, qty, result);
      successCount++;
    } catch (err) {
      console.error(`Error selling ${symbol}: ${err}`);
      failCount++;
    }
  });

  const msg = `Completed liquidation for ${accountName}\nSuccess: ${successCount}, Failed: ${failCount}`;
  ui.alert(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
}
