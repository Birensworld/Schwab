/**
 * Liquidation.gs — Sell all positions for a Schwab account at market
 * Version: 1.0 (2026-04-03)
 *
 * Entry points (assigned to sheet buttons):
 *   close_SW_Equity() — liquidates account 52172418 (SW Equity)
 *   close_SW_IRA()    — liquidates account 55262973 (SW IRA)
 *
 * Requires SCHWAB_ACCOUNT_MAP in Script Properties.
 * Run getSchwabAccountMap() once to populate it.
 *
 * Depends on: Authentication.gs → getValidAccessToken(), getAccessToken()
 */

// ─────────────────────────────────────────────────────────────────
// Button entry points — assign these to liquidation buttons on the sheet
// ─────────────────────────────────────────────────────────────────

function close_SW_Equity() { closePositionForAccount("52172418", "SW Equity"); }
function close_SW_IRA()    { closePositionForAccount("55262973", "SW IRA");    }

// ─────────────────────────────────────────────────────────────────
// Account map — run once to cache account number → hashed trading ID
// ─────────────────────────────────────────────────────────────────

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

  if (code === 401) {
    token = getAccessToken();
    if (!token) throw new Error("Unable to refresh token.");
    return getSchwabAccountMap();
  }
  if (code !== 200 || !text) throw new Error("Failed to fetch account map.");

  const json = JSON.parse(text);
  const arr  = Array.isArray(json.accounts) ? json.accounts : Array.isArray(json) ? json : [];
  const map  = {};
  arr.forEach(a => {
    if (a.accountNumber && (a.hashValue || a.accountId))
      map[a.accountNumber] = a.hashValue || a.accountId;
  });

  PropertiesService.getScriptProperties().setProperty("SCHWAB_ACCOUNT_MAP", JSON.stringify(map));
  Logger.log("✅ Saved Schwab Account Map: " + JSON.stringify(map));
  return map;
}

// ─────────────────────────────────────────────────────────────────
// Core: place a single market order via Schwab API
// ─────────────────────────────────────────────────────────────────

function placeSchwabOrder(acctId, token, instruction, symbol, qty) {
  const order = {
    orderType: "MARKET", orderStrategyType: "SINGLE", session: "NORMAL", duration: "DAY",
    orderLegCollection: [{
      instruction, quantity: qty,
      instrument: { symbol, assetType: "EQUITY" }
    }]
  };
  const resp = UrlFetchApp.fetch(
    `https://api.schwabapi.com/trader/v1/accounts/${acctId}/orders`, {
      method: "post",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type":  "application/json",
        "Accept":        "application/json"
      },
      payload: JSON.stringify(order),
      muteHttpExceptions: true
    });
  const code = resp.getResponseCode();
  const body = resp.getContentText();

  if (code === 401) throw new Error("Unauthorized. Refresh token and retry.");
  if (code >= 400)  throw new Error(`Order error [${code}]: ${body}`);
  if (!body || body.trim() === "") return { orderId: "", status: "ACCEPTED" };
  try { return JSON.parse(body); } catch (e) { return { orderId: "", status: "ACCEPTED" }; }
}

// ─────────────────────────────────────────────────────────────────
// Trade log — appends each executed order to the "Trade Log" sheet
// ─────────────────────────────────────────────────────────────────

function logTradeResultSchwab(accountNumber, accountName, symbol, qty, result) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log  = ss.getSheetByName("Trade Log") || ss.insertSheet("Trade Log");
  if (log.getLastRow() === 0)
    log.appendRow(["Timestamp","Broker","Account #","Account Name","Symbol","Qty",
                   "Filled Price","Order ID","Status"]);
  log.appendRow([new Date(), "Schwab", accountNumber, accountName, symbol, qty,
    result.price || "", result.orderId || "", result.status || ""]);
}

// ─────────────────────────────────────────────────────────────────
// Main liquidation function — sells all positions for one account
// ─────────────────────────────────────────────────────────────────

function closePositionForAccount(accountNumber, accountName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Schwab");
  const ui    = SpreadsheetApp.getUi();
  if (!sheet) return ui.alert("⚠️ Sheet 'Schwab' not found.");

  // Read positions from column I (hidden account ID written by AccountRefresh.gs)
  const data     = sheet.getDataRange().getValues().slice(1);
  const acctRows = data.filter(r => String(r[8]).trim() === String(accountNumber).trim());
  if (!acctRows.length)
    return ui.alert(`No positions found for ${accountName} (#${accountNumber}).`);

  const confirm = ui.alert(
    `Confirm liquidation for ${accountName} (#${accountNumber})?\n\n` +
    `${acctRows.length} positions will be SOLD at MARKET.`,
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
      Logger.log(`✅ Sold ${qty} ${symbol}`);
    } catch (err) {
      console.error(`Error selling ${symbol}: ${err}`);
      failCount++;
    }
  });

  const msg = `Completed liquidation for ${accountName}\nSuccess: ${successCount}, Failed: ${failCount}`;
  ui.alert(msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);
}
