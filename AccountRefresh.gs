/**
 * AccountRefresh.gs — Schwab portfolio refresh (balances + positions + totals)
 * Version: 1.2 (2026-04-03)
 *
 * Writes to the "Schwab" sheet:
 *   - One header row per account 
 *   - Position rows (symbol, qty, chg%, avg price, MV, P/L, P/L %, weight)
 *   - Totals row at the bottom (sum of all accounts)
 *
 * Depends on: Authentication.gs → getValidAccessToken()
 */

// ─────────────────────────────────────────────────────────────────
// Public entry point — called by Schwab → Refresh Portfolio menu item
// ─────────────────────────────────────────────────────────────────

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

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Schwab");
  if (!sheet) sheet = ss.insertSheet("Schwab");
  sheet.clear();
  let rowIndex = 1;
  if (!Array.isArray(data)) { console.error("Unexpected JSON:", data); return; }
  setSheetFont();

  // ── Build and sort account list ───────────────────────────────
  const accountOrderMap = { "SW Equity": 1, "SW IRA": 2, "SW Equity Sub": 3 };
  const allAccounts = [];

  data.forEach(wrapper => {
    const account   = wrapper.securitiesAccount;
    if (!account) return;
    const balances  = account.currentBalances || {};
    const acctId    = account.accountNumber || "";
    const acctLabel = getAccountLabel(acctId);
    const acctValue = balances.liquidationValue || 0;
    const cash      = balances.cashBalance || 0;
    const ytdPL = getYtdPLPercent_(acctId, acctValue);
    allAccounts.push({
      accountId:   acctId,
      label:       acctLabel,
      value:       acctValue,
      cash:        cash,
      allocPercent: (acctValue - cash) / acctValue,
      ytdPL:       ytdPL,
      order:       accountOrderMap[acctLabel] || 999,
      positions:   account.positions || []
    });
  });

  allAccounts.sort((a, b) => a.order - b.order);

  // ── Write each account block ──────────────────────────────────
  allAccounts.forEach(acct => {
    // Account header row
    const accountRange = sheet.getRange(rowIndex, 1, 1, 8);
    accountRange.setValues([[
  acct.label,
  acct.value,
  acct.cash,
  "",
  "ALLOC:",
  acct.allocPercent,
  "YTD P/L:",
  acct.ytdPL
  ]]);
    accountRange.setFontWeight("bold");
    sheet.getRange(rowIndex, 1).setBackground("#d9d9d9").setFontSize(12);
    sheet.getRange(rowIndex, 2, 1, 8).setBackground("#d0f0c0"); //Green
    sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat("#,##0.00");
    sheet.getRange(rowIndex, 6).setNumberFormat("0.00%");
    sheet.getRange(rowIndex, 8).setNumberFormat("0.00%");

    // 🔹 Conditional formatting for YTD P/L (Column H)
const ytdCell = sheet.getRange(rowIndex, 8);

const rules = sheet.getConditionalFormatRules();

// Positive → Dark Green
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor("#006400") // dark green
    .setBold(true)
    .setRanges([ytdCell])
    .build()
);

// Negative → Red
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor("red")
    .setBold(true)
    .setRanges([ytdCell])
    .build()
);

// Zero → Black
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setFontColor("black")
    .setBold(true)
    .setRanges([ytdCell])
    .build()
);

    sheet.setConditionalFormatRules(rules);
    
    rowIndex++;

    // Position header row
    const posHeader = ["Symbol","Qty","Chg %","Avg Price","MV","P/L","P/L %","Weight"];
    sheet.getRange(rowIndex, 1, 1, posHeader.length)
      .setValues([posHeader]).setFontWeight("bold").setBackground("#cfe2f3");
    rowIndex++;

    // Filter to equity-type positions, sort by market value descending
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
          qty, "",           // Chg % filled by GOOGLEFINANCE formula below
          avgPrice, mv, openPL,
          (avgPrice && qty) ? (openPL / (avgPrice * qty)) : 0,
          acct.value ? (mv / acct.value) : 0
        ];
      });

      // Write all position rows at once
      sheet.getRange(rowIndex, 1, posRows.length, posHeader.length).setValues(posRows);

      // Column formatting
      sheet.getRange(rowIndex, 1, posRows.length, 1).setFontWeight("bold");
      sheet.getRange(rowIndex, 2, posRows.length, 1).setNumberFormat("#,##0").setFontWeight("bold");
      sheet.getRange(rowIndex, 4, posRows.length, 3).setNumberFormat("#,##0.00");
      sheet.getRange(rowIndex, 5, posRows.length, 1).setFontWeight("bold");
      sheet.getRange(rowIndex, 7, posRows.length, 1).setNumberFormat("0.00%").setFontWeight("bold");
      sheet.getRange(rowIndex, 8, posRows.length, 1).setNumberFormat("0.00%").setFontWeight("bold");
      sheet.getRange(rowIndex, 9, posRows.length, 1)
        .setFontColor("#666666").setFontSize(9).setBackground("#f5f5f5");

      // GOOGLEFINANCE formulas for Chg % column
      sheet.getRange(rowIndex, 3, positions.length, 1).setFormulas(
        positions.map(pos =>
          [`=IFERROR(GOOGLEFINANCE("${pos.instrument?.symbol}", "changepct")/100, 0)`]));

      // Hidden account ID in column I (used by Liquidation.gs)
      sheet.getRange(rowIndex, 9, posRows.length, 1)
        .setValues(new Array(posRows.length).fill([acct.accountId]));

      // Conditional formatting — Chg %, P/L, P/L %
      const rules = sheet.getConditionalFormatRules();
      [3, 6, 7].forEach(col => {
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
      // P/L % special: highlight deep losses
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
    rowIndex++; // blank separator between accounts
  });

  // ── Totals row ────────────────────────────────────────────────
  const totalValue = allAccounts.reduce((s, a) => s + a.value, 0);
  const totalCash  = allAccounts.reduce((s, a) => s + a.cash,  0);
  const totalsRange = sheet.getRange(rowIndex, 1, 1, 4);
  totalsRange.setValues([["TOTALS", totalValue, totalCash, totalValue ? (totalCash / totalValue) : 0]]);
  totalsRange.setFontWeight("bold").setFontSize(11).setBackground("#d0f0c0");
  sheet.getRange(rowIndex, 2, 1, 2).setNumberFormat("#,##0.00");
  sheet.getRange(rowIndex, 4).setNumberFormat("0.00%");
}

// ─────────────────────────────────────────────────────────────────
// Helpers
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

function getAccountSuffix_(acctId) {
  return String(acctId || "").slice(-3);
}

function getStartOfYearNetLiq_(suffix) {
  const map = getNetLiqMap_(suffix);
  if (!map || Object.keys(map).length === 0) {
    console.log("[YTD DEBUG] No NetLiq data for suffix " + suffix);
    return null;
  }

  const currentYear = String(new Date().getFullYear());
  let earliestDateStr = null;
  let earliestValue = null;

  Object.keys(map).forEach(dateStr => {
    dateStr = String(dateStr).trim();

    // Expect plain YYYY-MM-DD strings only
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;
    if (!dateStr.startsWith(currentYear + "-")) return;

    const v = Number(map[dateStr]);
    if (!isFinite(v) || v <= 0) return;

    // Lexicographic comparison works for YYYY-MM-DD
    if (!earliestDateStr || dateStr < earliestDateStr) {
      earliestDateStr = dateStr;
      earliestValue = v;
    }
  });

  console.log(
    "[YTD DEBUG] Suffix: " + suffix +
    " | StartDate: " + earliestDateStr +
    " | StartValue: " + earliestValue
  );

  return earliestValue;
}

function getYtdPLPercent_(acctId, acctValue) {
  if (!acctValue) return 0;

  const suffix = getAccountSuffix_(acctId);
  const startValue = getStartOfYearNetLiq_(suffix);

  const ytd = (!startValue || startValue <= 0)
    ? 0
    : (acctValue - startValue) / startValue;

  console.log(
    "[YTD DEBUG] Account: " + acctId +
    " | Suffix: " + suffix +
    " | StartValue: " + startValue +
    " | CurrentValue: " + acctValue +
    " | YTD%: " + ytd
  );

  return ytd;
}
