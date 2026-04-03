/**
 * Authentication.gs — Schwab OAuth 2.0 token management (manual flow)
 * Version: 1.0 (2026-04-03)
 *
 * Flow:
 *   Step 1 — ShowAuthUrl()      → opens browser to Schwab login page
 *   Step 2 — doGet()            → receives auth code via redirect URI (web app)
 *   Step 3 — ExchangeAuthCode() → exchanges code for access + refresh tokens
 *   Auto   — getValidAccessToken() / getAccessToken() → refreshes silently
 *
 * ⚠️  Replace the three placeholders below with your actual Schwab app credentials
 *     before running. Never commit real credentials to source control.
 */

const SW_CLIENT_ID     = "YOUR_CLIENT_ID_HERE";
const SW_CLIENT_SECRET = "YOUR_CLIENT_SECRET_HERE";
const SW_REDIRECT_URI  = "YOUR_REDIRECT_URI_HERE";

const SW_TOKEN_URL        = "https://api.schwabapi.com/v1/oauth/token";
const SW_ACCESS_BUFFER_MS = 90 * 1000;  // refresh 90 sec before expiry

// ─────────────────────────────────────────────────────────────────
// Web app entry point — receives OAuth redirect from Schwab
// ─────────────────────────────────────────────────────────────────

function doGet(e) {
  const code = e.parameter.code;
  if (code) {
    PropertiesService.getScriptProperties().setProperty("SW_AUTH_CODE", code);
    return ContentService.createTextOutput("Schwab authorization successful. Code saved.");
  }
  return ContentService.createTextOutput("No code received from Schwab.");
}

// ─────────────────────────────────────────────────────────────────
// Step 1 — Generate authorization URL
// ─────────────────────────────────────────────────────────────────

function ShowAuthUrl() {
  const url =
    "https://api.schwabapi.com/v1/oauth/authorize" +
    "?response_type=code" +
    "&client_id="    + encodeURIComponent(SW_CLIENT_ID) +
    "&redirect_uri=" + encodeURIComponent(SW_REDIRECT_URI) +
    "&scope="        + encodeURIComponent("read trade:write");
  SpreadsheetApp.getUi().alert(
    "Open this URL in your browser, log in, and approve access:\n\n" + url);
}

// ─────────────────────────────────────────────────────────────────
// Step 2 — Exchange authorization code for tokens
// ─────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────
// Token refresh — called automatically when access token expires
// ─────────────────────────────────────────────────────────────────

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
// Public helper — used by all other files that need a Bearer token
// ─────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

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
