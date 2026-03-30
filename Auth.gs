/**
 * Auth.gs — Schwab OAuth 2.0 (Authorization Code flow)
 *
 * Uses the Apps Script OAuth2 library.
 * Library ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *
 * Credentials are stored in Script Properties (never in the sheet).
 * Tokens are stored in User Properties (per-user).
 */

// ─────────────────────────────────────────────────────────────────
// Public helpers
// ─────────────────────────────────────────────────────────────────

/** Opens the setup credentials dialog. */
function showSetupDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Setup')
    .setWidth(460)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Schwab API Setup');
}

/** Saves credentials entered in Setup.html and kicks off OAuth. */
function saveCredentials(clientId, clientSecret, accountNumber) {
  const props = PropertiesService.getScriptProperties();
  const toSave = {
    SCHWAB_CLIENT_ID:     clientId.trim(),
    SCHWAB_CLIENT_SECRET: clientSecret.trim(),
  };
  if (accountNumber && accountNumber.trim()) {
    toSave.SCHWAB_ACCOUNT_NUMBER = accountNumber.trim();
  }
  props.setProperties(toSave);

  // Clear any stale token so the user re-authorizes cleanly
  PropertiesService.getUserProperties().deleteProperty('oauth2.Schwab');

  return 'Credentials saved! Click the authorization link to complete setup.';
}

/** Shows current auth status in a simple alert. */
function checkAuthStatus() {
  const ui = SpreadsheetApp.getUi();
  try {
    const service = getSchwabService_();
    if (service.hasAccess()) {
      ui.alert('Authorization Status', '✅  Authorized with Schwab.', ui.ButtonSet.OK);
    } else {
      const result = ui.alert(
        'Authorization Status',
        '❌  Not authorized.\n\nWould you like to start the authorization flow now?',
        ui.ButtonSet.YES_NO
      );
      if (result === ui.Button.YES) startAuthFlow_();
    }
  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/** Called by Setup.html after credentials are saved to launch the OAuth flow. */
function authorizeWithSchwab() {
  startAuthFlow_();
}

/** OAuth2 callback — Schwab redirects here after the user approves. */
function authCallback(request) {
  const service = getSchwabService_();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif;padding:24px;text-align:center">
         <h2 style="color:#1a73e8">✅ Authorization Successful</h2>
         <p>Schwab access has been granted. You can close this tab and return to your spreadsheet.</p>
       </div>`
    );
  }
  return HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;padding:24px;text-align:center">
       <h2 style="color:#ea4335">❌ Authorization Failed</h2>
       <p>Please close this tab and try again from the Equity Curve menu.</p>
     </div>`
  );
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers (suffix _ = private)
// ─────────────────────────────────────────────────────────────────

function getSchwabService_() {
  const props = PropertiesService.getScriptProperties();
  const clientId     = props.getProperty('SCHWAB_CLIENT_ID');
  const clientSecret = props.getProperty('SCHWAB_CLIENT_SECRET');

  if (!clientId || !clientSecret) {
    throw new Error(
      'Schwab credentials not found. Please run "Setup & Authorize Schwab" from the menu first.'
    );
  }

  return OAuth2.createService('Schwab')
    .setAuthorizationBaseUrl('https://api.schwabapi.com/v1/oauth/authorize')
    .setTokenUrl('https://api.schwabapi.com/v1/oauth/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    // Schwab requires Basic auth on the token endpoint
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret),
    })
    .setScope('readonly');
}

function startAuthFlow_() {
  const service = getSchwabService_();
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Already authorized with Schwab!');
    return;
  }
  const authUrl = service.getAuthorizationUrl();
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;padding:20px">
       <h3 style="color:#1a73e8;margin-top:0">Authorize Schwab Access</h3>
       <p>Click the button below. A new tab will open where you log in to Schwab and approve access.</p>
       <a href="${authUrl}" target="_blank"
          style="display:inline-block;background:#1a73e8;color:white;padding:10px 22px;
                 text-decoration:none;border-radius:4px;font-size:14px">
         Authorize with Schwab
       </a>
       <p style="color:#888;font-size:12px;margin-top:16px">
         After approving, you can close this dialog. The callback page will confirm success.
       </p>
     </div>`
  ).setWidth(420).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Authorize Schwab');
}

/** Returns a valid Bearer token; throws if not authorized. */
function getAccessToken_() {
  const service = getSchwabService_();
  if (!service.hasAccess()) {
    throw new Error('Not authorized. Please run "Setup & Authorize Schwab" from the menu.');
  }
  return service.getAccessToken();
}
