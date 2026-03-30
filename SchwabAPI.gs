/**
 * SchwabAPI.gs — Thin wrappers around the Schwab Trader & Market Data APIs.
 *
 * Trader API base:    https://api.schwabapi.com/trader/v1
 * Market Data base:   https://api.schwabapi.com/marketdata/v1
 *
 * All public functions return parsed JSON. They throw on HTTP errors.
 */

const TRADER_API_BASE = 'https://api.schwabapi.com/trader/v1';
const MARKET_API_BASE = 'https://api.schwabapi.com/marketdata/v1';

// ─────────────────────────────────────────────────────────────────
// Account endpoints
// ─────────────────────────────────────────────────────────────────

/**
 * Returns the list of linked accounts with their encrypted hashes.
 * Response: [{ accountNumber, hashValue }, …]
 */
function getAccountNumbers() {
  return apiGet_(`${TRADER_API_BASE}/accounts/accountNumbers`);
}

/**
 * Returns full account details including current balances and positions.
 * @param {string} encryptedAccountNumber - The hashValue from getAccountNumbers()
 */
function getAccountDetails(encryptedAccountNumber) {
  return apiGet_(`${TRADER_API_BASE}/accounts/${encryptedAccountNumber}`, {
    fields: 'positions',
  });
}

/**
 * Returns the Net Liquidation Value for the configured account.
 * Resolves and caches the account hash automatically.
 * @returns {number}
 */
function fetchCurrentNetLiquidity() {
  const accountData = resolveAccountDetails_();
  return extractNetLiquidityFromAccount_(accountData);
}

/**
 * Returns all transactions for an account between startDate and endDate.
 * Dates are ISO-8601 strings, e.g. '2020-01-01T00:00:00.000Z'.
 * @param {string} encryptedAccountNumber
 * @param {string} startDate
 * @param {string} endDate
 * @returns {Array}
 */
function getTransactions(encryptedAccountNumber, startDate, endDate) {
  return apiGet_(
    `${TRADER_API_BASE}/accounts/${encryptedAccountNumber}/transactions`,
    {
      startDate,
      endDate,
      types: 'TRADE,RECEIVE_AND_DELIVER,DIVIDEND_OR_INTEREST,ELECTRONIC_FUND,OTHER',
    }
  );
}

// ─────────────────────────────────────────────────────────────────
// Market Data endpoints
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches daily OHLCV candles for a symbol between two dates.
 * @param {string} symbol  e.g. 'SPY'
 * @param {string} startDate  'YYYY-MM-DD'
 * @param {string} endDate    'YYYY-MM-DD'
 * @returns {{ candles: [{datetime, open, high, low, close, volume}], symbol, empty }}
 */
function getPriceHistory(symbol, startDate, endDate) {
  const startMs = new Date(startDate).getTime();
  const endMs   = new Date(endDate).getTime();

  return apiGet_(`${MARKET_API_BASE}/pricehistory`, {
    symbol,
    periodType:           'year',
    frequencyType:        'daily',
    frequency:            1,
    startDate:            startMs,
    endDate:              endMs,
    needExtendedHoursData: false,
  });
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Resolves the encrypted account hash (caches in Script Properties)
 * then returns the account details object.
 */
function resolveAccountDetails_() {
  const props = PropertiesService.getScriptProperties();
  let hash = props.getProperty('SCHWAB_ACCOUNT_HASH');

  if (!hash) {
    const accounts = getAccountNumbers();
    if (!accounts || accounts.length === 0) {
      throw new Error('No accounts found on this Schwab profile.');
    }

    const preferredAcctNum = props.getProperty('SCHWAB_ACCOUNT_NUMBER');
    const matched = preferredAcctNum
      ? accounts.find(a => a.accountNumber === preferredAcctNum)
      : null;

    const chosen = matched || accounts[0];
    hash = chosen.hashValue;
    props.setProperty('SCHWAB_ACCOUNT_HASH', hash);
  }

  return getAccountDetails(hash);
}

/**
 * Extracts the Net Liquidation Value from an account details response.
 * Schwab may return the value under slightly different field names
 * depending on account type (margin vs cash vs IRA).
 * @param {Object} accountData
 * @returns {number}
 */
function extractNetLiquidityFromAccount_(accountData) {
  const sec = accountData.securitiesAccount || accountData;
  const bal = sec.currentBalances || sec.projectedBalances || {};

  return (
    bal.liquidationValue     ??
    bal.netLiquidation       ??
    bal.totalAccountValue    ??
    bal.cashBalance          ??
    0
  );
}

/**
 * Generic authenticated GET request.
 * @param {string} url
 * @param {Object} [params]  Query-string key/value pairs
 * @returns {Object} Parsed JSON response
 */
function apiGet_(url, params) {
  const token = getAccessToken_();

  let fullUrl = url;
  if (params && Object.keys(params).length > 0) {
    const qs = Object.entries(params)
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&');
    fullUrl = `${url}?${qs}`;
  }

  const options = {
    method:           'GET',
    headers:          { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(fullUrl, options);
  const code     = response.getResponseCode();
  const body     = response.getContentText();

  if (code === 401) {
    // Token may have expired — clear it so the next call re-authenticates
    PropertiesService.getUserProperties().deleteProperty('oauth2.Schwab');
    throw new Error('Schwab session expired. Please re-authorize from the menu.');
  }
  if (code < 200 || code >= 300) {
    throw new Error(`Schwab API error ${code} on ${url}: ${body.substring(0, 300)}`);
  }

  return JSON.parse(body);
}
