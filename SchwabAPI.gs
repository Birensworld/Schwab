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
 *
 * If SCHWAB_ACCOUNT_HASH is already stored by your existing app, that
 * value is used directly. Otherwise we look up all linked accounts and
 * pick the one whose accountNumber matches SCHWAB_ACCOUNT_NUMBER (if set),
 * falling back to the account with the largest liquidation value so we
 * don't accidentally pick a small cash/IRA account.
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

    if (matched) {
      hash = matched.hashValue;
    } else if (accounts.length === 1) {
      hash = accounts[0].hashValue;
    } else {
      // Multiple accounts — pick the one with the highest liquidation value
      // to avoid landing on a small cash or IRA account by accident.
      hash = pickLargestAccount_(accounts);
    }

    props.setProperty('SCHWAB_ACCOUNT_HASH', hash);
  }

  return getAccountDetails(hash);
}

/**
 * Given the array from /accounts/accountNumbers, fetches each account's
 * balance and returns the hash of whichever has the highest liquidation value.
 * @param {Array<{accountNumber, hashValue}>} accounts
 * @returns {string} hashValue
 */
function pickLargestAccount_(accounts) {
  let bestHash  = accounts[0].hashValue;
  let bestValue = -Infinity;

  accounts.forEach(function(acct) {
    try {
      const details = getAccountDetails(acct.hashValue);
      const value   = extractNetLiquidityFromAccount_(details);
      if (value > bestValue) {
        bestValue = value;
        bestHash  = acct.hashValue;
      }
    } catch (e) {
      console.warn('Could not fetch details for account ' + acct.accountNumber + ': ' + e.message);
    }
  });

  return bestHash;
}

/**
 * Extracts the Net Liquidation Value from an account details response.
 * Schwab uses different field names for margin vs cash vs IRA accounts.
 * @param {Object} accountData
 * @returns {number}
 */
function extractNetLiquidityFromAccount_(accountData) {
  const sec = accountData.securitiesAccount || accountData;
  const bal = sec.currentBalances || sec.projectedBalances || {};

  // Try every known field name Schwab uses across account types
  const candidates = [
    bal.liquidationValue,
    bal.netLiquidation,
    bal.totalAccountValue,
    bal.accountValue,
    bal.equity,
  ];

  for (const v of candidates) {
    if (typeof v === 'number' && v > 0) return v;
  }
  return 0;
}

/**
 * Debug helper — run this once from the Apps Script editor (not the menu)
 * to inspect the raw account API response and confirm the correct field.
 * Check View → Logs after running.
 */
function debugAccountResponse() {
  const accounts = getAccountNumbers();
  console.log('Linked accounts: ' + JSON.stringify(accounts.map(a => a.accountNumber)));

  accounts.forEach(function(acct) {
    const details = getAccountDetails(acct.hashValue);
    const sec = details.securitiesAccount || details;
    console.log('\n=== Account ' + acct.accountNumber + ' ===');
    console.log('currentBalances: ' + JSON.stringify(sec.currentBalances));
  });
}

/**
 * Generic authenticated GET request.
 * @param {string} url
 * @param {Object} [params]  Query-string key/value pairs
 * @returns {Object} Parsed JSON response
 */
function apiGet_(url, params) {
  const token = getAccessToken();

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
