/**
 * SchwabAPI.gs — Thin wrappers around the Schwab Trader & Market Data APIs.
 *
 * Trader API base:    https://api.schwabapi.com/trader/v1
 * Market Data base:   https://api.schwabapi.com/marketdata/v1
 */

const TRADER_API_BASE = 'https://api.schwabapi.com/trader/v1';
const MARKET_API_BASE = 'https://api.schwabapi.com/marketdata/v1';

// ─────────────────────────────────────────────────────────────────
// Account endpoints
// ─────────────────────────────────────────────────────────────────

/**
 * Returns all linked accounts with their encrypted hashes.
 * Response: [{ accountNumber, hashValue }, …]
 */
function getAccountNumbers() {
  return apiGet_(`${TRADER_API_BASE}/accounts/accountNumbers`);
}

/**
 * Returns full account details including current balances and positions.
 * @param {string} encryptedAccountNumber  hashValue from getAccountNumbers()
 */
function getAccountDetails(encryptedAccountNumber) {
  return apiGet_(`${TRADER_API_BASE}/accounts/${encryptedAccountNumber}`, {
    fields: 'positions',
  });
}

/**
 * Returns the current Net Liquidation Value for the account identified
 * by its suffix (last 3 digits), e.g. '418'.
 * @param {string} suffix
 * @returns {number}
 */
function fetchNetLiqForSuffix_(suffix) {
  const hash    = getHashForSuffix_(suffix);
  const details = getAccountDetails(hash);
  return extractNetLiquidityFromAccount_(details);
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
  // Note: 'types' param does not accept comma-separated values.
  // Omitting it returns all transaction types, which we filter in NetLiquidity.gs.
  const result = apiGet_(
    `${TRADER_API_BASE}/accounts/${encryptedAccountNumber}/transactions`,
    {
      startDate,
      endDate,
    }
  );
  // API returns a direct array; guard against unexpected wrapper objects
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.transactions)) return result.transactions;
  return [];
}

/**
 * Debug helper — run from the Apps Script editor to test the transactions
 * endpoint directly. Check View → Logs after running.
 */
function debugTransactions() {
  const props = PropertiesService.getScriptProperties();
  var hash = props.getProperty('SCHWAB_ACCT_HASH_418');
  if (!hash) {
    hash = getHashForSuffix_('418');
  }

  var start = new Date();
  start.setDate(start.getDate() - 7);  // last 7 days

  console.log('Account hash (truncated): ' + hash.substring(0, 8) + '…');
  console.log('Querying transactions from ' + start.toISOString() + ' to now');

  try {
    var raw = apiGet_(
      TRADER_API_BASE + '/accounts/' + hash + '/transactions',
      {
        startDate: start.toISOString(),
        endDate:   new Date().toISOString(),
        types:     'TRADE,RECEIVE_AND_DELIVER,DIVIDEND_OR_INTEREST,ELECTRONIC_FUND,OTHER',
      }
    );
    console.log('Raw response type: ' + typeof raw);
    console.log('Is array: ' + Array.isArray(raw));
    console.log('Response (first 500 chars): ' + JSON.stringify(raw).substring(0, 500));
  } catch (e) {
    console.error('Transaction API error: ' + e.message);
  }
}

// ─────────────────────────────────────────────────────────────────
// Market Data endpoints
// ─────────────────────────────────────────────────────────────────

/**
 * Fetches daily OHLCV candles for a symbol between two dates.
 * @param {string} symbol       e.g. 'SPY'
 * @param {string} startDate    'YYYY-MM-DD'
 * @param {string} endDate      'YYYY-MM-DD'
 * @returns {{ candles: [{datetime, open, high, low, close, volume}], symbol, empty }}
 */
function getPriceHistory(symbol, startDate, endDate) {
  const startMs = new Date(startDate).getTime();
  const endMs   = new Date(endDate).getTime();

  return apiGet_(`${MARKET_API_BASE}/pricehistory`, {
    symbol,
    periodType:            'year',
    frequencyType:         'daily',
    frequency:             1,
    startDate:             startMs,
    endDate:               endMs,
    needExtendedHoursData: false,
  });
}

// ─────────────────────────────────────────────────────────────────
// Internal helpers
// ─────────────────────────────────────────────────────────────────

/**
 * Looks up and caches the encrypted account hash for a given suffix.
 * Hashes are stored in Script Properties as SCHWAB_ACCT_HASH_418 etc.
 * @param {string} suffix  e.g. '418'
 * @returns {string} hashValue
 */
function getHashForSuffix_(suffix) {
  const acctNum = ACCOUNT_CONFIGS[suffix];
  if (!acctNum) {
    throw new Error('Unknown account suffix "' + suffix + '". Add it to ACCOUNT_CONFIGS in Code.gs.');
  }

  const propKey = 'SCHWAB_ACCT_HASH_' + suffix;
  const props   = PropertiesService.getScriptProperties();
  let hash      = props.getProperty(propKey);

  if (!hash) {
    const accounts = getAccountNumbers();
    const matched  = accounts.find(function(a) { return a.accountNumber === acctNum; });
    if (!matched) {
      throw new Error('Account ' + acctNum + ' not found among linked accounts.');
    }
    hash = matched.hashValue;
    props.setProperty(propKey, hash);
  }

  return hash;
}

/**
 * Extracts the Net Liquidation Value from an account details response.
 * Tries every known Schwab field name in order of preference.
 * @param {Object} accountData
 * @returns {number}
 */
function extractNetLiquidityFromAccount_(accountData) {
  const sec = accountData.securitiesAccount || accountData;
  const bal = sec.currentBalances || sec.projectedBalances || {};

  const candidates = [
    bal.liquidationValue,
    bal.equity,
    bal.totalAccountValue,
    bal.accountValue,
    bal.netLiquidation,
  ];

  for (var i = 0; i < candidates.length; i++) {
    var v = candidates[i];
    if (typeof v === 'number' && v > 0) return v;
  }
  return 0;
}

/**
 * Debug helper — run once from the Apps Script editor to inspect raw
 * account API responses. Check View → Logs afterwards.
 */
function debugAccountResponse() {
  const accounts = getAccountNumbers();
  console.log('Linked accounts: ' + JSON.stringify(accounts.map(function(a) { return a.accountNumber; })));
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
    method:             'GET',
    headers:            { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(fullUrl, options);
  const code     = response.getResponseCode();
  const body     = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Schwab API error ${code} on ${url}: ${body.substring(0, 300)}`);
  }

  return JSON.parse(body);
}
