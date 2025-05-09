/*
Valid output types
Primitive Types:
    String: Returns text values.​
    Number: Returns numerical values.​
    Boolean: Returns true or false.​

Arrays:
    Array of Arrays: For multi-dimensional data, you can return a two-dimensional array (e.g., [[1, 2], [3, 4]]), which Excel will display across corresponding cell ranges.​

Specialized Data Types:
    Entity: Represents complex data structures with properties and optional display metadata.​
    FormattedNumber: Allows returning numbers with specific formatting, such as currency or percentages.​

https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-json-autogeneration
*/


/// <reference types="office-runtime" />
/// <reference types="office-js" />

import { create, Client, Config } from '@afintech/sdk/env/browser';
import { AccountPosition, Ticker } from 'node_modules/@afintech/sdk/dist/esm/graphql/graphql';


let config: Config = {
  host: 'https://app.architect.co/',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

let client: Client = (new Proxy({}, {
  get(_obj, _prop) {
    throw new Error('Client is not initialized');
  },
  set(_obj, _prop, _value) {
    throw new Error('Client is not initialized');
  }
}) as Client);


// ──────────────────────────────────────────────────────────────────────────────
//  Shared helper ─ add this in one place and import or copy to both files
// ──────────────────────────────────────────────────────────────────────────────
/** Excel can send a single cell, a 2-D array from an inline constant/range,
 *  or nothing at all.  Convert that to a flat string[] or `undefined`. */
const normalizeFields = (input: any): string[] | undefined => {
  if (input == null || (Array.isArray(input) && input.length === 0)) return undefined;

  // Inline/range arrays arrive as a 2-D matrix (rows × columns).
  if (Array.isArray(input) && Array.isArray(input[0])) {
    return (input as any[][]).flat().map(String).filter(Boolean);
  }

  // Already a flat array  ▸ ["bidPrice", "askPrice"]
  if (Array.isArray(input)) return (input as any[]).map(String);

  // Single cell ▸ "bidPrice"
  if (typeof input === "string") return [input];

  // Anything else: treat as “no fields” ⇒ default later on
  return undefined;
};


export function remakeClient(api_key: string, api_secret: string) {
  config.apiKey = api_key;
  config.apiSecret = api_secret;

  client = create(config);
  console.log("Client recreated with new config."); 
}

/**
 * Initialize the client with user-provided API key and secret.
 * This should run when the user enters their API key/secret.
 * Returns the user's email address.
 * @customfunction
 * @helpurl https://excel.architect.co/functions_help.html#INITIALIZECLIENT
 * @returns The user's email address
 */
export async function initializeClient() : Promise<string> {
  let apiKey: string | null;
  let apiSecret: string | null;
  try {
    const {
      ArchitectApiKey = null,
      ArchitectApiSecret = null,
    } = await OfficeRuntime.storage.getItems(['ArchitectApiKey', 'ArchitectApiSecret']);

    apiKey = ArchitectApiKey;
    apiSecret = ArchitectApiSecret;

    
  } catch (error) {
    console.log("Error accessing storage.");
    apiKey = null;
    apiSecret = null;
  }

  if (!apiKey) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "api_key has not been input"
    )
  }
  if (!apiSecret) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "api_secret has not been input"
    )
  }

  remakeClient(apiKey, apiSecret);

  try {
    let email = await client.userEmail();
    console.log("Client initialized successfully. User email:", email);
    return email;
  } catch (error) {
    console.error("Client failed to initialize. Please check your API key and secret: ", error);
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Client failed to initialize. Please check your API key and secret."
    )
  }
 }

/**
 * Returns the bid/ask prices of the given market.
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @helpurl https://excel.architect.co/functions_help.html#MARKETBBO
 * @volatile
 */
export async function marketBBO(symbol: string, venue: string): Promise<number[] []> {
  let snapshot: Ticker = await client.ticker([], symbol, venue)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }
  try {
    const bid: number = snapshot.bidPrice ? parseFloat(snapshot.bidPrice) : NaN;
    const ask: number = snapshot.askPrice ? parseFloat(snapshot.askPrice) : NaN;
    return [[bid, ask]]
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse bid/ask prices"
    )
  }
}


/**
 * Stream the bid/ask/last prices of a market in real-time. Or other fields if specified.
 * 
 * For the field params, the possible values are: askPrice, askSize, bidPrice, bidSize, dividend, dividendYield, epsAdj, high24h, lastPrice, lastSettlementPrice, lastSize, low24h, marketCap, open24h, openInterest, priceToEarnings, sessionHigh, sessionLow, sessionOpen, sessionVolume, sharesOutstandingWeightedAdj, symbol, timestamp, volume24h, volume30d
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @param [fields] List of fields to stream, default value is ["bidPrice","askPrice","lastPrice"]
 * @param invocation Streaming invocation object
 * @streaming
 */
export function streamMarketTicker(symbol: string, venue: string, fields: string[][] | undefined, invocation: CustomFunctions.StreamingInvocation<number[][]>): void {
  const defaultFields = ["bidPrice", "askPrice", "lastPrice"];
  const chosenFields = normalizeFields(fields) ?? defaultFields;

  // Map snapshot keys → numeric values we can push into Excel
  const pluck = (snap: Ticker, key: string): number =>
    snap[key as keyof Ticker] !== undefined
      ? Number(snap[key as keyof Ticker])
      : NaN;

  const intervalId = setInterval(async () => {
    try {
      const snap = await client.ticker([], symbol, venue);

      if (!snap) {
        invocation.setResult([Array(chosenFields.length).fill(NaN)]);
        return;
      }

      const row = chosenFields.map((f) => pluck(snap, f));
      invocation.setResult([row]);
    } catch (err) {
      console.error("streamMarketTicker:", err);
      invocation.setResult([Array(chosenFields.length).fill(NaN)]);
    }
  }, 1_000);                                // Update every second

  invocation.onCanceled = () => clearInterval(intervalId);
}

/**
 * Returns the mid price of a the given market.
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @helpurl https://excel.architect.co/functions_help.html#MARKETMID
 * @volatile
 */
export async function marketMid(symbol: string, venue: string): Promise<number> {
    let bbo = await marketBBO(symbol, venue);

    let ask = bbo[0][1];
    let bid = bbo[0][0];

    return isNaN(bid) || isNaN(ask) ? NaN : (bid + ask) / 2;
}


/**
 * Get the bid/ask/last price and size of a market.
 * Returns: bid price, bid size, ask price, ask size, last price, last size.
 * For the field params, the possible values are: askPrice, askSize, bidPrice, bidSize, dividend, dividendYield, epsAdj, high24h, lastPrice, lastSettlementPrice, lastSize, low24h, marketCap, open24h, openInterest, priceToEarnings, sessionHigh, sessionLow, sessionOpen, sessionVolume, sharesOutstandingWeightedAdj, symbol, timestamp, volume24h, volume30d
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @param [fields] List of fields to stream, default value is ["bidPrice","bidSize","askPrice","askSize","lastPrice","lastSize"]
 * @helpurl https://excel.architect.co/functions_help.html#MARKETTICKER
 * @volatile
 */
export async function marketTicker(symbol: string, venue: string, fields?: string[][] | undefined): Promise<number[] []> {
  const defaultFields = [
    "bidPrice",
    "bidSize",
    "askPrice",
    "askSize",
    "lastPrice",
    "lastSize",
  ];
  const chosenFields = normalizeFields(fields) ?? defaultFields;

  // Helper: pluck a numeric value (or NaN) from the snapshot
  const pluck = (snap: Ticker, key: string): number =>
    snap[key as keyof Ticker] !== undefined
      ? Number(snap[key as keyof Ticker])
      : NaN;

  try {
    const snap = await client.ticker([], symbol, venue);

    if (!snap) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "Received bad data from the server, please try again.",
      );
    }

    const row = chosenFields.map((f) => pluck(snap, f));
    return [row]; // one-row 2-D array → a single row in Excel
  } catch (err) {
    if (err instanceof CustomFunctions.Error) throw err; // preserve Excel error
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to fetch or parse market data.",
    );
  }
}


/**
 * Returns a list accounts for a given API key/secret.
 * @customfunction
 * @helpurl https://excel.architect.co/functions_help.html#ACCOUNTLIST
 */
export async function accountList(): Promise<string[][]> {
  const snapshot = await client.accounts([]);

  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    );
  }

  const rows: string [][] = [];

  rows.push(["Account Name", "Trader", "Trade Permission", "View Permission"]);

  snapshot.forEach(account => {
    rows.push([
      account.account.name,
      account.trader,
      account.permissions.trade.toString(),
      account.permissions.view.toString()
    ]);
  });

  try {
    return rows;
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse account data"
    );
  }
}

/**
 * Get positions for a given account.
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @helpurl https://excel.architect.co/functions_help.html#ACCOUNTPOSITIONS
 */
export async function accountPositions(account_name: string): Promise<string[][]> {
  let snapshot = await client.accountSummary([], account_name)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }

  try {
    const headers = [
      "Symbol",
      "Quantity",
      "Cost Basis",
      // "Break Even Price",
      // "Liquidation Price",
      // "Trade Time"
    ];
    const rows: string[][] = [[snapshot.timestamp, ...Array(headers.length - 1).fill("")]];

    rows.push(headers);

    snapshot.positions.forEach(position => {
      rows.push([
        position.symbol,
        position.quantity,
        position.costBasis ?? "NaN",
        // position.breakEvenPrice ?? "NaN",
        // position.liquidationPrice ?? "NaN",
        // position.tradeTime ?? ""
      ]);
    });

    return rows;
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse account summary snapshot"
    )
  }
}

/**
 * Stream the positions for a given account in real-time, ensuring the same structure as accountPositions.
 * Any symbols not in the account will be returned with zero values.
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @param symbols List of market symbols for the positions, e.g. ["ES 20250620 CME Future", "NQ 20250620 CME Future"].
 * @param show_all If true, show all positions in the account.
 * @param invocation Streaming invocation object
 * @helpurl https://excel.architect.co/functions_help.html#STREAMACCOUNTPOSITIONVALUES
 * @streaming
 */
export function streamAccountPositionValues(
  account_name: string,
  show_all: boolean,
  symbols: string[],
  invocation: CustomFunctions.StreamingInvocation<string[][]>
): void {
  // Hoist constants to avoid re‑allocating them each tick
  const headers = ["Symbol", "Quantity", "Cost Basis"];
  const headerRow = headers; 
  const headerCount = headers.length;

  const baseSymbols = show_all ? null : symbols;

  try {
    const intervalId = setInterval(async () => {
      try {
        const snapshot = await client.accountSummary([], account_name);
        if (!snapshot) {
          invocation.setResult([["Error: No data available"]]);
          return;
        }

        const posMap = new Map<string, AccountPosition>();
        for (const pos of snapshot.positions) {
          posMap.set(pos.symbol, pos);
        }

        let symbolList: string[];
        if (show_all) {
          // union of account symbols + requested symbols
          symbolList = Array.from(
            new Set([...posMap.keys(), ...symbols])
          );
        } else {
          // only the requested symbols
          symbolList = baseSymbols!;
        }

        const rows: string[][] = [];

        const tsRow = new Array<string>(headerCount);
        tsRow[0] = snapshot.timestamp;
        for (let i = 1; i < headerCount; i++) tsRow[i] = "";
        rows.push(tsRow);

        rows.push(headerRow);

        for (const sym of symbolList) {
          const p = posMap.get(sym);
          if (p) {
            rows.push([
              p.symbol,
              p.quantity.toString(),
              p.costBasis != null ? p.costBasis.toString() : "NaN",
            ]);
          } else {
            rows.push([sym, "0", "0"]);
          }
        }

        invocation.setResult(rows);
      } catch (err) {
        console.error("Error fetching account position values:", err);
        invocation.setResult([["Error fetching data"]]);
      }
    }, 1000);

    invocation.onCanceled = () => clearInterval(intervalId);
  } catch (err) {
    console.error("Error initializing streaming function:", err);
    invocation.setResult([["Error initializing function"]]);
  }
}



/**
 * Returns account Pnl information: cash excess, equity, position margin, purchasing power, realized pnl, unrealized pnl, total margin, yesterday equity
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @helpurl https://excel.architect.co/functions_help.html#ACCOUNTPNL
 * @volatile
 */
export async function accountPnl(account_name: string): Promise<number[] []> {
  let snapshot = await client.accountSummary([], account_name)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }

  let cashExcess = snapshot.cashExcess ? parseFloat(snapshot.cashExcess) : NaN;
  let equity = snapshot.equity ? parseFloat(snapshot.equity) : NaN;
  let positionMargin = snapshot.positionMargin ? parseFloat(snapshot.positionMargin) : NaN;
  let purchasingPower = snapshot.purchasingPower ? parseFloat(snapshot.purchasingPower) : NaN;
  let realizedPnl = snapshot.realizedPnl ? parseFloat(snapshot.realizedPnl) : NaN;
  let unrealizedPnl = snapshot.unrealizedPnl ? parseFloat(snapshot.unrealizedPnl) : NaN;
  let totalMargin = snapshot.totalMargin ? parseFloat(snapshot.totalMargin) : NaN;
  let yesterdayEquity = snapshot.yesterdayEquity ? parseFloat(snapshot.yesterdayEquity) : NaN;
  try {
    return [[cashExcess, equity, positionMargin, purchasingPower, realizedPnl, unrealizedPnl, totalMargin, yesterdayEquity]]
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse account summary snapshot"
    )
  }
}

/**
 * Get Account Balance.
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @helpurl https://excel.architect.co/functions_help.html#ACCOUNTBALANCE
 * @volatile
 */
export async function accountBalance(account_name: string): Promise<number> {
  let snapshot = await client.accountSummary([], account_name)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }

  let usd_balance: number = 0;

  snapshot.balances.forEach(balance => {
    if (balance.product == "USD") {
      usd_balance = parseFloat(balance.balance)
    }
  }
  )

  try {
    return usd_balance
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse account summary snapshot"
    )
  }
}

/**
 * Get the list of symbols available for a given market.
 * @customfunction
 * @param base_name Base name, e.g. "ES", "NQ", "RTY"
 * @param expiration Expiration date, accepts several formats. e.g. "20250620", or "M5"/"M25" for June 2025, "Jun25"
 * @helpurl https://excel.architect.co/functions_help.html#MARKETLIST
 */
export async function deriveFuturesSymbol(base_name: string, expiration: string): Promise<string> {

  const expirationString: string = getExpirationString(expiration);
  const searchString: string = `${base_name} ${expirationString}`;
  const symbols = await client.searchSymbols({ searchString });

  for (const symbol of symbols) {
    if (symbol.includes(base_name) && symbol.includes(expirationString)) {
      return symbol;
    }
  }

  throw new CustomFunctions.Error(
    CustomFunctions.ErrorCode.notAvailable,
    "No symbols found for the given base name and expiration date."
  );
}


function getExpirationString(expiration: string): string {
  const monthMap: { [key: string]: string } = {
    Jan: "01",
    Feb: "02",
    Mar: "03",
    Apr: "04",
    May: "05",
    Jun: "06",
    Jul: "07",
    Aug: "08",
    Sep: "09",
    Oct: "10",
    Nov: "11",
    Dec: "12",
    F: "01",
    G: "02",
    H: "03",
    J: "04",
    K: "05",
    M: "06",
    N: "07",
    Q: "08",
    U: "09",
    V: "10",
    X: "11",
    Z: "12"
  };
  const match = expiration.match(/([A-Za-z]+|M)(\d{1,2})/);
  if (!match) {
    throw new Error(`Invalid expiration format: ${expiration}`);
  }
  let month = match[1]; // Month part (e.g., "Jun" or "M")
  let year = match[2];  // Year part (e.g., "25" or "5")

  // Handle month abbreviations (e.g., "Jun", "June")
  if (month.length > 1) {
    month = month.substring(0, 3); // Take the first 3 letters
  }

  // Convert month to numeric format
  const numericMonth = monthMap[month];
  if (!numericMonth) {
    throw new Error(`Invalid month abbreviation: ${month}`);
  }

  // Convert year to 4-digit format
  const fullYear = year.length === 1 ? `202${year}` : `20${year}`;

  // Return the formatted expiration date (YYYYMM)
  return `${fullYear}${numericMonth}`;
}

/**
 * Search symbols by market name
 * @param market_name Market name, e.g. "ES", "NQ", "RTY"
 * @helpurl https://excel.architect.co/functions_help.html#SEARCHSYMBOLS
 * @customfunction 
 */
export async function searchSymbols(market_name: string): Promise<string [] []> {
  const symbols = await client.searchSymbols({ searchString: market_name});

  const result = symbols.map(symbol => [symbol]);
  return result;
}

Office.onReady(async (info) => {
  try {
    await initializeClient()
    console.log('Client initialized using saved API key/secret');
  } catch (error) {
    console.log(error)
  }
});
