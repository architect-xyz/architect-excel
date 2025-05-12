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
import { HistoricalFillsResponse, Ticker } from 'node_modules/@afintech/sdk/dist/esm/graphql/graphql';


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
 * For the field params, the possible values are: askPrice, askSize, bidPrice, bidSize, lastPrice, lastSize, lastSettlementPrice, timestamp, dividend, dividendYield, epsAdj, high24h, low24h, marketCap, open24h, openInterest, priceToEarnings, sessionHigh, sessionLow, sessionOpen, sessionVolume, sharesOutstandingWeightedAdj, symbol, timestamp, volume24h, volume30d
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @param [fields] List of fields to stream, default value is ["bidPrice","askPrice","lastPrice"]
 * @param invocation Streaming invocation object
 * @helpurl https://excel.architect.co/functions_help.html#STREAMMARKETTICKER
 * @streaming
 */
export function streamMarketTicker(
  symbol: string,
  venue: string,
  fields: string[][],
  invocation: CustomFunctions.StreamingInvocation<number[][]>,
): void {
  const defaultFields = ["bidPrice", "askPrice", "lastPrice"] as const;
  const chosenFields = (normalizeFields(fields) ?? [...defaultFields]) as (keyof Ticker)[];

  // Helper: pull numeric values (NaN if absent)
  const pluck = (snap: Pick<Ticker, keyof Ticker>, key: keyof Ticker): number =>
    snap[key] !== undefined ? Number(snap[key]) : NaN;

  const intervalId = setInterval(async () => {
    try {
      // 🔑  Only request the chosen fields
      const snap = await client.ticker(chosenFields, symbol, venue);

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
  }, 1_000); // update every second

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
 * For the field params, the possible values are: askPrice, askSize, bidPrice, bidSize, lastPrice, lastSize, lastSettlementPrice, timestamp, dividend, dividendYield, epsAdj, high24h, low24h, marketCap, open24h, openInterest, priceToEarnings, sessionHigh, sessionLow, sessionOpen, sessionVolume, sharesOutstandingWeightedAdj, symbol, timestamp, volume24h, volume30d
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @param [fields] List of fields to stream, default value is ["bidPrice","bidSize","askPrice","askSize","lastPrice","lastSize"]
 * @helpurl https://excel.architect.co/functions_help.html#MARKETTICKER
 * @volatile
 */
export async function marketTicker(symbol: string, venue: string, fields?: string[][]): Promise<number[] []> {
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

// /**
//  * Get positions for a given account.
//  * @customfunction
//  * @param account_name Account name, gotten from accountList function.
//  * @helpurl https://excel.architect.co/functions_help.html#ACCOUNTPOSITIONS
//  */
// export async function accountPositions(account_name: string): Promise<string[][]> {
//   let snapshot = await client.accountSummary([], account_name)
//   if (!snapshot) {
//     throw new CustomFunctions.Error(
//       CustomFunctions.ErrorCode.notAvailable,
//       "Received bad data from the server, please try again."
//     )
//   }
// 
//   try {
//     const headers = [
//       "Symbol",
//       "Quantity",
//       "Cost Basis",
//       // "Break Even Price",
//       // "Liquidation Price",
//       // "Trade Time"
//     ];
//     const rows: string[][] = [[snapshot.timestamp, ...Array(headers.length - 1).fill("")]];
// 
//     rows.push(headers);
// 
//     snapshot.positions.forEach(position => {
//       rows.push([
//         position.symbol,
//         position.quantity,
//         position.costBasis ?? "NaN",
//         // position.breakEvenPrice ?? "NaN",
//         // position.liquidationPrice ?? "NaN",
//         // position.tradeTime ?? ""
//       ]);
//     });
// 
//     return rows;
//   } catch (error) {
//     throw new CustomFunctions.Error(
//       CustomFunctions.ErrorCode.invalidValue,
//       "Failed to parse account summary snapshot"
//     )
//   }
// }

/**
 * Stream the positions for a given account in real-time, ensuring the same structure as accountPositions.
 * Any symbols not in the account will be returned with zero values.
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @param [symbols] List of market symbols for the positions, e.g. ["ES 20250620 CME Future", "NQ 20250620 CME Future"].
 * @param [show_all] If true, show all positions in the account.
 * @param invocation Streaming invocation object
 * @helpurl https://excel.architect.co/functions_help.html#STREAMACCOUNTPOSITIONVALUES
 * @streaming
 */
export function streamAccountPositionValues(
  account_name: string,
  symbols: string[][],
  show_all: boolean,
  invocation: CustomFunctions.StreamingInvocation<string[][]>
): void {
  // Hoist constants to avoid re‑allocating them each tick

  const headerRow      = ["Symbol", "NetPosition", "Cost Basis"];
  const chosenSymbols = normalizeFields(symbols) ?? [];
  const includeAll     = show_all ?? false;

  const intervalId = setInterval(async () => {
    try {
      const snapshot = await client.accountSummary([], account_name);
      if (!snapshot) {
        invocation.setResult([["Error: No data available"]]);
        return;
      }

      // Map existing positions for quick look-ups
      const posMap = new Map(snapshot.positions.map(p => [p.symbol, p]));

      // Decide which symbols to show this tick
      let symbolList: string[];
      if (includeAll) {
        symbolList = Array.from(
          new Set([...posMap.keys(), ...chosenSymbols]),
        );
      } else if (chosenSymbols.length) {
        symbolList = chosenSymbols;
      } else {
        // No list supplied → just dump the account’s current symbols
        symbolList = Array.from(posMap.keys());
      }

      // Build the 2-D result (timestamp row ▸ header ▸ data rows)
      const rows: string[][] = [
        [snapshot.timestamp, "", ""],          // timestamp row
        headerRow,
        ...symbolList.map(sym => {
          const p = posMap.get(sym);
          return p
            ? [p.symbol, p.quantity.toString(),
               p.costBasis != null ? p.costBasis.toString() : "NaN"]
            : [sym, "0", "0"];
        }),
      ];

      invocation.setResult(rows);

    } catch (err) {
      console.error("streamAccountPositionValues:", err);
      invocation.setResult([["Error fetching data"]]);
    }
  }, 1_000);                                        // refresh every second

  invocation.onCanceled = () => clearInterval(intervalId);
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

/**
 * Gets fills analysis for a given account and string of symbols for the current trading day.
 * @param accountName Account name, gotten from accountList function.
 * @param symbols     Market symbols, e.g. ["ES 20250620 CME Future", "NQ 20250620 CME Future"]
 * @customfunction
 */
export async function fillsAnalysis(
  accountName: string,
  symbols: string[][]
): Promise<string[][]> {
  /* ── 0. validate input ───────────────────────────────────────────── */
  if (!symbols?.length) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "No symbols provided."
    );
  }

  /* ── 1. normalise symbols & fetch fills ──────────────────────────── */
  const requested = new Set(normalizeFields(symbols) ?? []);        // unique, O(1) lookup
  const wantAll   = requested.size === 0;

  const fromInclusive = getStartOfTradingDate();                    // 17:00 ET previous calendar-day
  const snapshot: HistoricalFillsResponse = await client.historicalFills(
    [],
    { account: accountName, fromInclusive: fromInclusive.toISOString() }
  );

  console.log("fromInclusive:", fromInclusive);
  console.log("Fills snapshot:", snapshot);
  const fills = snapshot?.fills ?? [];

  /* ── 2. aggregate in one pass ────────────────────────────────────── */
  interface Agg {
    lastTs: string;
    tradeCount: number;
    buyQty: number;   buyPQ: number;        // Σ(qty·price) buys
    sellQty: number;  sellPQ: number;       // Σ(qty·price) sells
  }
  const agg: Record<string, Agg> = Object.create(null);

  for (let i = 0; i < fills.length; ++i) {
    const f = fills[i];
    if (!wantAll && !requested.has(f.symbol)) continue;

    let a = agg[f.symbol];
    if (!a) {
      a = agg[f.symbol] = {
        lastTs: "",
        tradeCount: 0,
        buyQty: 0,  buyPQ: 0,
        sellQty: 0, sellPQ: 0,
      };
    }

    ++a.tradeCount;

    const qty   = +f.quantity;                           // fast string→number
    if (qty === 0) continue;                             // ignore odd zero-qty fills
    const price = +f.price;

    if (f.dir === "buy") {
      a.buyQty += qty;
      a.buyPQ  += qty * price;
    } else {
      a.sellQty += qty;
      a.sellPQ  += qty * price;
    }

    const ts = (f.tradeTime ?? f.recvTime) as string | undefined;
    if (ts && ts > a.lastTs) a.lastTs = ts;
  }

  /* ── 3. shape output for Excel ───────────────────────────────────── */
  const header = [
    "Timestamp",
    "Symbols",
    "TradePosition",
    "TradeCount",
    "AverageBuyPrice",
    "BuyQuantity",
    "AverageSellPrice",
    "SellQuantity",
  ] as const;

  const rows: string[][] = [header.slice()];

  // 3a. rows for symbols that had at least one fill
  for (const sym in agg) {
    const a = agg[sym];
    rows.push([
      a.lastTs,
      sym,
      String(a.buyQty - a.sellQty),
      String(a.tradeCount),
      a.buyQty  ? (a.buyPQ  / a.buyQty ).toFixed(2) : "",
      String(a.buyQty),
      a.sellQty ? (a.sellPQ / a.sellQty).toFixed(2) : "",
      String(a.sellQty),
    ]);
  }

  // 3b. rows for requested symbols with *no* fills
  if (!wantAll) {
    for (const sym of requested) {
      if (agg[sym]) continue;                            // already included
      rows.push([
        "",  // Timestamp
        sym, // Symbols
        "", "", "", "", "", "",
      ]);
    }
  }

  return rows;
}


/**
 * Return the timestamp of the *current* trading-day open
 * (5 p.m. Eastern Time — i.e. America/New_York).
 *
 * • If it’s still **before** 17:00 ET, the trading day started
 *   at 17:00 ET **yesterday**.  
 * • Otherwise it started at 17:00 ET **today**.
 *
 */
function getStartOfTradingDate(): Date {
  const now = new Date();

  // How far New York time is ahead (+) or behind (–) the local clock *right now*
  const nyOffsetMs =
    new Date(now.toLocaleString("en-US", { timeZone: "America/New_York" })).getTime() -
    now.getTime();

  // `now` expressed on the New York wall-clock
  const nowNY = new Date(now.getTime() + nyOffsetMs);

  // 17:00 on that New York calendar date
  const startNY = new Date(nowNY);
  startNY.setHours(17, 0, 0, 0);

  // If we haven’t reached 17:00 in New York yet, the trading day started “yesterday”
  if (nowNY < startNY) {
    startNY.setDate(startNY.getDate() - 2);
  }

  // Convert the New York wall-clock time back to the correct absolute instant
  return new Date(startNY.getTime() - nyOffsetMs);
}


Office.onReady(async (info) => {
  try {
    await initializeClient()
    console.log('Client initialized using saved API key/secret');
  } catch (error) {
    console.log(error)
  }
});
