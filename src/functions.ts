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
*/


/// <reference types="office-runtime" />
/// <reference types="office-js" />

import { create, Client, Config } from '@afintech/sdk/env/browser';
import { Ticker } from 'node_modules/@afintech/sdk/dist/esm/graphql/graphql';


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


export function remakeClient(api_key: string, api_secret: string) {
  config.apiKey = api_key;
  config.apiSecret = api_secret;

  client = create(config);
  console.log("Client recreated with new config."); 
}

/**
 * Initialize the client with user-provided API key and secret.
 * This should run when the user enters their API key/secret.
 * @customfunction
 * @returns The user's email address
 * @helpurl https://excel.architect.co/docs/functions.html#INITIALIZECLIENT
 * @volatile
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
 * Get the bid/ask and last price of a market.
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @returns The bbo prices of the given market
 * @helpurl https://excel.architect.co/docs/functions.html#MARKETBBO
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
 * Get the mid price of a the given market.
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @returns The mid market price of the given market
 * @helpurl https://excel.architect.co/docs/functions.html#MARKETMID
 * @volatile
 */
export async function marketMid(symbol: string, venue: string): Promise<number> {
    let bbo = await marketBBO(symbol, venue);

    let ask = bbo[0][1];
    let bid = bbo[0][0];

    return isNaN(bid) || isNaN(ask) ? NaN : (bid + ask) / 2;
}


/**
 * Get the bid/ask/last price and size of a market
 * @customfunction
 * @param symbol Market symbol, e.g. "ES 20250620 CME Future"
 * @param venue Market venue, e.g. "CME"
 * @returns The ticker information: bid price, bid size, ask price, ask size, last price, last size
 * @helpurl https://excel.architect.co/docs/functions.html#MARKETTICKER
 * @volatile
 */
export async function marketTicker(symbol: string, venue: string): Promise<number[] []> {
  let snapshot: Ticker = await client.ticker([], symbol, venue)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }
  try {
    const bid_px: number = snapshot.bidPrice ? parseFloat(snapshot.bidPrice) : NaN;
    const bid_sz: number = snapshot.bidSize ? parseFloat(snapshot.bidSize) : NaN;
    const ask_px: number = snapshot.askPrice ? parseFloat(snapshot.askPrice) : NaN;
    const ask_sz: number = snapshot.askSize ? parseFloat(snapshot.askSize) : NaN;
    const last_px: number = snapshot.lastPrice ? parseFloat(snapshot.lastPrice) : NaN;
    const last_sz: number = snapshot.lastSize ? parseFloat(snapshot.lastSize) : NaN;
    return [[bid_px, bid_sz, ask_px, ask_sz, last_px, last_sz]]
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse bid/ask prices"
    )
  }
}


/**
 * Get accounts for a given API key/secret.
 * @customfunction
 * @returns List of accounts
 * @helpurl https://excel.architect.co/docs/functions.html#ACCOUNTLIST
 * @volatile
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
 * Get Positions
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @returns The position information
 * @helpurl https://excel.architect.co/docs/functions.html#ACCOUNTPOSITIONS
 * @volatile
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
 * Get Daily PnL
 * @customfunction
 * @param account_name Account name, gotten from accountList function.
 * @returns Account Pnl information: cash excess, equity, position margin, purchasing power, realized pnl, unrealized pnl, total margin, yesterday equity
 * @helpurl https://excel.architect.co/docs/functions.html#ACCOUNTPNL
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
 * @returns Account balances
 * @helpurl https://excel.architect.co/docs/functions.html#ACCOUNTBALANCE
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
 * Search symbols by market name
 * @param market_name Market name, e.g. "ES", "NQ", "RTY"
 * @helpurl https://excel.architect.co/docs/functions.html#SEARCHSYMBOLS
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
