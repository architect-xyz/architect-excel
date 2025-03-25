// Excel Plugin Example: Query Backend API Using TypeScript


/*
important functions:
close
open
day high
day low

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

import { create, Client } from '@afintech/sdk/env/browser';
import { getStorageItem, config } from './client';
import { Ticker } from 'node_modules/@afintech/sdk/dist/esm/graphql/graphql';



let client: Client = (new Proxy({}, {
  get(_obj, _prop) {
    throw new Error('Client is not initialized');
  },
  set(_obj, _prop, _value) {
    throw new Error('Client is not initialized');
  }
}) as Client);


/**
 * Initialize the client with user-provided API key and secret.
 * This should run when the user enters their API key/secret.
 * @customfunction
 * @returns The user's email address
 * @volatile
 */
export async function initializeClient() : Promise<string> {
  const apiKey = await getStorageItem('ArchitectApiKey');
  const apiSecret = await getStorageItem('ArchitectApiSecret');

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

  config.apiKey = apiKey;
  config.apiSecret = apiSecret;
  client = create(config);

  try {
    return await client.userEmail();
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to initialize client"
    )
  }
 }

/**
 * Get the bid/ask and last price of a market
 * @customfunction
 * @param symbol Market symbol
 * @param venue Market venue
 * @returns The bbo prices of the given market
 * @volatile
 */
export async function getMarketBBO(symbol: string, venue: string): Promise<number[] []> {
  // let snapshot: Ticker = await client.ticker(["symbol", "bidPrice", "askPrice", "lastPrice"], symbol, venue)
  let snapshot: Ticker = await client.ticker([], symbol, venue)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }
  try {
    console.log(snapshot)
    console.log(snapshot.bidPrice, snapshot.askPrice, snapshot.lastPrice)
    const bid: number = snapshot.bidPrice ? parseFloat(snapshot.bidPrice) : NaN;
    const ask: number = snapshot.askPrice ? parseFloat(snapshot.askPrice) : NaN;
    const last: number = snapshot.lastPrice ? parseFloat(snapshot.lastPrice) : NaN;
    return [[bid, ask, last]]
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse bid/ask prices"
    )
  }
}



/**
 * Fetch market snapshot and populate Excel worksheet
 * @customfunction
 * @param symbol Market symbol
 * @returns The mid market price of the given market
 * @volatile
 */
export async function getMarketMid(symbol: string, venue: string): Promise<number> {
    let bbo = await getMarketBBO(symbol, venue);

    let ask = bbo[0][1];
    let bid = bbo[0][0];

    return isNaN(bid) || isNaN(ask) ? NaN : (bid + ask) / 2;
}




/**
 * Get the bid/ask/last price and size of a market
 * @customfunction
 * @param symbol Market symbol
 * @param venue Market venue
 * @returns The ticker information: bid price, bid size, ask price, ask size, last price, last size
 * @volatile
 */
export async function getTicker(symbol: string, venue: string): Promise<number[] []> {
  // let snapshot: Ticker = await client.ticker(["symbol", "bidPrice", "bidSize", "askPrice", "askSize", "lastPrice", "lastSize"], symbol, venue)
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
 * Get accounts
 * @customfunction
 * @param [header] add the header
 * @returns List of accounts
 * @volatile
 */
export async function getAccounts(header?: boolean): Promise<string[][]> {
  const snapshot = await client.accounts([]);

  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    );
  }

  const rows: string [][] = [];

  if (header) {
    rows.push(["Account Name", "Trader", "Trade Permission", "View Permission"]);
  }

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
 * @param account_name Account name
 * @returns The position information
 * @volatile
 */
export async function getPositions(account_name: string): Promise<string [] []> {
  let snapshot = await client.accountSummary([], account_name)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }

  let timestamp = new Date(snapshot.timestamp).toLocaleString()

  let breakEvenPrice: string[] = [];
  let costBasis: string[] = [];
  let liquidationPrice: string[] = [];
  let symbol: string[] = [];
  let qty: string[] = [];
  let tradeTime: string[] = [];

  snapshot.positions.forEach (position => {
    breakEvenPrice.push(position.breakEvenPrice ?? "NaN")
    costBasis.push(position.costBasis ?? "NaN")
    liquidationPrice.push(position.liquidationPrice ?? "NaN")
    symbol.push(position.symbol)
    qty.push(position.quantity)
    tradeTime.push(position.tradeTime ? new Date(position.tradeTime).toLocaleString() : "")
  })

  try {
    return [[timestamp], breakEvenPrice, costBasis, liquidationPrice, symbol, qty, tradeTime]
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
 * @param account_name Account name
 * @returns account pnl
 * @volatile
 */
export async function getPnL(account_name: string): Promise<number[] []> {
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
 * Get Account Balance
 * @customfunction
 * @param account_name Account name
 * @returns Account balances
 * @volatile
 */
export async function getAccountBalance(account_name: string): Promise<string [] []> {
  let snapshot = await client.accountSummary([], account_name)
  if (!snapshot) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }

  let products: string[] = [];
  let balances: string[] = [];

  snapshot.balances.forEach(balance => {
    products.push(balance.product)
    balances.push(balance.balance)
  }
  )

  try {
    return [products, balances]
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Failed to parse account summary snapshot"
    )
  }
}


/**
 * Search symbols by market name
 * @customfunction 
 */
export async function searchSymbols(market_name: string): Promise<string [] []> {
  const symbols = await client.searchSymbols({ searchString: market_name});

  const result = symbols.map(symbol => [symbol]);
  return result;
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    let success = await initializeClient()
    success ? console.log('Client initialized using saved API key/secret') : console.log('Client not initialized because of missing API key or secret');
  }
});
