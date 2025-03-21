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
    Stack Overflow+3learn.microsoft.com+3learn.microsoft.com+3

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
 */
export async function initializeClient() : Promise<boolean> {
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
  return true;
}
/**
 * Get the last price of a market
 * @customfunction
 * @param market Market symbol
 * @returns The last price of the given market
 * @volatile
 */
export async function getMarketLast(market: string): Promise<number | undefined> {
  throw new CustomFunctions.Error(
    CustomFunctions.ErrorCode.notAvailable,
    'Not implemented'
  );
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
  let snapshot: Ticker = await client.ticker(["symbol", "bidPrice", "askPrice", "lastPrice"], symbol, venue)
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
 * maybe a streaming function?
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
 * @returns The ticker information
 * @volatile
 */
export async function getTicker(symbol: string, venue: string): Promise<number[] []> {
  let snapshot: Ticker = await client.ticker(["symbol", "bidPrice", "bidSize", "askPrice", "askSize", "lastPrice", "lastSize"], symbol, venue)
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