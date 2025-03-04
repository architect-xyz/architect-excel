// Excel Plugin Example: Query Backend API Using TypeScript


/*
important functions:
close
open
day high
day low
*/



/// <reference types="office-runtime" />
/// <reference types="office-js" />

import { create, Client } from '@afintech/sdk/env/browser';
import { getStorageItem, config } from './client';



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
 * Get the bid price of a market
 * @customfunction
 * @param market Market symbol
 * @returns The bbo prices of the given market
 * @volatile
 */
export async function getMarketBBO(market: string): Promise<number[] []> {
  const snapshot = await client.marketSnapshot([], market);
  if (!snapshot || !snapshot.bidPrice || !snapshot.askPrice) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "Received bad data from the server, please try again."
    )
  }
  try {
    const bid = parseFloat(snapshot.bidPrice);
    const ask = parseFloat(snapshot.askPrice);
    return [[bid, ask]]
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
 * @param market Market symbol
 * @returns The mid market price of the given market
 * @volatile
 * maybe a streaming function?
 */
export async function getMarketMid(market: string): Promise<number> {
  let snapshot;
  try {
    snapshot = await client.marketSnapshot([], market);
  } catch (error) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "Error getting market snapshot"
    )
  }


    if (!snapshot || !snapshot.bidPrice || !snapshot.askPrice) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "Received bad data from the server, please try again."
      )
    }

    const bid = parseFloat(snapshot.bidPrice);
    const ask = parseFloat(snapshot.askPrice);

    return isNaN(bid) || isNaN(ask) ? NaN : (bid + ask) / 2;
}

/**
 * returns the market name
 * @customfunction 
 */
export async function testClient(): Promise<string> {
  const market_name = "ES 20250321 CME Future";
  const snapshot = await client.filterMarkets([], {
    venue: 'CME',
    base: 'MES',
    quote: '',
    underlying: '',
    maxResults: 1,
    resultsOffset: 0,
    searchString: '',
    onlyFavorites: false,
    sortByVolumeDesc: true,
  });

  const market = snapshot[0].exchangeSymbol;

  return market;
}


Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    let success = await initializeClient()
    success ? console.log('Client initialized using saved API key/secret') : console.log('Client not initialized because of missing API key or secret');
  }
});