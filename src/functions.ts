// Excel Plugin Example: Query Backend API Using TypeScript

/*
important functions:

last price
mid price
bid price
ask price
close
open
day high
day low
*/

/// <reference types="office-runtime" />
/// <reference types="office-js" />
import { create, type Config, type Client } from '@afintech/sdk/env/browser';

console.log("loading 1")

let config: Config = {
  host: 'https://app.architect.co',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

console.log("loading 2")

let client: Client = (new Proxy({}, {
  get(_obj, _prop) {
    throw new Error('Client is not initialized');
  },
  set(_obj, _prop, _value) {
    throw new Error('Client is not initialized');
  }
}) as Client);

console.log("loading 3")

/**
 * Helper function to set an item in storage
 */
export async function setStorageItem(key: string, value: string): Promise<void> {
  if (typeof Office !== 'undefined' && Office.context) {
    try {
      await OfficeRuntime.storage.setItem(key, value);
    } catch (error) {
      console.error('Error setting storage item:', error);
    }
  } else {
    localStorage.setItem(key, value);
  }
}

/**
 * Helper function to get an item from storage
 */
export async function getStorageItem(key: string): Promise<string | null> {
  if (typeof Office !== 'undefined' && Office.context) {
    return await OfficeRuntime.storage.getItem(key);
  } else {
    return localStorage.getItem(key);
  }
}


/**
 * Initialize the client with user-provided API key and secret
 * @customfunction
 */
export async function initializeClient() {
  const apiKey = await getStorageItem('ArchitectApiKey');
  const apiSecret = await getStorageItem('ArchitectApiSecret');

  if (!apiKey || !apiSecret) {
    throw new Error('API Key and Secret must be provided.');
  }

  config.apiKey = apiKey;
  config.apiSecret = apiSecret;
  client = create(config);
}

/**
 * Fetch market snapshot and populate Excel worksheet
 * @customfunction
 * @param market Market identifier
 * @volatile
 * maybe a streaming function?
 */
export async function getMarketMid(market: string): Promise<number | undefined> {
  try {
    const snapshot = await client.marketSnapshot([], market);

    if (!snapshot || !snapshot.bidPrice || !snapshot.askPrice) {
      console.error('Invalid or missing snapshot data');
      return NaN;
    }

    const bid = parseFloat(snapshot.bidPrice);
    const ask = parseFloat(snapshot.askPrice);

    return isNaN(bid) || isNaN(ask) ? NaN : (bid + ask) / 2;
  } catch (error) {
    console.error('Error fetching market snapshot:', error);
    return undefined;
  }
}

/**
 * Returns a string for testing purposes
 * @customfunction 
 */
export function testFunction(): string {
  return "Hello World!";
}

/**
 * validates API key
 * @customfunction 
 */
export function validateAPIKey(): boolean {
  const apiKey = localStorage.getItem('ArchitectApiKey');
  const apiSecret = localStorage.getItem('ArchitectApiSecret');
  if (!apiKey || !apiSecret) {
    return false
  }
  return true
}

/**
 * returns the market name
 * @customfunction 
 */
export async function testClient(): Promise<string> {
  const market_name = 'MES 20250321 CME Future/USD*CME/CQG';

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

console.log("loading 2:07")