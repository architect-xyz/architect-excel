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


let config: Config = {
  host: 'https://app.architect.co',
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

/**
 * Helper function to set an item in storage
 */
export async function setStorageItem(key: string, value: string): Promise<void> {
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined') {
    await OfficeRuntime.storage.setItem(key, value);
  } else if (typeof localStorage !== 'undefined') {
    localStorage.setItem(key, value);
  } else {
    throw new Error('No available storage method to set to.');
  }
}

/**
 * Helper function to get an item from storage
 */
export async function getStorageItem(key: string): Promise<string | null> {
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined') {
    return await OfficeRuntime.storage.getItem(key);
  } else if (typeof localStorage !== 'undefined') {
    return localStorage.getItem(key);
  } else {
    throw new Error('No available storage method to get from.');
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
 * Test error
 * @customfunction
 */
export function testError(): string {
  throw new Error('Test error');
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
  const apiKey = getStorageItem('ArchitectApiKey');
  const apiSecret = getStorageItem('ArchitectApiSecret');
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


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeClient().catch(error => console.error(error));
  }
});