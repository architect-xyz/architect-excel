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
 * Initialize the client with user-provided API key and secret
 * @customfunction
 */
export function initializeClient() {
  const apiKey = localStorage.getItem('ArchitectApiKey');
  const apiSecret = localStorage.getItem('ArchitectApiSecret');

  if (!apiKey || !apiSecret) {
    throw new Error('API Key and Secret must be provided.');
  }

  config.apiKey = apiKey;
  config.apiSecret = apiSecret;
  client = create(config);
}

/**
 * Fetch market snapshot and populate Excel worksheet
 * @param market Market identifier
 */
async function getMarketMid(market: string): Promise<number | undefined> {
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

function testAPI(): string {
  const apiKey = localStorage.getItem('ArchitectApiKey');
  return apiKey ?? "No Key";
}

async function testClient(): Promise<string> {
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


export { getMarketMid, testAPI, testClient };
