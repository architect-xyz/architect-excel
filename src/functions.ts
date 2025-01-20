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

let client: Client | null = null;

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

function clientCheck(): Client{
  if (!client) {
    throw new Error('Client is not initialized.');
  }
  return client
}

/**
 * Fetch market snapshot and populate Excel worksheet
 * @param market Market identifier
 */
async function getMarketMid(market: string): Promise<number | undefined> {
  try {
    client = clientCheck();

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
 * Main function to initialize and execute the plugin
 */
async function main() {
  try {
    const sheetName = 'ARCHITECT_CONFIG';
    const market = 'MES 20250321 CME Future/USD*CME/CQG';
    await getMarketMid(market);
  } catch (error) {
    console.error('Error in main:', error);
  }
}

// Entry point for the Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('Excel Add-in ready.');
    // main();
  }
});
export { getMarketMid };
