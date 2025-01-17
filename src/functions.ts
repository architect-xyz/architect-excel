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
import { create } from '@afintech/sdk';
import type {Config, Client} from '@afintech/sdk/dist/esm/sdk';


let config: Config = {
  host: 'https://app.architect.co',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

let client: Client;


/**
 * Uses the saved API key/secret for a custom function.
 * @customfunction
 */
export async function fetchData(): Promise<string> {
  const apiKey = Office.context.document.settings.get('apiKey');
  const apiSecret = Office.context.document.settings.get('apiSecret');

  if (!apiKey || !apiSecret) {
    throw new Error("API key or secret not set. Use the ribbon to configure them.");
  }

  // Example usage with the API
  return `Using API Key: ${apiKey}, Secret: ${apiSecret}`;
}

/**
 * Initialize the client with user-provided API key and secret
 * @customfunction
 */
export function initializeClient() {
  const apiKey = Office.context.document.settings.get('apiKey');
  const apiSecret = Office.context.document.settings.get('apiSecret');

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


/**
 * Main function to initialize and execute the plugin
 */
async function main() {
  try {
    const sheetName = 'ARCHITECT_CONFIG';
    const market = 'MES 20250321 CME Future/USD*CME/CQG';
    await initializeClient();
    await fetchMarketSnapshot(market);
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
export { fetchMarketSnapshot };