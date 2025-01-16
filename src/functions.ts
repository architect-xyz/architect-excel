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
 * @param sheetName Name of the worksheet containing API credentials
 */
async function initializeClient(sheetName: string) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    let worksheet = workbook.worksheets.getItemOrNullObject(sheetName);
    await context.sync();

    // Create API Config worksheet if it doesn't exist
    if (worksheet.isNullObject) {
      worksheet = workbook.worksheets.add(sheetName);
      worksheet.getRange('A1').values = [['API_KEY:']];
      worksheet.getRange('B1').values = [['']];
      worksheet.getRange('A2').values = [['API_SECRET:']];
      worksheet.getRange('B2').values = [['']];
      console.log(`Created ${sheetName} worksheet. Please fill in your API key and secret.`);
      await context.sync();
      return;
    }

    const apiKey = worksheet.getRange('B1').values[0][0] as string;
    const apiSecret = worksheet.getRange('B2').values[0][0] as string;

    if (!apiKey || !apiSecret) {
      throw new Error('API Key and Secret must be provided in cells B1 and B2.');
    }

    config.apiKey = apiKey;
    config.apiSecret = apiSecret;
    client = create(config);
  });
}

/**
 * Fetch market snapshot and populate Excel worksheet
 * @param market Market identifier
 */
async function fetchMarketSnapshot(market: string) {
  try {
    const snapshot = await client.marketSnapshot([], market);

    if (!snapshot) {
      console.error('No snapshot data received');
      return;
    }
  }
  catch (error) {
    console.error('Error fetching market snapshot:', error);
  }
}

/**
 * Main function to initialize and execute the plugin
 */
async function main() {
  try {
    const sheetName = 'ARCHITECT_CONFIG';
    const market = 'MES 20250321 CME Future/USD*CME/CQG';
    await initializeClient(sheetName);
    await fetchMarketSnapshot(market);
  } catch (error) {
    console.error('Error in main:', error);
  }
}

// Entry point for the Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('Excel Add-in ready.');
    main();
  }
});
export { fetchMarketSnapshot };