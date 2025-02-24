
import { type Config, } from '@afintech/sdk/env/browser';

let config: Config = {
  host: 'https://app.architect.co/',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

export {config};



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
