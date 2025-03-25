
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
  // Office add-in environment
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    await OfficeRuntime.storage.setItem(key, value);
  }
  // Browser environment with optional partition key support
  else if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      localStorage.setItem(storageKey, value);
    } catch (error) {
      console.error('Failed to set item in localStorage:', error);
      throw error;
    }
  } else {
    throw new Error('No available storage method to set data.');
  }
}

/**
 * Helper function to get an item from storage
 */
export async function getStorageItem(key: string): Promise<string | null> {
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    return await OfficeRuntime.storage.getItem(key);
  } else if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      return localStorage.getItem(storageKey);
    } catch (error) {
      console.error('Failed to get item from localStorage:', error);
      throw error;
    }
  } else {
    throw new Error('No available storage method to retrieve data.');
  }
}
