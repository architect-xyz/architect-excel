import { type Config } from '@afintech/sdk/env/browser';

let config: Config = {
  host: 'https://app.architect.co/',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

export { config };

/**
 * Helper function to set an item in storage with fallback.
 */
export async function setStorageItem(key: string, value: string): Promise<void> {
  if (typeof localStorage !== 'undefined') {
    try {
      localStorage.setItem(key, value);
      console.log("Data saved to localStorage.");
      return;
    } catch (error) {
      console.error('Failed to set item in localStorage:', error);
      throw error;
    }
  }
  
  throw new Error('No available storage method to set data.');
}

/**
 * Helper function to get an item from storage with fallback.
 */
export async function getStorageItem(key: string): Promise<string | null> {
  if (typeof localStorage !== 'undefined') {
    try {
      return localStorage.getItem(key);
    } catch (error) {
      console.error('Failed to get item from localStorage:', error);
      throw error;
    }
  }
  console.log("No available storage method to retrieve data.");
  
  throw new Error('No available storage method to retrieve data.');
}

/**
 * Helper function to remove an item from storage with fallback.
 */
export async function removeStorageItem(key: string): Promise<void> {
  if (typeof localStorage !== 'undefined') {
    try {
      localStorage.removeItem(key);
      return;
    } catch (error) {
      console.error('Failed to remove item from localStorage:', error);
      throw error;
    }
  }
  
  throw new Error('No available storage method to remove data.');
}
