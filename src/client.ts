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
  // Try using OfficeRuntime.storage first.
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      await OfficeRuntime.storage.setItem(key, value);
      return;
    } catch (error) {
      console.warn("OfficeRuntime.storage.setItem failed, falling back to localStorage.", error);
    }
  }
  
  // Fallback to localStorage if available.
  if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      localStorage.setItem(storageKey, value);
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
  // Try using OfficeRuntime.storage first.
  console.log("TEST A")
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      return await OfficeRuntime.storage.getItem(key);
    } catch (error) {
      console.warn("OfficeRuntime.storage.getItem failed, falling back to localStorage.", error);
    }
  }
  
  // Fallback to localStorage if available.
  if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      return localStorage.getItem(storageKey);
    } catch (error) {
      console.error('Failed to get item from localStorage:', error);
      throw error;
    }
  }
  
  throw new Error('No available storage method to retrieve data.');
}

/**
 * Helper function to remove an item from storage with fallback.
 */
export async function removeStorageItem(key: string): Promise<void> {
  // Try using OfficeRuntime.storage first.
  if (typeof Office !== 'undefined' && Office.context && typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
    try {
      await OfficeRuntime.storage.removeItem(key);
      return;
    } catch (error) {
      console.warn("OfficeRuntime.storage.removeItem failed, falling back to localStorage.", error);
    }
  }
  
  // Fallback to localStorage if available.
  if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      localStorage.removeItem(storageKey);
      return;
    } catch (error) {
      console.error('Failed to remove item from localStorage:', error);
      throw error;
    }
  }
  
  throw new Error('No available storage method to remove data.');
}
