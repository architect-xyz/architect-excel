import { type Config } from '@afintech/sdk/env/browser';

let config: Config = {
  host: 'https://app.architect.co/',
  apiKey: '',
  apiSecret: '',
  tradingMode: 'live',
};

export { config };


export async function setOfficeSetting(key: string, value: string): Promise<void> {
  Office.context.document.settings
}

/**
 * Helper function to set an item in storage with fallback.
 */
export async function setStorageItem(key: string, value: string): Promise<void> {
  // Try using OfficeRuntime.storage first.
  if (typeof OfficeRuntime !== 'undefined' && typeof OfficeRuntime.storage !== 'undefined') {
    try {
      await OfficeRuntime.storage.setItem(key, value);
      console.log("Data saved to OfficeRuntime.storage.");
      return;
    } catch (error) {
      console.warn("OfficeRuntime.storage.setItem failed, falling back to localStorage.", error);
    }
  }
  console.log("Using localStorage to save data.");
  
  // Fallback to localStorage if available.
  if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      localStorage.setItem(storageKey, value);
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
  // Try using OfficeRuntime.storage first.
  if (typeof OfficeRuntime !== 'undefined' && typeof OfficeRuntime.storage !== 'undefined') {
    try {
      return await OfficeRuntime.storage.getItem(key);
    } catch (error) {
      console.warn("OfficeRuntime.storage.getItem failed, falling back to localStorage.", error);
    }
  }

  console.log("No OfficeRuntime.storage available, falling back to localStorage.");
  
  // Fallback to localStorage if available.
  if (typeof localStorage !== 'undefined') {
    try {
      const partitionKey = Office?.context?.partitionKey;
      const storageKey = partitionKey ? `${partitionKey}${key}` : key;
      console.log("Using localStorage to retrieve data.");
      return localStorage.getItem(storageKey);
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
  // Try using OfficeRuntime.storage first.
  if (typeof OfficeRuntime !== 'undefined' && typeof OfficeRuntime.storage !== 'undefined') {
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
