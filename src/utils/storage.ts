/**
 * IndexedDB and Browser Cache Storage Service for GM Fashions Inventory & Sales Suite.
 * Stores stock items, sales items, search history, and file metadata locally in the browser's
 * high-speed IndexedDB cache so switching between views or refreshing never loses data.
 */

import {
  MappedInventoryItem,
  MappedSalesItem,
  UploadedFileInfo,
  SearchState,
  StoreName,
} from '../types/inventory';

const DB_NAME = 'gm_fashions_pos_db';
const DB_VERSION = 2;
const STORE_CACHE = 'session_cache';

export interface SalesSearchState {
  query: string;
  stores: StoreName[];
  isAllStores: boolean;
}

interface CachedSessionData {
  id: string;
  stockItems: MappedInventoryItem[];
  stockFileInfo: UploadedFileInfo | null;
  stockSearchState?: SearchState | null;
  salesItems: MappedSalesItem[];
  salesFileInfo: UploadedFileInfo | null;
  salesSearchState?: SalesSearchState | null;
  savedAt: number;
}

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      reject(new Error('IndexedDB is not supported in this browser.'));
      return;
    }

    const request = window.indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_CACHE)) {
        db.createObjectStore(STORE_CACHE, { keyPath: 'id' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/**
 * Save Stock data to browser IndexedDB storage
 */
export async function saveStockDataToStorage(
  items: MappedInventoryItem[],
  info: UploadedFileInfo | null
): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);

    // Get existing to merge
    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      const existing: CachedSessionData = getReq.result || {
        id: 'current_session',
        stockItems: [],
        stockFileInfo: null,
        salesItems: [],
        salesFileInfo: null,
        savedAt: Date.now(),
      };

      existing.stockItems = items;
      existing.stockFileInfo = info;
      existing.savedAt = Date.now();

      store.put(existing);
    };
  } catch (err) {
    console.warn('Failed to persist stock data in browser storage:', err);
  }
}

/**
 * Save Stock search criteria to browser IndexedDB storage
 */
export async function saveStockSearchToStorage(
  searchState: SearchState | null
): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);

    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      if (getReq.result) {
        const existing: CachedSessionData = getReq.result;
        existing.stockSearchState = searchState;
        existing.savedAt = Date.now();
        store.put(existing);
      }
    };
  } catch (err) {
    console.warn('Failed to persist stock search in browser storage:', err);
  }
}

/**
 * Save Sales data to browser IndexedDB storage
 */
export async function saveSalesDataToStorage(
  items: MappedSalesItem[],
  info: UploadedFileInfo | null
): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);

    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      const existing: CachedSessionData = getReq.result || {
        id: 'current_session',
        stockItems: [],
        stockFileInfo: null,
        salesItems: [],
        salesFileInfo: null,
        savedAt: Date.now(),
      };

      existing.salesItems = items;
      existing.salesFileInfo = info;
      existing.savedAt = Date.now();

      store.put(existing);
    };
  } catch (err) {
    console.warn('Failed to persist sales data in browser storage:', err);
  }
}

/**
 * Save Sales search criteria (Toon Label query, store filters) to browser IndexedDB storage
 */
export async function saveSalesSearchToStorage(
  searchState: SalesSearchState | null
): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);

    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      if (getReq.result) {
        const existing: CachedSessionData = getReq.result;
        existing.salesSearchState = searchState;
        existing.savedAt = Date.now();
        store.put(existing);
      }
    };
  } catch (err) {
    console.warn('Failed to persist sales search in browser storage:', err);
  }
}

/**
 * Load both Stock and Sales data + search states from browser storage on app boot
 */
export async function loadSessionFromStorage(): Promise<{
  stockItems: MappedInventoryItem[];
  stockFileInfo: UploadedFileInfo | null;
  stockSearchState: SearchState | null;
  salesItems: MappedSalesItem[];
  salesFileInfo: UploadedFileInfo | null;
  salesSearchState: SalesSearchState | null;
} | null> {
  try {
    const db = await openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_CACHE, 'readonly');
      const store = tx.objectStore(STORE_CACHE);
      const req = store.get('current_session');

      req.onsuccess = () => {
        const data = req.result as CachedSessionData | undefined;
        if (data) {
          resolve({
            stockItems: data.stockItems || [],
            stockFileInfo: data.stockFileInfo || null,
            stockSearchState: data.stockSearchState || null,
            salesItems: data.salesItems || [],
            salesFileInfo: data.salesFileInfo || null,
            salesSearchState: data.salesSearchState || null,
          });
        } else {
          resolve(null);
        }
      };

      req.onerror = () => resolve(null);
    });
  } catch (err) {
    console.warn('Failed to load session from browser storage:', err);
    return null;
  }
}

/**
 * Clear full session from browser storage
 */
export async function clearStorageSession(): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);
    store.delete('current_session');
  } catch (err) {
    console.warn('Failed to clear browser storage:', err);
  }
}

/**
 * Clear only Stock data
 */
export async function clearStockStorage(): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);
    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      if (getReq.result) {
        const data = getReq.result as CachedSessionData;
        data.stockItems = [];
        data.stockFileInfo = null;
        data.stockSearchState = null;
        store.put(data);
      }
    };
  } catch (err) {
    console.warn('Failed to clear stock from browser storage:', err);
  }
}

/**
 * Clear only Sales data
 */
export async function clearSalesStorage(): Promise<void> {
  try {
    const db = await openDB();
    const tx = db.transaction(STORE_CACHE, 'readwrite');
    const store = tx.objectStore(STORE_CACHE);
    const getReq = store.get('current_session');
    getReq.onsuccess = () => {
      if (getReq.result) {
        const data = getReq.result as CachedSessionData;
        data.salesItems = [];
        data.salesFileInfo = null;
        data.salesSearchState = null;
        store.put(data);
      }
    };
  } catch (err) {
    console.warn('Failed to clear sales from browser storage:', err);
  }
}
