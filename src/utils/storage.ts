/**
 * High-Capacity IndexedDB & Browser Cache Storage Service for GM Fashions Inventory & Sales Suite.
 * Configured to support up to 1 GB (1,024 MB) of uploaded report cache data using chunked IndexedDB
 * storage + persistent storage quota negotiation, while starting fresh on browser reload.
 */

import {
  MappedInventoryItem,
  MappedSalesItem,
  UploadedFileInfo,
  SearchState,
  StoreName,
} from '../types/inventory';

export const MAX_CACHE_MEMORY_BYTES = 1024 * 1024 * 1024; // 1 GB (1,073,741,824 bytes)
export const MAX_FILE_UPLOAD_BYTES = 1024 * 1024 * 1024;  // 1 GB (1,073,741,824 bytes)

const DB_NAME = 'gm_fashions_pos_db';
const DB_VERSION = 3;
const STORE_META = 'cache_meta';
const STORE_CHUNKS = 'cache_chunks';
const LEGACY_STORE_CACHE = 'session_cache';

const CACHE_CHUNK_SIZE = 10000; // 10,000 records per chunk to stay well under structured clone limits

export interface SalesSearchState {
  query: string;
  stores: StoreName[];
  isAllStores: boolean;
}

export interface StorageCapacityInfo {
  persisted: boolean;
  quotaBytes: number;
  usageBytes: number;
  maxCacheLimitBytes: number;
}

interface CacheMetaRecord {
  id: 'stock_meta' | 'sales_meta' | 'stock_search' | 'sales_search';
  fileInfo?: UploadedFileInfo | null;
  totalChunks?: number;
  totalItems?: number;
  estimatedBytes?: number;
  stockSearchState?: SearchState | null;
  salesSearchState?: SalesSearchState | null;
  savedAt: number;
}

interface CacheChunkRecord {
  id: string; // e.g. "stock_0", "sales_12"
  type: 'stock' | 'sales';
  chunkIndex: number;
  items: MappedInventoryItem[] | MappedSalesItem[];
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Formats byte sizes cleanly into KB, MB, or GB.
 */
export function formatByteSize(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 B';
  if (bytes >= 1024 * 1024 * 1024) {
    return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
  }
  if (bytes >= 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }
  if (bytes >= 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${bytes} B`;
}

/**
 * Requests browser persistent storage permission so IndexedDB can cache up to 1 GB+
 * without premature browser eviction.
 */
export async function requestHighCapacityStorage(): Promise<StorageCapacityInfo> {
  let persisted = false;
  let quotaBytes = MAX_CACHE_MEMORY_BYTES;
  let usageBytes = 0;

  try {
    if (typeof navigator !== 'undefined' && navigator.storage) {
      if (typeof navigator.storage.persist === 'function') {
        persisted = await navigator.storage.persist().catch(() => false);
      }
      if (typeof navigator.storage.estimate === 'function') {
        const estimate = await navigator.storage.estimate();
        if (estimate.quota) quotaBytes = estimate.quota;
        if (estimate.usage) usageBytes = estimate.usage;
      }
    }
  } catch (err) {
    console.warn('Storage quota check skipped:', err);
  }

  return {
    persisted,
    quotaBytes,
    usageBytes,
    maxCacheLimitBytes: MAX_CACHE_MEMORY_BYTES,
  };
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
      if (!db.objectStoreNames.contains(STORE_META)) {
        db.createObjectStore(STORE_META, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_CHUNKS)) {
        const chunkStore = db.createObjectStore(STORE_CHUNKS, { keyPath: 'id' });
        chunkStore.createIndex('by_type', 'type', { unique: false });
      }
      if (!db.objectStoreNames.contains(LEGACY_STORE_CACHE)) {
        db.createObjectStore(LEGACY_STORE_CACHE, { keyPath: 'id' });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function runTx(
  db: IDBDatabase,
  storeNames: string[],
  mode: IDBTransactionMode,
  executor: (stores: Record<string, IDBObjectStore>) => void
): Promise<void> {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeNames, mode);
    const stores: Record<string, IDBObjectStore> = {};
    for (const name of storeNames) {
      stores[name] = tx.objectStore(name);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error('IndexedDB transaction aborted'));
    executor(stores);
  });
}

async function clearChunksByType(db: IDBDatabase, type: 'stock' | 'sales'): Promise<void> {
  return new Promise((resolve, reject) => {
    const tx = db.transaction([STORE_CHUNKS, STORE_META], 'readwrite');
    const chunkStore = tx.objectStore(STORE_CHUNKS);
    const metaStore = tx.objectStore(STORE_META);
    const index = chunkStore.index('by_type');
    const range = IDBKeyRange.only(type);

    const cursorReq = index.openCursor(range);
    cursorReq.onsuccess = (e) => {
      const cursor = (e.target as IDBRequest<IDBCursorWithValue | null>).result;
      if (cursor) {
        cursor.delete();
        cursor.continue();
      }
    };

    metaStore.delete(type === 'stock' ? 'stock_meta' : 'sales_meta');

    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function estimateDatasetBytes(items: unknown[], fileInfo: UploadedFileInfo | null): number {
  const fileBytes = fileInfo?.fileSize || 0;
  if (items.length === 0) return fileBytes;
  // Estimate ~280 bytes per mapped record in memory/IDB, take max with raw file size
  return Math.max(fileBytes, items.length * 280);
}

/**
 * Save Stock data in 10,000-row chunks into IndexedDB session cache (supports up to 1 GB).
 * Executes non-blocking across event loop ticks so huge datasets never freeze the UI.
 */
export async function saveStockDataToStorage(
  items: MappedInventoryItem[],
  info: UploadedFileInfo | null
): Promise<void> {
  try {
    await requestHighCapacityStorage();
    const estimatedBytes = estimateDatasetBytes(items, info);
    if (estimatedBytes > MAX_CACHE_MEMORY_BYTES) {
      console.warn('Stock dataset exceeds 1 GB cache ceiling; keeping active session in RAM only.');
      return;
    }

    const db = await openDB();
    await clearChunksByType(db, 'stock');

    if (items.length === 0) {
      db.close();
      return;
    }

    const totalChunks = Math.ceil(items.length / CACHE_CHUNK_SIZE);

    for (let chunkIdx = 0; chunkIdx < totalChunks; chunkIdx++) {
      const start = chunkIdx * CACHE_CHUNK_SIZE;
      const slice = items.slice(start, start + CACHE_CHUNK_SIZE);
      const chunkRecord: CacheChunkRecord = {
        id: `stock_${chunkIdx}`,
        type: 'stock',
        chunkIndex: chunkIdx,
        items: slice,
      };
      await runTx(db, [STORE_CHUNKS], 'readwrite', (stores) => {
        stores[STORE_CHUNKS].put(chunkRecord);
      });
      if (chunkIdx % 2 === 1) {
        await yieldToEventLoop();
      }
    }

    const metaRecord: CacheMetaRecord = {
      id: 'stock_meta',
      fileInfo: info,
      totalChunks,
      totalItems: items.length,
      estimatedBytes,
      savedAt: Date.now(),
    };
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].put(metaRecord);
    });

    db.close();
  } catch (err) {
    console.warn('Chunked stock cache write skipped (continuing in-memory):', err);
  }
}

/**
 * Save Stock search criteria to browser IndexedDB session cache
 */
export async function saveStockSearchToStorage(
  searchState: SearchState | null
): Promise<void> {
  try {
    const db = await openDB();
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].put({
        id: 'stock_search',
        stockSearchState: searchState,
        savedAt: Date.now(),
      } as CacheMetaRecord);
    });
    db.close();
  } catch (err) {
    console.warn('Failed to save stock search to cache:', err);
  }
}

/**
 * Save Sales data in 10,000-row chunks into IndexedDB session cache (supports up to 1 GB).
 */
export async function saveSalesDataToStorage(
  items: MappedSalesItem[],
  info: UploadedFileInfo | null
): Promise<void> {
  try {
    await requestHighCapacityStorage();
    const estimatedBytes = estimateDatasetBytes(items, info);
    if (estimatedBytes > MAX_CACHE_MEMORY_BYTES) {
      console.warn('Sales dataset exceeds 1 GB cache ceiling; keeping active session in RAM only.');
      return;
    }

    const db = await openDB();
    await clearChunksByType(db, 'sales');

    if (items.length === 0) {
      db.close();
      return;
    }

    const totalChunks = Math.ceil(items.length / CACHE_CHUNK_SIZE);

    for (let chunkIdx = 0; chunkIdx < totalChunks; chunkIdx++) {
      const start = chunkIdx * CACHE_CHUNK_SIZE;
      const slice = items.slice(start, start + CACHE_CHUNK_SIZE);
      const chunkRecord: CacheChunkRecord = {
        id: `sales_${chunkIdx}`,
        type: 'sales',
        chunkIndex: chunkIdx,
        items: slice,
      };
      await runTx(db, [STORE_CHUNKS], 'readwrite', (stores) => {
        stores[STORE_CHUNKS].put(chunkRecord);
      });
      if (chunkIdx % 2 === 1) {
        await yieldToEventLoop();
      }
    }

    const metaRecord: CacheMetaRecord = {
      id: 'sales_meta',
      fileInfo: info,
      totalChunks,
      totalItems: items.length,
      estimatedBytes,
      savedAt: Date.now(),
    };
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].put(metaRecord);
    });

    db.close();
  } catch (err) {
    console.warn('Chunked sales cache write skipped (continuing in-memory):', err);
  }
}

/**
 * Save Sales search criteria to browser IndexedDB session cache
 */
export async function saveSalesSearchToStorage(
  searchState: SalesSearchState | null
): Promise<void> {
  try {
    const db = await openDB();
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].put({
        id: 'sales_search',
        salesSearchState: searchState,
        savedAt: Date.now(),
      } as CacheMetaRecord);
    });
    db.close();
  } catch (err) {
    console.warn('Failed to save sales search to cache:', err);
  }
}

/**
 * Load session data on app boot - always returns clean fresh state so the app starts fresh on reload
 */
export async function loadSessionFromStorage(): Promise<{
  stockItems: MappedInventoryItem[];
  stockFileInfo: UploadedFileInfo | null;
  stockSearchState: SearchState | null;
  salesItems: MappedSalesItem[];
  salesFileInfo: UploadedFileInfo | null;
  salesSearchState: SalesSearchState | null;
} | null> {
  return {
    stockItems: [],
    stockFileInfo: null,
    stockSearchState: null,
    salesItems: [],
    salesFileInfo: null,
    salesSearchState: null,
  };
}

/**
 * Clear full session from browser storage (all chunks + metadata + legacy store)
 */
export async function clearStorageSession(): Promise<void> {
  try {
    const db = await openDB();
    await runTx(db, [STORE_META, STORE_CHUNKS, LEGACY_STORE_CACHE], 'readwrite', (stores) => {
      stores[STORE_META].clear();
      stores[STORE_CHUNKS].clear();
      stores[LEGACY_STORE_CACHE].clear();
    });
    db.close();
  } catch (err) {
    console.warn('Failed to clear browser storage:', err);
  }
}

/**
 * Clear only Stock data from cache
 */
export async function clearStockStorage(): Promise<void> {
  try {
    const db = await openDB();
    await clearChunksByType(db, 'stock');
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].delete('stock_search');
    });
    db.close();
  } catch (err) {
    console.warn('Failed to clear stock from browser storage:', err);
  }
}

/**
 * Clear only Sales data from cache
 */
export async function clearSalesStorage(): Promise<void> {
  try {
    const db = await openDB();
    await clearChunksByType(db, 'sales');
    await runTx(db, [STORE_META], 'readwrite', (stores) => {
      stores[STORE_META].delete('sales_search');
    });
    db.close();
  } catch (err) {
    console.warn('Failed to clear sales from browser storage:', err);
  }
}
