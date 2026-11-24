import {
  XLSX,
  styleMatrixWorksheet,
  styleShuffledMatrixWorksheet,
  styleDetailedWorksheet,
  calculateColumnWidths,
  MatrixColorGroup,
} from './excelStyler';
import {
  MappedInventoryItem,
  StoreName,
  ALL_STORES,
  TOON_EXCEL_STORES,
  ToonMatrixRow,
} from '../types/inventory';
import {
  ShuffledMatrixRow,
  StoreCellShuffleInfo,
  MatrixStoreKey,
  STORE_SHORTCUTS,
} from './stockShuffle';
import { SAMPLE_WONDERSOFT_DATA } from './sampleData';
import { MAX_FILE_UPLOAD_BYTES, formatByteSize } from './storage';
import {
  ParseProgress,
  streamCsvFileInBatches,
  yieldToMainThread,
} from './streamingCsvParser';

/**
 * Normalizes header keys by lowercasing and stripping all non-alphanumeric chars
 */
export function normalizeKey(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Robust EAN string extractor:
 * Handles scientific notation, decimals (.0), trailing spaces, leading quotes, and string conversions
 */
export function cleanEanString(val: unknown): string {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (!str) return '';

  // Remove leading/trailing quotes often present in Excel text export e.g. '8907701122334
  str = str.replace(/^['"]+/, '').replace(/['"]+$/, '').trim();

  // If already pure integer string of 8-18 digits
  if (/^\d{8,18}$/.test(str)) {
    return str;
  }

  // Handle scientific notation e.g. "8.907701122334e+12", "8.9077E12", "8.907701122334E+012"
  if (/^[0-9.]+[eE][+-]?[0-9]+$/i.test(str)) {
    try {
      const num = Number(str);
      if (!isNaN(num) && isFinite(num)) {
        str = BigInt(Math.round(num)).toString();
      }
    } catch {
      // fallback
    }
  }

  // Handle decimals from Excel floating number e.g. "8907701122334.0", "8907701122334.00"
  if (str.includes('.') && !str.toLowerCase().includes('e')) {
    const [intPart, decPart] = str.split('.');
    if (!decPart || /^0+$/.test(decPart)) {
      str = intPart;
    }
  }

  return str.trim();
}

/**
 * Robust number cleaner: parses strings, commas, rupee symbols, etc.
 */
function cleanNumber(val: unknown, defaultVal = 0): number {
  if (val === null || val === undefined || val === '') return defaultVal;
  if (typeof val === 'number') return isNaN(val) ? defaultVal : val;
  const cleaned = String(val).replace(/[^0-9.-]/g, '');
  const parsed = parseFloat(cleaned);
  return isNaN(parsed) ? defaultVal : parsed;
}

/**
 * Attempts to match a raw string (branch, outlet, sheet name, column header) to a known StoreName.
 * Returns null if the string is generic (e.g. "Sales", "Report", "Sheet1", "GM Fashions") and does not explicitly name a store.
 */
export function tryMatchStoreName(rawStore: string): StoreName | null {
  if (!rawStore) return null;
  const str = String(rawStore).trim();
  const norm = str.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!norm) return null;

  // 1. Kootapalli FIRST and comprehensively
  // Supports: Kootapalli, Koottappalli, Koottapalli, Kootappalli, Koothapalli,
  // Koota Palli, Kootta Palli, Kuttapalli, Koottampalli, KTP, GM Fashions Koottappalli, etc.
  if (
    norm.includes('koot') ||
    norm.includes('kooth') ||
    norm.includes('koott') ||
    norm.includes('kuttap') ||
    norm.includes('kootam') ||
    norm.includes('koottam') ||
    norm.startsWith('ktp') ||
    norm.includes('ktp')
  ) {
    return 'Kootapalli';
  }

  // 2. Salem
  if (norm.includes('salem') || norm.includes('slm')) return 'Salem';

  // 3. Karur
  if (norm.includes('karur') || norm.includes('krr')) return 'Karur';

  // 4. Mallur
  if (norm.includes('mallur') || norm.includes('mlr')) return 'Mallur';

  // 5. Namakkal
  if (norm.includes('namakkal') || norm.includes('nmkl') || norm.includes('namakal')) return 'Namakkal';

  // 6. Kumbakonam
  if (norm.includes('kumba') || norm.includes('kumbakonam') || norm.includes('kmb')) return 'Kumbakonam';

  // 7. Thiruvannamalai
  if (
    norm.includes('thiruvan') ||
    norm.includes('tiruvan') ||
    norm.includes('tvm') ||
    norm.includes('thiruvanna') ||
    norm.includes('tiruvanna')
  ) {
    return 'Thiruvannamalai';
  }

  // 8. GM Fashions Warehouse - ONLY matched if explicitly containing warehouse / godown / central terms
  // NEVER match apparel categories ('kidswear', 'menswear', 'nightwear') or tax terms ('centraltax', 'cgst')
  const isExcludedApparelOrTax =
    norm.includes('kidswear') ||
    norm.includes('menswear') ||
    norm.includes('womenswear') ||
    norm.includes('partywear') ||
    norm.includes('innerwear') ||
    norm.includes('nightwear') ||
    norm.includes('footwear') ||
    norm.includes('casualwear') ||
    norm.includes('sportswear') ||
    norm.includes('software') ||
    norm.includes('hardware') ||
    norm.includes('centraltax') ||
    norm.includes('centralgst') ||
    norm.includes('centralcess');

  if (!isExcludedApparelOrTax) {
    if (
      norm.includes('warehouse') ||
      norm.includes('godown') ||
      norm.includes('mainwh') ||
      norm.includes('headoffice') ||
      norm.includes('gmfashionswarehouse') ||
      norm.includes('centralwarehouse') ||
      norm.includes('centralgodown') ||
      norm.includes('centralwh') ||
      norm === 'wh' ||
      norm === 'central' ||
      norm === 'warehouse'
    ) {
      return 'GM Fashions Warehouse';
    }
  }

  // Exact store name match
  const directMatch = ALL_STORES.find(
    (s) => s.toLowerCase().replace(/[^a-z0-9]/g, '') === norm
  );
  if (directMatch) return directMatch;

  return null;
}

/**
 * Robust store name matching (case-insensitive & fuzzy) with fallback
 */
export function matchStoreName(rawStore: string, fallback: StoreName = 'Kootapalli'): StoreName {
  const matched = tryMatchStoreName(rawStore);
  if (matched) return matched;

  const norm = String(rawStore || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (
    norm.includes('warehouse') ||
    norm.includes('godown') ||
    norm.includes('gmfashionswarehouse') ||
    norm.includes('centralwarehouse') ||
    norm.includes('centralgodown')
  ) {
    return 'GM Fashions Warehouse';
  }

  return fallback;
}

/**
 * Detects the real header row index from a 2D array of rows (works for streaming CSV or dense sheets).
 */
export function findHeaderRowIndexFromRows(rows: unknown[][]): number {
  const maxSearch = Math.min(rows.length - 1, 30);
  let bestRow = 0;
  let bestScore = 0;

  const keywordCategories = [
    ['ean', 'barcode', 'upc', 'gtin', 'sku', 'bcode'],
    ['productgroup', 'productgrp', 'prodgroup', 'producttype', 'group', 'category'],
    ['department', 'dept', 'itemname', 'productname', 'description'],
    ['class', 'subclass', 'brand'],
    ['color', 'colour', 'shade'],
    ['size', 'itemsize'],
    ['garment', 'garments', 'fabric', 'material'],
    ['mrp', 'price', 'rate', 'salerate', 'rsp'],
    ['stock', 'qty', 'quantity', 'closingstock', 'balqty', 'totalqty'],
    ['toon', 'style', 'styleno', 'itemcode', 'design'],
    ['store', 'branch', 'outlet', 'location', 'salem', 'karur', 'mallur', 'namakkal', 'kumbakonam', 'kootapalli', 'koottappalli', 'koottapalli', 'koot', 'warehouse', 'thiruvannamalai']
  ];

  for (let r = 0; r <= maxSearch; r++) {
    const row = rows[r];
    if (!row || !Array.isArray(row)) continue;

    let nonBlankCount = 0;
    const cellValues: string[] = [];

    for (let c = 0; c < row.length; c++) {
      const v = row[c];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        nonBlankCount++;
        cellValues.push(String(v).toLowerCase().replace(/[^a-z0-9]/g, ''));
      }
    }

    if (nonBlankCount < 3) continue;

    let rowScore = 0;
    for (const group of keywordCategories) {
      if (cellValues.some((val) => group.some((kw) => val.includes(kw)))) {
        rowScore += 3;
      }
    }

    if (cellValues.some((val) => val.includes('ean') || val.includes('barcode'))) {
      rowScore += 8;
    }

    if (rowScore > bestScore) {
      bestScore = rowScore;
      bestRow = r;
    }
  }

  return bestScore >= 6 ? bestRow : 0;
}

/**
 * Detects the real header row index in a worksheet (supports both standard and dense SheetJS modes).
 * In many real Wondersoft/ERP exports, rows 0-4 contain company titles, dates, or branch names.
 */
export function findHeaderRowIndex(worksheet: XLSX.WorkSheet): number {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
  const maxSearch = Math.min(range.e.r, 30);
  const denseWs = worksheet as unknown as Record<number, (XLSX.CellObject | undefined)[] | undefined>;

  let bestRow = 0;
  let bestScore = 0;

  for (let r = range.s.r; r <= maxSearch; r++) {
    let nonBlankCount = 0;
    const cellValues: string[] = [];
    const denseRow = denseWs[r];

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = denseRow ? denseRow[c] : worksheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
        nonBlankCount++;
        cellValues.push(String(cell.v).toLowerCase().replace(/[^a-z0-9]/g, ''));
      }
    }

    // A real header row must span at least 3 distinct non-empty columns (skipping single banner titles)
    if (nonBlankCount < 3) {
      continue;
    }

    const keywordCategories = [
      ['ean', 'barcode', 'upc', 'gtin', 'sku', 'bcode'],
      ['productgroup', 'productgrp', 'prodgroup', 'producttype', 'group', 'category'],
      ['department', 'dept', 'itemname', 'productname', 'description'],
      ['class', 'subclass', 'brand'],
      ['color', 'colour', 'shade'],
      ['size', 'itemsize'],
      ['garment', 'garments', 'fabric', 'material'],
      ['mrp', 'price', 'rate', 'salerate', 'rsp'],
      ['stock', 'qty', 'quantity', 'closingstock', 'balqty', 'totalqty'],
      ['toon', 'style', 'styleno', 'itemcode', 'design'],
      ['store', 'branch', 'outlet', 'location', 'salem', 'karur', 'mallur', 'namakkal', 'kumbakonam', 'kootapalli', 'koottappalli', 'koottapalli', 'koot', 'warehouse', 'thiruvannamalai']
    ];

    let rowScore = 0;
    for (const group of keywordCategories) {
      if (cellValues.some((val) => group.some((kw) => val.includes(kw)))) {
        rowScore += 3;
      }
    }

    // Extra weight if the row explicitly contains an EAN or Barcode column header
    if (cellValues.some((val) => val.includes('ean') || val.includes('barcode'))) {
      rowScore += 8;
    }

    if (rowScore > bestScore) {
      bestScore = rowScore;
      bestRow = r;
    }
  }

  return bestScore >= 6 ? bestRow : 0;
}

/**
 * Formats store name into the standardized uppercase display used in store count boxes
 * (matches user image: KOOTTAPPALLI, KUMBAKONAM, MALLUR, NAMAKKAL, SALEM, TIRUVANNAMALAI, KARUR, WAREHOUSE)
 */
export function formatStoreDisplayName(store: string): string {
  const norm = String(store || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (
    norm.includes('koot') ||
    norm.includes('kooth') ||
    norm.includes('koott') ||
    norm.includes('kuttap') ||
    norm.includes('kootam') ||
    norm.includes('koottam') ||
    norm.startsWith('ktp') ||
    norm.includes('ktp')
  ) {
    return 'KOOTTAPPALLI';
  }
  if (norm.includes('kumba') || norm.includes('kumbakonam') || norm.includes('kmb')) return 'KUMBAKONAM';
  if (norm.includes('mallur') || norm.includes('mlr')) return 'MALLUR';
  if (norm.includes('namakkal') || norm.includes('namakal') || norm.includes('nmkl')) return 'NAMAKKAL';
  if (norm.includes('salem') || norm.includes('slm')) return 'SALEM';
  if (norm.includes('tiruvan') || norm.includes('thiruvan') || norm.includes('tvm') || norm.includes('thiruvanna')) return 'TIRUVANNAMALAI';
  if (norm.includes('karur') || norm.includes('krr')) return 'KARUR';
  if (
    norm.includes('ware') ||
    norm === 'wh' ||
    norm.includes('central') ||
    norm.includes('godown') ||
    norm.includes('gmfashionswarehouse') ||
    norm === 'warehouse'
  ) {
    return 'GM FASHIONS WAREHOUSE';
  }
  return String(store || '').toUpperCase().trim();
}

/**
 * Finds all store quantity columns in a wide-format row (supports variations like KOOTTAPPALLI, TIRUVANNAMALAI)
 */
export function findStoreColumnsInRow(
  lookup: Record<string, unknown>
): { store: StoreName; displayName: string; qty: number }[] {
  const storeEntries: { store: StoreName; displayName: string; qty: number }[] = [];
  const checkedKeys = new Set<string>();

  for (const [rawKey, val] of Object.entries(lookup)) {
    const norm = normalizeKey(rawKey);
    if (checkedKeys.has(norm)) continue;

    // Skip non-store metadata columns
    if (
      norm.includes('product') ||
      norm.includes('dept') ||
      norm.includes('class') ||
      norm.includes('brand') ||
      norm.includes('color') ||
      norm.includes('colour') ||
      norm.includes('size') ||
      norm.includes('garment') ||
      norm.includes('fabric') ||
      norm.includes('ean') ||
      norm.includes('barcode') ||
      norm.includes('mrp') ||
      norm.includes('price') ||
      norm.includes('rate') ||
      norm.includes('toon') ||
      norm.includes('style') ||
      norm.includes('image') ||
      norm === 'total' ||
      norm === 'totalqty' ||
      norm === 'totalstock' ||
      norm === 'totalquantity'
    ) {
      continue;
    }

    if (
      norm.includes('koot') ||
      norm.includes('kooth') ||
      norm.includes('koott') ||
      norm.includes('kuttap') ||
      norm.includes('kootam') ||
      norm.includes('koottam') ||
      norm.startsWith('ktp') ||
      norm.includes('ktp')
    ) {
      storeEntries.push({ store: 'Kootapalli', displayName: 'KOOTTAPPALLI', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('kumba') || norm.includes('kumbakonam') || norm.includes('kmb')) {
      storeEntries.push({ store: 'Kumbakonam', displayName: 'KUMBAKONAM', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('mallur') || norm.includes('mlr')) {
      storeEntries.push({ store: 'Mallur', displayName: 'MALLUR', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('namakkal') || norm.includes('namakal') || norm.includes('nmkl')) {
      storeEntries.push({ store: 'Namakkal', displayName: 'NAMAKKAL', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('salem') || norm.includes('slm')) {
      storeEntries.push({ store: 'Salem', displayName: 'SALEM', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('tiruvan') || norm.includes('thiruvan') || norm.includes('tvm') || norm.includes('thiruvanna')) {
      storeEntries.push({ store: 'Thiruvannamalai', displayName: 'TIRUVANNAMALAI', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (norm.includes('karur') || norm.includes('krr')) {
      storeEntries.push({ store: 'Karur', displayName: 'KARUR', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    } else if (
      norm.includes('ware') ||
      norm === 'wh' ||
      norm.includes('godown') ||
      norm.includes('central') ||
      norm.includes('gmfashionswarehouse')
    ) {
      storeEntries.push({ store: 'GM Fashions Warehouse', displayName: 'GM FASHIONS WAREHOUSE', qty: cleanNumber(val, 0) });
      checkedKeys.add(norm);
    }
  }

  return storeEntries;
}

/**
 * Calculates and attaches totalQuantity and storeBreakdown across all stores for each unique item/EAN.
 * Mutates items in-place and shares storeBreakdown references to keep RAM usage minimal for 1 GB uploads.
 */
export function calculateTotalQuantities(items: MappedInventoryItem[]): MappedInventoryItem[] {
  const totalsMap = new Map<string, number>();
  const storeBreakdownsMap = new Map<string, Record<string, number>>();

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = item.ean
      ? item.ean.trim()
      : `${item.department}|${item.productName}|${item.colour}|${item.size}|${item.toonLabel}`.toLowerCase();

    const curr = totalsMap.get(key) || 0;
    totalsMap.set(key, curr + (item.stockQuantity || 0));

    let bd = storeBreakdownsMap.get(key);
    if (!bd) {
      bd = {
        'GM FASHIONS WAREHOUSE': 0,
      };
      storeBreakdownsMap.set(key, bd);
    }

    if (item.storeBreakdown && Object.keys(item.storeBreakdown).length > 0) {
      for (const [st, q] of Object.entries(item.storeBreakdown)) {
        const sDisplay = formatStoreDisplayName(st);
        bd[sDisplay] = Math.max(bd[sDisplay] || 0, q);
      }
    } else {
      const sDisplay = formatStoreDisplayName(item.store);
      bd[sDisplay] = (bd[sDisplay] || 0) + (item.stockQuantity || 0);
    }

    if (bd['GM FASHIONS WAREHOUSE'] === undefined) {
      bd['GM FASHIONS WAREHOUSE'] = 0;
    }
  }

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = item.ean
      ? item.ean.trim()
      : `${item.department}|${item.productName}|${item.colour}|${item.size}|${item.toonLabel}`.toLowerCase();

    const breakdown = storeBreakdownsMap.get(key) || {
      [formatStoreDisplayName(item.store)]: item.stockQuantity || 0,
      'GM FASHIONS WAREHOUSE': 0,
    };

    if (breakdown['GM FASHIONS WAREHOUSE'] === undefined) {
      breakdown['GM FASHIONS WAREHOUSE'] = 0;
    }

    item.totalQuantity = totalsMap.get(key) || item.stockQuantity || 0;
    item.storeBreakdown = breakdown;
  }

  return items;
}

/**
 * Maps a single row or converts wide multi-store rows into MappedInventoryItem[]
 *
 * Explicit mapping requested by user:
 * - product group  ->  department
 * - department     ->  product name
 * - class          ->  brand
 * - color          ->  colour
 * - size           ->  size
 * - garments       ->  fabric
 * - EAN code       ->  EAN
 * - mrp            ->  selling price
 * - store / branch ->  store
 * - stock / qty    ->  stock quantity (store-wise quantity)
 * - total quantity ->  total stock quantity across stores
 */
export function mapWondersoftRow(
  raw: Record<string, unknown>,
  index: number
): MappedInventoryItem[] {
  const lookup: Record<string, unknown> = {};
  for (const [key, val] of Object.entries(raw)) {
    lookup[normalizeKey(key)] = val;
  }

  // 1. Department (from Wondersoft: Product Group)
  const department = String(
    lookup['productgroup'] ||
      lookup['productgrp'] ||
      lookup['prodgroup'] ||
      lookup['group'] ||
      lookup['producttype'] ||
      lookup['category'] ||
      lookup['pg'] ||
      'Apparel'
  ).trim();

  // 2. Product Name (from Wondersoft: Department)
  const productName = String(
    lookup['department'] ||
      lookup['dept'] ||
      lookup['productname'] ||
      lookup['itemname'] ||
      lookup['itemdescription'] ||
      lookup['description'] ||
      lookup['desc'] ||
      lookup['product'] ||
      'Fashion Wear'
  ).trim();

  // 3. Brand (from Wondersoft: Class)
  const brand = String(
    lookup['class'] ||
      lookup['brand'] ||
      lookup['brandname'] ||
      lookup['subclass'] ||
      lookup['categoryclass'] ||
      'GM Fashions'
  ).trim();

  // 4. Colour (from Wondersoft: Color / Colour)
  const colour = String(
    lookup['color'] ||
      lookup['colour'] ||
      lookup['shade'] ||
      lookup['colorname'] ||
      lookup['colourname'] ||
      'Standard'
  ).trim();

  // 5. Size (from Wondersoft: Size)
  const size = String(
    lookup['size'] ||
      lookup['itemsize'] ||
      lookup['sizedesc'] ||
      'Free'
  ).trim();

  // 6. Fabric (from Wondersoft: Garments / Garment)
  const fabric = String(
    lookup['garments'] ||
      lookup['garment'] ||
      lookup['fabric'] ||
      lookup['material'] ||
      lookup['fabrictype'] ||
      lookup['garmenttype'] ||
      'Cotton'
  ).trim();

  // 7. EAN (from Wondersoft: EAN code, Barcode, etc.)
  let rawEanVal: unknown =
    lookup['eancode'] ||
    lookup['ean'] ||
    lookup['eanno'] ||
    lookup['eannumber'] ||
    lookup['barcode'] ||
    lookup['barcodeno'] ||
    lookup['itembarcode'] ||
    lookup['eanbarcode'] ||
    lookup['barcodeean'] ||
    lookup['barcodenumber'] ||
    lookup['upc'] ||
    lookup['gtin'] ||
    '';

  // If not found by direct key, search all keys containing ean, barcode, bcode, upc, gtin
  if (!rawEanVal) {
    for (const [k, v] of Object.entries(lookup)) {
      if (
        v !== undefined &&
        v !== null &&
        v !== '' &&
        (k.includes('ean') ||
          k.includes('barcode') ||
          k.includes('bcode') ||
          k.includes('upc') ||
          k.includes('gtin'))
      ) {
        rawEanVal = v;
        break;
      }
    }
  }

  // If still not found, check itemcode / productcode / code / sku if it looks like an EAN barcode (digits length >= 8)
  if (!rawEanVal) {
    for (const [k, v] of Object.entries(lookup)) {
      if (
        v !== undefined &&
        v !== null &&
        (k.includes('itemcode') ||
          k.includes('itemno') ||
          k.includes('prodcode') ||
          k.includes('productcode') ||
          k.includes('code') ||
          k.includes('sku') ||
          k.includes('plu') ||
          k.includes('artcode'))
      ) {
        const cleaned = String(v).trim().replace(/[^0-9]/g, '');
        if (cleaned.length >= 8 && cleaned.length <= 18) {
          rawEanVal = v;
          break;
        }
      }
    }
  }

  const ean = cleanEanString(rawEanVal);

  // 8. Selling Price (from Wondersoft: MRP)
  const rawMrp =
    lookup['mrp'] ||
    lookup['sellingprice'] ||
    lookup['price'] ||
    lookup['rate'] ||
    lookup['salerate'] ||
    lookup['rsp'] ||
    0;
  const sellingPrice = cleanNumber(rawMrp, 0);

  // 9. Toon Label (from Wondersoft: Toon Label / Toon Lable / Tool Label / Style)
  let toonLabel = String(
    lookup['toonlabel'] ||
      lookup['toonlable'] ||
      lookup['toollabel'] ||
      lookup['toollable'] ||
      lookup['toon'] ||
      lookup['toonno'] ||
      lookup['tooncode'] ||
      lookup['style'] ||
      lookup['styleno'] ||
      lookup['stylecode'] ||
      lookup['stylename'] ||
      lookup['design'] ||
      lookup['designno'] ||
      lookup['artno'] ||
      lookup['articleno'] ||
      lookup['subgroup'] ||
      ''
  ).trim();

  // If toonLabel not found, only fallback to itemcode if it wasn't already used as the EAN
  if (!toonLabel) {
    const rawItemCode = String(lookup['itemcode'] || lookup['itemno'] || '').trim();
    if (rawItemCode && rawItemCode !== String(rawEanVal).trim()) {
      toonLabel = rawItemCode;
    }
  }

  // 10. Check if this file has separate store columns (Wide format)
  // e.g. columns named "Salem", "Karur", "KOOTTAPPALLI", "Namakkal", "Kumbakonam", "TIRUVANNAMALAI", "Mallur", "Warehouse"
  const storeColumnsFound = findStoreColumnsInRow(lookup);

  // If the row explicitly has wide store columns (e.g. Salem, Karur, Mallur...)
  if (storeColumnsFound.length >= 2) {
    const totalRowQty = cleanNumber(
      lookup['total'] || lookup['totalqty'] || lookup['totalstock'] || lookup['totalquantity'] || 0,
      storeColumnsFound.reduce((acc, s) => acc + s.qty, 0)
    );

    const breakdownObj: Record<string, number> = {};
    for (const sc of storeColumnsFound) {
      breakdownObj[sc.displayName] = sc.qty;
    }

    // If all are 0 or empty, create a single row for GM Fashions Warehouse with 0
    const nonZeroStores = storeColumnsFound.filter((s) => s.qty > 0);
    const targetStores = nonZeroStores.length > 0 ? nonZeroStores : [{ store: 'GM Fashions Warehouse' as StoreName, displayName: 'GM FASHIONS WAREHOUSE', qty: 0 }];

    return targetStores.map((st, subIdx) => ({
      id: `item-${index}-${subIdx}-${ean || Math.random().toString(36).slice(2, 8)}`,
      department,
      productName,
      brand,
      colour,
      size,
      fabric,
      ean,
      sellingPrice,
      stockQuantity: st.qty,
      totalQuantity: totalRowQty,
      store: st.store,
      toonLabel,
      storeBreakdown: breakdownObj,
    }));
  }

  // Otherwise, standard single-store row (Tall format)
  // Retail branch/location MUST take precedence over warehouse/company
  const branchCandidate = String(
    lookup['branch'] ||
      lookup['branchname'] ||
      lookup['location'] ||
      lookup['outlet'] ||
      lookup['site'] ||
      lookup['shop'] ||
      lookup['counter'] ||
      ''
  ).trim();

  const storeCandidate = String(
    lookup['store'] ||
      lookup['storename'] ||
      lookup['company'] ||
      ''
  ).trim();

  let store: StoreName;
  const matchedBranch = branchCandidate ? tryMatchStoreName(branchCandidate) : null;
  const matchedStore = storeCandidate ? tryMatchStoreName(storeCandidate) : null;

  if (matchedBranch && matchedBranch !== 'GM Fashions Warehouse') {
    store = matchedBranch;
  } else if (matchedStore && matchedStore !== 'GM Fashions Warehouse') {
    store = matchedStore;
  } else if (matchedBranch) {
    store = matchedBranch;
  } else if (matchedStore) {
    store = matchedStore;
  } else {
    store = matchStoreName(branchCandidate || storeCandidate || 'Kootapalli');
  }

  const rawStockVal =
    lookup['stock'] ||
    lookup['stockquantity'] ||
    lookup['stockqty'] ||
    lookup['qty'] ||
    lookup['quantity'] ||
    lookup['closingstock'] ||
    lookup['balqty'] ||
    lookup['balanceqty'] ||
    lookup['currentstock'] ||
    lookup['availableqty'] ||
    0;
  const stockQuantity = cleanNumber(rawStockVal, 0);

  const rawTotalQty = cleanNumber(
    lookup['totalquantity'] || lookup['totalqty'] || lookup['totalstock'] || stockQuantity,
    stockQuantity
  );

  return [
    {
      id: `item-${index}-${ean || Math.random().toString(36).slice(2, 8)}`,
      department,
      productName,
      brand,
      colour,
      size,
      fabric,
      ean,
      sellingPrice,
      stockQuantity,
      totalQuantity: rawTotalQty,
      store,
      toonLabel,
    },
  ];
}

interface StockColumnLayout {
  deptIdx: number;
  prodIdx: number;
  brandIdx: number;
  colourIdx: number;
  sizeIdx: number;
  fabricIdx: number;
  eanIdx: number;
  eanFallbackIndices: number[];
  codeFallbackIndices: number[];
  mrpIdx: number;
  toonIdx: number;
  itemCodeIdx: number;
  branchIdx: number;
  storeIdx: number;
  stockIdx: number;
  totalQtyIdx: number;
  wideStores: { colIndex: number; store: StoreName; displayName: string }[];
}

function buildStockColumnLayout(headers: string[]): StockColumnLayout {
  const normToIdx = new Map<string, number>();
  headers.forEach((h, idx) => {
    const norm = normalizeKey(String(h || ''));
    if (norm && !normToIdx.has(norm)) {
      normToIdx.set(norm, idx);
    }
  });

  const findFirst = (keys: string[]): number => {
    for (const k of keys) {
      const idx = normToIdx.get(k);
      if (idx !== undefined) return idx;
    }
    return -1;
  };

  const eanFallbackIndices: number[] = [];
  const codeFallbackIndices: number[] = [];
  headers.forEach((h, idx) => {
    const k = normalizeKey(String(h || ''));
    if (!k) return;
    if (k.includes('ean') || k.includes('barcode') || k.includes('bcode') || k.includes('upc') || k.includes('gtin')) {
      eanFallbackIndices.push(idx);
    }
    if (
      k.includes('itemcode') ||
      k.includes('itemno') ||
      k.includes('prodcode') ||
      k.includes('productcode') ||
      k.includes('code') ||
      k.includes('sku') ||
      k.includes('plu') ||
      k.includes('artcode')
    ) {
      codeFallbackIndices.push(idx);
    }
  });

  const wideStores: { colIndex: number; store: StoreName; displayName: string }[] = [];
  const seenStoreNorms = new Set<string>();
  headers.forEach((h, idx) => {
    const norm = normalizeKey(String(h || ''));
    if (!norm || seenStoreNorms.has(norm)) return;
    if (
      norm.includes('product') ||
      norm.includes('dept') ||
      norm.includes('class') ||
      norm.includes('brand') ||
      norm.includes('color') ||
      norm.includes('colour') ||
      norm.includes('size') ||
      norm.includes('garment') ||
      norm.includes('fabric') ||
      norm.includes('ean') ||
      norm.includes('barcode') ||
      norm.includes('mrp') ||
      norm.includes('price') ||
      norm.includes('rate') ||
      norm.includes('toon') ||
      norm.includes('style') ||
      norm.includes('image') ||
      norm === 'total' ||
      norm === 'totalqty' ||
      norm === 'totalstock' ||
      norm === 'totalquantity'
    ) {
      return;
    }

    if (
      norm.includes('koot') ||
      norm.includes('kooth') ||
      norm.includes('koott') ||
      norm.includes('kuttap') ||
      norm.includes('kootam') ||
      norm.includes('koottam') ||
      norm.startsWith('ktp') ||
      norm.includes('ktp')
    ) {
      wideStores.push({ colIndex: idx, store: 'Kootapalli', displayName: 'KOOTTAPPALLI' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('kumba') || norm.includes('kumbakonam') || norm.includes('kmb')) {
      wideStores.push({ colIndex: idx, store: 'Kumbakonam', displayName: 'KUMBAKONAM' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('mallur') || norm.includes('mlr')) {
      wideStores.push({ colIndex: idx, store: 'Mallur', displayName: 'MALLUR' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('namakkal') || norm.includes('namakal') || norm.includes('nmkl')) {
      wideStores.push({ colIndex: idx, store: 'Namakkal', displayName: 'NAMAKKAL' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('salem') || norm.includes('slm')) {
      wideStores.push({ colIndex: idx, store: 'Salem', displayName: 'SALEM' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('tiruvan') || norm.includes('thiruvan') || norm.includes('tvm') || norm.includes('thiruvanna')) {
      wideStores.push({ colIndex: idx, store: 'Thiruvannamalai', displayName: 'TIRUVANNAMALAI' });
      seenStoreNorms.add(norm);
    } else if (norm.includes('karur') || norm.includes('krr')) {
      wideStores.push({ colIndex: idx, store: 'Karur', displayName: 'KARUR' });
      seenStoreNorms.add(norm);
    } else if (
      norm.includes('ware') ||
      norm === 'wh' ||
      norm.includes('godown') ||
      norm.includes('central') ||
      norm.includes('gmfashionswarehouse')
    ) {
      wideStores.push({ colIndex: idx, store: 'GM Fashions Warehouse', displayName: 'GM FASHIONS WAREHOUSE' });
      seenStoreNorms.add(norm);
    }
  });

  return {
    deptIdx: findFirst(['productgroup', 'productgrp', 'prodgroup', 'group', 'producttype', 'category', 'pg']),
    prodIdx: findFirst(['department', 'dept', 'productname', 'itemname', 'itemdescription', 'description', 'desc', 'product']),
    brandIdx: findFirst(['class', 'brand', 'brandname', 'subclass', 'categoryclass']),
    colourIdx: findFirst(['color', 'colour', 'shade', 'colorname', 'colourname']),
    sizeIdx: findFirst(['size', 'itemsize', 'sizedesc']),
    fabricIdx: findFirst(['garments', 'garment', 'fabric', 'material', 'fabrictype', 'garmenttype']),
    eanIdx: findFirst(['eancode', 'ean', 'eanno', 'eannumber', 'barcode', 'barcodeno', 'itembarcode', 'eanbarcode', 'barcodeean', 'barcodenumber', 'upc', 'gtin']),
    eanFallbackIndices,
    codeFallbackIndices,
    mrpIdx: findFirst(['mrp', 'sellingprice', 'price', 'rate', 'salerate', 'rsp']),
    toonIdx: findFirst(['toonlabel', 'toonlable', 'toollabel', 'toollable', 'toon', 'toonno', 'tooncode', 'style', 'styleno', 'stylecode', 'stylename', 'design', 'designno', 'artno', 'articleno', 'subgroup']),
    itemCodeIdx: findFirst(['itemcode', 'itemno']),
    branchIdx: findFirst(['branch', 'branchname', 'location', 'outlet', 'site', 'shop', 'counter']),
    storeIdx: findFirst(['store', 'storename', 'company']),
    stockIdx: findFirst(['stock', 'stockquantity', 'stockqty', 'qty', 'quantity', 'closingstock', 'balqty', 'balanceqty', 'currentstock', 'availableqty']),
    totalQtyIdx: findFirst(['totalquantity', 'totalqty', 'totalstock', 'total']),
    wideStores,
  };
}

function mapStockArrayRow(
  row: unknown[],
  layout: StockColumnLayout,
  index: number,
  out: MappedInventoryItem[]
): void {
  if (!row || row.length === 0) return;
  let hasContent = false;
  for (let i = 0; i < row.length; i++) {
    const v = row[i];
    if (v !== '' && v !== null && v !== undefined) {
      hasContent = true;
      break;
    }
  }
  if (!hasContent) return;

  const strAt = (idx: number, fallback: string): string => {
    if (idx < 0 || idx >= row.length) return fallback;
    const v = row[idx];
    if (v === null || v === undefined || v === '') return fallback;
    const s = String(v).trim();
    return s || fallback;
  };

  const department = strAt(layout.deptIdx, 'Apparel');
  const productName = strAt(layout.prodIdx, 'Fashion Wear');
  const brand = strAt(layout.brandIdx, 'GM Fashions');
  const colour = strAt(layout.colourIdx, 'Standard');
  const size = strAt(layout.sizeIdx, 'Free');
  const fabric = strAt(layout.fabricIdx, 'Cotton');

  let rawEanVal: unknown = layout.eanIdx >= 0 ? row[layout.eanIdx] : '';
  if (rawEanVal === undefined || rawEanVal === null || rawEanVal === '') {
    for (const idx of layout.eanFallbackIndices) {
      const v = row[idx];
      if (v !== undefined && v !== null && v !== '') {
        rawEanVal = v;
        break;
      }
    }
  }
  if (rawEanVal === undefined || rawEanVal === null || rawEanVal === '') {
    for (const idx of layout.codeFallbackIndices) {
      const v = row[idx];
      if (v !== undefined && v !== null) {
        const cleaned = String(v).trim().replace(/[^0-9]/g, '');
        if (cleaned.length >= 8 && cleaned.length <= 18) {
          rawEanVal = v;
          break;
        }
      }
    }
  }

  const ean = cleanEanString(rawEanVal);
  const sellingPrice = layout.mrpIdx >= 0 ? cleanNumber(row[layout.mrpIdx], 0) : 0;

  let toonLabel = strAt(layout.toonIdx, '');
  if (!toonLabel && layout.itemCodeIdx >= 0) {
    const rawItemCode = strAt(layout.itemCodeIdx, '');
    if (rawItemCode && rawItemCode !== String(rawEanVal || '').trim()) {
      toonLabel = rawItemCode;
    }
  }

  if (layout.wideStores.length >= 2) {
    let sumQty = 0;
    const breakdownObj: Record<string, number> = {};
    for (const sc of layout.wideStores) {
      const q = cleanNumber(row[sc.colIndex], 0);
      breakdownObj[sc.displayName] = q;
      sumQty += q;
    }
    const totalRowQty =
      layout.totalQtyIdx >= 0 ? cleanNumber(row[layout.totalQtyIdx], sumQty) : sumQty;

    let pushedAny = false;
    for (let sIdx = 0; sIdx < layout.wideStores.length; sIdx++) {
      const sc = layout.wideStores[sIdx];
      const q = breakdownObj[sc.displayName];
      if (q > 0) {
        pushedAny = true;
        out.push({
          id: `item-${index}-${sIdx}-${ean || index}`,
          department,
          productName,
          brand,
          colour,
          size,
          fabric,
          ean,
          sellingPrice,
          stockQuantity: q,
          totalQuantity: totalRowQty,
          store: sc.store,
          toonLabel,
          storeBreakdown: breakdownObj,
        });
      }
    }
    if (!pushedAny) {
      out.push({
        id: `item-${index}-0-${ean || index}`,
        department,
        productName,
        brand,
        colour,
        size,
        fabric,
        ean,
        sellingPrice,
        stockQuantity: 0,
        totalQuantity: totalRowQty,
        store: 'GM Fashions Warehouse',
        toonLabel,
        storeBreakdown: breakdownObj,
      });
    }
    return;
  }

  const branchCandidate = strAt(layout.branchIdx, '');
  const storeCandidate = strAt(layout.storeIdx, '');
  let store: StoreName;
  const matchedBranch = branchCandidate ? tryMatchStoreName(branchCandidate) : null;
  const matchedStore = storeCandidate ? tryMatchStoreName(storeCandidate) : null;

  if (matchedBranch && matchedBranch !== 'GM Fashions Warehouse') {
    store = matchedBranch;
  } else if (matchedStore && matchedStore !== 'GM Fashions Warehouse') {
    store = matchedStore;
  } else if (matchedBranch) {
    store = matchedBranch;
  } else if (matchedStore) {
    store = matchedStore;
  } else {
    store = matchStoreName(branchCandidate || storeCandidate || 'Kootapalli');
  }

  const stockQuantity = layout.stockIdx >= 0 ? cleanNumber(row[layout.stockIdx], 0) : 0;
  const rawTotalQty =
    layout.totalQtyIdx >= 0 ? cleanNumber(row[layout.totalQtyIdx], stockQuantity) : stockQuantity;

  out.push({
    id: `item-${index}-${ean || index}`,
    department,
    productName,
    brand,
    colour,
    size,
    fabric,
    ean,
    sellingPrice,
    stockQuantity,
    totalQuantity: rawTotalQty,
    store,
    toonLabel,
  });
}

/**
 * High-capacity parser supporting Stock Reports up to 1 GB (.csv streamed in 4 MB chunks, .xlsx/.xls via dense SheetJS mode).
 */
export async function parseInventoryFile(
  file: File,
  onProgress?: (progress: ParseProgress) => void
): Promise<MappedInventoryItem[]> {
  if (file.size > MAX_FILE_UPLOAD_BYTES) {
    throw new Error(
      `File size (${formatByteSize(file.size)}) exceeds the 1 GB (${formatByteSize(MAX_FILE_UPLOAD_BYTES)}) maximum limit.`
    );
  }

  const lowerName = file.name.toLowerCase();

  // Streaming chunked path for CSV files (supports up to 1 GB with minimal memory footprint)
  if (lowerName.endsWith('.csv')) {
    const allItems: MappedInventoryItem[] = [];
    const headerScanBuffer: string[][] = [];
    let layout: StockColumnLayout | null = null;
    let rowCounter = 0;

    await streamCsvFileInBatches(
      file,
      (batch) => {
        if (!layout) {
          for (let i = 0; i < batch.length; i++) {
            headerScanBuffer.push(batch[i]);
          }
          if (headerScanBuffer.length >= 2) {
            const headerIdx = findHeaderRowIndexFromRows(headerScanBuffer);
            const headers = (headerScanBuffer[headerIdx] || []).map((h) => String(h || '').trim());
            layout = buildStockColumnLayout(headers);
            for (let r = headerIdx + 1; r < headerScanBuffer.length; r++) {
              mapStockArrayRow(headerScanBuffer[r], layout, rowCounter++, allItems);
            }
            headerScanBuffer.length = 0;
          }
        } else {
          for (let i = 0; i < batch.length; i++) {
            mapStockArrayRow(batch[i], layout, rowCounter++, allItems);
          }
        }
      },
      onProgress
    );

    if (allItems.length === 0) {
      throw new Error('The CSV file is empty. No valid inventory rows found.');
    }

    if (onProgress) {
      onProgress({
        percent: 96,
        loadedBytes: file.size,
        totalBytes: file.size,
        rowsParsed: allItems.length,
        stage: 'Indexing store quantities...',
      });
    }
    await yieldToMainThread();

    const result = calculateTotalQuantities(allItems);
    if (onProgress) {
      onProgress({
        percent: 100,
        loadedBytes: file.size,
        totalBytes: file.size,
        rowsParsed: result.length,
        stage: 'Completed',
      });
    }
    return result;
  }

  // Excel (.xlsx / .xls) path with dense memory-saving SheetJS options + batched processing
  if (onProgress) {
    onProgress({
      percent: 15,
      loadedBytes: 0,
      totalBytes: file.size,
      rowsParsed: 0,
      stage: 'Reading workbook into memory...',
    });
  }

  const buffer = await file.arrayBuffer();
  const result = await parseExcelFile(buffer, (p) => {
    if (onProgress) {
      onProgress({ ...p, loadedBytes: file.size, totalBytes: file.size });
    }
  });
  return result;
}

/**
 * Parses an Excel or CSV file (ArrayBuffer) into MappedInventoryItem[] using dense low-memory SheetJS mode.
 */
export async function parseExcelFile(
  fileBuffer: ArrayBuffer,
  onProgress?: (progress: ParseProgress) => void
): Promise<MappedInventoryItem[]> {
  const workbook = XLSX.read(fileBuffer, {
    type: 'array',
    dense: true,
    cellDates: true,
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellStyles: false,
    cellText: false,
  });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('The uploaded Excel file contains no worksheets.');
  }

  // Pick the largest sheet
  let bestSheetName = '';
  let bestRowEstimate = -1;
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const est = range.e.r - range.s.r;
    if (est > bestRowEstimate) {
      bestRowEstimate = est;
      bestSheetName = sheetName;
    } else {
      delete workbook.Sheets[sheetName];
    }
  }

  if (!bestSheetName) {
    throw new Error('The worksheet is empty. No inventory rows found.');
  }

  const targetWs = workbook.Sheets[bestSheetName];
  const headerIdx = findHeaderRowIndex(targetWs);
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(targetWs, {
    header: 1,
    range: headerIdx,
    defval: '',
  });
  // Release SheetJS worksheet structure immediately for GC
  delete workbook.Sheets[bestSheetName];

  if (!rawRows || rawRows.length < 2) {
    throw new Error('The worksheet is empty. No inventory rows found.');
  }

  const headers = (rawRows[0] || []).map((h) => String(h || '').trim());
  const layout = buildStockColumnLayout(headers);
  const allItems: MappedInventoryItem[] = [];
  const totalDataRows = rawRows.length - 1;
  const BATCH_SIZE = 10000;

  for (let r = 1; r < rawRows.length; r++) {
    mapStockArrayRow(rawRows[r], layout, r - 1, allItems);
    if (r % BATCH_SIZE === 0) {
      if (onProgress) {
        onProgress({
          percent: 30 + Math.round((r / totalDataRows) * 62),
          loadedBytes: fileBuffer.byteLength,
          totalBytes: fileBuffer.byteLength,
          rowsParsed: allItems.length,
          stage: `Mapping records (${allItems.length.toLocaleString()})...`,
        });
      }
      await yieldToMainThread();
    }
  }

  if (allItems.length === 0) {
    throw new Error('The worksheet is empty. No inventory rows found.');
  }

  if (onProgress) {
    onProgress({
      percent: 95,
      loadedBytes: fileBuffer.byteLength,
      totalBytes: fileBuffer.byteLength,
      rowsParsed: allItems.length,
      stage: 'Indexing store quantities...',
    });
  }
  await yieldToMainThread();

  return calculateTotalQuantities(allItems);
}

/**
 * Load default GM Fashions sample data
 */
export function loadSampleInventory(): MappedInventoryItem[] {
  const items: MappedInventoryItem[] = [];
  SAMPLE_WONDERSOFT_DATA.forEach((row, idx) => {
    items.push(...mapWondersoftRow(row as unknown as Record<string, unknown>, idx));
  });
  return calculateTotalQuantities(items);
}

/**
 * Search Logic:
 * - EAN Code: Matches against the clean EAN/barcode string with bi-directional and leading-zero tolerance.
 * - Toon Label: Partial and exact matching (handles case, spaces, and hyphens flexibly).
 * - Store selection: Filters by selected stores
 */
export function filterInventory(
  items: MappedInventoryItem[],
  type: 'EAN' | 'TOON',
  query: string,
  selectedStores: StoreName[],
  isAllStores: boolean
): MappedInventoryItem[] {
  const cleanQuery = query.trim();
  if (!cleanQuery) return [];

  // Allowed stores: when 'ALL' stores is selected, only the 8 stores in ALL_STORES are returned.
  // Any unwanted/extraneous stores are completely excluded.
  // When specific stores are selected, only those chosen stores (strictly within ALL_STORES) are returned.
  const allowedStoresSet = new Set<StoreName>(
    isAllStores ? ALL_STORES : selectedStores.filter((s) => ALL_STORES.includes(s))
  );

  const filtered = items.filter((item) => {
    // 1. Store filtering: strictly allow ONLY valid configured stores from the dropdown
    if (!allowedStoresSet.has(item.store)) {
      return false;
    }

    // 2. Query filtering
    if (type === 'EAN') {
      const rawItemEan = String(item.ean || '').trim();
      const rawQueryEan = String(cleanQuery || '').trim();

      if (!rawItemEan || !rawQueryEan) return false;

      const cleanItemEan = cleanEanString(rawItemEan).replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
      const cleanTargetEan = cleanEanString(rawQueryEan).replace(/[^0-9a-zA-Z]/g, '').toLowerCase();

      if (!cleanItemEan || !cleanTargetEan) return false;

      // 1. Direct match
      if (cleanItemEan === cleanTargetEan) return true;

      // 2. Strip leading zeros match (e.g. 08907701122334 vs 8907701122334)
      const noZeroItem = cleanItemEan.replace(/^0+/, '');
      const noZeroTarget = cleanTargetEan.replace(/^0+/, '');
      if (noZeroItem && noZeroTarget && noZeroItem === noZeroTarget) {
        return true;
      }

      // 3. Bi-directional substring match (minimum 3 characters)
      if (cleanTargetEan.length >= 3 && cleanItemEan.includes(cleanTargetEan)) {
        return true;
      }
      if (cleanItemEan.length >= 3 && cleanTargetEan.includes(cleanItemEan)) {
        return true;
      }

      return false;
    } else {
      // Toon Label: partial and exact matching (case-insensitive, ignoring hyphens/spaces)
      if (!item.toonLabel) return false;

      const lowerItemToon = item.toonLabel.toLowerCase();
      const lowerQuery = cleanQuery.toLowerCase();

      if (lowerItemToon.includes(lowerQuery)) return true;

      const strippedItem = lowerItemToon.replace(/[^a-z0-9]/g, '');
      const strippedQuery = lowerQuery.replace(/[^a-z0-9]/g, '');

      return strippedItem.length > 0 && strippedItem.includes(strippedQuery);
    }
  });

  // When searching by Toon Label (including partial match):
  // 1st Preference: Toon Label (Ascending order)
  // 2nd Preference: Colour (Ascending order)
  // 3rd Preference: Size (Ascending order e.g. 32, 34, 36, 38 or S, M, L, XL, 2XL)
  if (type === 'TOON') {
    filtered.sort((a, b) => compareToonColourAndSize(a, b));
  }

  return filtered;
}

/**
 * Standard garment size ordering priority
 * S, M, L, XL, 2XL, 3XL... or numeric 28, 30, 32, 34, 36...
 */
const STANDARD_SIZE_RANK: Record<string, number> = {
  'fs': 1,
  'free': 1,
  'free size': 1,
  'freesize': 1,
  'one size': 1,
  'os': 1,
  'xxs': 5,
  '2xs': 5,
  'xs': 10,
  'extra small': 10,
  's': 20,
  'small': 20,
  'm': 30,
  'med': 30,
  'medium': 30,
  'l': 40,
  'large': 40,
  'xl': 50,
  'x-large': 50,
  'extra large': 50,
  'xxl': 60,
  '2xl': 60,
  'xxxl': 70,
  '3xl': 70,
  'xxxxl': 80,
  '4xl': 80,
  '5xl': 90,
  '6xl': 100,
};

export function compareSizes(sizeA: string, sizeB: string): number {
  const rawA = (sizeA || '').trim().toLowerCase();
  const rawB = (sizeB || '').trim().toLowerCase();

  if (rawA === rawB) return 0;

  // 1. Direct number check (e.g. 28, 30, 32, 34, 36, 38, 40, 42, 44 or "32cm", "34 inch")
  const numMatchA = rawA.match(/^\s*(\d+(\.\d+)?)/);
  const numMatchB = rawB.match(/^\s*(\d+(\.\d+)?)/);
  if (numMatchA && numMatchB) {
    const numA = parseFloat(numMatchA[1]);
    const numB = parseFloat(numMatchB[1]);
    if (numA !== numB) return numA - numB;
    return rawA.localeCompare(rawB, undefined, { numeric: true });
  }

  // 2. Standard garment alpha size check (XS, S, M, L, XL, XXL / 2XL, etc.)
  const cleanA = rawA.replace(/[^a-z0-9]/g, '');
  const cleanB = rawB.replace(/[^a-z0-9]/g, '');

  const rankA = STANDARD_SIZE_RANK[rawA] ?? STANDARD_SIZE_RANK[cleanA];
  const rankB = STANDARD_SIZE_RANK[rawB] ?? STANDARD_SIZE_RANK[cleanB];

  if (rankA !== undefined && rankB !== undefined) {
    return rankA - rankB;
  }
  if (rankA !== undefined) return -1;
  if (rankB !== undefined) return 1;

  // 3. Embedded numbers if any
  const matchA = rawA.match(/\d+(\.\d+)?/);
  const matchB = rawB.match(/\d+(\.\d+)?/);
  const nA = matchA ? parseFloat(matchA[0]) : NaN;
  const nB = matchB ? parseFloat(matchB[0]) : NaN;

  if (!isNaN(nA) && !isNaN(nB)) {
    if (nA !== nB) return nA - nB;
    return rawA.localeCompare(rawB, undefined, { numeric: true });
  }
  if (!isNaN(nA)) return -1;
  if (!isNaN(nB)) return 1;

  return rawA.localeCompare(rawB, undefined, { numeric: true });
}

/**
 * Strict 3-level ordering for Toon Label queries:
 * 1st preference: Toon Label in Ascending Order (e.g., natural alphanumeric A-Z, 0-9)
 * 2nd preference: Colour in Ascending Order (e.g., Blue, Green, Red...)
 * 3rd preference: Size in Ascending Order (e.g., 32, 34, 36, 38 or S, M, L, XL, 2XL)
 */
export function compareToonColourAndSize(
  a: { toonLabel?: string; colour: string; size: string },
  b: { toonLabel?: string; colour: string; size: string }
): number {
  // 1st Preference: Toon Label (Ascending order with natural numeric collation)
  const toonA = (a.toonLabel || '').trim();
  const toonB = (b.toonLabel || '').trim();
  const toonDiff = toonA.localeCompare(toonB, undefined, { numeric: true, sensitivity: 'base' });
  if (toonDiff !== 0) {
    return toonDiff;
  }

  // 2nd Preference: Colour (Ascending order)
  const colA = (a.colour || '').trim().toLowerCase();
  const colB = (b.colour || '').trim().toLowerCase();
  const colorDiff = colA.localeCompare(colB, undefined, { numeric: true, sensitivity: 'base' });
  if (colorDiff !== 0) {
    return colorDiff;
  }

  // 3rd Preference: Size (Ascending order, e.g. 32, 34, 36, 38 or S, M, L, XL, 2XL)
  return compareSizes(a.size, b.size);
}

/**
 * Sort primarily by Colour (alphabetical: Blue, Green, Red...),
 * and secondarily by Size (S, M, L, XL, 2XL...) within each colour.
 */
export function compareColourAndSize(
  a: { colour: string; size: string },
  b: { colour: string; size: string }
): number {
  const colA = (a.colour || '').trim().toLowerCase();
  const colB = (b.colour || '').trim().toLowerCase();

  const colorDiff = colA.localeCompare(colB);
  if (colorDiff !== 0) {
    return colorDiff;
  }

  return compareSizes(a.size, b.size);
}

/**
 * Generates the strictly formatted Toon Label Matrix Excel file
 * Required columns:
 * [Color, Toon Label, Size, Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, Size Total]
 */
export function generateToonLabelMatrix(
  items: MappedInventoryItem[]
): ToonMatrixRow[] {
  // Aggregate by [colour, toonLabel, size]
  const matrixMap = new Map<string, ToonMatrixRow>();

  for (const item of items) {
    const key = `${item.colour}__${item.toonLabel}__${item.size}`;
    let row = matrixMap.get(key);

    if (!row) {
      row = {
        colour: item.colour,
        toonLabel: item.toonLabel,
        size: item.size,
        kootapalli: 0,
        karur: 0,
        salem: 0,
        namakkal: 0,
        kumbakonam: 0,
        thiruvannamalai: 0,
        mallur: 0,
        gmFashionsWarehouse: 0,
        sizeTotal: 0,
      };
      matrixMap.set(key, row);
    }

    const qty = Number(item.stockQuantity) || 0;
    switch (item.store) {
      case 'Kootapalli':
        row.kootapalli += qty;
        break;
      case 'Karur':
        row.karur += qty;
        break;
      case 'Salem':
        row.salem += qty;
        break;
      case 'Namakkal':
        row.namakkal += qty;
        break;
      case 'Kumbakonam':
        row.kumbakonam += qty;
        break;
      case 'Thiruvannamalai':
        row.thiruvannamalai += qty;
        break;
      case 'Mallur':
        row.mallur += qty;
        break;
      case 'GM Fashions Warehouse':
        row.gmFashionsWarehouse += qty;
        break;
      default:
        break;
    }

    row.sizeTotal =
      row.kootapalli +
      row.karur +
      row.salem +
      row.namakkal +
      row.kumbakonam +
      row.thiruvannamalai +
      row.mallur +
      row.gmFashionsWarehouse;
  }

  const rows = Array.from(matrixMap.values());
  rows.sort((a, b) => compareToonColourAndSize(a, b));
  return rows;
}

/**
 * Downloads Excel file for Toon Label search with strict Matrix format:
 * [Color, Toon Label, Size, Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, GM Fashions Warehouse, Size Total]
 */
export function exportToonLabelExcel(
  items: MappedInventoryItem[],
  searchQuery: string
): string {
  const matrixRows = generateToonLabelMatrix(items);

  // Group consecutive rows with the same (toonLabel, colour)
  const groups: MatrixColorGroup[] = [];
  let currentGroup: MatrixColorGroup | null = null;

  matrixRows.forEach((row, idx) => {
    const key = `${(row.toonLabel || '').trim()}___${(row.colour || '').trim()}`;
    const prevKey = currentGroup
      ? `${currentGroup.toonLabel.trim()}___${currentGroup.colour.trim()}`
      : '';

    if (!currentGroup || key !== prevKey) {
      currentGroup = {
        colour: row.colour,
        toonLabel: row.toonLabel,
        startIndex: idx,
        endIndex: idx,
        count: 1,
      };
      groups.push(currentGroup);
    } else {
      currentGroup.endIndex = idx;
      currentGroup.count++;
    }
  });

  // Group start indices set for quick lookup
  const groupStartIndices = new Set(groups.map((g) => g.startIndex));

  // If a color repeats across multiple sizes (e.g. Blue across 10 sizes),
  // show the color name once on the first row of that group, and merge cells in Excel
  const exportData = matrixRows.map((row, idx) => ({
    'Color': groupStartIndices.has(idx) ? row.colour : '',
    'Toon Label': groupStartIndices.has(idx) ? row.toonLabel : '',
    'Size': row.size,
    'Kootapalli': row.kootapalli,
    'Karur': row.karur,
    'Salem': row.salem,
    'Namakkal': row.namakkal,
    'Kumbakonam': row.kumbakonam,
    'Thiruvannamalai': row.thiruvannamalai,
    'Mallur': row.mallur,
    'GM Fashions Warehouse': row.gmFashionsWarehouse,
    'Size Total': row.sizeTotal,
  }));

  if (exportData.length > 0) {
    const totals = exportData.reduce(
      (acc, r) => {
        acc.Kootapalli += r.Kootapalli;
        acc.Karur += r.Karur;
        acc.Salem += r.Salem;
        acc.Namakkal += r.Namakkal;
        acc.Kumbakonam += r.Kumbakonam;
        acc.Thiruvannamalai += r.Thiruvannamalai;
        acc.Mallur += r.Mallur;
        acc.Warehouse += r['GM Fashions Warehouse'];
        acc.SizeTotal += r['Size Total'];
        return acc;
      },
      {
        Kootapalli: 0,
        Karur: 0,
        Salem: 0,
        Namakkal: 0,
        Kumbakonam: 0,
        Thiruvannamalai: 0,
        Mallur: 0,
        Warehouse: 0,
        SizeTotal: 0,
      }
    );

    exportData.push({
      'Color': 'TOTAL',
      'Toon Label': '-',
      'Size': '-',
      'Kootapalli': totals.Kootapalli,
      'Karur': totals.Karur,
      'Salem': totals.Salem,
      'Namakkal': totals.Namakkal,
      'Kumbakonam': totals.Kumbakonam,
      'Thiruvannamalai': totals.Thiruvannamalai,
      'Mallur': totals.Mallur,
      'GM Fashions Warehouse': totals.Warehouse,
      'Size Total': totals.SizeTotal,
    });
  }

  const headers = [
    'Color',
    'Toon Label',
    'Size',
    'Kootapalli',
    'Karur',
    'Salem',
    'Namakkal',
    'Kumbakonam',
    'Thiruvannamalai',
    'Mallur',
    'GM Fashions Warehouse',
    'Size Total',
  ];

  const worksheet = XLSX.utils.json_to_sheet(exportData, { header: headers });
  worksheet['!cols'] = calculateColumnWidths(headers, exportData);
  styleMatrixWorksheet(worksheet, headers, exportData.length, 'stock', groups);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Toon Label Stock Matrix');

  const safeQuery = searchQuery.replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = `GM_Fashion_Toon_${safeQuery || 'Results'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(workbook, filename);
  return filename;
}

function formatShuffledCellForExcel(info: StoreCellShuffleInfo): string | number {
  const lines: string[] = [];

  if (info.borrowedDetails && info.borrowedDetails.length > 0) {
    for (const b of info.borrowedDetails) {
      lines.push(`${info.originalQty} [+${b.qty}] ${b.sourceShortcut}`);
    }
  }

  if (info.donatedToStores && info.donatedToStores.length > 0) {
    for (const d of info.donatedToStores) {
      lines.push(`${info.originalQty} [-${d.qty}] → ${d.shortcut}`);
    }
  }

  if (lines.length > 0) {
    return lines.join('\n');
  }

  return info.shuffledQty;
}

const MATRIX_STORE_ORDER: MatrixStoreKey[] = [
  'kootapalli',
  'karur',
  'salem',
  'namakkal',
  'kumbakonam',
  'thiruvannamalai',
  'mallur',
  'gmFashionsWarehouse',
];

function formatRowTransferSummary(row: ShuffledMatrixRow): string {
  const movements: string[] = [];
  for (const storeKey of MATRIX_STORE_ORDER) {
    const info = row.cellInfo[storeKey];
    if (info && info.borrowedDetails) {
      for (const b of info.borrowedDetails) {
        movements.push(`${b.sourceShortcut} → ${STORE_SHORTCUTS[storeKey]} (+${b.qty})`);
      }
    }
  }
  return movements.length > 0 ? movements.join(', ') : '—';
}

/**
 * Downloads a 2-sheet styled Excel workbook for Auto Shuffle or Manual Shuffle:
 * 1. "Shuffled Action Matrix": exact UI lent/bought details (e.g. "0 [+1] KR", "3 [-1] -> SLM") + Transfer Movements column
 * 2. "Final Prediction Matrix": clean post-shuffle numeric target inventory per store
 */
export function exportShuffledMatrixExcel(
  shuffledRows: ShuffledMatrixRow[],
  searchQuery: string,
  shuffleMode: 'auto' | 'manual'
): string {
  // Build color groups matching UI row merges
  const groups: MatrixColorGroup[] = [];
  let currentGroup: MatrixColorGroup | null = null;
  const groupStartIndices = new Set<number>();

  shuffledRows.forEach((row, idx) => {
    const key = `${(row.toonLabel || '').trim()}___${(row.colour || '').trim()}`;
    const prevKey = currentGroup
      ? `${currentGroup.toonLabel.trim()}___${currentGroup.colour.trim()}`
      : '';

    if (!currentGroup || key !== prevKey) {
      currentGroup = {
        colour: row.colour,
        toonLabel: row.toonLabel,
        startIndex: idx,
        endIndex: idx,
        count: 1,
      };
      groups.push(currentGroup);
      groupStartIndices.add(idx);
    } else {
      currentGroup.endIndex = idx;
      currentGroup.count++;
    }
  });

  // Sheet 1: Shuffled Action Matrix (UI lent / bought details)
  const actionHeaders = [
    'Color',
    'Toon Label',
    'Size',
    'Kootapalli',
    'Karur',
    'Salem',
    'Namakkal',
    'Kumbakonam',
    'Thiruvannamalai',
    'Mallur',
    'GM Fashions Warehouse',
    'Size Total',
    'Transfer Movements',
  ];

  const actionData: Record<string, string | number>[] = shuffledRows.map((row, idx) => ({
    'Color': groupStartIndices.has(idx) ? row.colour : '',
    'Toon Label': groupStartIndices.has(idx) ? row.toonLabel : '',
    'Size': row.size,
    'Kootapalli': formatShuffledCellForExcel(row.cellInfo.kootapalli),
    'Karur': formatShuffledCellForExcel(row.cellInfo.karur),
    'Salem': formatShuffledCellForExcel(row.cellInfo.salem),
    'Namakkal': formatShuffledCellForExcel(row.cellInfo.namakkal),
    'Kumbakonam': formatShuffledCellForExcel(row.cellInfo.kumbakonam),
    'Thiruvannamalai': formatShuffledCellForExcel(row.cellInfo.thiruvannamalai),
    'Mallur': formatShuffledCellForExcel(row.cellInfo.mallur),
    'GM Fashions Warehouse': formatShuffledCellForExcel(row.cellInfo.gmFashionsWarehouse),
    'Size Total': row.sizeTotal,
    'Transfer Movements': formatRowTransferSummary(row),
  }));

  const totals = shuffledRows.reduce(
    (acc, r) => {
      acc.Kootapalli += r.kootapalli;
      acc.Karur += r.karur;
      acc.Salem += r.salem;
      acc.Namakkal += r.namakkal;
      acc.Kumbakonam += r.kumbakonam;
      acc.Thiruvannamalai += r.thiruvannamalai;
      acc.Mallur += r.mallur;
      acc.Warehouse += r.gmFashionsWarehouse;
      acc.SizeTotal += r.sizeTotal;
      acc.MovedUnits += r.totalRedirectedInRow;
      return acc;
    },
    {
      Kootapalli: 0,
      Karur: 0,
      Salem: 0,
      Namakkal: 0,
      Kumbakonam: 0,
      Thiruvannamalai: 0,
      Mallur: 0,
      Warehouse: 0,
      SizeTotal: 0,
      MovedUnits: 0,
    }
  );

  if (actionData.length > 0) {
    actionData.push({
      'Color': 'SHUFFLED TOTAL',
      'Toon Label': '-',
      'Size': '-',
      'Kootapalli': totals.Kootapalli,
      'Karur': totals.Karur,
      'Salem': totals.Salem,
      'Namakkal': totals.Namakkal,
      'Kumbakonam': totals.Kumbakonam,
      'Thiruvannamalai': totals.Thiruvannamalai,
      'Mallur': totals.Mallur,
      'GM Fashions Warehouse': totals.Warehouse,
      'Size Total': totals.SizeTotal,
      'Transfer Movements': `${totals.MovedUnits} units moved`,
    });
  }

  const actionWorksheet = XLSX.utils.json_to_sheet(actionData, { header: actionHeaders });
  actionWorksheet['!cols'] = calculateColumnWidths(actionHeaders, actionData);
  styleShuffledMatrixWorksheet(actionWorksheet, actionHeaders, actionData.length, groups);

  // Sheet 2: Final Prediction Matrix (Clean Target Stocks)
  const predHeaders = [
    'Color',
    'Toon Label',
    'Size',
    'Kootapalli',
    'Karur',
    'Salem',
    'Namakkal',
    'Kumbakonam',
    'Thiruvannamalai',
    'Mallur',
    'GM Fashions Warehouse',
    'Size Total',
  ];

  const predData: Record<string, string | number>[] = shuffledRows.map((row, idx) => ({
    'Color': groupStartIndices.has(idx) ? row.colour : '',
    'Toon Label': groupStartIndices.has(idx) ? row.toonLabel : '',
    'Size': row.size,
    'Kootapalli': row.kootapalli,
    'Karur': row.karur,
    'Salem': row.salem,
    'Namakkal': row.namakkal,
    'Kumbakonam': row.kumbakonam,
    'Thiruvannamalai': row.thiruvannamalai,
    'Mallur': row.mallur,
    'GM Fashions Warehouse': row.gmFashionsWarehouse,
    'Size Total': row.sizeTotal,
  }));

  if (predData.length > 0) {
    predData.push({
      'Color': 'PREDICTED TOTAL',
      'Toon Label': '-',
      'Size': '-',
      'Kootapalli': totals.Kootapalli,
      'Karur': totals.Karur,
      'Salem': totals.Salem,
      'Namakkal': totals.Namakkal,
      'Kumbakonam': totals.Kumbakonam,
      'Thiruvannamalai': totals.Thiruvannamalai,
      'Mallur': totals.Mallur,
      'GM Fashions Warehouse': totals.Warehouse,
      'Size Total': totals.SizeTotal,
    });
  }

  const predWorksheet = XLSX.utils.json_to_sheet(predData, { header: predHeaders });
  predWorksheet['!cols'] = calculateColumnWidths(predHeaders, predData);
  styleMatrixWorksheet(predWorksheet, predHeaders, predData.length, 'stock', groups);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    actionWorksheet,
    shuffleMode === 'manual' ? 'Manual Shuffle Actions' : 'Shuffled Action Matrix'
  );
  XLSX.utils.book_append_sheet(workbook, predWorksheet, 'Final Prediction Matrix');

  const safeQuery = searchQuery.replace(/[^a-zA-Z0-9_-]/g, '_');
  const modeTag = shuffleMode === 'manual' ? 'Manual_Shuffle' : 'Auto_Shuffle';
  const filename = `GM_Fashion_${modeTag}_${safeQuery || 'Results'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(workbook, filename);
  return filename;
}

/**
 * Downloads Excel file with custom columns requested by user:
 * [Product Type, Department, Class, Color, Size, Fabric, EAN Code, Toon Label, Price, Total Store Qty, Total Stock Unit, Available Stores]
 */
export function exportEanExcel(
  items: MappedInventoryItem[],
  eanQuery: string
): string {
  // Consolidate into single common rows per unique product/EAN
  const distinctMap = new Map<string, MappedInventoryItem>();
  for (const item of items) {
    const key = item.ean
      ? item.ean.trim()
      : `${item.department}|${item.productName}|${item.colour}|${item.size}|${item.toonLabel}`.toLowerCase();
    if (!distinctMap.has(key)) {
      distinctMap.set(key, item);
    }
  }

  const distinctList = Array.from(distinctMap.values());
  distinctList.sort((a, b) => compareToonColourAndSize(a, b));

  const exportData = distinctList.map((item) => {
    // Collect available stores (qty > 0)
    const storeMap = new Map<string, number>();
    if (item.storeBreakdown && Object.keys(item.storeBreakdown).length > 0) {
      for (const [st, q] of Object.entries(item.storeBreakdown)) {
        if (q > 0) {
          storeMap.set(formatStoreDisplayName(st), q);
        }
      }
    } else {
      const targetEan = item.ean ? item.ean.trim() : '';
      const matching = items.filter((i) => {
        if (targetEan && i.ean) return i.ean.trim() === targetEan;
        return (
          i.department === item.department &&
          i.productName === item.productName &&
          i.colour === item.colour &&
          i.size === item.size
        );
      });
      for (const m of matching) {
        if (m.stockQuantity > 0) {
          const sDisplay = formatStoreDisplayName(m.store);
          storeMap.set(sDisplay, (storeMap.get(sDisplay) || 0) + m.stockQuantity);
        }
      }
    }

    const available = Array.from(storeMap.entries()).filter(([_, q]) => q > 0);
    const totalUnits = available.reduce((sum, [_, q]) => sum + q, 0);
    const availableStoresSummary = available
      .map(([s, q]) => `${s}: ${q}`)
      .join(', ');

    return {
      'Product Type': item.department,
      'Department': item.productName,
      'Class': item.brand,
      'Color': item.colour,
      'Size': item.size,
      'Fabric': item.fabric,
      'EAN Code': item.ean,
      'Toon Label': item.toonLabel || '—',
      'Price': item.sellingPrice,
      'Total Store Qty': available.length,
      'Total Stock Unit': totalUnits > 0 ? totalUnits : (item.totalQuantity ?? item.stockQuantity),
      'Available Stores': availableStoresSummary || 'None',
    };
  });

  const headers = [
    'Product Type',
    'Department',
    'Class',
    'Color',
    'Size',
    'Fabric',
    'EAN Code',
    'Toon Label',
    'Price',
    'Total Store Qty',
    'Total Stock Unit',
    'Available Stores',
  ];

  const worksheet = XLSX.utils.json_to_sheet(exportData, { header: headers });
  worksheet['!cols'] = calculateColumnWidths(headers, exportData);
  // Column index 10 is 'Total Stock Unit' which gets highlighted
  styleDetailedWorksheet(worksheet, headers, exportData.length, 'stock', 10);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Inventory Stock Details');

  const safeEan = eanQuery.replace(/[^0-9a-zA-Z_-]/g, '_');
  const filename = `GM_Fashion_EAN_${safeEan || 'Results'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(workbook, filename);
  return filename;
}

/**
 * Downloads the sample Wondersoft Excel template
 */
export function downloadSampleWondersoftExcel(): string {
  const worksheet = XLSX.utils.json_to_sheet(SAMPLE_WONDERSOFT_DATA);
  const headers = Object.keys(SAMPLE_WONDERSOFT_DATA[0] || {});
  worksheet['!cols'] = calculateColumnWidths(headers, SAMPLE_WONDERSOFT_DATA as any);
  styleDetailedWorksheet(worksheet, headers, SAMPLE_WONDERSOFT_DATA.length, 'stock');

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Wondersoft_Export');

  const filename = 'GM_Fashion_Wondersoft_Sample_Export.xlsx';
  XLSX.writeFile(workbook, filename);
  return filename;
}
