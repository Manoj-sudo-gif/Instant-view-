import {
  XLSX,
  styleMatrixWorksheet,
  styleDetailedWorksheet,
  calculateColumnWidths,
  MatrixColorGroup,
} from './excelStyler';
import {
  MappedSalesItem,
  SalesMatrixRow,
  StoreName,
  ALL_STORES,
  UploadedFileInfo,
} from '../types/inventory';
import {
  cleanEanString,
  matchStoreName,
  tryMatchStoreName,
  formatStoreDisplayName,
  findHeaderRowIndex,
  findHeaderRowIndexFromRows,
  compareToonColourAndSize,
} from './excelParser';
import { MAX_FILE_UPLOAD_BYTES, formatByteSize } from './storage';
import {
  ParseProgress,
  streamCsvFileInBatches,
  yieldToMainThread,
} from './streamingCsvParser';

function cleanNumber(val: unknown, defaultVal = 0): number {
  if (val === null || val === undefined || val === '') return defaultVal;
  if (typeof val === 'number') return isNaN(val) ? defaultVal : val;
  const str = String(val).trim();
  if (!str) return defaultVal;
  const isNegative = str.startsWith('-') || (str.startsWith('(') && str.endsWith(')'));
  const cleaned = str.replace(/[^0-9.]/g, '');
  if (!cleaned) return defaultVal;
  const parsed = parseFloat(cleaned);
  if (isNaN(parsed)) return defaultVal;
  return isNegative ? -parsed : parsed;
}

/**
 * Normalizes header keys by lowercasing and stripping all non-alphanumeric chars
 */
function normalizeKey(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Specifically finds the column index for "Quantity" (Sales Quantity).
 * Prioritizes the exact column name 'quantity' as specified by user.
 */
function findQuantityColumnIndex(headers: string[]): number {
  // 1. Exact matches (priority: 'quantity' at the top!)
  const exactKeywords = [
    'quantity',
    'qty',
    'salesquantity',
    'salesqty',
    'soldquantity',
    'soldqty',
    'billquantity',
    'billqty',
    'netquantity',
    'netqty',
    'saleqty',
    'salequantity',
    'itemqty',
    'totqty',
    'totalqty',
    'nos',
    'pieces',
  ];

  for (const kw of exactKeywords) {
    const idx = headers.findIndex((h) => normalizeKey(h) === kw);
    if (idx !== -1) return idx;
  }

  // 2. Contains 'quantity' (e.g., 'Sales Quantity', 'Bill Quantity', 'Net Quantity')
  const idxQtyWord = headers.findIndex((h) => {
    const norm = normalizeKey(h);
    return (
      norm.includes('quantity') &&
      !norm.includes('min') &&
      !norm.includes('max') &&
      !norm.includes('reorder') &&
      !norm.includes('order')
    );
  });
  if (idxQtyWord !== -1) return idxQtyWord;

  // 3. Contains 'qty' or 'sold'
  const idxQtyShort = headers.findIndex((h) => {
    const norm = normalizeKey(h);
    return (
      (norm.includes('qty') || norm.includes('sold')) &&
      !norm.includes('min') &&
      !norm.includes('max') &&
      !norm.includes('reorder') &&
      !norm.includes('order')
    );
  });
  if (idxQtyShort !== -1) return idxQtyShort;

  return -1;
}

/**
 * Specifically finds the column index for "Toon Label" (or style / item code)
 */
function findToonLabelColumnIndex(headers: string[]): number {
  const exactKeywords = [
    'toonlable',
    'toonlabel',
    'toon',
    'style',
    'styleno',
    'stylenumber',
    'itemcode',
    'item_code',
    'designno',
    'design',
    'subclass',
    'class',
  ];

  for (const kw of exactKeywords) {
    const idx = headers.findIndex((h) => normalizeKey(h) === kw);
    if (idx !== -1) return idx;
  }

  const idxContains = headers.findIndex((h) => {
    const norm = normalizeKey(h);
    return norm.includes('toon') || norm.includes('style') || norm.includes('design');
  });
  if (idxContains !== -1) return idxContains;

  return -1;
}

/**
 * Specifically finds candidate column indices for "Branch" / "Location" / "Outlet" / "Store"
 * Prioritizes branch/outlet/location before generic company/store names
 */
function findStoreCandidateColumnIndices(headers: string[]): number[] {
  const indices: number[] = [];
  const priorityKeywords = [
    'branch',
    'branchname',
    'branchdesc',
    'branch_name',
    'location',
    'locationname',
    'location_name',
    'outlet',
    'outletname',
    'outlet_name',
    'store',
    'storename',
    'store_name',
    'shop',
    'shopname',
    'counter',
    'pos',
    'site',
    'sitename',
  ];

  headers.forEach((h, idx) => {
    const norm = normalizeKey(h);
    if (!norm) return;
    if (priorityKeywords.some((kw) => norm === kw || norm.includes(kw))) {
      indices.push(idx);
    }
  });

  return indices;
}

/**
 * Specifically finds the primary column index for "Store" / "Branch"
 */
function findStoreColumnIndex(headers: string[]): number {
  const priorityKeywords = [
    'branch',
    'branchname',
    'branchdesc',
    'location',
    'locationname',
    'outlet',
    'outletname',
    'store',
    'storename',
    'shop',
    'shopname',
    'counter',
    'pos',
    'site',
    'sitename',
  ];

  for (const kw of priorityKeywords) {
    const idx = headers.findIndex((h) => normalizeKey(h) === kw);
    if (idx !== -1) return idx;
  }

  const idxContains = headers.findIndex((h) => {
    const norm = normalizeKey(h);
    return (
      norm.includes('branch') ||
      norm.includes('location') ||
      norm.includes('outlet') ||
      norm.includes('store')
    );
  });
  if (idxContains !== -1) return idxContains;

  return -1;
}

/**
 * Scans rows before the header row for store/branch titles or banners.
 * Retail stores (especially Koottappalli) take strict precedence over generic warehouse titles.
 */
function extractStoreFromPreHeader(rawRows: unknown[][], headerRowIdx: number): StoreName | null {
  const maxScan = Math.min(headerRowIdx, rawRows.length);
  let candidateWarehouse: StoreName | null = null;

  for (let r = 0; r < maxScan; r++) {
    const row = rawRows[r];
    if (!row || !Array.isArray(row)) continue;
    for (const cell of row) {
      if (cell !== undefined && cell !== null && String(cell).trim()) {
        const matched = tryMatchStoreName(String(cell).trim());
        if (matched) {
          if (matched !== 'GM Fashions Warehouse') {
            return matched; // Immediate return for retail store (e.g. Koottappalli)
          } else {
            candidateWarehouse = matched;
          }
        }
      }
    }
  }

  return candidateWarehouse;
}

interface SalesSheetContext {
  sheetName: string;
  headers: string[];
  qtyColIdx: number;
  toonColIdx: number;
  storeColIdx: number;
  storeCandidateColIndices: number[];
  toonFallbackIdx: number;
  colourIdx: number;
  sizeIdx: number;
  productNameIdx: number;
  brandIdx: number;
  eanIdx: number;
  mrpIdx: number;
  billNoIdx: number;
  dateIdx: number;
  sheetLevelStore: StoreName | null;
  storeColIndices: { store: StoreName; colIndex: number; displayName: string }[];
  retailStoreCols: { store: StoreName; colIndex: number; displayName: string }[];
  isWideFormat: boolean;
}

function buildSalesSheetContext(
  sheetName: string,
  headers: string[],
  rawRowsForPreHeader: unknown[][],
  headerRowIdx: number,
  fileStoreMatch: StoreName | null
): SalesSheetContext {
  const headerIndexLookup: Record<string, number> = {};
  headers.forEach((h, idx) => {
    if (h) {
      const norm = normalizeKey(h);
      if (!(norm in headerIndexLookup)) {
        headerIndexLookup[norm] = idx;
      }
    }
  });

  const findColIdx = (...keys: string[]): number => {
    for (const k of keys) {
      const norm = normalizeKey(k);
      if (norm in headerIndexLookup) {
        return headerIndexLookup[norm];
      }
    }
    for (const k of keys) {
      const norm = normalizeKey(k);
      for (const [hNorm, cIdx] of Object.entries(headerIndexLookup)) {
        if (hNorm.includes(norm) || norm.includes(hNorm)) {
          return cIdx;
        }
      }
    }
    return -1;
  };

  const qtyColIdx = findQuantityColumnIndex(headers);
  const toonColIdx = findToonLabelColumnIndex(headers);
  const storeColIdx = findStoreColumnIndex(headers);
  const storeCandidateColIndices = findStoreCandidateColumnIndices(headers);

  const sheetStoreMatch = tryMatchStoreName(sheetName);
  const preHeaderStoreMatch = extractStoreFromPreHeader(rawRowsForPreHeader, headerRowIdx);
  const sheetLevelStore: StoreName | null =
    (sheetStoreMatch && sheetStoreMatch !== 'GM Fashions Warehouse' ? sheetStoreMatch : null) ||
    (preHeaderStoreMatch && preHeaderStoreMatch !== 'GM Fashions Warehouse' ? preHeaderStoreMatch : null) ||
    (fileStoreMatch && fileStoreMatch !== 'GM Fashions Warehouse' ? fileStoreMatch : null) ||
    sheetStoreMatch ||
    preHeaderStoreMatch ||
    fileStoreMatch;

  const storeColIndices: { store: StoreName; colIndex: number; displayName: string }[] = [];
  headers.forEach((h, idx) => {
    const norm = normalizeKey(h);
    if (!norm) return;

    if (
      norm.includes('toon') ||
      norm.includes('style') ||
      norm.includes('color') ||
      norm.includes('colour') ||
      norm.includes('size') ||
      norm.includes('mrp') ||
      norm.includes('price') ||
      norm.includes('rate') ||
      norm.includes('ean') ||
      norm.includes('barcode') ||
      norm.includes('date') ||
      norm.includes('bill') ||
      norm.includes('invoice') ||
      norm.includes('itemname') ||
      norm.includes('product') ||
      norm.includes('dept') ||
      norm.includes('class') ||
      norm.includes('brand') ||
      norm === 'total' ||
      norm === 'totalqty' ||
      norm === 'totalstock' ||
      norm === 'totalquantity' ||
      norm === 'totalsales' ||
      norm === 'salesqty'
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
      storeColIndices.push({ store: 'Kootapalli', colIndex: idx, displayName: 'KOOTTAPPALLI' });
      return;
    }
    if (norm.includes('kumba') || norm.includes('kumbakonam') || norm.includes('kmb')) {
      storeColIndices.push({ store: 'Kumbakonam', colIndex: idx, displayName: 'KUMBAKONAM' });
      return;
    }
    if (norm.includes('namakkal') || norm.includes('namakal') || norm.includes('nmkl')) {
      storeColIndices.push({ store: 'Namakkal', colIndex: idx, displayName: 'NAMAKKAL' });
      return;
    }
    if (norm.includes('salem') || norm.includes('slm')) {
      storeColIndices.push({ store: 'Salem', colIndex: idx, displayName: 'SALEM' });
      return;
    }
    if (norm.includes('karur') || norm.includes('krr')) {
      storeColIndices.push({ store: 'Karur', colIndex: idx, displayName: 'KARUR' });
      return;
    }
    if (norm.includes('mallur') || norm.includes('mlr')) {
      storeColIndices.push({ store: 'Mallur', colIndex: idx, displayName: 'MALLUR' });
      return;
    }
    if (norm.includes('thiru') || norm.includes('tiru') || norm.includes('tvm')) {
      storeColIndices.push({ store: 'Thiruvannamalai', colIndex: idx, displayName: 'TIRUVANNAMALAI' });
      return;
    }

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

    if (isExcludedApparelOrTax) return;

    if (
      norm.includes('warehouse') ||
      norm.includes('godown') ||
      norm.includes('gmfashionswarehouse') ||
      norm.includes('centralwarehouse') ||
      norm.includes('centralgodown') ||
      norm === 'wh'
    ) {
      storeColIndices.push({ store: 'GM Fashions Warehouse', colIndex: idx, displayName: 'GM FASHIONS WAREHOUSE' });
    }
  });

  const hasRowStoreColumn = storeColIdx !== -1 || storeCandidateColIndices.length > 0;
  const retailStoreCols = storeColIndices.filter((sc) => sc.store !== 'GM Fashions Warehouse');
  const isWideFormat = !hasRowStoreColumn && retailStoreCols.length >= 2;

  return {
    sheetName,
    headers,
    qtyColIdx,
    toonColIdx,
    storeColIdx,
    storeCandidateColIndices,
    toonFallbackIdx: findColIdx(
      'toonlable',
      'toonlabel',
      'toon_lable',
      'toon_label',
      'toon',
      'style',
      'styleno',
      'style_no',
      'itemcode',
      'designno',
      'design',
      'subclass',
      'class',
      'itemname',
      'productname'
    ),
    colourIdx: findColIdx('colour', 'color', 'shade'),
    sizeIdx: findColIdx('size', 'itemsize'),
    productNameIdx: findColIdx('itemname', 'productname', 'department', 'dept', 'description', 'productgroup'),
    brandIdx: findColIdx('brand', 'class', 'subclass'),
    eanIdx: findColIdx('ean', 'barcode', 'upc', 'bcode'),
    mrpIdx: findColIdx('mrp', 'price', 'rate', 'salerate', 'amount', 'netamount'),
    billNoIdx: findColIdx('billno', 'invoiceno', 'bill', 'invno'),
    dateIdx: findColIdx('billdate', 'date', 'saledate'),
    sheetLevelStore,
    storeColIndices,
    retailStoreCols,
    isWideFormat,
  };
}

function mapSalesArrayRow(
  row: unknown[],
  ctx: SalesSheetContext,
  rowIndex: number,
  items: MappedSalesItem[],
  uniqueToonSet: Set<string>,
  uniqueEanSet: Set<string>
): void {
  if (!row || !Array.isArray(row)) return;
  let hasContent = false;
  for (let i = 0; i < row.length; i++) {
    const v = row[i];
    if (v !== '' && v !== null && v !== undefined) {
      hasContent = true;
      break;
    }
  }
  if (!hasContent) return;

  const strAt = (idx: number): string => {
    if (idx < 0 || idx >= row.length) return '';
    const v = row[idx];
    return v !== undefined && v !== null ? String(v).trim() : '';
  };

  let toonLabel = ctx.toonColIdx !== -1 ? strAt(ctx.toonColIdx) : '';
  if (!toonLabel && ctx.toonFallbackIdx !== -1) {
    toonLabel = strAt(ctx.toonFallbackIdx);
  }
  if (!toonLabel) return;

  uniqueToonSet.add(toonLabel);

  const colour = strAt(ctx.colourIdx) || 'Assorted';
  const size = strAt(ctx.sizeIdx) || 'Free';
  const productName = strAt(ctx.productNameIdx) || toonLabel;
  const brand = strAt(ctx.brandIdx) || 'GM Fashions';
  const rawEan = strAt(ctx.eanIdx);
  const ean = cleanEanString(rawEan);
  if (ean) uniqueEanSet.add(ean);

  const mrp = ctx.mrpIdx !== -1 ? cleanNumber(row[ctx.mrpIdx], 0) : 0;
  const billNo = strAt(ctx.billNoIdx);
  const date = strAt(ctx.dateIdx);

  if (ctx.isWideFormat) {
    const breakdown: Record<string, number> = {};
    let hasAnySales = false;

    for (const sc of ctx.storeColIndices) {
      const qty = cleanNumber(row[sc.colIndex], 0);
      if (qty > 0) {
        breakdown[sc.displayName] = (breakdown[sc.displayName] || 0) + qty;
        hasAnySales = true;

        items.push({
          id: `sales_${ctx.sheetName}_${rowIndex}_${sc.store}_${items.length}`,
          toonLabel,
          colour,
          size,
          productName,
          brand,
          ean,
          store: sc.store,
          salesQty: qty,
          mrp,
          billNo,
          date,
          storeBreakdown: breakdown,
        });
      }
    }

    if (!hasAnySales) {
      items.push({
        id: `sales_${ctx.sheetName}_${rowIndex}_zero_${items.length}`,
        toonLabel,
        colour,
        size,
        productName,
        brand,
        ean,
        store: ctx.storeColIndices[0]?.store || 'Kootapalli',
        salesQty: 0,
        mrp,
        billNo,
        date,
        storeBreakdown: breakdown,
      });
    }
    return;
  }

  const salesQty = ctx.qtyColIdx !== -1 ? cleanNumber(row[ctx.qtyColIdx], 0) : 0;

  let rowStore: StoreName | null = null;
  let candidateWarehouse: StoreName | null = null;

  if (ctx.storeColIdx !== -1) {
    const directVal = strAt(ctx.storeColIdx);
    if (directVal) {
      const matched = tryMatchStoreName(directVal);
      if (matched) {
        if (matched !== 'GM Fashions Warehouse') {
          rowStore = matched;
        } else {
          candidateWarehouse = matched;
        }
      }
    }
  }

  if (!rowStore) {
    for (const cIdx of ctx.storeCandidateColIndices) {
      const cellStr = strAt(cIdx);
      if (cellStr) {
        const matched = tryMatchStoreName(cellStr);
        if (matched) {
          if (matched !== 'GM Fashions Warehouse') {
            rowStore = matched;
            break;
          } else {
            candidateWarehouse = matched;
          }
        }
      }
    }
  }

  if (!rowStore && ctx.sheetLevelStore && ctx.sheetLevelStore !== 'GM Fashions Warehouse') {
    rowStore = ctx.sheetLevelStore;
  }

  if (!rowStore && ctx.retailStoreCols.length === 1) {
    rowStore = ctx.retailStoreCols[0].store;
  }

  if (!rowStore && candidateWarehouse) {
    rowStore = candidateWarehouse;
  }

  if (!rowStore) {
    rowStore = 'Kootapalli';
  }

  const storeDisplay = formatStoreDisplayName(rowStore);

  items.push({
    id: `sales_${ctx.sheetName}_${rowIndex}_${items.length}`,
    toonLabel,
    colour,
    size,
    productName,
    brand,
    ean,
    store: rowStore,
    salesQty,
    mrp,
    billNo,
    date,
    storeBreakdown: { [storeDisplay]: salesQty },
  });
}

/**
 * Parses an uploaded Sales Report Excel/CSV file (up to 1 GB) from Wondersoft POS.
 * Streams CSV files in 4 MB chunks and uses dense memory-saving SheetJS mode for Excel files.
 */
export async function parseSalesExcelFile(
  file: File,
  onProgress?: (progress: ParseProgress) => void
): Promise<{ items: MappedSalesItem[]; fileInfo: UploadedFileInfo }> {
  if (file.size > MAX_FILE_UPLOAD_BYTES) {
    throw new Error(
      `File size (${formatByteSize(file.size)}) exceeds the 1 GB (${formatByteSize(MAX_FILE_UPLOAD_BYTES)}) maximum limit.`
    );
  }

  const items: MappedSalesItem[] = [];
  const uniqueToonSet = new Set<string>();
  const uniqueEanSet = new Set<string>();
  let detectedQtyColName = 'Quantity';

  const fileStoreMatch = tryMatchStoreName(file.name);
  const lowerName = file.name.toLowerCase();

  if (lowerName.endsWith('.csv')) {
    const headerScanBuffer: string[][] = [];
    let ctx: SalesSheetContext | null = null;
    let rowCounter = 0;

    await streamCsvFileInBatches(
      file,
      (batch) => {
        if (!ctx) {
          for (let i = 0; i < batch.length; i++) {
            headerScanBuffer.push(batch[i]);
          }
          if (headerScanBuffer.length >= 2) {
            const headerRowIdx = findHeaderRowIndexFromRows(headerScanBuffer);
            const headers = (headerScanBuffer[headerRowIdx] || []).map((h) => String(h || '').trim());
            ctx = buildSalesSheetContext('CSV', headers, headerScanBuffer, headerRowIdx, fileStoreMatch);
            if (ctx.qtyColIdx !== -1 && ctx.headers[ctx.qtyColIdx]) {
              detectedQtyColName = ctx.headers[ctx.qtyColIdx];
            }
            for (let r = headerRowIdx + 1; r < headerScanBuffer.length; r++) {
              mapSalesArrayRow(headerScanBuffer[r], ctx, rowCounter++, items, uniqueToonSet, uniqueEanSet);
            }
            headerScanBuffer.length = 0;
          }
        } else {
          for (let i = 0; i < batch.length; i++) {
            mapSalesArrayRow(batch[i], ctx, rowCounter++, items, uniqueToonSet, uniqueEanSet);
          }
        }
      },
      onProgress
    );
  } else {
    if (onProgress) {
      onProgress({
        percent: 15,
        loadedBytes: 0,
        totalBytes: file.size,
        rowsParsed: 0,
        stage: 'Reading sales workbook...',
      });
    }

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, {
      type: 'array',
      dense: true,
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
      cellText: false,
    });

    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('The selected workbook has no sheets.');
    }

    const totalSheets = workbook.SheetNames.length;
    const BATCH_SIZE = 10000;

    for (let sIdx = 0; sIdx < totalSheets; sIdx++) {
      const sheetName = workbook.SheetNames[sIdx];
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet || !worksheet['!ref']) continue;

      const rawRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
        header: 1,
        defval: '',
      });
      const headerRowIdx = findHeaderRowIndex(worksheet);
      // Free worksheet memory immediately
      delete workbook.Sheets[sheetName];

      if (!rawRows || rawRows.length < 2) continue;

      const rawHeaderRow = (rawRows[headerRowIdx] as unknown[]) || [];
      const headers = rawHeaderRow.map((h) => String(h || '').trim());
      const ctx = buildSalesSheetContext(sheetName, headers, rawRows, headerRowIdx, fileStoreMatch);

      if (ctx.qtyColIdx !== -1 && headers[ctx.qtyColIdx]) {
        detectedQtyColName = headers[ctx.qtyColIdx];
      }

      const dataRowCount = rawRows.length - (headerRowIdx + 1);
      for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        mapSalesArrayRow(rawRows[r], ctx, r, items, uniqueToonSet, uniqueEanSet);
        if ((r - headerRowIdx) % BATCH_SIZE === 0) {
          if (onProgress) {
            const sheetFrac = (sIdx + (r - headerRowIdx) / Math.max(dataRowCount, 1)) / totalSheets;
            onProgress({
              percent: 25 + Math.round(sheetFrac * 68),
              loadedBytes: file.size,
              totalBytes: file.size,
              rowsParsed: items.length,
              stage: `Processing ${items.length.toLocaleString()} sales rows...`,
            });
          }
          await yieldToMainThread();
        }
      }
    }
  }

  if (items.length === 0) {
    throw new Error(
      'No valid sales records could be parsed. Ensure the file includes a Toon Label and Quantity column.'
    );
  }

  if (onProgress) {
    onProgress({
      percent: 96,
      loadedBytes: file.size,
      totalBytes: file.size,
      rowsParsed: items.length,
      stage: 'Aggregating store sales breakdowns...',
    });
  }
  await yieldToMainThread();

  const unifiedItems = calculateTotalSalesQuantities(items);
  let totalSalesUnits = 0;
  for (let i = 0; i < unifiedItems.length; i++) {
    totalSalesUnits += Number(unifiedItems[i].salesQty) || 0;
  }

  const fileInfo: UploadedFileInfo = {
    fileName: file.name,
    fileSize: file.size,
    totalRows: unifiedItems.length,
    uniqueEANs: uniqueEanSet.size,
    uniqueToonLabels: uniqueToonSet.size,
    uploadedAt: new Date().toLocaleTimeString(),
    totalSalesUnits,
    detectedQuantityColumn: detectedQtyColName,
  };

  if (onProgress) {
    onProgress({
      percent: 100,
      loadedBytes: file.size,
      totalBytes: file.size,
      rowsParsed: unifiedItems.length,
      stage: 'Completed',
    });
  }

  return { items: unifiedItems, fileInfo };
}

/**
 * Calculates and attaches complete storeBreakdown across all stores for each unique item/EAN in sales data.
 * Mutates items in-place to avoid doubling RAM usage on massive datasets.
 */
export function calculateTotalSalesQuantities(items: MappedSalesItem[]): MappedSalesItem[] {
  const storeBreakdownsMap = new Map<string, Record<string, number>>();

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = item.ean && item.ean.trim()
      ? item.ean.trim()
      : `${(item.colour || '').trim()}__${(item.toonLabel || '').trim()}__${(item.size || '').trim()}`.toLowerCase();

    let bd = storeBreakdownsMap.get(key);
    if (!bd) {
      bd = {
        'KOOTTAPPALLI': 0,
        'KARUR': 0,
        'SALEM': 0,
        'NAMAKKAL': 0,
        'KUMBAKONAM': 0,
        'TIRUVANNAMALAI': 0,
        'MALLUR': 0,
        'GM FASHIONS WAREHOUSE': 0,
      };
      storeBreakdownsMap.set(key, bd);
    }

    const sDisplay = formatStoreDisplayName(item.store);
    bd[sDisplay] = (bd[sDisplay] || 0) + (Number(item.salesQty) || 0);
  }

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = item.ean && item.ean.trim()
      ? item.ean.trim()
      : `${(item.colour || '').trim()}__${(item.toonLabel || '').trim()}__${(item.size || '').trim()}`.toLowerCase();

    item.storeBreakdown = storeBreakdownsMap.get(key) || {
      [formatStoreDisplayName(item.store)]: Number(item.salesQty) || 0,
      'KOOTTAPPALLI': 0,
      'GM FASHIONS WAREHOUSE': 0,
    };
  }

  return items;
}

/**
 * Generates the Store Matrix for Sales data:
 * Aggregates sales count grouped by [Color, Toon Label, Size]
 * and returns the count of sales per store:
 * Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, GM Fashions Warehouse, and Total Sales!
 */
export function generateSalesMatrix(items: MappedSalesItem[]): SalesMatrixRow[] {
  const matrixMap = new Map<string, SalesMatrixRow>();

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
        totalSales: 0,
      };
      matrixMap.set(key, row);
    }

    const qty = Number(item.salesQty) || 0;
    const rawStore = String(item.store || '').trim();
    const matchedStore = tryMatchStoreName(rawStore);
    const store = matchedStore || matchStoreName(rawStore, 'Kootapalli');
    switch (store) {
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

    row.totalSales =
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
 * Downloads Excel file for Store Matrix in Sales Report:
 * [Color, Toon Label, Size, Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, GM Fashions Warehouse, Total Sales]
 */
export function exportSalesMatrixExcel(
  items: MappedSalesItem[],
  searchQuery: string
): string {
  const matrixRows = generateSalesMatrix(items);

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
    'Total Sales': row.totalSales,
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
        acc.TotalSales += r['Total Sales'];
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
        TotalSales: 0,
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
      'Total Sales': totals.TotalSales,
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
    'Total Sales',
  ];

  const worksheet = XLSX.utils.json_to_sheet(exportData, { header: headers });
  worksheet['!cols'] = calculateColumnWidths(headers, exportData);
  styleMatrixWorksheet(worksheet, headers, exportData.length, 'sales', groups);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sales Store Matrix');

  const safeQuery = searchQuery.replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = `GM_Fashions_Sales_Matrix_${safeQuery || 'All'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(workbook, filename);
  return filename;
}

/**
 * Downloads Excel file for Detailed Sales Table (Inventory Table view)
 */
export function exportDetailedSalesExcel(
  items: MappedSalesItem[],
  searchQuery: string
): string {
  const exportData = items.map((item) => ({
    'Toon Label': item.toonLabel,
    'Product Description': item.productName,
    'Brand': item.brand,
    'Color': item.colour,
    'Size': item.size,
    'Store': item.store,
    'Sales Qty': item.salesQty,
    'MRP': item.mrp,
    'EAN': item.ean,
    'Bill No': item.billNo || '-',
    'Date': item.date || '-',
  }));

  const headers = [
    'Toon Label',
    'Product Description',
    'Brand',
    'Color',
    'Size',
    'Store',
    'Sales Qty',
    'MRP',
    'EAN',
    'Bill No',
    'Date',
  ];

  const worksheet = XLSX.utils.json_to_sheet(exportData, { header: headers });
  worksheet['!cols'] = calculateColumnWidths(headers, exportData);
  // Column 6 is 'Sales Qty' which gets highlighted
  styleDetailedWorksheet(worksheet, headers, exportData.length, 'sales', 6);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sales Table Details');

  const safeQuery = searchQuery.replace(/[^a-zA-Z0-9_-]/g, '_');
  const filename = `GM_Fashions_Sales_Table_${safeQuery || 'All'}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  XLSX.writeFile(workbook, filename);
  return filename;
}
