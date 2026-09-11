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
  compareToonColourAndSize,
} from './excelParser';

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

/**
 * Parses an uploaded Sales Report Excel/CSV file from Wondersoft POS.
 * Explicitly extracts the 'Quantity' column as sales quantity.
 * Supports single-sheet or multi-sheet workbooks, and wide store matrices.
 */
export async function parseSalesExcelFile(
  file: File
): Promise<{ items: MappedSalesItem[]; fileInfo: UploadedFileInfo }> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('The selected workbook has no sheets.');
  }

  const items: MappedSalesItem[] = [];
  const uniqueToonSet = new Set<string>();
  const uniqueEanSet = new Set<string>();
  let detectedQtyColName = 'Quantity';

  // Check if filename itself contains a store name (e.g. Koottappalli_Sales.xlsx)
  const fileStoreMatch = tryMatchStoreName(file.name);

  // Process sheets (supports multi-sheet store exports or single sheet)
  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet || !worksheet['!ref']) continue;

    const rawRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      defval: '',
    });

    if (!rawRows || rawRows.length < 2) continue;

    const headerRowIdx = findHeaderRowIndex(worksheet);
    const rawHeaderRow = (rawRows[headerRowIdx] as unknown[]) || [];
    const headers = rawHeaderRow.map((h) => String(h || '').trim());

    // Build column index lookup
    const headerIndexLookup: Record<string, number> = {};
    headers.forEach((h, idx) => {
      if (h) {
        const norm = normalizeKey(h);
        if (!(norm in headerIndexLookup)) {
          headerIndexLookup[norm] = idx;
        }
      }
    });

    // Detect exact column indices
    const qtyColIdx = findQuantityColumnIndex(headers);
    const toonColIdx = findToonLabelColumnIndex(headers);
    const storeColIdx = findStoreColumnIndex(headers);
    const storeCandidateColIndices = findStoreCandidateColumnIndices(headers);

    if (qtyColIdx !== -1 && headers[qtyColIdx]) {
      detectedQtyColName = headers[qtyColIdx];
    }

    // Check if sheet name indicates a specific store
    const sheetStoreMatch = tryMatchStoreName(sheetName);

    // Check pre-header rows (banner/header title rows above column headers)
    const preHeaderStoreMatch = extractStoreFromPreHeader(rawRows, headerRowIdx);

    // Determine default store for this sheet: retail store ALWAYS beats warehouse
    const sheetLevelStore: StoreName | null =
      (sheetStoreMatch && sheetStoreMatch !== 'GM Fashions Warehouse' ? sheetStoreMatch : null) ||
      (preHeaderStoreMatch && preHeaderStoreMatch !== 'GM Fashions Warehouse' ? preHeaderStoreMatch : null) ||
      (fileStoreMatch && fileStoreMatch !== 'GM Fashions Warehouse' ? fileStoreMatch : null) ||
      sheetStoreMatch ||
      preHeaderStoreMatch ||
      fileStoreMatch;

    // Detect wide store columns (e.g. KOOTTAPPALLI, SALEM, KARUR, etc. as separate column headers)
    const storeColIndices: { store: StoreName; colIndex: number; displayName: string }[] = [];
    headers.forEach((h, idx) => {
      const norm = normalizeKey(h);
      if (!norm) return;

      // Skip non-store metadata columns
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

      // Check Kootapalli FIRST and comprehensively
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
        storeColIndices.push({
          store: 'Kootapalli',
          colIndex: idx,
          displayName: 'KOOTTAPPALLI',
        });
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
      // Exclude garment category wear terms, tax terms, software, etc. from matching as store columns
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

      if (isExcludedApparelOrTax) {
        return;
      }

      if (
        norm.includes('warehouse') ||
        norm.includes('godown') ||
        norm.includes('gmfashionswarehouse') ||
        norm.includes('centralwarehouse') ||
        norm.includes('centralgodown') ||
        norm === 'wh'
      ) {
        storeColIndices.push({ store: 'GM Fashions Warehouse', colIndex: idx, displayName: 'GM FASHIONS WAREHOUSE' });
        return;
      }
    });

    // An Excel file with a 'StoreName', 'BranchName', or 'Store' column is NEVER a wide matrix.
    // Wide matrix format ONLY applies when there is NO row store column AND at least 2 distinct RETAIL store columns.
    const hasRowStoreColumn = storeColIdx !== -1 || storeCandidateColIndices.length > 0;
    const retailStoreCols = storeColIndices.filter(
      (sc) => sc.store !== 'GM Fashions Warehouse'
    );
    const isWideFormat = !hasRowStoreColumn && retailStoreCols.length >= 2;

    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (
        !row ||
        !Array.isArray(row) ||
        row.every((c) => c === '' || c === null || c === undefined)
      ) {
        continue;
      }

      // Helper to get cell by header variations (exact first, then fuzzy)
      const getVal = (...keys: string[]): string => {
        for (const k of keys) {
          const norm = normalizeKey(k);
          if (norm in headerIndexLookup) {
            const cIdx = headerIndexLookup[norm];
            const val = row[cIdx];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              return String(val).trim();
            }
          }
        }
        for (const k of keys) {
          const norm = normalizeKey(k);
          for (const [hNorm, cIdx] of Object.entries(headerIndexLookup)) {
            if (hNorm.includes(norm) || norm.includes(hNorm)) {
              const val = row[cIdx];
              if (val !== undefined && val !== null && String(val).trim() !== '') {
                return String(val).trim();
              }
            }
          }
        }
        return '';
      };

      // Extract Toon Label (either from direct column or fallback lookup)
      let toonLabel = '';
      if (toonColIdx !== -1 && row[toonColIdx] !== undefined && row[toonColIdx] !== null) {
        toonLabel = String(row[toonColIdx]).trim();
      }
      if (!toonLabel) {
        toonLabel = (
          getVal(
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
            'design'
          ) || getVal('subclass', 'class', 'itemname', 'productname')
        ).trim();
      }

      if (!toonLabel) {
        continue;
      }

      uniqueToonSet.add(toonLabel);

      // Extract Color, Size, Brand, Product Name, EAN, MRP
      const colour = (getVal('colour', 'color', 'shade') || 'Assorted').trim();
      const size = (getVal('size', 'itemsize') || 'Free').trim();
      const productName = (
        getVal('itemname', 'productname', 'department', 'dept', 'description', 'productgroup') ||
        toonLabel
      ).trim();
      const brand = (getVal('brand', 'class', 'subclass') || 'GM Fashions').trim();
      const rawEan = getVal('ean', 'barcode', 'upc', 'bcode');
      const ean = cleanEanString(rawEan);
      if (ean) uniqueEanSet.add(ean);

      const mrp = cleanNumber(
        getVal('mrp', 'price', 'rate', 'salerate', 'amount', 'netamount'),
        0
      );
      const billNo = getVal('billno', 'invoiceno', 'bill', 'invno');
      const date = getVal('billdate', 'date', 'saledate');

      if (isWideFormat) {
        // Wide Format: read quantity from each store column
        const breakdown: Record<string, number> = {};
        let hasAnySales = false;

        for (const sc of storeColIndices) {
          const rawVal = row[sc.colIndex];
          const qty = cleanNumber(rawVal, 0);
          if (qty > 0) {
            breakdown[sc.displayName] = (breakdown[sc.displayName] || 0) + qty;
            hasAnySales = true;

            items.push({
              id: `sales_${sheetName}_${r}_${sc.store}_${items.length}`,
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
            id: `sales_${sheetName}_${r}_zero_${items.length}`,
            toonLabel,
            colour,
            size,
            productName,
            brand,
            ean,
            store: storeColIndices[0]?.store || 'Kootapalli',
            salesQty: 0,
            mrp,
            billNo,
            date,
            storeBreakdown: breakdown,
          });
        }
      } else {
        // Standard Row Format: Sales Quantity extracted directly from 'Quantity' column!
        let salesQty = 0;
        if (qtyColIdx !== -1 && row[qtyColIdx] !== undefined && row[qtyColIdx] !== null) {
          salesQty = cleanNumber(row[qtyColIdx], 0);
        } else {
          const rawQty = getVal(
            'quantity',
            'qty',
            'salesquantity',
            'salesqty',
            'soldquantity',
            'soldqty',
            'billquantity',
            'billqty',
            'netquantity',
            'netqty'
          );
          salesQty = cleanNumber(rawQty, 0);
        }

        // Determine Store for this row
        // Retail store (especially Koottappalli) MUST take precedence over Warehouse
        let rowStore: StoreName | null = null;
        let candidateWarehouse: StoreName | null = null;

        // 1. Direct Primary Store Column (e.g. Column B: "StoreName" having "KOOTTAPPALLI" or "MALLUR")
        if (storeColIdx !== -1 && row[storeColIdx] !== undefined && row[storeColIdx] !== null) {
          const directVal = String(row[storeColIdx]).trim();
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

        // 2. Check candidate columns (branch, location, outlet, store)
        if (!rowStore) {
          for (const cIdx of storeCandidateColIndices) {
            const cell = row[cIdx];
            if (cell !== undefined && cell !== null && String(cell).trim()) {
              const matched = tryMatchStoreName(String(cell).trim());
              if (matched) {
                if (matched !== 'GM Fashions Warehouse') {
                  rowStore = matched;
                  break; // Retail store found!
                } else {
                  candidateWarehouse = matched;
                }
              }
            }
          }
        }

        // 3. Check getVal for specific branch keywords first
        if (!rowStore) {
          const rawBranch = getVal(
            'branch',
            'branchname',
            'branchdesc',
            'branch_name',
            'location',
            'locationname',
            'outlet',
            'outletname',
            'shop',
            'counter'
          );
          if (rawBranch) {
            const matched = tryMatchStoreName(rawBranch);
            if (matched && matched !== 'GM Fashions Warehouse') {
              rowStore = matched;
            } else if (matched) {
              candidateWarehouse = matched;
            }
          }
        }

        // 4. Check getVal for company/store keywords
        if (!rowStore) {
          const rawStore = getVal('store', 'storename', 'company', 'site');
          if (rawStore) {
            const matched = tryMatchStoreName(rawStore);
            if (matched && matched !== 'GM Fashions Warehouse') {
              rowStore = matched;
            } else if (matched) {
              candidateWarehouse = matched;
            }
          }
        }

        // 5. Use sheet-level / pre-header / file-level detected store ONLY IF IT IS A RETAIL STORE
        if (!rowStore && sheetLevelStore && sheetLevelStore !== 'GM Fashions Warehouse') {
          rowStore = sheetLevelStore;
        }

        // 6. If headers had a single store column (e.g. KOOTTAPPALLI)
        if (!rowStore && retailStoreCols.length === 1) {
          rowStore = retailStoreCols[0].store;
        }

        // 7. If row had a candidate warehouse and no retail store found
        if (!rowStore && candidateWarehouse) {
          rowStore = candidateWarehouse;
        }

        // 8. Fallback to Koottappalli
        if (!rowStore) {
          rowStore = 'Kootapalli';
        }

        const storeDisplay = formatStoreDisplayName(rowStore);

        items.push({
          id: `sales_${sheetName}_${r}_${items.length}`,
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
    }
  }

  if (items.length === 0) {
    throw new Error(
      'No valid sales records could be parsed. Ensure the file includes a Toon Label and Quantity column.'
    );
  }

  // Calculate and attach complete store breakdown for each unique item/EAN
  const unifiedItems = calculateTotalSalesQuantities(items);
  const totalSalesUnits = unifiedItems.reduce((sum, item) => sum + (Number(item.salesQty) || 0), 0);

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

  return { items: unifiedItems, fileInfo };
}

/**
 * Calculates and attaches complete storeBreakdown across all stores for each unique item/EAN in sales data
 */
export function calculateTotalSalesQuantities(items: MappedSalesItem[]): MappedSalesItem[] {
  const totalsMap = new Map<string, number>();
  const storeBreakdownsMap = new Map<string, Record<string, number>>();

  for (const item of items) {
    const key = item.ean && item.ean.trim()
      ? item.ean.trim()
      : `${(item.colour || '').trim()}__${(item.toonLabel || '').trim()}__${(item.size || '').trim()}`.toLowerCase();

    const curr = totalsMap.get(key) || 0;
    totalsMap.set(key, curr + (Number(item.salesQty) || 0));

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

  return items.map((item) => {
    const key = item.ean && item.ean.trim()
      ? item.ean.trim()
      : `${(item.colour || '').trim()}__${(item.toonLabel || '').trim()}__${(item.size || '').trim()}`.toLowerCase();

    const bd = storeBreakdownsMap.get(key) || {
      [formatStoreDisplayName(item.store)]: Number(item.salesQty) || 0,
      'KOOTTAPPALLI': 0,
      'GM FASHIONS WAREHOUSE': 0,
    };

    return {
      ...item,
      storeBreakdown: bd,
    };
  });
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
