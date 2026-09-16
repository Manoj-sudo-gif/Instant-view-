import { useState, useEffect, useMemo, Fragment } from 'react';
import {
  FileSpreadsheet,
  Download,
  Grid3X3,
  Search,
  ArrowUpDown,
  ChevronDown,
  ChevronUp,
  Boxes,
  Store,
  Shuffle,
  RotateCcw,
} from 'lucide-react';
import {
  MappedInventoryItem,
  SearchState,
  StoreName,
  ALL_STORES,
} from '../types/inventory';
import {
  exportToonLabelExcel,
  exportEanExcel,
  generateToonLabelMatrix,
  formatStoreDisplayName,
  compareColourAndSize,
  compareSizes,
  compareToonColourAndSize,
} from '../utils/excelParser';
import {
  shuffleMatrixData,
  ShuffledMatrixRow,
  MatrixStoreKey,
} from '../utils/stockShuffle';
import { StoreStockCards } from './StoreStockCards';

interface RoughViewTableProps {
  items: MappedInventoryItem[];
  searchState: SearchState;
  lastExportFilename: string | null;
  onDownloadExcel?: (mode?: 'mapped' | 'matrix') => void;
}

const ALLOWED_STORE_DISPLAYS = new Set([
  'KOOTTAPPALLI',
  'KUMBAKONAM',
  'MALLUR',
  'NAMAKKAL',
  'SALEM',
  'TIRUVANNAMALAI',
  'KARUR',
  'GM FASHIONS WAREHOUSE',
]);

const PREFERRED_ORDER = [
  'KOOTTAPPALLI',
  'KUMBAKONAM',
  'MALLUR',
  'NAMAKKAL',
  'SALEM',
  'TIRUVANNAMALAI',
  'KARUR',
  'GM FASHIONS WAREHOUSE',
];

export interface AvailableStoreStock {
  storeName: string;
  qty: number;
}

export function getItemStoreStats(
  item: MappedInventoryItem,
  allItems: MappedInventoryItem[] = []
): {
  availableStores: AvailableStoreStock[];
  totalStoreQty: number;
  totalUnits: number;
} {
  const map = new Map<string, number>();

  // 1. Direct store breakdown from item if available
  if (item.storeBreakdown && Object.keys(item.storeBreakdown).length > 0) {
    for (const [st, q] of Object.entries(item.storeBreakdown)) {
      if (q > 0) {
        const sDisplay = formatStoreDisplayName(st);
        if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
          map.set(sDisplay, q);
        }
      }
    }
  }

  // 2. Aggregate from allItems for matching product identity
  if (allItems.length > 0) {
    const targetEan = item.ean ? item.ean.trim() : '';
    const matching = allItems.filter((i) => {
      if (targetEan && i.ean) {
        return i.ean.trim() === targetEan;
      }
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
        if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
          const prev = map.get(sDisplay) || 0;
          map.set(sDisplay, Math.max(prev, m.stockQuantity));
        }
      }
    }
  }

  // 3. Fallback to item itself
  if (map.size === 0 && item.stockQuantity > 0) {
    const sDisplay = formatStoreDisplayName(item.store);
    if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
      map.set(sDisplay, item.stockQuantity);
    }
  }

  // Strictly filter only stores with available stock (> 0)
  const availableStores: AvailableStoreStock[] = Array.from(map.entries())
    .filter(([_, qty]) => qty > 0)
    .map(([storeName, qty]) => ({
      storeName,
      qty,
    }));

  availableStores.sort((a, b) => {
    const idxA = PREFERRED_ORDER.indexOf(a.storeName);
    const idxB = PREFERRED_ORDER.indexOf(b.storeName);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.storeName.localeCompare(b.storeName);
  });

  // Calculate sum of store quantities
  const totalUnits = availableStores.reduce((acc, curr) => acc + curr.qty, 0);
  const totalStoreQty = availableStores.length;

  return {
    availableStores,
    totalStoreQty,
    totalUnits: totalUnits > 0 ? totalUnits : (item.totalQuantity || item.stockQuantity || 0),
  };
}

export function RoughViewTable({
  items,
  searchState,
  onDownloadExcel,
}: RoughViewTableProps) {
  const [activeStoreFilter, setActiveStoreFilter] = useState<StoreName | 'ALL'>('ALL');
  const [tableSearch, setTableSearch] = useState('');
  const [sortField, setSortField] = useState<keyof MappedInventoryItem>(
    searchState.type === 'TOON' ? 'toonLabel' : 'totalQuantity'
  );
  const [sortAsc, setSortAsc] = useState(searchState.type === 'TOON' ? true : false);
  const [viewMode, setViewMode] = useState<'table' | 'matrix'>('table');
  const [expandedItemId, setExpandedItemId] = useState<string | null>(null);
  const [isShuffled, setIsShuffled] = useState(false);

  // Sync sort when search type or query changes:
  // TOON searches automatically prioritize 1st: Toon Label (asc), 2nd: Colour (asc), 3rd: Size (32, 34, 36... or S, M, L...)
  useEffect(() => {
    setIsShuffled(false);
    if (searchState.type === 'TOON') {
      setSortField('toonLabel');
      setSortAsc(true);
    } else {
      setSortField('totalQuantity');
      setSortAsc(false);
    }
  }, [searchState.type, searchState.query, searchState.timestamp]);

  // Manual Excel download handler
  const handleDownload = () => {
    if (onDownloadExcel) {
      onDownloadExcel(viewMode === 'matrix' ? 'matrix' : 'mapped');
      return;
    }
    if (searchState.type === 'TOON' && viewMode === 'matrix') {
      exportToonLabelExcel(items, searchState.query);
    } else {
      exportEanExcel(items, searchState.query);
    }
  };

  // Filter items based on store card and search text
  const filteredItems = useMemo(() => {
    let result = [...items];

    if (activeStoreFilter !== 'ALL') {
      result = result.filter((item) => item.store === activeStoreFilter);
    } else {
      result = result.filter((item) => ALL_STORES.includes(item.store));
    }

    if (tableSearch.trim()) {
      const q = tableSearch.toLowerCase().trim();
      result = result.filter(
        (item) =>
          item.productName.toLowerCase().includes(q) ||
          item.department.toLowerCase().includes(q) ||
          item.brand.toLowerCase().includes(q) ||
          item.colour.toLowerCase().includes(q) ||
          item.size.toLowerCase().includes(q) ||
          item.fabric.toLowerCase().includes(q) ||
          item.ean.toLowerCase().includes(q) ||
          item.toonLabel.toLowerCase().includes(q) ||
          item.store.toLowerCase().includes(q)
      );
    }

    return result;
  }, [items, activeStoreFilter, tableSearch]);

  // Consolidate into single common rows (no repetitive store list rows)
  const distinctItems = useMemo(() => {
    const map = new Map<string, MappedInventoryItem>();
    for (const item of filteredItems) {
      const key = item.ean
        ? item.ean.trim()
        : `${item.department}|${item.productName}|${item.colour}|${item.size}|${item.toonLabel}`.toLowerCase();
      if (!map.has(key)) {
        map.set(key, item);
      }
    }
    const list = Array.from(map.values());

    // Sort distinct items
    list.sort((a, b) => {
      // 1. Toon Label sorting:
      // 1st Preference: Toon Label (Ascending order)
      // 2nd Preference: Colour (Ascending order)
      // 3rd Preference: Size (Ascending order: 32, 34, 36, 38... or S, M, L, XL...)
      if (sortField === 'toonLabel') {
        const diff = compareToonColourAndSize(a, b);
        return sortAsc ? diff : -diff;
      }

      // 2. Colour sorting: Sorts by Colour first, and within each colour, sorts by Size order (32, 34, 36... or S, M, L, XL, 2XL...)
      if (sortField === 'colour') {
        const diff = compareColourAndSize(a, b);
        return sortAsc ? diff : -diff;
      }

      // 3. Size sorting: Sorts by standard garment/numeric size order, and then by Colour, then Toon Label
      if (sortField === 'size') {
        const sizeDiff = compareSizes(a.size, b.size);
        if (sizeDiff !== 0) {
          return sortAsc ? sizeDiff : -sizeDiff;
        }
        const colDiff = (a.colour || '').localeCompare(b.colour || '');
        if (colDiff !== 0) {
          return sortAsc ? colDiff : -colDiff;
        }
        const toonDiff = (a.toonLabel || '').localeCompare(b.toonLabel || '', undefined, { numeric: true });
        return sortAsc ? toonDiff : -toonDiff;
      }

      // 4. Total Stock Unit sorting
      if (sortField === 'totalQuantity') {
        const statsA = getItemStoreStats(a, items);
        const statsB = getItemStoreStats(b, items);
        return sortAsc ? statsA.totalUnits - statsB.totalUnits : statsB.totalUnits - statsA.totalUnits;
      }

      // 5. Total Store Qty sorting
      if (sortField === 'store') {
        const statsA = getItemStoreStats(a, items);
        const statsB = getItemStoreStats(b, items);
        return sortAsc ? statsA.totalStoreQty - statsB.totalStoreQty : statsB.totalStoreQty - statsA.totalStoreQty;
      }

      const valA = a[sortField];
      const valB = b[sortField];

      if (typeof valA === 'number' && typeof valB === 'number') {
        return sortAsc ? valA - valB : valB - valA;
      }
      return sortAsc
        ? String(valA || '').localeCompare(String(valB || ''))
        : String(valB || '').localeCompare(String(valA || ''));
    });

    return list;
  }, [filteredItems, items, sortField, sortAsc]);

  // Overall totals
  const totalStockQuantity = useMemo(
    () => items.reduce((acc, curr) => acc + (Number(curr.stockQuantity) || 0), 0),
    [items]
  );

  const totalValuation = useMemo(
    () =>
      items.reduce(
        (acc, curr) =>
          acc + (Number(curr.stockQuantity) || 0) * (Number(curr.sellingPrice) || 0),
        0
      ),
    [items]
  );

  // Matrix view data for Toon Label
  const matrixData = useMemo(() => {
    if (searchState.type !== 'TOON' && viewMode !== 'matrix') return [];
    return generateToonLabelMatrix(filteredItems);
  }, [filteredItems, searchState.type, viewMode]);

  // Pre-calculate shuffled matrix data with stock redistribution
  const shuffledResult = useMemo(() => {
    return shuffleMatrixData(matrixData);
  }, [matrixData]);

  // Group matrixData by (toonLabel, colour) so repeated colors (e.g. 10 sizes of Blue) are visually merged
  const matrixGroups = useMemo(() => {
    const groups: {
      colour: string;
      toonLabel: string;
      startIndex: number;
      endIndex: number;
      count: number;
    }[] = [];
    let cur: { colour: string; toonLabel: string; startIndex: number; endIndex: number; count: number } | null = null;
    matrixData.forEach((row, idx) => {
      const key = `${(row.toonLabel || '').trim()}___${(row.colour || '').trim()}`;
      const prevKey = cur ? `${cur.toonLabel.trim()}___${cur.colour.trim()}` : '';
      if (!cur || key !== prevKey) {
        cur = { colour: row.colour, toonLabel: row.toonLabel, startIndex: idx, endIndex: idx, count: 1 };
        groups.push(cur);
      } else {
        cur.endIndex = idx;
        cur.count++;
      }
    });
    return groups;
  }, [matrixData]);

  const groupInfoMap = useMemo(() => {
    const map = new Map<number, { isFirst: boolean; count: number; groupIndex: number }>();
    matrixGroups.forEach((g, gIdx) => {
      for (let i = g.startIndex; i <= g.endIndex; i++) {
        map.set(i, {
          isFirst: i === g.startIndex,
          count: g.count,
          groupIndex: gIdx,
        });
      }
    });
    return map;
  }, [matrixGroups]);

  // Totals for original store matrix
  const matrixTotals = useMemo(() => {
    return matrixData.reduce(
      (acc, r) => {
        acc.kootapalli += r.kootapalli;
        acc.karur += r.karur;
        acc.salem += r.salem;
        acc.namakkal += r.namakkal;
        acc.kumbakonam += r.kumbakonam;
        acc.thiruvannamalai += r.thiruvannamalai;
        acc.mallur += r.mallur;
        acc.gmFashionsWarehouse += r.gmFashionsWarehouse;
        acc.sizeTotal += r.sizeTotal;
        return acc;
      },
      {
        kootapalli: 0,
        karur: 0,
        salem: 0,
        namakkal: 0,
        kumbakonam: 0,
        thiruvannamalai: 0,
        mallur: 0,
        gmFashionsWarehouse: 0,
        sizeTotal: 0,
      }
    );
  }, [matrixData]);

  // Totals for shuffled store matrix
  const shuffledTotals = useMemo(() => {
    return shuffledResult.shuffledRows.reduce(
      (acc, r) => {
        acc.kootapalli += r.kootapalli;
        acc.karur += r.karur;
        acc.salem += r.salem;
        acc.namakkal += r.namakkal;
        acc.kumbakonam += r.kumbakonam;
        acc.thiruvannamalai += r.thiruvannamalai;
        acc.mallur += r.mallur;
        acc.gmFashionsWarehouse += r.gmFashionsWarehouse;
        acc.sizeTotal += r.sizeTotal;
        return acc;
      },
      {
        kootapalli: 0,
        karur: 0,
        salem: 0,
        namakkal: 0,
        kumbakonam: 0,
        thiruvannamalai: 0,
        mallur: 0,
        gmFashionsWarehouse: 0,
        sizeTotal: 0,
      }
    );
  }, [shuffledResult.shuffledRows]);

  // Helper to render cell in the Shuffled Matrix with clear distinction for borrowed / donated / untouched
  const renderShuffledMatrixCell = (
    storeKey: MatrixStoreKey,
    shuffledRow: ShuffledMatrixRow
  ) => {
    const cellInfo = shuffledRow.cellInfo[storeKey];

    // 1. If stock was 0 and borrowed (redirected) from another store
    if (cellInfo.isRedirected) {
      return (
        <span
          className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-md bg-purple-100/90 text-purple-900 font-extrabold border-2 border-purple-400 shadow-xs ring-1 ring-purple-300"
          title={`Borrowed 1 unit from ${cellInfo.sourceStoreName} (${cellInfo.sourceShortcut})`}
        >
          <span className="text-xs text-purple-950">{cellInfo.shuffledQty}</span>
          <span className="text-[10px] px-1.5 py-0.5 rounded bg-purple-700 text-white font-black uppercase tracking-wide shadow-2xs">
            {cellInfo.sourceShortcut}
          </span>
        </span>
      );
    }

    // 2. If this store donated (lent) stock to a 0-stock store
    if (cellInfo.donatedCount > 0) {
      return (
        <span
          className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-amber-50/80 text-slate-800 font-bold border border-amber-300"
          title={`Lent ${cellInfo.donatedCount} unit(s) to 0-stock store(s)`}
        >
          <span className="text-emerald-700">{cellInfo.shuffledQty}</span>
          <span className="text-[9.5px] font-bold text-amber-700 bg-amber-100 px-1 py-0.2 rounded">
            (-{cellInfo.donatedCount})
          </span>
        </span>
      );
    }

    // 3. Regular store with existing stock (untouched)
    if (cellInfo.shuffledQty > 0) {
      return (
        <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
          {cellInfo.shuffledQty}
        </span>
      );
    }

    // 4. Store remains 0
    return <span className="text-slate-300">0</span>;
  };

  const getColourBadgeClass = (colour: string) => {
    const c = (colour || '').toUpperCase().trim();
    if (/BLUE|NAVY|ROYAL|SKY|DENIM|INDIGO/.test(c)) return 'bg-blue-100 text-blue-800 border-blue-300';
    if (/GREEN|MINT|OLIVE|PISTA|SAGE|BOTTLE|EMERALD/.test(c)) return 'bg-emerald-100 text-emerald-800 border-emerald-300';
    if (/YELLOW|GOLD|MUSTARD|LEMON/.test(c)) return 'bg-amber-100 text-amber-900 border-amber-300';
    if (/RED|MAROON|WINE|BURGUNDY|CRIMSON/.test(c)) return 'bg-rose-100 text-rose-800 border-rose-300';
    if (/ORANGE|PEACH|CORAL|RUST/.test(c)) return 'bg-orange-100 text-orange-800 border-orange-300';
    if (/PURPLE|VIOLET|LAVENDER|LILAC|PLUM/.test(c)) return 'bg-purple-100 text-purple-800 border-purple-300';
    if (/PINK|ROSE|MAGENTA/.test(c)) return 'bg-pink-100 text-pink-800 border-pink-300';
    if (/BLACK|CHARCOAL|GREY|GRAY|MELANGE/.test(c)) return 'bg-slate-200 text-slate-800 border-slate-300';
    if (/WHITE|CREAM|BEIGE|IVORY|KHAKI/.test(c)) return 'bg-amber-50 text-amber-900 border-amber-200';
    return 'bg-indigo-100 text-indigo-800 border-indigo-200';
  };

  const handleSort = (field: keyof MappedInventoryItem) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      // For Toon Label, Colour, and Size, default to ascending
      setSortAsc(field === 'toonLabel' || field === 'colour' || field === 'size' ? true : false);
    }
  };

  const toggleRowExpansion = (itemId: string) => {
    setExpandedItemId((prev) => (prev === itemId ? null : itemId));
  };

  return (
    <div id="rough-view-section" className="space-y-4">
      {/* Header & View Switcher Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-200">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-lg font-bold text-slate-900 tracking-tight flex items-center gap-2">
              <Boxes className="w-5 h-5 text-indigo-600" />
              <span>Inventory Stock Details</span>
            </h2>
            <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
              {distinctItems.length} Products ({items.length} records)
            </span>
          </div>
          <p className="text-xs text-slate-500 mt-0.5">
            {searchState.type === 'EAN' ? (
              <span>
                Filtered by EAN:{' '}
                <strong className="font-mono text-slate-800">{searchState.query}</strong>. Click
                any row to view available store stock below.
              </span>
            ) : (
              <span>
                Filtered by Toon Label:{' '}
                <strong className="font-mono text-slate-800">{searchState.query}</strong>. Click
                any row to view available store stock below.
              </span>
            )}
          </p>
        </div>

        {/* View Mode Controls & Manual Excel Download Button */}
        <div className="flex flex-wrap items-center gap-2">
          {searchState.type === 'TOON' && (
            <div className="flex items-center rounded-lg bg-slate-100 p-0.5 border border-slate-200 text-xs">
              <button
                type="button"
                id="view-mode-table-btn"
                onClick={() => setViewMode('table')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md font-semibold transition-all cursor-pointer ${
                  viewMode === 'table'
                    ? 'bg-white text-emerald-700 shadow-xs border border-slate-200/90'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Boxes className="w-3.5 h-3.5" />
                <span>Inventory Table</span>
              </button>

              <button
                type="button"
                id="view-mode-matrix-btn"
                onClick={() => setViewMode('matrix')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md font-semibold transition-all cursor-pointer ${
                  viewMode === 'matrix'
                    ? 'bg-white text-indigo-700 shadow-xs border border-slate-200/90'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Grid3X3 className="w-3.5 h-3.5" />
                <span>Store Matrix</span>
              </button>
            </div>
          )}

          {/* Manual Excel Download Button */}
          <button
            type="button"
            id="download-excel-btn"
            onClick={handleDownload}
            className="flex items-center gap-2 px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white text-xs font-semibold shadow-xs transition-all shrink-0 cursor-pointer"
            title="Click to download Excel file (.xlsx)"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Download Excel</span>
            <Download className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Overview Stock Cards */}
      <StoreStockCards
        items={items}
        selectedFilterStore={activeStoreFilter}
        onSelectStoreFilter={setActiveStoreFilter}
      />

      {/* Table Filter Controls */}
      <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 pt-1">
        <div className="flex items-center gap-2">
          <div className="relative w-full sm:w-80">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              id="table-quick-search-input"
              value={tableSearch}
              onChange={(e) => setTableSearch(e.target.value)}
              placeholder="Search by product, color, size, brand..."
              className="w-full bg-white border border-slate-300 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500"
            />
          </div>
          {activeStoreFilter !== 'ALL' && (
            <button
              type="button"
              onClick={() => setActiveStoreFilter('ALL')}
              className="text-xs text-indigo-700 hover:text-indigo-900 bg-indigo-50 border border-indigo-200 px-2 py-1 rounded-md font-medium cursor-pointer"
            >
              Store: {activeStoreFilter} ✕
            </button>
          )}
        </div>

        <div className="text-xs text-slate-500 text-right flex items-center justify-end gap-3">
          <span className="text-slate-400">Total Valuation:</span>
          <span className="font-mono text-slate-800 font-bold">
            ₹{totalValuation.toLocaleString()}
          </span>
          <span className="text-slate-300">•</span>
          <span className="text-slate-400">Units:</span>
          <span className="font-mono text-emerald-700 font-bold">
            {totalStockQuantity.toLocaleString()} units
          </span>
        </div>
      </div>

      {/* PRIMARY CONSOLIDATED TABLE VIEW */}
      {viewMode === 'table' && (
        <div
          id="consolidated-inventory-table-wrapper"
          className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-xs"
        >
          <table
            id="consolidated-inventory-table"
            className="w-full text-left text-xs border-collapse"
          >
            <thead>
              <tr className="bg-slate-50 text-slate-600 font-bold uppercase text-[11px] tracking-wider border-b border-slate-200 select-none">
                <th
                  onClick={() => handleSort('department')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900"
                >
                  <div className="flex items-center gap-1">
                    <span>Product Type</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('productName')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900"
                >
                  <div className="flex items-center gap-1">
                    <span>Department</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('brand')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900"
                >
                  <div className="flex items-center gap-1">
                    <span>Class</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('colour')}
                  className={`px-3.5 py-3 cursor-pointer select-none transition-colors ${
                    sortField === 'colour'
                      ? 'text-indigo-700 bg-indigo-50/70 font-extrabold'
                      : 'hover:text-slate-900'
                  }`}
                  title="Sort by Colour order (with Size order S, M, L... within each colour)"
                >
                  <div className="flex items-center gap-1">
                    <span>Color</span>
                    <ArrowUpDown className={`w-3 h-3 ${sortField === 'colour' ? 'text-indigo-600' : 'text-slate-400'}`} />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('size')}
                  className={`px-3.5 py-3 cursor-pointer select-none text-center transition-colors ${
                    sortField === 'size'
                      ? 'text-indigo-700 bg-indigo-50/70 font-extrabold'
                      : 'hover:text-slate-900'
                  }`}
                  title="Sort by Size order (S, M, L, XL, 2XL...)"
                >
                  <div className="flex items-center justify-center gap-1">
                    <span>Size</span>
                    <ArrowUpDown className={`w-3 h-3 ${sortField === 'size' ? 'text-indigo-600' : 'text-slate-400'}`} />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('fabric')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900"
                >
                  <div className="flex items-center gap-1">
                    <span>Fabric</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('ean')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900 font-mono"
                >
                  <div className="flex items-center gap-1">
                    <span>EAN Code</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('toonLabel')}
                  className={`px-3.5 py-3 cursor-pointer select-none font-mono transition-colors ${
                    sortField === 'toonLabel'
                      ? 'text-indigo-700 bg-indigo-50/70 font-extrabold'
                      : 'hover:text-slate-900'
                  }`}
                  title="1st: Toon Label (Ascending), 2nd: Colour (Ascending), 3rd: Size (32, 34, 36... or S, M, L...)"
                >
                  <div className="flex items-center gap-1">
                    <span>Toon Label</span>
                    <ArrowUpDown
                      className={`w-3 h-3 ${
                        sortField === 'toonLabel' ? 'text-indigo-600' : 'text-slate-400'
                      }`}
                    />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('sellingPrice')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900 text-right"
                >
                  <div className="flex items-center justify-end gap-1">
                    <span>Price</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('store')}
                  className="px-3.5 py-3 cursor-pointer hover:text-slate-900 text-center"
                >
                  <div className="flex items-center justify-center gap-1">
                    <span>Total Store Qty</span>
                    <ArrowUpDown className="w-3 h-3 text-slate-400" />
                  </div>
                </th>
                <th
                  onClick={() => handleSort('totalQuantity')}
                  className="px-3.5 py-3 cursor-pointer hover:text-emerald-800 text-right bg-emerald-50/70 font-bold text-emerald-800"
                >
                  <div className="flex items-center justify-end gap-1">
                    <span>Total Stock Unit</span>
                    <ArrowUpDown className="w-3 h-3 text-emerald-600" />
                  </div>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
              {distinctItems.length === 0 ? (
                <tr>
                  <td colSpan={11} className="py-12 text-center text-slate-400">
                    No products matching current filter criteria.
                  </td>
                </tr>
              ) : (
                distinctItems.map((item) => {
                  const stats = getItemStoreStats(item, filteredItems);
                  const isExpanded = expandedItemId === item.id;

                  return (
                    <Fragment key={item.id}>
                      <tr
                        onClick={() => toggleRowExpansion(item.id)}
                        className={`transition-colors cursor-pointer select-none ${
                          isExpanded
                            ? 'bg-emerald-50/60 border-l-4 border-l-emerald-600'
                            : 'hover:bg-slate-50/80 border-l-4 border-l-transparent'
                        }`}
                      >
                        {/* Product Type */}
                        <td className="px-3.5 py-3 font-semibold text-slate-800">
                          {item.department}
                        </td>

                        {/* Department */}
                        <td className="px-3.5 py-3 text-slate-900 font-bold">
                          {item.productName}
                        </td>

                        {/* Class */}
                        <td className="px-3.5 py-3 text-slate-700">
                          <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200 text-[11px] font-medium">
                            {item.brand}
                          </span>
                        </td>

                        {/* Color */}
                        <td className="px-3.5 py-3 text-slate-700 font-medium">
                          {item.colour}
                        </td>

                        {/* Size */}
                        <td className="px-3.5 py-3 text-center font-mono font-bold text-slate-800">
                          <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200">
                            {item.size}
                          </span>
                        </td>

                        {/* Fabric */}
                        <td className="px-3.5 py-3 text-slate-600">
                          {item.fabric}
                        </td>

                        {/* EAN Code */}
                        <td className="px-3.5 py-3 font-mono font-bold text-indigo-700">
                          {item.ean}
                        </td>

                        {/* Toon Label */}
                        <td className="px-3.5 py-3 font-mono text-slate-700">
                          {item.toonLabel || '—'}
                        </td>

                        {/* Price */}
                        <td className="px-3.5 py-3 text-right font-mono font-bold text-slate-800">
                          ₹{(item.sellingPrice || 0).toLocaleString()}
                        </td>

                        {/* Total Store Qty */}
                        <td className="px-3.5 py-3 text-center font-mono font-bold">
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              toggleRowExpansion(item.id);
                            }}
                            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-bold transition-all cursor-pointer ${
                              isExpanded
                                ? 'bg-amber-500 text-white shadow-xs'
                                : stats.totalStoreQty > 0
                                ? 'bg-amber-50 text-amber-800 border border-amber-300 hover:bg-amber-100'
                                : 'bg-slate-100 text-slate-400 border border-slate-200'
                            }`}
                            title={
                              stats.totalStoreQty > 0
                                ? `Click to view breakdown across ${stats.totalStoreQty} stores`
                                : 'No store has stock'
                            }
                          >
                            <Store className="w-3.5 h-3.5" />
                            <span>{stats.totalStoreQty}</span>
                            {stats.totalStoreQty > 0 &&
                              (isExpanded ? (
                                <ChevronUp className="w-3 h-3" />
                              ) : (
                                <ChevronDown className="w-3 h-3" />
                              ))}
                          </button>
                        </td>

                        {/* Total Stock Unit */}
                        <td className="px-3.5 py-3 text-right font-mono font-extrabold text-sm bg-emerald-50/40 text-emerald-800">
                          <span className="px-2.5 py-1 rounded-md bg-emerald-100 border border-emerald-300 shadow-2xs">
                            {stats.totalUnits}
                          </span>
                        </td>
                      </tr>

                      {/* INLINE EXPANDED VIEW: Renders directly below the clicked data row */}
                      {isExpanded && (
                        <tr className="bg-slate-50/90 border-b border-emerald-200">
                          <td colSpan={11} className="p-3 sm:p-4">
                            <div className="bg-white rounded-xl border border-emerald-200/90 p-4 sm:p-5 shadow-xs space-y-3">
                              {/* Sub-header Bar */}
                              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 pb-2.5 border-b border-slate-200">
                                <div className="flex items-center gap-2">
                                  <div className="p-1.5 rounded-lg bg-emerald-100 text-emerald-800">
                                    <Store className="w-4 h-4" />
                                  </div>
                                  <div>
                                    <div className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                                      <span>Available Store Stock</span>
                                      <span className="px-2 py-0.5 rounded-full text-[11px] font-extrabold bg-amber-100 text-amber-800 border border-amber-300">
                                        {stats.totalStoreQty} {stats.totalStoreQty === 1 ? 'Store' : 'Stores'}
                                      </span>
                                    </div>
                                    <div className="text-[11px] text-slate-500 font-medium mt-0.5">
                                      {item.productName} • Size <strong className="text-slate-800">{item.size}</strong> • Color <strong className="text-slate-800">{item.colour}</strong> • Fabric <strong className="text-slate-800">{item.fabric}</strong>
                                    </div>
                                  </div>
                                </div>

                                <div className="flex items-center gap-2 self-start sm:self-auto bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200">
                                  <span className="text-xs text-emerald-800 font-semibold">Total Stock Unit:</span>
                                  <span className="text-sm font-mono font-extrabold text-emerald-800">
                                    {stats.totalUnits} units
                                  </span>
                                </div>
                              </div>

                              {/* Separate Store Boxes: ONLY AVAILABLE STORES (qty > 0) */}
                              {stats.availableStores.length > 0 ? (
                                <div className="flex flex-wrap gap-3 items-stretch pt-1">
                                  {stats.availableStores.map((st) => (
                                    <div
                                      key={st.storeName}
                                      className="bg-[#f0fdf4] border border-[#a7f3d0] rounded-xl px-4 py-3 text-center flex flex-col items-center justify-center min-w-[130px] sm:min-w-[145px] flex-1 max-w-[200px] shadow-xs hover:shadow-sm hover:border-emerald-400 hover:-translate-y-0.5 transition-all"
                                    >
                                      <div className="text-[11px] sm:text-xs font-bold text-slate-800 uppercase tracking-wider mb-1.5 flex items-center justify-center gap-1">
                                        <span>{st.storeName}</span>
                                        {st.storeName === 'GM FASHIONS WAREHOUSE' && (
                                          <span className="text-[9px] px-1 py-0.2 rounded font-bold uppercase bg-emerald-200 text-emerald-800">
                                            WH
                                          </span>
                                        )}
                                      </div>
                                      <div className="flex items-baseline gap-1">
                                        <span className="text-2xl sm:text-3xl font-extrabold font-mono text-[#059669] leading-none">
                                          {st.qty}
                                        </span>
                                        <span className="text-xs font-medium text-slate-500">units</span>
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <div className="text-xs text-slate-400 italic py-2 text-center">
                                  No stores currently have available stock for this item.
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* VIEW 2: TOON LABEL STORE MATRIX */}
      {viewMode === 'matrix' && (
        <div className="space-y-6">
          {/* 1. ORIGINAL STORE MATRIX (Always visible in Matrix mode) */}
          <div className="space-y-2.5">
            <div className="flex items-center justify-between gap-3 px-1">
              <div className="flex items-center gap-2">
                <span className="inline-block w-2.5 h-2.5 rounded-full bg-indigo-600"></span>
                <h3 className="text-sm font-bold text-slate-800 tracking-tight">
                  Original Store Matrix
                </h3>
                <span className="text-[11px] text-slate-500 font-medium">
                  (Actual live inventory across all stores)
                </span>
              </div>
              <div className="text-xs text-slate-500 font-medium">
                Total Rows: <span className="font-bold text-slate-800">{matrixData.length}</span>
              </div>
            </div>

            <div
              id="toon-matrix-table-wrapper"
              className="overflow-x-auto rounded-xl border border-slate-200 bg-white shadow-xs"
            >
              <table
                id="toon-matrix-table"
                className="w-full text-left text-xs border-collapse"
              >
                <thead>
                  <tr className="bg-slate-50 text-slate-500 font-bold uppercase text-[11px] tracking-wider border-b border-slate-200">
                    <th className="px-3.5 py-3">Color</th>
                    <th className="px-3.5 py-3 font-mono">Toon Label</th>
                    <th className="px-3.5 py-3 text-center">Size</th>
                    <th className="px-3.5 py-3 text-right">Kootapalli</th>
                    <th className="px-3.5 py-3 text-right">Karur</th>
                    <th className="px-3.5 py-3 text-right">Salem</th>
                    <th className="px-3.5 py-3 text-right">Namakkal</th>
                    <th className="px-3.5 py-3 text-right">Kumbakonam</th>
                    <th className="px-3.5 py-3 text-right">Thiruvannamalai</th>
                    <th className="px-3.5 py-3 text-right">Mallur</th>
                    <th className="px-3.5 py-3 text-right">GM Fashions WH</th>
                    <th className="px-3.5 py-3 text-right font-bold text-indigo-700 bg-indigo-50/70">
                      Size Total
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
                  {matrixData.length === 0 ? (
                    <tr>
                      <td colSpan={12} className="py-8 text-center text-slate-400">
                        No matrix data available.
                      </td>
                    </tr>
                  ) : (
                    matrixData.map((row, idx) => {
                      const gInfo = groupInfoMap.get(idx);
                      const isFirst = gInfo ? gInfo.isFirst : true;
                      const count = gInfo ? gInfo.count : 1;

                      return (
                        <tr
                          key={`orig-matrix-${idx}`}
                          className={`hover:bg-slate-50 transition-colors ${
                            isFirst && idx > 0 ? 'border-t-2 border-slate-200' : ''
                          }`}
                        >
                          {isFirst && (
                            <td
                              rowSpan={count}
                              className="px-3.5 py-3 text-center align-middle border-r border-slate-200 bg-white font-medium"
                            >
                              <span
                                className={`inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold border shadow-2xs ${getColourBadgeClass(
                                  row.colour
                                )}`}
                              >
                                {row.colour}
                              </span>
                            </td>
                          )}
                          {isFirst && (
                            <td
                              rowSpan={count}
                              className="px-3.5 py-3 text-center align-middle font-mono font-bold text-indigo-700 bg-indigo-50/30 border-r border-slate-200"
                            >
                              {row.toonLabel}
                            </td>
                          )}
                          <td className="px-3.5 py-2.5 text-center font-mono font-bold text-slate-800 border-r border-slate-100">
                            <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200">
                              {row.size}
                            </span>
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.kootapalli > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.kootapalli}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.karur > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.karur}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.salem > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.salem}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.namakkal > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.namakkal}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.kumbakonam > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.kumbakonam}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.thiruvannamalai > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.thiruvannamalai}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.mallur > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.mallur}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {row.gmFashionsWarehouse > 0 ? (
                              <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">
                                {row.gmFashionsWarehouse}
                              </span>
                            ) : (
                              <span className="text-slate-300">0</span>
                            )}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono font-bold text-indigo-700 bg-indigo-50/50">
                            {row.sizeTotal}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
                {matrixData.length > 0 && (
                  <tfoot className="bg-amber-50/70 border-t-2 border-slate-300 font-bold text-xs text-slate-900">
                    <tr>
                      <td className="px-3.5 py-3 uppercase tracking-wider font-extrabold text-amber-900">
                        TOTAL
                      </td>
                      <td className="px-3.5 py-3 text-center text-slate-400 font-mono">-</td>
                      <td className="px-3.5 py-3 text-center text-slate-400 font-mono">-</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.kootapalli}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.karur}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.salem}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.namakkal}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.kumbakonam}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.thiruvannamalai}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.mallur}</td>
                      <td className="px-3.5 py-3 text-right font-mono">{matrixTotals.gmFashionsWarehouse}</td>
                      <td className="px-3.5 py-3 text-right font-mono font-black text-amber-950 bg-amber-100/80">
                        {matrixTotals.sizeTotal}
                      </td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </div>

          {/* 2. SHUFFLE BUTTON ONLY (Directly below Store Matrix) */}
          {matrixData.length > 0 && (
            <div
              id="matrix-shuffle-panel"
              className="flex flex-wrap items-center justify-between gap-3 py-2 px-1"
            >
              <div>
                {!isShuffled ? (
                  <button
                    type="button"
                    id="btn-shuffle-stocks"
                    onClick={() => setIsShuffled(true)}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white font-bold text-sm shadow-md transition-all cursor-pointer"
                  >
                    <Shuffle className="w-4 h-4" />
                    <span>Shuffle</span>
                  </button>
                ) : (
                  <button
                    type="button"
                    id="btn-reset-shuffle"
                    onClick={() => setIsShuffled(false)}
                    className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg bg-white hover:bg-slate-100 active:scale-95 text-slate-700 font-bold text-sm border border-slate-300 shadow-xs transition-all cursor-pointer"
                  >
                    <RotateCcw className="w-4 h-4 text-slate-500" />
                    <span>Hide Shuffled Matrix</span>
                  </button>
                )}
              </div>

              {isShuffled && (
                <div className="flex items-center gap-2 text-xs font-semibold text-purple-900 bg-purple-50 border border-purple-200 px-3 py-1.5 rounded-lg">
                  <span className="inline-block w-2 h-2 rounded-full bg-purple-600"></span>
                  <span>Shuffled Matrix Active ({shuffledResult.totalRedirectedUnits} units borrowed)</span>
                </div>
              )}
            </div>
          )}

          {/* 3. SHUFFLED MATRIX (Visible below when Shuffle is clicked) */}
          {isShuffled && matrixData.length > 0 && (
            <div id="shuffled-matrix-section" className="space-y-2.5 pt-2 border-t-2 border-dashed border-purple-200">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 px-1">
                <div className="flex items-center gap-2">
                  <span className="inline-block w-3 h-3 rounded-full bg-purple-600 animate-pulse"></span>
                  <h3 className="text-sm font-extrabold text-purple-900 tracking-tight flex items-center gap-2">
                    <span>Shuffled Store Matrix</span>
                    <span className="text-xs px-2.5 py-0.5 rounded-full bg-purple-100 text-purple-800 border border-purple-300 font-bold">
                      Stock Balanced
                    </span>
                  </h3>
                </div>
                <div className="text-xs text-purple-800 font-medium">
                  Borrowed stock highlighted in <strong className="text-purple-950 bg-purple-100 px-1.5 py-0.5 rounded border border-purple-300">Purple with Store Code (e.g., 1 KR, 1 KP, 1 MLR)</strong>
                </div>
              </div>

              <div
                id="shuffled-toon-matrix-table-wrapper"
                className="overflow-x-auto rounded-xl border-2 border-purple-300 bg-white shadow-md ring-1 ring-purple-100"
              >
                <table
                  id="shuffled-toon-matrix-table"
                  className="w-full text-left text-xs border-collapse"
                >
                  <thead>
                    <tr className="bg-gradient-to-r from-purple-50 via-indigo-50/50 to-purple-50 text-purple-900 font-bold uppercase text-[11px] tracking-wider border-b border-purple-200">
                      <th className="px-3.5 py-3">Color</th>
                      <th className="px-3.5 py-3 font-mono">Toon Label</th>
                      <th className="px-3.5 py-3 text-center">Size</th>
                      <th className="px-3.5 py-3 text-right">Kootapalli</th>
                      <th className="px-3.5 py-3 text-right">Karur</th>
                      <th className="px-3.5 py-3 text-right">Salem</th>
                      <th className="px-3.5 py-3 text-right">Namakkal</th>
                      <th className="px-3.5 py-3 text-right">Kumbakonam</th>
                      <th className="px-3.5 py-3 text-right">Thiruvannamalai</th>
                      <th className="px-3.5 py-3 text-right">Mallur</th>
                      <th className="px-3.5 py-3 text-right">GM Fashions WH</th>
                      <th className="px-3.5 py-3 text-right font-bold text-purple-900 bg-purple-100/80">
                        Size Total
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-purple-100 text-xs text-slate-700">
                    {shuffledResult.shuffledRows.map((row, idx) => {
                      const gInfo = groupInfoMap.get(idx);
                      const isFirst = gInfo ? gInfo.isFirst : true;
                      const count = gInfo ? gInfo.count : 1;

                      return (
                        <tr
                          key={`shuffled-matrix-${idx}`}
                          className={`hover:bg-purple-50/30 transition-colors ${
                            isFirst && idx > 0 ? 'border-t-2 border-purple-200' : ''
                          }`}
                        >
                          {isFirst && (
                            <td
                              rowSpan={count}
                              className="px-3.5 py-3 text-center align-middle border-r border-purple-100 bg-white font-medium"
                            >
                              <span
                                className={`inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold border shadow-2xs ${getColourBadgeClass(
                                  row.colour
                                )}`}
                              >
                                {row.colour}
                              </span>
                            </td>
                          )}
                          {isFirst && (
                            <td
                              rowSpan={count}
                              className="px-3.5 py-3 text-center align-middle font-mono font-bold text-purple-900 bg-purple-50/40 border-r border-purple-100"
                            >
                              {row.toonLabel}
                            </td>
                          )}
                          <td className="px-3.5 py-2.5 text-center font-mono font-bold text-slate-800 border-r border-purple-100">
                            <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200">
                              {row.size}
                            </span>
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('kootapalli', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('karur', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('salem', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('namakkal', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('kumbakonam', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('thiruvannamalai', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('mallur', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono">
                            {renderShuffledMatrixCell('gmFashionsWarehouse', row)}
                          </td>
                          <td className="px-3.5 py-2.5 text-right font-mono font-bold text-purple-900 bg-purple-50/70">
                            {row.sizeTotal}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  {shuffledResult.shuffledRows.length > 0 && (
                    <tfoot className="bg-purple-100/70 border-t-2 border-purple-300 font-bold text-xs text-purple-950">
                      <tr>
                        <td className="px-3.5 py-3 uppercase tracking-wider font-black text-purple-950">
                          SHUFFLED TOTAL
                        </td>
                        <td className="px-3.5 py-3 text-center text-purple-400 font-mono">-</td>
                        <td className="px-3.5 py-3 text-center text-purple-400 font-mono">-</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.kootapalli}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.karur}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.salem}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.namakkal}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.kumbakonam}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.thiruvannamalai}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.mallur}</td>
                        <td className="px-3.5 py-3 text-right font-mono">{shuffledTotals.gmFashionsWarehouse}</td>
                        <td className="px-3.5 py-3 text-right font-mono font-black text-purple-950 bg-purple-200/90">
                          {shuffledTotals.sizeTotal}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
