import { useState, useMemo } from 'react';
import {
  FileSpreadsheet,
  Download,
  Grid3X3,
  Search,
  Boxes,
  Store,
  CheckCircle2,
} from 'lucide-react';
import {
  MappedSalesItem,
  SalesMatrixRow,
  StoreName,
  ALL_STORES,
} from '../types/inventory';
import {
  generateSalesMatrix,
  exportSalesMatrixExcel,
  exportDetailedSalesExcel,
} from '../utils/salesParser';
import { matchStoreName, tryMatchStoreName } from '../utils/excelParser';

interface SalesViewTableProps {
  items: MappedSalesItem[];
  searchQuery: string;
  selectedStores: StoreName[];
}

export function SalesViewTable({
  items,
  searchQuery,
  selectedStores,
}: SalesViewTableProps) {
  const [activeStoreFilter, setActiveStoreFilter] = useState<StoreName | 'ALL'>('ALL');
  const [tableSearch, setTableSearch] = useState('');
  const [viewMode, setViewMode] = useState<'table' | 'matrix'>('matrix');

  // Filter items based on selected stores, search query, and inner table search
  const filteredItems = useMemo(() => {
    let result = [...items];

    if (activeStoreFilter !== 'ALL') {
      result = result.filter((i) => i.store === activeStoreFilter);
    } else if (selectedStores.length > 0 && selectedStores.length < ALL_STORES.length) {
      result = result.filter((i) => selectedStores.includes(i.store));
    }

    if (tableSearch.trim()) {
      const q = tableSearch.toLowerCase().trim();
      result = result.filter(
        (i) =>
          i.toonLabel.toLowerCase().includes(q) ||
          i.productName.toLowerCase().includes(q) ||
          i.colour.toLowerCase().includes(q) ||
          i.size.toLowerCase().includes(q) ||
          i.brand.toLowerCase().includes(q) ||
          i.store.toLowerCase().includes(q) ||
          (i.ean && i.ean.toLowerCase().includes(q))
      );
    }

    return result;
  }, [items, activeStoreFilter, selectedStores, tableSearch]);

  // Generate Matrix rows from filtered items
  const matrixRows = useMemo(() => {
    return generateSalesMatrix(filteredItems);
  }, [filteredItems]);

  // Group matrixRows by (toonLabel, colour) so repeated colors (e.g. 10 sizes of Blue) are visually merged
  const matrixGroups = useMemo(() => {
    const groups: {
      colour: string;
      toonLabel: string;
      startIndex: number;
      endIndex: number;
      count: number;
    }[] = [];
    let cur: { colour: string; toonLabel: string; startIndex: number; endIndex: number; count: number } | null = null;
    matrixRows.forEach((row, idx) => {
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
  }, [matrixRows]);

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

  // Compute Store Matrix column totals
  const matrixTotals = useMemo(() => {
    return matrixRows.reduce(
      (acc, r) => {
        acc.kootapalli += r.kootapalli;
        acc.karur += r.karur;
        acc.salem += r.salem;
        acc.namakkal += r.namakkal;
        acc.kumbakonam += r.kumbakonam;
        acc.thiruvannamalai += r.thiruvannamalai;
        acc.mallur += r.mallur;
        acc.gmFashionsWarehouse += r.gmFashionsWarehouse;
        acc.totalSales += r.totalSales;
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
        totalSales: 0,
      }
    );
  }, [matrixRows]);

  // Store sales summary counts for the top cards
  const storeSalesSummary = useMemo(() => {
    const summary: Record<string, number> = {
      Kootapalli: 0,
      Karur: 0,
      Salem: 0,
      Namakkal: 0,
      Kumbakonam: 0,
      Thiruvannamalai: 0,
      Mallur: 0,
      'GM Fashions Warehouse': 0,
    };

    for (const item of items) {
      const rawStore = String(item.store || '').trim();
      const matched = tryMatchStoreName(rawStore);
      const storeKey = (item.store in summary ? item.store : null) || matched || matchStoreName(rawStore, 'Kootapalli');
      if (storeKey in summary) {
        summary[storeKey] += Number(item.salesQty) || 0;
      }
    }

    const totalAll = Object.values(summary).reduce((a, b) => a + b, 0);
    return { summary, totalAll };
  }, [items]);

  // Download Handler based on the ACTIVE view mode requested by user!
  const [downloadNotice, setDownloadNotice] = useState<{
    show: boolean;
    filename: string;
  } | null>(null);

  const handleDownload = () => {
    let filename = '';
    if (viewMode === 'matrix') {
      filename = exportSalesMatrixExcel(filteredItems, searchQuery);
    } else {
      filename = exportDetailedSalesExcel(filteredItems, searchQuery);
    }
    setDownloadNotice({ show: true, filename });
    setTimeout(() => {
      setDownloadNotice(null);
    }, 4500);
  };

  return (
    <div id="sales-view-table-container" className="space-y-4">
      {/* Header & View Mode Switcher */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-200">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-lg font-bold text-slate-900 tracking-tight flex items-center gap-2">
              <Boxes className="w-5 h-5 text-indigo-600" />
              <span>Sales Report Analysis</span>
            </h2>
            <span className="text-[11px] font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200">
              {matrixRows.length} Product Variants ({filteredItems.length} Sales Lines)
            </span>
          </div>
          <p className="text-xs text-slate-500 mt-0.5">
            Filtered by Toon Label: <strong className="font-mono text-slate-800">{searchQuery}</strong>.
            Showing store-wise sales quantities.
          </p>
        </div>

        {/* View mode toggle & Manual Download Button */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Toggle between Store Matrix and Inventory Table */}
          <div className="flex items-center rounded-lg bg-slate-100 p-0.5 border border-slate-200 text-xs">
            <button
              type="button"
              id="sales-view-mode-matrix-btn"
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

            <button
              type="button"
              id="sales-view-mode-table-btn"
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
          </div>

          {/* Download Excel Button - Downloads according to active view */}
          <button
            type="button"
            id="sales-download-excel-btn"
            onClick={handleDownload}
            className="flex items-center gap-2 px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white text-xs font-semibold shadow-xs transition-all shrink-0 cursor-pointer"
            title={`Download ${viewMode === 'matrix' ? 'Store Matrix' : 'Sales Table'} Excel`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Download Excel ({viewMode === 'matrix' ? 'Matrix' : 'Table'})</span>
            <Download className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Store Sales Summary Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 lg:grid-cols-9 gap-2">
        <button
          type="button"
          onClick={() => setActiveStoreFilter('ALL')}
          className={`p-2.5 rounded-lg border text-left transition-all cursor-pointer ${
            activeStoreFilter === 'ALL'
              ? 'bg-indigo-50/70 border-indigo-500 ring-1 ring-indigo-500 text-indigo-950'
              : 'bg-white border-slate-200 hover:border-slate-300 text-slate-700'
          }`}
        >
          <div className="text-[10px] font-bold uppercase tracking-wider text-slate-500">
            ALL STORES
          </div>
          <div className="text-base font-extrabold text-indigo-700 mt-0.5">
            {storeSalesSummary.totalAll.toLocaleString()}
          </div>
          <div className="text-[10px] text-slate-400">Total Sales</div>
        </button>

        {ALL_STORES.map((store) => {
          const qty = storeSalesSummary.summary[store] || 0;
          const isSelected = activeStoreFilter === store;
          return (
            <button
              key={store}
              type="button"
              onClick={() => setActiveStoreFilter(isSelected ? 'ALL' : store)}
              className={`p-2.5 rounded-lg border text-left transition-all cursor-pointer ${
                isSelected
                  ? 'bg-indigo-50/70 border-indigo-500 ring-1 ring-indigo-500 text-indigo-950'
                  : 'bg-white border-slate-200 hover:border-slate-300 text-slate-700'
              }`}
            >
              <div className="text-[10px] font-bold uppercase tracking-wider text-slate-500 truncate">
                {store === 'GM Fashions Warehouse' ? 'Warehouse' : (store === 'Kootapalli' ? 'Koottappalli' : store)}
              </div>
              <div className={`text-base font-bold mt-0.5 ${qty > 0 ? 'text-slate-900' : 'text-slate-400'}`}>
                {qty.toLocaleString()}
              </div>
              <div className="text-[10px] text-slate-400">Sales Qty</div>
            </button>
          );
        })}
      </div>

      {/* Quick Search inside Results */}
      <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 pt-1">
        <div className="flex items-center gap-2">
          <div className="relative w-full sm:w-80">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              id="sales-table-quick-search-input"
              value={tableSearch}
              onChange={(e) => setTableSearch(e.target.value)}
              placeholder="Search by color, size, description..."
              className="w-full bg-white border border-slate-300 rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
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

        <div className="text-xs text-slate-500 text-right flex items-center justify-end gap-2">
          <span>Active View:</span>
          <span className="font-semibold text-slate-800">
            {viewMode === 'matrix' ? 'Store Matrix Table' : 'Inventory Table Details'}
          </span>
        </div>
      </div>

      {/* VIEW 1: STORE MATRIX (Exactly requested: store-wise sales count with Toon Label) */}
      {viewMode === 'matrix' ? (
        <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left border-collapse">
              <thead className="bg-slate-100/90 text-slate-700 font-bold border-b border-slate-200 uppercase tracking-wider text-[11px]">
                <tr>
                  <th className="px-3 py-3 text-left">Color</th>
                  <th className="px-3 py-3 text-left font-mono">Toon Label</th>
                  <th className="px-3 py-3 text-center">Size</th>
                  <th className="px-3 py-3 text-right">Koottappalli</th>
                  <th className="px-3 py-3 text-right">Karur</th>
                  <th className="px-3 py-3 text-right">Salem</th>
                  <th className="px-3 py-3 text-right">Namakkal</th>
                  <th className="px-3 py-3 text-right">Kumbakonam</th>
                  <th className="px-3 py-3 text-right">Thiruvannamalai</th>
                  <th className="px-3 py-3 text-right">Mallur</th>
                  <th className="px-3 py-3 text-right">GM Fashions WH</th>
                  <th className="px-3.5 py-3 text-right font-extrabold bg-indigo-50/70 text-indigo-900 border-l border-indigo-100">
                    Total Sales
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {matrixRows.length > 0 ? (
                  matrixRows.map((row, idx) => {
                    const gInfo = groupInfoMap.get(idx);
                    const isFirst = gInfo ? gInfo.isFirst : true;
                    const count = gInfo ? gInfo.count : 1;

                    return (
                      <tr
                        key={`${row.colour}-${row.toonLabel}-${row.size}-${idx}`}
                        className={`hover:bg-slate-50/80 transition-colors ${
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
                            className="px-3.5 py-3 text-center align-middle font-mono font-bold text-emerald-800 bg-emerald-50/30 border-r border-slate-200"
                          >
                            {row.toonLabel}
                          </td>
                        )}
                        <td className="px-3 py-2.5 text-center font-semibold text-slate-700 border-r border-slate-100">
                          <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200">
                            {row.size}
                          </span>
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.kootapalli > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.kootapalli}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.karur > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.karur}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.salem > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.salem}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.namakkal > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.namakkal}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.kumbakonam > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.kumbakonam}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.thiruvannamalai > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.thiruvannamalai}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.mallur > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.mallur}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono">
                          {row.gmFashionsWarehouse > 0 ? (
                            <span className="inline-block px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold border border-emerald-200">
                              {row.gmFashionsWarehouse}
                            </span>
                          ) : (
                            <span className="text-slate-300">0</span>
                          )}
                        </td>
                        <td className="px-3.5 py-2.5 text-right font-mono font-extrabold text-emerald-800 bg-emerald-50/60 border-l border-emerald-100">
                          {row.totalSales}
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td colSpan={12} className="px-6 py-8 text-center text-slate-500">
                      No sales records found for this query in the selected stores.
                    </td>
                  </tr>
                )}
              </tbody>
              {/* Sticky / Permanent TOTAL row at the bottom */}
              {matrixRows.length > 0 && (
                <tfoot className="bg-slate-100 font-bold border-t-2 border-slate-300 text-slate-900 text-xs">
                  <tr>
                    <td className="px-3 py-3 uppercase tracking-wider font-extrabold">TOTAL</td>
                    <td className="px-3 py-3 font-mono text-slate-500">-</td>
                    <td className="px-3 py-3 text-center text-slate-500">-</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.kootapalli}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.karur}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.salem}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.namakkal}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.kumbakonam}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.thiruvannamalai}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.mallur}</td>
                    <td className="px-3 py-3 text-right font-mono">{matrixTotals.gmFashionsWarehouse}</td>
                    <td className="px-3.5 py-3 text-right font-mono font-black text-indigo-900 bg-indigo-100/70 border-l border-indigo-200">
                      {matrixTotals.totalSales}
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      ) : (
        /* VIEW 2: INVENTORY TABLE (Detailed row-by-row table) */
        <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left border-collapse">
              <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200 uppercase tracking-wider text-[11px]">
                <tr>
                  <th className="px-3.5 py-3 text-left font-mono">Toon Label</th>
                  <th className="px-3.5 py-3 text-left">Product / Description</th>
                  <th className="px-3.5 py-3 text-left">Color</th>
                  <th className="px-3.5 py-3 text-center">Size</th>
                  <th className="px-3.5 py-3 text-left">Store</th>
                  <th className="px-3.5 py-3 text-right font-bold text-indigo-800">Sales Qty</th>
                  {filteredItems.some((i) => i.mrp > 0) && (
                    <th className="px-3.5 py-3 text-right">MRP</th>
                  )}
                  {filteredItems.some((i) => !!i.ean) && (
                    <th className="px-3.5 py-3 text-left font-mono">EAN</th>
                  )}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredItems.length > 0 ? (
                  filteredItems.map((item) => (
                    <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                      <td className="px-3.5 py-2.5 font-mono font-semibold text-indigo-700">
                        {item.toonLabel}
                      </td>
                      <td className="px-3.5 py-2.5 font-medium text-slate-800">
                        {item.productName}
                      </td>
                      <td className="px-3.5 py-2.5 text-slate-700">
                        {item.colour}
                      </td>
                      <td className="px-3.5 py-2.5 text-center font-semibold text-slate-700">
                        {item.size}
                      </td>
                      <td className="px-3.5 py-2.5 text-slate-800 font-medium">
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[11px]">
                          <Store className="w-3 h-3 text-slate-400" />
                          <span>{item.store}</span>
                        </span>
                      </td>
                      <td className="px-3.5 py-2.5 text-right font-mono font-bold text-slate-900">
                        {item.salesQty}
                      </td>
                      {filteredItems.some((i) => i.mrp > 0) && (
                        <td className="px-3.5 py-2.5 text-right font-mono text-slate-600">
                          ₹{item.mrp || 0}
                        </td>
                      )}
                      {filteredItems.some((i) => !!i.ean) && (
                        <td className="px-3.5 py-2.5 font-mono text-slate-500 text-[11px]">
                          {item.ean || '-'}
                        </td>
                      )}
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={8} className="px-6 py-8 text-center text-slate-500">
                      No matching sales items found.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Floating Download Toast Notification */}
      {downloadNotice && downloadNotice.show && (
        <div
          id="sales-export-toast-notification"
          className="fixed bottom-6 right-6 z-50 bg-white border border-emerald-400 shadow-xl rounded-xl p-4 max-w-sm text-xs animate-bounce-short"
        >
          <div className="flex items-start gap-3">
            <div className="w-8 h-8 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-600 shrink-0">
              <CheckCircle2 className="w-4 h-4" />
            </div>
            <div className="space-y-1">
              <p className="font-bold text-slate-900">Color-Highlighted Excel Exported</p>
              <p className="text-slate-600">
                Downloaded with styled headers, store badges &amp; total row highlights.
              </p>
              <p className="font-mono text-[11px] text-emerald-700 font-semibold truncate">
                {downloadNotice.filename}
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
