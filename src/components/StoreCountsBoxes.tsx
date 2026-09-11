import React, { useMemo } from 'react';
import { MappedInventoryItem } from '../types/inventory';
import { formatStoreDisplayName } from '../utils/excelParser';

interface StoreCountsBoxesProps {
  item: MappedInventoryItem;
  allItems?: MappedInventoryItem[];
  className?: string;
  selectedStore?: string | null;
  onSelectStore?: (storeName: string) => void;
}

// The 8 allowed stores strictly matching user's ERP standard
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

export function StoreCountsBoxes({
  item,
  allItems = [],
  className = '',
  selectedStore,
  onSelectStore,
}: StoreCountsBoxesProps) {
  // Aggregate store counts for this specific item / EAN
  const storeCounts: { storeName: string; qty: number; price: number }[] = useMemo(() => {
    const map = new Map<string, number>();

    // 1. Direct breakdown if already attached
    if (item.storeBreakdown && Object.keys(item.storeBreakdown).length > 0) {
      for (const [s, q] of Object.entries(item.storeBreakdown)) {
        const sDisplay = formatStoreDisplayName(s);
        if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
          map.set(sDisplay, q);
        }
      }
    }

    // 2. Also aggregate from allItems matching this EAN or item identity
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
        const sDisplay = formatStoreDisplayName(m.store);
        if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
          const prev = map.get(sDisplay) || 0;
          map.set(sDisplay, Math.max(prev, m.stockQuantity || 0));
        }
      }
    }

    // 3. Fallback to item itself if map is empty
    if (map.size === 0 && item.stockQuantity > 0) {
      const sDisplay = formatStoreDisplayName(item.store);
      if (ALLOWED_STORE_DISPLAYS.has(sDisplay)) {
        map.set(sDisplay, item.stockQuantity);
      }
    }

    // Filter strictly:
    // - Only the 8 valid stores
    // - Only stores that actually have available stock (qty > 0)
    // - Unavailable / zero stock stores are excluded as requested
    const result = Array.from(map.entries())
      .filter(([storeName, qty]) => {
        if (!ALLOWED_STORE_DISPLAYS.has(storeName)) return false;
        return qty > 0;
      })
      .map(([storeName, qty]) => ({
        storeName,
        qty,
        price: item.sellingPrice || 0,
      }));

    result.sort((a, b) => {
      const idxA = PREFERRED_ORDER.indexOf(a.storeName);
      const idxB = PREFERRED_ORDER.indexOf(b.storeName);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.storeName.localeCompare(b.storeName);
    });

    return result;
  }, [item, allItems]);

  const totalItemStock = useMemo(() => {
    return storeCounts.reduce((acc, curr) => acc + curr.qty, 0);
  }, [storeCounts]);

  return (
    <div
      id={`store-counts-container-${item.id}`}
      className={`bg-white rounded-xl border border-slate-200/90 p-4 sm:p-5 shadow-xs space-y-3.5 transition-all ${className}`}
    >
      {/* Top Header Bar matching user screenshot: "Store Counts:" on left, Pipe-separated counts on right */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 px-4 py-2.5 bg-slate-50/90 rounded-lg border border-slate-200/80">
        <div className="text-xs sm:text-sm font-bold text-slate-500 tracking-wide shrink-0">
          Store Counts:
        </div>
        <div className="text-xs sm:text-[13px] font-mono font-bold text-slate-800 tracking-wide flex flex-wrap items-center gap-1.5 sm:justify-end">
          {storeCounts.length > 0 ? (
            storeCounts.map((sc, idx) => {
              const isZero = sc.qty === 0;
              return (
                <React.Fragment key={sc.storeName}>
                  {idx > 0 && <span className="text-slate-300 font-normal px-1">|</span>}
                  <span className="hover:text-emerald-700 transition-colors">
                    {sc.storeName}:{' '}
                    <span className={isZero ? 'text-slate-500 font-bold' : 'text-slate-900 font-extrabold'}>
                      {sc.qty}
                    </span>
                  </span>
                </React.Fragment>
              );
            })
          ) : (
            <span className="text-slate-400 font-normal italic">0 stores with stock</span>
          )}
        </div>
      </div>

      {/* Sequential Store Cards matching user uploaded image (image.png) */}
      {storeCounts.length > 0 ? (
        <div className="flex flex-wrap gap-3 items-stretch">
          {storeCounts.map((sc) => {
            const isHighlighted = selectedStore === sc.storeName;
            const isZero = sc.qty === 0;
            return (
              <div
                key={sc.storeName}
                onClick={() => onSelectStore?.(sc.storeName)}
                className={`rounded-xl px-4 py-3 text-center flex flex-col items-center justify-center min-w-[130px] sm:min-w-[145px] flex-1 shadow-xs transition-all hover:shadow-sm cursor-pointer ${
                  isZero
                    ? 'bg-slate-50 border border-slate-200 text-slate-500 hover:border-slate-300'
                    : 'bg-[#f0fdf4] border border-[#a7f3d0] hover:border-emerald-400 hover:-translate-y-0.5'
                } ${
                  isHighlighted ? 'ring-2 ring-emerald-500 shadow-md bg-emerald-50' : ''
                }`}
                title={`Click to focus ${sc.storeName}`}
              >
                {/* Store Name Header */}
                <div className="text-[11px] sm:text-xs font-bold text-slate-800 uppercase tracking-wider mb-1.5 flex items-center justify-center gap-1">
                  <span>{sc.storeName}</span>
                  {sc.storeName === 'GM FASHIONS WAREHOUSE' && (
                    <span
                      className={`text-[9px] px-1 py-0.2 rounded font-bold uppercase ${
                        isZero ? 'bg-slate-200 text-slate-600' : 'bg-emerald-200 text-emerald-800'
                      }`}
                    >
                      WH
                    </span>
                  )}
                </div>

                {/* Units Count */}
                <div className="flex items-baseline justify-center gap-1 my-0.5">
                  <span
                    className={`text-2xl sm:text-3xl font-extrabold font-mono leading-none ${
                      isZero ? 'text-slate-400' : 'text-[#059669]'
                    }`}
                  >
                    {sc.qty}
                  </span>
                  <span className="text-xs font-medium text-slate-500">units</span>
                </div>

                {/* Selling Price */}
                <div className="text-xs font-semibold text-slate-600 mt-1 font-mono">
                  ₹{sc.price.toLocaleString()}
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <div className="py-6 text-center text-slate-400 text-xs italic bg-slate-50 rounded-xl border border-dashed border-slate-200">
          No stock currently recorded across any store for this item.
        </div>
      )}

      {/* Subtle Metadata line matching user image bottom text */}
      <div className="text-[11px] text-slate-500 pt-2 border-t border-slate-100 flex flex-wrap items-center gap-x-3 gap-y-1">
        <span>
          Original Wondersoft Dept:{' '}
          <strong className="text-slate-800 font-semibold">{item.productName || item.department}</strong>
        </span>
        <span className="text-slate-300">•</span>
        <span>
          Product Type:{' '}
          <strong className="text-slate-800 font-semibold">{item.department}</strong>
        </span>
        <span className="text-slate-300">•</span>
        <span>
          Brand / Class: <strong className="text-slate-700">{item.brand}</strong>
        </span>
        <span className="text-slate-300">•</span>
        <span>
          Color: <strong className="text-slate-700">{item.colour}</strong>
        </span>
        <span className="text-slate-300">•</span>
        <span>
          Size: <strong className="text-slate-700">{item.size}</strong>
        </span>
        <span className="text-slate-300">•</span>
        <span>
          Total Units:{' '}
          <strong className="text-emerald-700 font-bold">
            {totalItemStock} units
          </strong>{' '}
          across {storeCounts.length} {storeCounts.length === 1 ? 'store' : 'stores'}
        </span>
      </div>
    </div>
  );
}
