import { useState, useEffect, KeyboardEvent } from 'react';
import {
  Tag,
  ArrowRight,
  AlertCircle,
  RotateCcw,
  X,
  Search,
} from 'lucide-react';
import { StoreName, ALL_STORES } from '../types/inventory';
import { StoreDropdown } from './StoreDropdown';

const DEFAULT_ALL_STORES: StoreName[] = [...ALL_STORES];

interface SalesSearchSectionProps {
  onExecuteSearch: (
    query: string,
    stores: StoreName[],
    isAllStores: boolean
  ) => void;
  onClearSearch?: () => void;
  availableToonLabels?: string[];
  initialQuery?: string;
  initialStores?: StoreName[];
  initialIsAllStores?: boolean;
}

export function SalesSearchSection({
  onExecuteSearch,
  onClearSearch,
  availableToonLabels = [],
  initialQuery = '',
  initialStores = DEFAULT_ALL_STORES,
  initialIsAllStores = true,
}: SalesSearchSectionProps) {
  const [toonQuery, setToonQuery] = useState(initialQuery);
  const [toonStores, setToonStores] = useState<StoreName[]>(initialStores);
  const [isToonAllStores, setIsToonAllStores] = useState(initialIsAllStores);
  const [toonError, setToonError] = useState<string | null>(null);

  // Synchronize if initial values change (e.g. from cache restoration or reset)
  // Use primitive dependencies to prevent infinite re-renders
  const storesKey = (initialStores || DEFAULT_ALL_STORES).join(',');

  useEffect(() => {
    setToonQuery(initialQuery || '');
  }, [initialQuery]);

  useEffect(() => {
    if (initialStores && initialStores.length > 0) {
      setToonStores(initialStores);
    } else {
      setToonStores(DEFAULT_ALL_STORES);
    }
  }, [storesKey]);

  useEffect(() => {
    setIsToonAllStores(initialIsAllStores ?? true);
  }, [initialIsAllStores]);

  const handleClearAll = () => {
    setToonQuery('');
    setToonError(null);
    setToonStores(DEFAULT_ALL_STORES);
    setIsToonAllStores(true);
    if (onClearSearch) {
      onClearSearch();
    }
  };

  const handleSubmitToon = () => {
    const trimmed = toonQuery.trim();
    if (!trimmed) {
      setToonError('Please enter a Toon Label.');
      return;
    }
    if (toonStores.length === 0) {
      setToonError('Please select at least one store location.');
      return;
    }

    setToonError(null);
    onExecuteSearch(trimmed, toonStores, isToonAllStores);
  };

  return (
    <div
      id="sales-search-criteria-section"
      className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-sm space-y-6"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <span>Sales Inventory Lookup</span>
            <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-100">
              Step 2
            </span>
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Search sales by Toon Label with store filtering.
          </p>
        </div>

        <button
          type="button"
          id="clear-all-sales-search-btn"
          onClick={handleClearAll}
          className="self-start sm:self-auto flex items-center gap-1.5 px-3 py-1.5 text-xs text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors cursor-pointer border border-transparent hover:border-slate-200"
          title="Reset search criteria"
        >
          <RotateCcw className="w-3.5 h-3.5" />
          <span>Clear All Criteria</span>
        </button>
      </div>

      {/* TOON LABEL SEARCH ROW */}
      <div id="sales-toon-search-row" className="space-y-2">
        <div className="flex items-center justify-between">
          <label
            htmlFor="sales-toon-input-field"
            className="text-xs font-bold uppercase tracking-widest text-slate-500 flex items-center gap-1.5"
          >
            <Tag className="w-3.5 h-3.5 text-indigo-600" />
            <span>Search: Toon Label (Partial or Exact Match)</span>
          </label>
        </div>

        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
          {/* Input field */}
          <div className="relative flex-1">
            <input
              type="text"
              id="sales-toon-input-field"
              value={toonQuery}
              placeholder="Enter Toon Label"
              onChange={(e) => {
                setToonQuery(e.target.value);
                if (toonError) setToonError(null);
              }}
              onKeyDown={(e: KeyboardEvent<HTMLInputElement>) => {
                if (e.key === 'Enter') handleSubmitToon();
              }}
              className="w-full bg-white border border-slate-300 rounded-lg px-3.5 py-2.5 text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-all"
            />
            {toonQuery && (
              <button
                type="button"
                id="clear-sales-toon-input-btn"
                onClick={() => {
                  setToonQuery('');
                  setToonError(null);
                }}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 rounded-md cursor-pointer"
                title="Clear Toon Label"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {/* Store Multi-Select Dropdown */}
          <StoreDropdown
            idPrefix="sales-toon"
            selectedStores={toonStores}
            onChange={(stores, isAll) => {
              setToonStores(stores);
              setIsToonAllStores(isAll);
              if (toonError) setToonError(null);
            }}
            isAllStores={isToonAllStores}
          />

          {/* OK Submit Button */}
          <button
            type="button"
            id="sales-toon-search-ok-btn"
            onClick={handleSubmitToon}
            className="flex items-center justify-center gap-2 px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white text-sm font-bold rounded-lg shadow-md shadow-indigo-100 transition-all shrink-0 cursor-pointer"
          >
            <span>OK</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>

        {/* Inline Error */}
        {toonError && (
          <div
            id="sales-toon-validation-error"
            className="flex items-center gap-1.5 text-xs text-rose-600 mt-1 font-medium"
          >
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            <span>{toonError}</span>
          </div>
        )}
      </div>
    </div>
  );
}
