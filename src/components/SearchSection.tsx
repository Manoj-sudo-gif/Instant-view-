import { useState, useEffect, KeyboardEvent } from 'react';
import {
  Barcode,
  Tag,
  ArrowRight,
  AlertCircle,
  CheckCircle2,
  RotateCcw,
  X,
} from 'lucide-react';
import { StoreName, ALL_STORES, SearchType, SearchState } from '../types/inventory';
import { StoreDropdown } from './StoreDropdown';

interface SearchSectionProps {
  onExecuteSearch: (
    type: SearchType,
    query: string,
    stores: StoreName[],
    isAllStores: boolean
  ) => void;
  onClearSearch?: () => void;
  availableToonLabels?: string[];
  availableEans?: string[];
  initialSearchState?: SearchState | null;
}

const DEFAULT_ALL_STORES: StoreName[] = [...ALL_STORES];

export function SearchSection({
  onExecuteSearch,
  onClearSearch,
  availableToonLabels = [],
  availableEans = [],
  initialSearchState = null,
}: SearchSectionProps) {
  // Row 1: EAN Search State
  const [eanQuery, setEanQuery] = useState(
    initialSearchState?.type === 'EAN' ? initialSearchState.query : ''
  );
  const [eanStores, setEanStores] = useState<StoreName[]>(
    initialSearchState?.type === 'EAN' ? initialSearchState.selectedStores : DEFAULT_ALL_STORES
  );
  const [isEanAllStores, setIsEanAllStores] = useState(
    initialSearchState?.type === 'EAN' ? initialSearchState.isAllStores : true
  );
  const [eanError, setEanError] = useState<string | null>(null);

  // Row 2: Toon Label Search State
  const [toonQuery, setToonQuery] = useState(
    initialSearchState?.type === 'TOON' ? initialSearchState.query : ''
  );
  const [toonStores, setToonStores] = useState<StoreName[]>(
    initialSearchState?.type === 'TOON' ? initialSearchState.selectedStores : DEFAULT_ALL_STORES
  );
  const [isToonAllStores, setIsToonAllStores] = useState(
    initialSearchState?.type === 'TOON' ? initialSearchState.isAllStores : true
  );
  const [toonError, setToonError] = useState<string | null>(null);

  // Primitive key to avoid infinite re-renders on object reference shifts
  const searchStateKey = initialSearchState
    ? `${initialSearchState.type}_${initialSearchState.query}_${(initialSearchState.selectedStores || []).join(',')}_${initialSearchState.isAllStores}_${initialSearchState.timestamp || 0}`
    : '';

  // Synchronize when initialSearchState changes
  useEffect(() => {
    if (initialSearchState) {
      if (initialSearchState.type === 'EAN') {
        setEanQuery(initialSearchState.query);
        setEanStores(initialSearchState.selectedStores);
        setIsEanAllStores(initialSearchState.isAllStores);
      } else if (initialSearchState.type === 'TOON') {
        setToonQuery(initialSearchState.query);
        setToonStores(initialSearchState.selectedStores);
        setIsToonAllStores(initialSearchState.isAllStores);
      }
    }
  }, [searchStateKey]);

  // Common Clear All Handler
  const handleClearAll = () => {
    setEanQuery('');
    setToonQuery('');
    setEanError(null);
    setToonError(null);
    setEanStores([...ALL_STORES]);
    setIsEanAllStores(true);
    setToonStores([...ALL_STORES]);
    setIsToonAllStores(true);
    if (onClearSearch) {
      onClearSearch();
    }
  };

  // Validate and submit EAN
  const handleSubmitEan = () => {
    const trimmed = eanQuery.trim();
    if (!trimmed) {
      setEanError('Please enter an EAN code or barcode to search.');
      return;
    }
    const cleanDigits = trimmed.replace(/[\s_\-]/g, '');
    if (cleanDigits.length < 2) {
      setEanError('EAN code must contain at least 2 digits/characters.');
      return;
    }
    if (eanStores.length === 0) {
      setEanError('Please select at least one store location.');
      return;
    }

    setEanError(null);
    onExecuteSearch('EAN', trimmed, eanStores, isEanAllStores);
  };

  // Submit Toon Label
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
    onExecuteSearch('TOON', trimmed, toonStores, isToonAllStores);
  };

  const isEanValid = eanQuery.trim().length >= 2;

  return (
    <div
      id="search-criteria-section"
      className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-sm space-y-6"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <span>Instant Inventory Lookup</span>
            <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-100">
              Step 2
            </span>
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Search by EAN code or Toon Label with store filtering.
          </p>
        </div>

        {/* Common Clear Search Button */}
        <button
          type="button"
          id="instant-lookup-clear-all-btn"
          onClick={handleClearAll}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 active:bg-slate-300 text-slate-700 border border-slate-300 text-xs font-semibold transition-all cursor-pointer self-start sm:self-auto"
          title="Clear all search inputs, store filters, and results"
        >
          <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
          <span>Clear Search</span>
        </button>
      </div>

      {/* SEARCH ROW 1: EAN CODE */}
      <div id="ean-search-row" className="space-y-2">
        <div className="flex items-center justify-between">
          <label
            htmlFor="ean-input-field"
            className="text-xs font-bold uppercase tracking-widest text-slate-500 flex items-center gap-1.5"
          >
            <Barcode className="w-3.5 h-3.5 text-indigo-600" />
            <span>Search 1: EAN Code</span>
          </label>
          <div className="flex items-center gap-2 text-xs">
            {eanQuery.length > 0 && (
              <span
                className={`font-mono text-[11px] px-1.5 py-0.5 rounded font-semibold ${
                  isEanValid
                    ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                    : 'bg-amber-50 text-amber-700 border border-amber-200'
                }`}
              >
                {eanQuery.trim().length} chars
              </span>
            )}
            {isEanValid && (
              <span className="text-emerald-600 flex items-center gap-1 text-[11px] font-medium">
                <CheckCircle2 className="w-3 h-3" /> Ready
              </span>
            )}
          </div>
        </div>

        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
          {/* EAN Input Field with inline clear button */}
          <div className="relative flex-1">
            <input
              type="text"
              id="ean-input-field"
              value={eanQuery}
              maxLength={30}
              placeholder="Enter EAN code"
              onChange={(e) => {
                setEanQuery(e.target.value);
                if (eanError) setEanError(null);
              }}
              onKeyDown={(e: KeyboardEvent<HTMLInputElement>) => {
                if (e.key === 'Enter') handleSubmitEan();
              }}
              className={`w-full bg-white border text-slate-800 placeholder-slate-400 rounded-lg pl-3.5 pr-8 py-2.5 text-sm font-mono tracking-wide focus:outline-none focus:ring-2 transition-all ${
                eanError
                  ? 'border-rose-300 focus:ring-rose-200'
                  : isEanValid
                  ? 'border-emerald-400 focus:ring-emerald-200'
                  : 'border-slate-300 focus:ring-indigo-500 focus:border-indigo-500'
              }`}
            />
            {eanQuery && (
              <button
                type="button"
                id="clear-ean-input-btn"
                onClick={() => {
                  setEanQuery('');
                  setEanError(null);
                }}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 rounded-md cursor-pointer"
                title="Clear EAN"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          {/* Store Multi-Select Dropdown next to EAN Input */}
          <StoreDropdown
            idPrefix="ean"
            selectedStores={eanStores}
            onChange={(stores, isAll) => {
              setEanStores(stores);
              setIsEanAllStores(isAll);
              if (eanError) setEanError(null);
            }}
            isAllStores={isEanAllStores}
          />

          {/* OK Submit Button matching Sleek Interface styling */}
          <button
            type="button"
            id="ean-search-ok-btn"
            onClick={handleSubmitEan}
            className="flex items-center justify-center gap-2 px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white text-sm font-bold rounded-lg shadow-md shadow-indigo-100 transition-all shrink-0 cursor-pointer"
          >
            <span>OK</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>

        {/* EAN Inline Validation Error */}
        {eanError && (
          <div
            id="ean-validation-error"
            className="flex items-center gap-1.5 text-xs text-rose-600 mt-1 font-medium"
          >
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            <span>{eanError}</span>
          </div>
        )}
      </div>

      <div className="h-px bg-slate-200 my-4" />

      {/* SEARCH ROW 2: TOON LABEL */}
      <div id="toon-search-row" className="space-y-2">
        <div className="flex items-center justify-between">
          <label
            htmlFor="toon-input-field"
            className="text-xs font-bold uppercase tracking-widest text-slate-500 flex items-center gap-1.5"
          >
            <Tag className="w-3.5 h-3.5 text-indigo-600" />
            <span>Search 2: Toon Label (Partial or Exact Match)</span>
          </label>
        </div>

        <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
          {/* Toon Label Input Field with inline clear button */}
          <div className="relative flex-1">
            <input
              type="text"
              id="toon-input-field"
              value={toonQuery}
              placeholder="Enter Toon Label"
              onChange={(e) => {
                setToonQuery(e.target.value);
                if (toonError) setToonError(null);
              }}
              onKeyDown={(e: KeyboardEvent<HTMLInputElement>) => {
                if (e.key === 'Enter') handleSubmitToon();
              }}
              className={`w-full bg-white border text-slate-800 placeholder-slate-400 rounded-lg pl-3.5 pr-8 py-2.5 text-sm font-mono tracking-wide focus:outline-none focus:ring-2 transition-all ${
                toonError
                  ? 'border-rose-300 focus:ring-rose-200'
                  : 'border-slate-300 focus:ring-indigo-500 focus:border-indigo-500'
              }`}
            />
            {toonQuery && (
              <button
                type="button"
                id="clear-toon-input-btn"
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

          {/* Store Multi-Select Dropdown next to Toon Label Input */}
          <StoreDropdown
            idPrefix="toon"
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
            id="toon-search-ok-btn"
            onClick={handleSubmitToon}
            className="flex items-center justify-center gap-2 px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white text-sm font-bold rounded-lg shadow-md shadow-indigo-100 transition-all shrink-0 cursor-pointer"
          >
            <span>OK</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>

        {/* Toon Inline Error */}
        {toonError && (
          <div
            id="toon-validation-error"
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
