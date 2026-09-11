import { useState, useRef, useEffect } from 'react';
import { ChevronDown, Check, Store } from 'lucide-react';
import { StoreName, ALL_STORES } from '../types/inventory';

interface StoreDropdownProps {
  idPrefix: string;
  selectedStores: StoreName[];
  onChange: (stores: StoreName[], isAll: boolean) => void;
  isAllStores: boolean;
}

export function StoreDropdown({
  idPrefix,
  selectedStores,
  onChange,
  isAllStores,
}: StoreDropdownProps) {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close when clicking outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setIsOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleToggleAll = () => {
    if (isAllStores || selectedStores.length === ALL_STORES.length) {
      // Uncheck all
      onChange([], false);
    } else {
      // Check all
      onChange([...ALL_STORES], true);
    }
  };

  const handleToggleStore = (store: StoreName) => {
    let updated: StoreName[];
    if (selectedStores.includes(store)) {
      updated = selectedStores.filter((s) => s !== store);
    } else {
      updated = [...selectedStores, store];
    }
    const allSelected = updated.length === ALL_STORES.length;
    onChange(updated, allSelected);
  };

  // Label text for trigger button
  const getButtonLabel = () => {
    if (isAllStores || selectedStores.length === ALL_STORES.length) {
      return 'All Stores (8)';
    }
    if (selectedStores.length === 0) {
      return 'Select Stores (0)';
    }
    const formatName = (s: StoreName) => (s === 'Kootapalli' ? 'Koottappalli' : s);
    if (selectedStores.length === 1) {
      return formatName(selectedStores[0]);
    }
    if (selectedStores.length === 2) {
      return `${formatName(selectedStores[0])}, ${formatName(selectedStores[1])}`;
    }
    return `${selectedStores.length} Stores Selected`;
  };

  return (
    <div className="relative w-full sm:w-64" ref={dropdownRef}>
      {/* Dropdown Trigger Button */}
      <button
        type="button"
        id={`${idPrefix}-store-dropdown-btn`}
        onClick={() => setIsOpen((prev) => !prev)}
        className={`w-full flex items-center justify-between gap-2 px-3.5 py-2.5 rounded-lg border text-sm font-medium transition-all cursor-pointer ${
          isOpen
            ? 'bg-white border-indigo-500 ring-2 ring-indigo-100 text-slate-900'
            : selectedStores.length === 0
            ? 'bg-rose-50 border-rose-300 text-rose-700 hover:bg-rose-100/60'
            : 'bg-slate-50 border-slate-300 text-slate-700 hover:bg-slate-100 hover:border-slate-400'
        }`}
      >
        <div className="flex items-center gap-2 truncate">
          <Store className="w-4 h-4 text-slate-500 shrink-0" />
          <span className="truncate">{getButtonLabel()}</span>
        </div>
        <ChevronDown
          className={`w-4 h-4 text-slate-400 shrink-0 transition-transform duration-200 ${
            isOpen ? 'rotate-180 text-indigo-600' : ''
          }`}
        />
      </button>

      {/* Dropdown Popup */}
      {isOpen && (
        <div
          id={`${idPrefix}-store-menu`}
          className="absolute z-50 mt-1.5 w-72 left-0 sm:right-0 sm:left-auto bg-white border border-slate-200 rounded-xl shadow-xl p-2 max-h-80 overflow-y-auto"
        >
          <div className="px-2 py-1.5 text-xs font-bold text-slate-500 tracking-wider uppercase flex items-center justify-between border-b border-slate-200 pb-2 mb-1">
            <span>Filter Store Locations</span>
            <span className="text-[11px] text-indigo-600 font-semibold font-mono">
              {selectedStores.length}/{ALL_STORES.length} selected
            </span>
          </div>

          {/* Option: "All Stores" */}
          <label
            id={`${idPrefix}-option-all-stores`}
            className="flex items-center gap-2.5 px-2.5 py-2 rounded-lg hover:bg-indigo-50 cursor-pointer text-sm font-semibold text-indigo-700 transition-colors"
          >
            <input
              type="checkbox"
              id={`${idPrefix}-checkbox-all`}
              checked={isAllStores || selectedStores.length === ALL_STORES.length}
              onChange={handleToggleAll}
              className="w-4 h-4 rounded border-slate-300 bg-white text-indigo-600 focus:ring-indigo-500 cursor-pointer"
            />
            <span className="flex-1">All Stores (Select All)</span>
            {(isAllStores || selectedStores.length === ALL_STORES.length) && (
              <Check className="w-3.5 h-3.5 text-indigo-600" />
            )}
          </label>

          <div className="h-px bg-slate-200 my-1" />

          {/* Individual Store Options */}
          <div className="space-y-0.5">
            {ALL_STORES.map((store) => {
              const isChecked = selectedStores.includes(store);
              return (
                <label
                  key={store}
                  id={`${idPrefix}-option-store-${store.toLowerCase()}`}
                  className={`flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg cursor-pointer text-sm transition-colors ${
                    isChecked
                      ? 'bg-indigo-50/80 text-indigo-900 font-medium'
                      : 'hover:bg-slate-50 text-slate-700'
                  }`}
                >
                  <input
                    type="checkbox"
                    id={`${idPrefix}-checkbox-${store.toLowerCase()}`}
                    checked={isChecked}
                    onChange={() => handleToggleStore(store)}
                    className="w-4 h-4 rounded border-slate-300 bg-white text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                  />
                  <span className="flex-1">{store === 'Kootapalli' ? 'Koottappalli' : store}</span>
                  {(store === 'GM Fashions Warehouse' || store.toLowerCase().includes('warehouse')) && (
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 border border-slate-200">
                      Central
                    </span>
                  )}
                </label>
              );
            })}
          </div>

          {/* Quick Actions Footer */}
          <div className="mt-2 pt-2 border-t border-slate-200 flex items-center justify-between text-xs px-1">
            <button
              type="button"
              id={`${idPrefix}-select-all-btn`}
              onClick={() => onChange([...ALL_STORES], true)}
              className="text-indigo-600 hover:text-indigo-700 font-semibold py-1 px-1.5 rounded hover:bg-indigo-50 cursor-pointer"
            >
              Select All
            </button>
            <button
              type="button"
              id={`${idPrefix}-clear-all-btn`}
              onClick={() => onChange([], false)}
              className="text-slate-500 hover:text-slate-700 font-semibold py-1 px-1.5 rounded hover:bg-slate-100 cursor-pointer"
            >
              Clear Selection
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
