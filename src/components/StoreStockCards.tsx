import { StoreName, MappedInventoryItem, ALL_STORES } from '../types/inventory';
import { Store, Layers, Building2 } from 'lucide-react';

interface StoreStockCardsProps {
  items: MappedInventoryItem[];
  selectedFilterStore: StoreName | 'ALL';
  onSelectStoreFilter: (store: StoreName | 'ALL') => void;
}

export function StoreStockCards({
  items,
  selectedFilterStore,
  onSelectStoreFilter,
}: StoreStockCardsProps) {
  // Aggregate stock by store
  const storeCounts: Record<string, number> = {};
  let totalStock = 0;

  ALL_STORES.forEach((s) => {
    storeCounts[s] = 0;
  });

  items.forEach((item) => {
    const qty = Number(item.stockQuantity) || 0;
    totalStock += qty;
    if (storeCounts[item.store] !== undefined) {
      storeCounts[item.store] += qty;
    } else {
      storeCounts[item.store] = qty;
    }
  });

  return (
    <div id="store-stock-cards-container" className="space-y-3">
      <div className="flex items-center justify-between">
        <h3 className="text-xs font-bold uppercase tracking-widest text-slate-500 flex items-center gap-1.5">
          <Building2 className="w-3.5 h-3.5 text-indigo-600" />
          <span>Store-Wise Stock Counts</span>
        </h3>
        <span className="text-[11px] text-slate-400 hidden sm:inline">
          Click any card to filter the table below
        </span>
      </div>

      {/* Grid of Store Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 lg:grid-cols-9 gap-2.5">
        {/* Total Aggregate Card */}
        <button
          type="button"
          id="store-card-all"
          onClick={() => onSelectStoreFilter('ALL')}
          className={`p-3 rounded-lg border text-left transition-all cursor-pointer shadow-sm ${
            selectedFilterStore === 'ALL'
              ? 'bg-indigo-50/70 border-indigo-600 ring-2 ring-indigo-200'
              : 'bg-white border-slate-200 hover:border-indigo-300'
          }`}
        >
          <div className="flex items-center justify-between text-[10px] uppercase font-bold text-slate-400">
            <span className="truncate">Total Stock</span>
            <Layers className="w-3 h-3 text-indigo-600" />
          </div>
          <div className="mt-1 text-xl font-bold text-indigo-600 font-mono">
            {totalStock.toLocaleString()}
          </div>
          <div className="text-[10px] text-slate-500 mt-0.5">
            {items.length} records
          </div>
        </button>

        {/* Individual Store Cards */}
        {ALL_STORES.map((store) => {
          const count = storeCounts[store] || 0;
          const isSelected = selectedFilterStore === store;

          return (
            <button
              key={store}
              type="button"
              id={`store-card-${store.toLowerCase()}`}
              onClick={() => onSelectStoreFilter(store)}
              className={`p-3 rounded-lg border text-left transition-all cursor-pointer shadow-sm ${
                isSelected
                  ? 'bg-indigo-50/70 border-indigo-600 ring-2 ring-indigo-200'
                  : 'bg-white border-slate-200 hover:border-indigo-300'
              }`}
            >
              <div className="flex items-center justify-between text-[10px] uppercase font-bold text-slate-400">
                <span className="truncate">{store}</span>
                <Store className="w-3 h-3 text-slate-400" />
              </div>
              <div
                className={`mt-1 text-xl font-bold font-mono ${
                  count > 0 ? 'text-slate-800' : 'text-slate-400'
                }`}
              >
                {count.toLocaleString()}
              </div>
              <div className="text-[10px] text-slate-500 mt-0.5">
                {count > 0 ? (
                  <span className="text-emerald-700 font-medium">In Stock</span>
                ) : (
                  <span>Out of Stock</span>
                )}
              </div>
            </button>
          );
        })}
      </div>
    </div>
  );
}
