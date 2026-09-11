import { useState, useMemo } from 'react';
import {
  TrendingUp,
  ArrowLeft,
  AlertTriangle,
  RotateCw,
} from 'lucide-react';
import {
  UploadedFileInfo,
  MappedSalesItem,
  StoreName,
  ALL_STORES,
} from '../types/inventory';
import { SalesDataUpload } from './SalesDataUpload';
import { SalesSearchSection } from './SalesSearchSection';
import { SalesViewTable } from './SalesViewTable';

export interface SalesSearchCriteria {
  query: string;
  stores: StoreName[];
  isAllStores: boolean;
}

interface ReportViewProps {
  onSwitchToInstant: () => void;
  onBackToHome?: () => void;
  salesItems: MappedSalesItem[];
  salesFileInfo: UploadedFileInfo | null;
  onDataLoaded: (items: MappedSalesItem[], info: UploadedFileInfo) => void;
  onResetData: () => void;
  searchState: SalesSearchCriteria | null;
  onExecuteSearch: (
    query: string,
    stores: StoreName[],
    isAllStores: boolean
  ) => void;
  onClearSearch: () => void;
}

export function ReportView({
  onSwitchToInstant,
  onBackToHome,
  salesItems,
  salesFileInfo,
  onDataLoaded,
  onResetData,
  searchState,
  onExecuteSearch,
  onClearSearch,
}: ReportViewProps) {
  // When sales report is loaded from Step 1
  const handleDataLoaded = (items: MappedSalesItem[], info: UploadedFileInfo) => {
    onDataLoaded(items, info);
    onClearSearch();
  };

  // Reset sales data
  const handleResetData = () => {
    onResetData();
    onClearSearch();
  };

  // Available unique Toon Labels for suggestions
  const availableToonLabels = useMemo(() => {
    const set = new Set<string>();
    for (const item of salesItems) {
      if (item.toonLabel) {
        set.add(item.toonLabel);
      }
    }
    return Array.from(set).sort();
  }, [salesItems]);

  // Filtered sales items based on Toon Label (partial or exact) and store selection
  const searchResults = useMemo(() => {
    if (!searchState) return [];

    const q = searchState.query.toLowerCase();
    const isAll = searchState.isAllStores || searchState.stores.length === ALL_STORES.length;

    return salesItems.filter((item) => {
      // Check Toon Label match (partial or exact)
      const matchesQuery =
        !q ||
        item.toonLabel.toLowerCase().includes(q) ||
        item.colour.toLowerCase().includes(q) ||
        item.size.toLowerCase().includes(q) ||
        item.productName.toLowerCase().includes(q);

      if (!matchesQuery) return false;

      // Check store filter
      if (!isAll && !searchState.stores.includes(item.store)) {
        return false;
      }

      return true;
    });
  }, [salesItems, searchState]);

  return (
    <div id="sales-report-workflow-container" className="space-y-6">
      {/* Top Breadcrumb & Switcher Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          {onBackToHome && (
            <button
              type="button"
              id="sales-back-to-home-btn"
              onClick={onBackToHome}
              className="inline-flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-900 transition-colors cursor-pointer font-medium"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Back to Home</span>
            </button>
          )}

          <div className="text-xs text-slate-500">
            <span className="text-indigo-600 font-semibold">
              Sales Report Workflow
            </span>
          </div>
        </div>

        {/* Quick Stock Report Switch */}
        <button
          type="button"
          id="switch-to-stock-report-btn"
          onClick={onSwitchToInstant}
          className="self-start sm:self-auto text-xs px-3.5 py-2 rounded-lg border border-slate-200 hover:border-slate-300 hover:bg-slate-50 text-slate-700 font-medium transition-colors cursor-pointer flex items-center gap-1.5"
        >
          <span>Switch to Stock Report</span>
        </button>
      </div>

      {/* STEP 1: Upload Sales Report */}
      <SalesDataUpload
        onDataLoaded={handleDataLoaded}
        fileInfo={salesFileInfo}
        onReset={handleResetData}
      />

      {/* STEP 2: Reveal Toon Label Search ONLY after successful upload */}
      {salesFileInfo && (
        <SalesSearchSection
          onExecuteSearch={onExecuteSearch}
          onClearSearch={onClearSearch}
          availableToonLabels={availableToonLabels}
          initialQuery={searchState?.query}
          initialStores={searchState?.stores}
          initialIsAllStores={searchState?.isAllStores}
        />
      )}

      {/* STEP 3: Results Area - Inventory Table & Store Matrix with store-wise sales count */}
      {searchState && (
        <>
          {searchResults.length > 0 ? (
            <SalesViewTable
              items={searchResults}
              searchQuery={searchState.query}
              selectedStores={searchState.stores}
            />
          ) : (
            <div
              id="sales-no-results-box"
              className="bg-white border border-slate-200 rounded-xl p-8 text-center space-y-3 shadow-sm"
            >
              <div className="w-12 h-12 rounded-xl bg-amber-50 border border-amber-200 text-amber-600 flex items-center justify-center mx-auto">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="text-base font-bold text-slate-900">
                No matching sales records found
              </h3>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                No sales records matched your Toon Label query &quot;{searchState.query}&quot; in the selected stores.
                Try searching with a broader prefix like &quot;6-G028&quot; or check your uploaded file.
              </p>
            </div>
          )}
        </>
      )}
    </div>
  );
}
