import { useState, useMemo, useEffect } from 'react';
import {
  MappedInventoryItem,
  MappedSalesItem,
  SearchState,
  SearchType,
  StoreName,
  UploadedFileInfo,
} from './types/inventory';
import { Header } from './components/Header';
import { HomeScreen } from './components/HomeScreen';
import { DataUpload } from './components/DataUpload';
import { SearchSection } from './components/SearchSection';
import { RoughViewTable } from './components/RoughViewTable';
import { ReportView } from './components/ReportView';
import {
  filterInventory,
  exportToonLabelExcel,
  exportEanExcel,
} from './utils/excelParser';
import {
  loadSessionFromStorage,
  saveStockDataToStorage,
  saveStockSearchToStorage,
  saveSalesDataToStorage,
  saveSalesSearchToStorage,
  clearStorageSession,
  clearStockStorage,
  clearSalesStorage,
  SalesSearchState,
} from './utils/storage';
import { CheckCircle2, AlertTriangle, ArrowLeft, Heart } from 'lucide-react';

export default function App() {
  // Navigation state: 'home' | 'instant' | 'report'
  const [currentView, setCurrentView] = useState<'home' | 'instant' | 'report'>('home');

  // Parsed inventory items and file metadata (STOCK)
  const [inventoryItems, setInventoryItems] = useState<MappedInventoryItem[]>([]);
  const [fileInfo, setFileInfo] = useState<UploadedFileInfo | null>(null);

  // Parsed sales items and file metadata (SALES) - kept in App state so switching views never loses data!
  const [salesItems, setSalesItems] = useState<MappedSalesItem[]>([]);
  const [salesFileInfo, setSalesFileInfo] = useState<UploadedFileInfo | null>(null);

  // Search and results state for Stock Report
  const [searchState, setSearchState] = useState<SearchState | null>(null);
  const [searchResults, setSearchResults] = useState<MappedInventoryItem[]>([]);
  const [lastExportFilename, setLastExportFilename] = useState<string | null>(null);

  // Search state for Sales Report (standardized & persisted)
  const [salesSearchState, setSalesSearchState] = useState<SalesSearchState | null>(null);

  const [downloadToast, setDownloadToast] = useState<{
    show: boolean;
    message: string;
    filename: string;
  } | null>(null);

  // Load persisted session on initial mount
  useEffect(() => {
    loadSessionFromStorage().then((session) => {
      if (session) {
        if (session.stockItems && session.stockItems.length > 0) {
          setInventoryItems(session.stockItems);
          setFileInfo(session.stockFileInfo);
          if (session.stockSearchState) {
            setSearchState(session.stockSearchState);
            const results = filterInventory(
              session.stockItems,
              session.stockSearchState.type,
              session.stockSearchState.query,
              session.stockSearchState.selectedStores,
              session.stockSearchState.isAllStores
            );
            setSearchResults(results);
          }
          console.log(
            `%c[GM Fashions Storage Cache] Restored ${session.stockItems.length} Stock items from browser cache.`,
            'color: #10b981; font-weight: bold;'
          );
        }
        if (session.salesItems && session.salesItems.length > 0) {
          setSalesItems(session.salesItems);
          setSalesFileInfo(session.salesFileInfo);
          if (session.salesSearchState) {
            setSalesSearchState(session.salesSearchState);
          }
          console.log(
            `%c[GM Fashions Storage Cache] Restored ${session.salesItems.length} Sales items from browser cache.`,
            'color: #10b981; font-weight: bold;'
          );
        }
      }
    });
  }, []);

  // Unique Toon Labels and EANs for autocomplete & quick sample testing
  const availableToonLabels = useMemo(() => {
    const set = new Set<string>();
    inventoryItems.forEach((item) => {
      if (item.toonLabel) set.add(item.toonLabel);
    });
    return Array.from(set);
  }, [inventoryItems]);

  const availableEans = useMemo(() => {
    const set = new Set<string>();
    inventoryItems.forEach((item) => {
      if (item.ean) set.add(item.ean);
    });
    return Array.from(set);
  }, [inventoryItems]);

  // Handle stock data upload completion
  const handleDataLoaded = (items: MappedInventoryItem[], info: UploadedFileInfo) => {
    setInventoryItems(items);
    setFileInfo(info);
    saveStockDataToStorage(items, info);
    console.log(
      `%c[GM Fashions Storage Cache] Stored ${items.length} Stock items to browser cache.`,
      'color: #6366f1; font-weight: bold;'
    );
    // Clear any previous search results
    setSearchState(null);
    setSearchResults([]);
    setLastExportFilename(null);
    saveStockSearchToStorage(null);
  };

  // Handle sales data upload completion
  const handleSalesDataLoaded = (items: MappedSalesItem[], info: UploadedFileInfo) => {
    setSalesItems(items);
    setSalesFileInfo(info);
    saveSalesDataToStorage(items, info);
    console.log(
      `%c[GM Fashions Storage Cache] Stored ${items.length} Sales items to browser cache.`,
      'color: #6366f1; font-weight: bold;'
    );
  };

  // Reset / Replace data (Clear full session for both)
  const handleResetData = () => {
    setInventoryItems([]);
    setFileInfo(null);
    setSalesItems([]);
    setSalesFileInfo(null);
    setSearchState(null);
    setSearchResults([]);
    setSalesSearchState(null);
    setLastExportFilename(null);
    clearStorageSession();
    console.log(
      '%c[GM Fashions Storage Cache] Cleared all session data from browser cache.',
      'color: #ef4444; font-weight: bold;'
    );
  };

  // Clear only Stock data
  const handleResetStockData = () => {
    setInventoryItems([]);
    setFileInfo(null);
    setSearchState(null);
    setSearchResults([]);
    setLastExportFilename(null);
    clearStockStorage();
  };

  // Clear only Sales data
  const handleResetSalesData = () => {
    setSalesItems([]);
    setSalesFileInfo(null);
    setSalesSearchState(null);
    clearSalesStorage();
  };

  // Clear only search criteria and current search results
  const handleClearSearch = () => {
    setSearchState(null);
    setSearchResults([]);
    setLastExportFilename(null);
    saveStockSearchToStorage(null);
  };

  // Execute Search: Filter inventory and display Rough View (NO auto-download)
  const handleExecuteSearch = (
    type: SearchType,
    query: string,
    stores: StoreName[],
    isAllStores: boolean
  ) => {
    const results = filterInventory(
      inventoryItems,
      type,
      query,
      stores,
      isAllStores
    );

    const newState: SearchState = {
      type,
      query,
      selectedStores: stores,
      isAllStores,
      timestamp: Date.now(),
    };

    setSearchState(newState);
    setSearchResults(results);
    saveStockSearchToStorage(newState);
    // Note: Excel download will only trigger when user explicitly clicks the Download Excel button
  };

  // Handle sales search execution (Toon Label query, stores filter)
  const handleExecuteSalesSearch = (
    query: string,
    stores: StoreName[],
    isAllStores: boolean
  ) => {
    const newSalesSearch: SalesSearchState = {
      query: query.trim(),
      stores,
      isAllStores,
    };
    setSalesSearchState(newSalesSearch);
    saveSalesSearchToStorage(newSalesSearch);
  };

  // Clear sales search
  const handleClearSalesSearch = () => {
    setSalesSearchState(null);
    saveSalesSearchToStorage(null);
  };

  // Explicit manual Excel (.xlsx) download triggered ONLY upon user click
  const handleDownloadExcel = (mode?: 'mapped' | 'matrix') => {
    if (!searchState || searchResults.length === 0) return;
    let exportedFile = '';
    try {
      if (searchState.type === 'TOON' && mode === 'matrix') {
        // Strict format: [Color, Toon Label, Size, Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, Size Total]
        exportedFile = exportToonLabelExcel(searchResults, searchState.query);
      } else {
        // Mapped custom columns: Department, Product Name, Brand, Colour, Size, Fabric, EAN, Selling Price, Store, Store Quantity, Total Quantity
        exportedFile = exportEanExcel(searchResults, searchState.query);
      }
      setLastExportFilename(exportedFile);

      // Trigger user confirmation toast
      setDownloadToast({
        show: true,
        message:
          searchState.type === 'TOON' && mode === 'matrix'
            ? `Color-highlighted Toon Label stock matrix exported (${searchResults.length} rows)`
            : `Color-highlighted stock details exported (${searchResults.length} rows)`,
        filename: exportedFile,
      });

      // Hide toast after 5 seconds
      setTimeout(() => {
        setDownloadToast((prev) => (prev ? { ...prev, show: false } : null));
      }, 5000);
    } catch (err) {
      console.error('Error generating Excel file:', err);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans selection:bg-indigo-500 selection:text-white">
      {/* Global Application Header */}
      <Header
        currentView={currentView}
        onSelectView={setCurrentView}
        fileInfo={fileInfo}
        salesFileInfo={salesFileInfo}
        onResetData={handleResetData}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* VIEW 1: HOME SCREEN (Phase 1 Selection) */}
        {currentView === 'home' && (
          <HomeScreen
            onSelectInstantView={() => setCurrentView('instant')}
            onSelectReportView={() => setCurrentView('report')}
            fileInfo={fileInfo}
            salesFileInfo={salesFileInfo}
          />
        )}

        {/* VIEW 2: INSTANT VIEW DASHBOARD (Stock Report Flow) */}
        <div className={currentView === 'instant' ? 'space-y-6' : 'hidden'}>
          {/* Breadcrumb / Back button */}
          <div className="flex items-center justify-between">
            <button
              type="button"
              onClick={() => setCurrentView('home')}
              className="inline-flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-900 transition-colors cursor-pointer font-medium"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Back to Home</span>
            </button>

            <div className="text-xs text-slate-500">
              <span className="text-indigo-600 font-semibold">
                Stock Report Workflow
              </span>
            </div>
          </div>

          {/* STEP 1: Data Upload Section */}
          <DataUpload
            onDataLoaded={handleDataLoaded}
            fileInfo={fileInfo}
            onReset={handleResetStockData}
          />

          {/* STEP 2: Reveal Search Fields & Store Selection ONLY after successful upload */}
          {fileInfo && (
            <SearchSection
              onExecuteSearch={handleExecuteSearch}
              onClearSearch={handleClearSearch}
              availableToonLabels={availableToonLabels}
              availableEans={availableEans}
              initialSearchState={searchState}
            />
          )}

          {/* STEP 3: Rough View (UI Table) & Results below search area */}
          {searchState && (
            <>
              {searchResults.length > 0 ? (
                <RoughViewTable
                  items={searchResults}
                  searchState={searchState}
                  lastExportFilename={lastExportFilename}
                  onDownloadExcel={handleDownloadExcel}
                />
              ) : (
                <div
                  id="no-results-box"
                  className="bg-white border border-slate-200 rounded-xl p-8 text-center space-y-3 shadow-sm"
                >
                  <div className="w-12 h-12 rounded-xl bg-amber-50 border border-amber-200 text-amber-600 flex items-center justify-center mx-auto">
                    <AlertTriangle className="w-6 h-6" />
                  </div>
                  <h3 className="text-base font-bold text-slate-900">
                    No matching records found
                  </h3>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    No inventory matched your query &quot;{searchState.query}&quot; in the selected stores.
                    {searchState.type === 'EAN'
                      ? ' Note that EAN search requires exact 13 digits.'
                      : ' Try a partial Toon Label prefix like "6-G028".'}
                  </p>
                </div>
              )}
            </>
          )}
        </div>

        {/* VIEW 3: SALES REPORT VIEW */}
        <div className={currentView === 'report' ? 'space-y-6' : 'hidden'}>
          <ReportView
            onSwitchToInstant={() => setCurrentView('instant')}
            onBackToHome={() => setCurrentView('home')}
            salesItems={salesItems}
            salesFileInfo={salesFileInfo}
            onDataLoaded={handleSalesDataLoaded}
            onResetData={handleResetSalesData}
            searchState={salesSearchState}
            onExecuteSearch={handleExecuteSalesSearch}
            onClearSearch={handleClearSalesSearch}
          />
        </div>
      </main>

      {/* Floating Download Toast Notification */}
      {downloadToast && downloadToast.show && (
        <div
          id="export-toast-notification"
          className="fixed bottom-6 right-6 z-50 bg-white border border-emerald-300 shadow-xl rounded-xl p-4 max-w-sm text-xs animate-bounce-short"
        >
          <div className="flex items-start gap-3">
            <div className="w-8 h-8 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-center text-emerald-600 shrink-0">
              <CheckCircle2 className="w-4 h-4" />
            </div>
            <div className="space-y-1">
              <p className="font-bold text-slate-900">Color-Highlighted Excel Exported</p>
              <p className="text-slate-600">{downloadToast.message}</p>
              <p className="font-mono text-[11px] text-emerald-700 font-semibold truncate">
                {downloadToast.filename}
              </p>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer id="app-footer" className="border-t border-slate-200 bg-white py-4 mt-12">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-800">Instant View</span>
            <span>•</span>
            <span>Inventory &amp; Reporting Suite</span>
          </div>
          <div
            id="footer-made-by-manoj"
            className="flex items-center gap-1.5 font-medium text-slate-600"
          >
            <span>Made by Manoj</span>
            <Heart className="w-3.5 h-3.5 text-red-500 fill-red-500 inline shrink-0" />
          </div>
        </div>
      </footer>
    </div>
  );
}
