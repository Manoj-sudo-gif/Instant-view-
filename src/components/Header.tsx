import { Shirt, BarChart3, Search, Database, Layers, RotateCcw } from 'lucide-react';
import { UploadedFileInfo } from '../types/inventory';

interface HeaderProps {
  currentView: 'home' | 'instant' | 'report';
  onSelectView: (view: 'home' | 'instant' | 'report') => void;
  fileInfo: UploadedFileInfo | null;
  salesFileInfo?: UploadedFileInfo | null;
  onResetData?: () => void;
}

export function Header({
  currentView,
  onSelectView,
  fileInfo,
  salesFileInfo,
  onResetData,
}: HeaderProps) {
  return (
    <header
      id="app-header"
      className="sticky top-0 z-40 h-16 bg-white border-b border-slate-200 px-4 sm:px-8 flex items-center justify-between"
    >
      <div className="w-full flex items-center justify-between">
        {/* Brand Logo & Title with Sleek Interface underline styling */}
        <div className="flex items-center gap-6 sm:gap-8">
          <button
            type="button"
            id="brand-logo-container"
            onClick={() => onSelectView('home')}
            className="flex items-center gap-2.5 cursor-pointer text-left group"
          >
            <span className="text-xl font-bold tracking-tight text-indigo-700 underline decoration-indigo-300 decoration-2 underline-offset-4 group-hover:text-indigo-800 transition-colors">
              INSTANT VIEW
            </span>
          </button>

          {/* Primary View Switcher Navigation matching Sleek Interface buttons */}
          <nav id="primary-navigation" className="flex items-center gap-1">
            <button
              type="button"
              id="nav-stock-report-btn"
              onClick={() => onSelectView('instant')}
              className={`px-3.5 sm:px-4 py-2 rounded-lg text-sm transition-all cursor-pointer flex items-center gap-1.5 ${
                currentView === 'instant'
                  ? 'bg-indigo-50 text-indigo-700 font-semibold border border-indigo-100'
                  : 'text-slate-500 hover:bg-slate-50 font-medium'
              }`}
            >
              <span>Stock Report</span>
              {fileInfo && (
                <span className="w-2 h-2 rounded-full bg-emerald-500" />
              )}
            </button>

            <button
              type="button"
              id="nav-sales-report-btn"
              onClick={() => onSelectView('report')}
              className={`px-3.5 sm:px-4 py-2 rounded-lg text-sm transition-all cursor-pointer flex items-center gap-1.5 ${
                currentView === 'report'
                  ? 'bg-indigo-50 text-indigo-700 font-semibold border border-indigo-100'
                  : 'text-slate-500 hover:bg-slate-50 font-medium'
              }`}
            >
              <span>Sales Report</span>
              {salesFileInfo && (
                <span className="w-2 h-2 rounded-full bg-emerald-500" />
              )}
            </button>
          </nav>
        </div>

        {/* Right side: File Status and Admin Panel avatar matching Sleek Interface */}
        <div className="flex items-center gap-3">
          {/* Active loaded badge for Stock or Sales */}
          {(fileInfo || salesFileInfo) && (
            <div className="flex items-center gap-2">
              {currentView === 'report' && salesFileInfo ? (
                <div
                  id="sales-file-loaded-badge"
                  className="hidden md:flex items-center gap-2 px-2.5 py-1 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-medium"
                >
                  <Database className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="max-w-[130px] truncate">{salesFileInfo.fileName}</span>
                  <span className="text-emerald-700 font-mono">({salesFileInfo.totalRows} sales)</span>
                </div>
              ) : fileInfo ? (
                <div
                  id="file-loaded-badge"
                  className="hidden md:flex items-center gap-2 px-2.5 py-1 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-medium"
                >
                  <Database className="w-3.5 h-3.5 text-emerald-600" />
                  <span className="max-w-[130px] truncate">{fileInfo.fileName}</span>
                  <span className="text-emerald-700 font-mono">({fileInfo.totalRows} stock)</span>
                </div>
              ) : null}

              {onResetData && (
                <button
                  type="button"
                  id="header-clear-session-btn"
                  onClick={onResetData}
                  className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-semibold transition-all cursor-pointer"
                  title="Clear Session and reset all loaded data"
                >
                  <RotateCcw className="w-3 h-3 text-rose-600" />
                  <span className="hidden sm:inline">Clear Session</span>
                </button>
              )}
            </div>
          )}

          <div className="flex items-center gap-2.5">
            <div className="h-8 w-8 rounded-full bg-slate-200 border border-slate-300 flex items-center justify-center text-xs font-bold text-slate-600">
              AD
            </div>
            <span className="text-sm font-medium text-slate-600 hidden sm:inline">
              Admin Panel
            </span>
          </div>
        </div>
      </div>
    </header>
  );
}
