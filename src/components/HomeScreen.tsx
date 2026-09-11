import { ArrowRight, Search, BarChart3, Store } from 'lucide-react';
import { UploadedFileInfo } from '../types/inventory';

interface HomeScreenProps {
  onSelectInstantView: () => void;
  onSelectReportView: () => void;
  fileInfo: UploadedFileInfo | null;
  salesFileInfo?: UploadedFileInfo | null;
}

export function HomeScreen({
  onSelectInstantView,
  onSelectReportView,
  fileInfo,
  salesFileInfo,
}: HomeScreenProps) {
  return (
    <div className="max-w-3xl mx-auto py-10 sm:py-16 px-4 space-y-8">
      {/* Title */}
      <div className="text-center space-y-3 max-w-md mx-auto">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-indigo-50 border border-indigo-200 text-indigo-700 text-xs font-semibold">
          <Store className="w-3.5 h-3.5" />
          <span>Instant View Platform</span>
        </div>

        <h1 className="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">
          Inventory &amp; Reporting Portal
        </h1>
      </div>

      {/* Compact Square Cards: Stock Report & Sales Report */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-6 max-w-md mx-auto">
        {/* OPTION 1: STOCK REPORT */}
        <button
          type="button"
          id="option-stock-report-card"
          onClick={onSelectInstantView}
          className="group bg-white border-2 border-indigo-600 hover:border-indigo-700 hover:bg-indigo-50/20 rounded-xl p-6 shadow-sm hover:shadow transition-all duration-200 cursor-pointer flex flex-col items-center justify-center text-center gap-3 aspect-square sm:aspect-[4/3]"
        >
          <div className="w-12 h-12 rounded-xl bg-indigo-600 flex items-center justify-center text-white shadow-sm shadow-indigo-100 group-hover:scale-105 transition-transform">
            <Search className="w-6 h-6" />
          </div>
          <div className="space-y-1">
            <h2 className="text-base sm:text-lg font-bold text-slate-900 group-hover:text-indigo-700 transition-colors">
              Stock Report
            </h2>
            <span className="inline-flex items-center gap-1.5 text-xs text-indigo-600 font-semibold">
              <span>{fileInfo ? 'Continue Session' : 'Launch'}</span>
              <ArrowRight className="w-3.5 h-3.5 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>

        {/* OPTION 2: SALES REPORT */}
        <button
          type="button"
          id="option-sales-report-card"
          onClick={onSelectReportView}
          className={`group bg-white rounded-xl p-6 shadow-sm hover:shadow transition-all duration-200 cursor-pointer flex flex-col items-center justify-center text-center gap-3 aspect-square sm:aspect-[4/3] ${
            salesFileInfo
              ? 'border-2 border-emerald-600 hover:border-emerald-700 hover:bg-emerald-50/20'
              : 'border border-slate-200 hover:border-slate-300 hover:bg-slate-50'
          }`}
        >
          <div
            className={`w-12 h-12 rounded-xl flex items-center justify-center transition-transform group-hover:scale-105 ${
              salesFileInfo
                ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-100'
                : 'bg-slate-100 border border-slate-200 text-slate-600 group-hover:text-slate-900'
            }`}
          >
            <BarChart3 className="w-6 h-6" />
          </div>
          <div className="space-y-1">
            <h2 className="text-base sm:text-lg font-bold text-slate-800 group-hover:text-slate-900 transition-colors">
              Sales Report
            </h2>
            <span
              className={`inline-flex items-center gap-1.5 text-xs font-semibold ${
                salesFileInfo ? 'text-emerald-600' : 'text-slate-500'
              }`}
            >
              <span>{salesFileInfo ? 'Continue Session' : 'Launch'}</span>
              <ArrowRight className="w-3.5 h-3.5 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </button>
      </div>
    </div>
  );
}
