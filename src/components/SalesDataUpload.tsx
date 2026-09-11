import { useState, useRef, ChangeEvent, DragEvent } from 'react';
import {
  UploadCloud,
  CheckCircle2,
  AlertCircle,
  RotateCw,
  Trash2,
  ArrowRight,
  TrendingUp,
} from 'lucide-react';
import { UploadedFileInfo, MappedSalesItem } from '../types/inventory';
import { parseSalesExcelFile } from '../utils/salesParser';

interface SalesDataUploadProps {
  onDataLoaded: (items: MappedSalesItem[], fileInfo: UploadedFileInfo) => void;
  fileInfo: UploadedFileInfo | null;
  onReset: () => void;
}

export function SalesDataUpload({
  onDataLoaded,
  fileInfo,
  onReset,
}: SalesDataUploadProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = async (file: File) => {
    setErrorMessage(null);
    setIsProcessing(true);

    try {
      if (
        !file.name.endsWith('.xlsx') &&
        !file.name.endsWith('.xls') &&
        !file.name.endsWith('.csv')
      ) {
        throw new Error(
          'Unsupported file format. Please upload a Sales Excel (.xlsx, .xls) or .csv file exported from Wondersoft POS.'
        );
      }

      const { items, fileInfo: info } = await parseSalesExcelFile(file);

      if (items.length === 0) {
        throw new Error('No valid sales records found in the uploaded file.');
      }

      onDataLoaded(items, info);
    } catch (err: unknown) {
      console.error(err);
      setErrorMessage(
        err instanceof Error ? err.message : 'Failed to parse sales report file.'
      );
    } finally {
      setIsProcessing(false);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processFile(file);
    }
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  return (
    <div
      id="sales-data-upload-container"
      className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm space-y-4"
    >
      {/* Header section */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-4">
        <div>
          <div className="flex items-center gap-2">
            <TrendingUp className="w-5 h-5 text-indigo-600" />
            <h2 className="text-lg font-bold text-slate-800 tracking-tight">
              Upload Sales Report
            </h2>
            <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-100">
              Step 1
            </span>
          </div>
          <p className="text-xs text-slate-500 mt-0.5">
            Upload the sales report spreadsheet exported from Wondersoft POS software (.xlsx, .xls, .csv).
          </p>
        </div>
      </div>

      {/* When File is already parsed - Sleek Interface Emerald Banner */}
      {fileInfo ? (
        <div
          id="sales-file-uploaded-banner"
          className="rounded-xl border border-emerald-200 bg-emerald-50/70 p-4 sm:p-5 text-emerald-950 flex flex-col md:flex-row md:items-center justify-between gap-4 transition-all"
        >
          <div className="flex items-start sm:items-center gap-3.5">
            <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm shadow-emerald-200">
              <CheckCircle2 className="w-5 h-5" />
            </div>
            <div className="space-y-1 min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="font-bold text-sm sm:text-base text-slate-900 truncate max-w-[280px] sm:max-w-md">
                  {fileInfo.fileName}
                </span>
                <span className="inline-flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-full bg-emerald-200/70 text-emerald-900">
                  Sales Report Loaded
                </span>
              </div>
              <div className="flex items-center gap-2 sm:gap-4 text-xs text-emerald-900/80 flex-wrap">
                <span>
                  <strong>{fileInfo.totalRows.toLocaleString()}</strong> Sales Records
                </span>
                {fileInfo.totalSalesUnits !== undefined && (
                  <>
                    <span>•</span>
                    <span>
                      <strong>{fileInfo.totalSalesUnits.toLocaleString()}</strong> Total Sold Units (from &quot;{fileInfo.detectedQuantityColumn || 'Quantity'}&quot;)
                    </span>
                  </>
                )}
                <span>•</span>
                <span>
                  <strong>{fileInfo.uniqueToonLabels.toLocaleString()}</strong> Unique Toon Labels
                </span>
                {fileInfo.uniqueEANs > 0 && (
                  <>
                    <span>•</span>
                    <span>
                      <strong>{fileInfo.uniqueEANs.toLocaleString()}</strong> Barcodes / EANs
                    </span>
                  </>
                )}
                <span>•</span>
                <span className="text-slate-500 font-mono">
                  {fileInfo.uploadedAt}
                </span>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0 self-end md:self-center">
            <button
              type="button"
              id="change-sales-file-btn"
              onClick={() => fileInputRef.current?.click()}
              className="text-xs px-3.5 py-1.5 rounded-lg border border-emerald-300 hover:bg-emerald-100/60 font-medium text-emerald-900 transition-colors cursor-pointer"
            >
              Change File
            </button>
            <button
              type="button"
              id="reset-sales-file-btn"
              onClick={onReset}
              className="text-xs px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-white text-slate-600 hover:text-rose-600 transition-colors cursor-pointer flex items-center gap-1"
              title="Clear sales data"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Clear</span>
            </button>
          </div>
        </div>
      ) : (
        /* Drag & Drop Upload Zone */
        <div
          id="sales-dropzone-target"
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => !isProcessing && fileInputRef.current?.click()}
          className={`relative border-2 border-dashed rounded-xl p-8 sm:p-10 text-center transition-all cursor-pointer ${
            isDragging
              ? 'border-indigo-600 bg-indigo-50/40 scale-[0.99]'
              : 'border-slate-300 hover:border-indigo-500 bg-slate-50/50 hover:bg-slate-50'
          } ${isProcessing ? 'pointer-events-none opacity-60' : ''}`}
        >
          <input
            ref={fileInputRef}
            type="file"
            id="sales-file-input"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileChange}
            className="hidden"
          />

          <div className="flex flex-col items-center justify-center space-y-3">
            <div
              className={`w-14 h-14 rounded-2xl flex items-center justify-center transition-transform ${
                isDragging
                  ? 'bg-indigo-600 text-white scale-110 shadow-md shadow-indigo-100'
                  : 'bg-indigo-50 border border-indigo-100 text-indigo-600'
              }`}
            >
              {isProcessing ? (
                <RotateCw className="w-6 h-6 animate-spin text-indigo-600" />
              ) : (
                <UploadCloud className="w-7 h-7" />
              )}
            </div>

            <div className="space-y-1">
              <p className="text-sm font-semibold text-slate-800">
                {isProcessing
                  ? 'Processing and validating sales records...'
                  : 'Drop your Wondersoft Sales Report Excel file here'}
              </p>
              <p className="text-xs text-slate-500 max-w-sm mx-auto">
                Supports Excel (.xlsx, .xls) and CSV sales exports with Toon Label and Quantity columns.
              </p>
            </div>

            <button
              type="button"
              id="browse-sales-file-btn"
              disabled={isProcessing}
              className="mt-1 inline-flex items-center gap-1.5 px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white text-xs font-semibold shadow-sm shadow-indigo-100 transition-all cursor-pointer"
            >
              <span>Browse File</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      )}

      {/* Error alert */}
      {errorMessage && (
        <div
          id="sales-upload-error-alert"
          className="flex items-start gap-2.5 p-3.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-800 text-xs font-medium"
        >
          <AlertCircle className="w-4 h-4 shrink-0 text-rose-600 mt-0.5" />
          <div className="space-y-0.5">
            <span className="font-bold">Upload Failed</span>
            <p className="text-rose-700">{errorMessage}</p>
          </div>
        </div>
      )}
    </div>
  );
}
