import { useState, useRef, ChangeEvent, DragEvent } from 'react';
import {
  UploadCloud,
  CheckCircle2,
  AlertCircle,
  RotateCw,
  Trash2,
  ArrowRight,
} from 'lucide-react';
import { UploadedFileInfo, MappedInventoryItem } from '../types/inventory';
import {
  parseExcelFile,
} from '../utils/excelParser';

interface DataUploadProps {
  onDataLoaded: (items: MappedInventoryItem[], fileInfo: UploadedFileInfo) => void;
  fileInfo: UploadedFileInfo | null;
  onReset: () => void;
}

export function DataUpload({
  onDataLoaded,
  fileInfo,
  onReset,
}: DataUploadProps) {
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
          'Unsupported file format. Please upload an Excel (.xlsx, .xls) or .csv file generated from Wondersoft.'
        );
      }

      const buffer = await file.arrayBuffer();
      const mappedItems = await parseExcelFile(buffer);

      if (mappedItems.length === 0) {
        throw new Error('No valid inventory records found in the uploaded file.');
      }

      const uniqueEans = new Set(mappedItems.map((i) => i.ean).filter(Boolean)).size;
      const uniqueToons = new Set(mappedItems.map((i) => i.toonLabel).filter(Boolean)).size;

      const info: UploadedFileInfo = {
        fileName: file.name,
        fileSize: file.size,
        totalRows: mappedItems.length,
        uniqueEANs: uniqueEans,
        uniqueToonLabels: uniqueToons,
        uploadedAt: new Date().toLocaleTimeString(),
      };

      onDataLoaded(mappedItems, info);
    } catch (err: unknown) {
      console.error(err);
      setErrorMessage(
        err instanceof Error ? err.message : 'Failed to parse Excel file.'
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
    // reset input value so re-selecting same file triggers change
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
      id="data-upload-container"
      className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-sm"
    >
      {/* Section Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 pb-4 mb-5">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <span>Data Upload (Wondersoft Source)</span>
            <span
              className={`text-xs font-semibold px-2.5 py-0.5 rounded-full ${
                fileInfo
                  ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                  : 'bg-indigo-50 text-indigo-700 border border-indigo-100'
              }`}
            >
              {fileInfo ? '✓ Parsed Successfully' : 'Step 1: Upload Excel'}
            </span>
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Upload the inventory spreadsheet exported from legacy Wondersoft POS software.
          </p>
        </div>
      </div>

      {/* When File is already parsed - Sleek Interface Emerald Banner */}
      {fileInfo ? (
        <div
          id="parsed-file-summary-card"
          className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 bg-emerald-50 border border-emerald-200 px-4 py-3 rounded-lg"
        >
          <div className="flex items-center gap-3">
            <div className="w-7 h-7 rounded-full bg-emerald-500 flex items-center justify-center shrink-0">
              <CheckCircle2 className="w-4 h-4 text-white" />
            </div>
            <div>
              <span className="text-sm font-semibold text-emerald-900">
                Success: {fileInfo.fileName} parsed ({fileInfo.totalRows.toLocaleString()} rows)
              </span>
              <div className="flex items-center gap-3 text-xs text-emerald-800 mt-0.5 font-mono flex-wrap">
                <span className={fileInfo.uniqueEANs > 0 ? 'text-emerald-900 font-bold' : 'text-amber-900 font-bold bg-amber-100 px-1.5 py-0.5 rounded'}>
                  Unique EANs: {fileInfo.uniqueEANs} {fileInfo.uniqueEANs === 0 ? '(⚠️ check column header)' : '✓'}
                </span>
                <span>•</span>
                <span>Toon Styles: <strong>{fileInfo.uniqueToonLabels}</strong></span>
                <span>•</span>
                <span>Uploaded at: {fileInfo.uploadedAt}</span>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 self-end sm:self-auto">
            <button
              type="button"
              id="clear-session-btn"
              onClick={onReset}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-white hover:bg-rose-50 text-rose-700 border border-rose-200 text-xs font-semibold shadow-sm transition-colors cursor-pointer"
              title="Clear all parsed inventory and reset session"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-600" />
              <span>Clear Session</span>
            </button>
            <button
              type="button"
              id="change-reupload-file-btn"
              onClick={onReset}
              className="text-xs font-medium text-slate-700 hover:text-slate-900 px-3 py-1.5 rounded-md hover:bg-slate-100 border border-slate-300 transition-colors cursor-pointer"
            >
              Re-upload File
            </button>
          </div>
        </div>
      ) : (
        /* Upload Drag-and-Drop Area */
        <div className="space-y-4">
          <div
            id="dropzone-area"
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-8 sm:p-10 text-center cursor-pointer transition-all duration-200 ${
              isDragging
                ? 'border-indigo-500 bg-indigo-50/60 scale-[1.005]'
                : 'border-slate-300 hover:border-indigo-400 bg-slate-50/50 hover:bg-slate-50'
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              id="wondersoft-file-input"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
              className="hidden"
            />

            <div className="max-w-md mx-auto space-y-3">
              <div className="w-14 h-14 mx-auto rounded-2xl bg-indigo-50 border border-indigo-100 flex items-center justify-center text-indigo-600">
                {isProcessing ? (
                  <RotateCw className="w-7 h-7 animate-spin text-indigo-600" />
                ) : (
                  <UploadCloud className="w-7 h-7 text-indigo-600" />
                )}
              </div>

              <div>
                <p className="text-sm sm:text-base font-semibold text-slate-800">
                  {isProcessing
                    ? 'Parsing Wondersoft inventory data...'
                    : 'Click to select or drag & drop Wondersoft Excel file'}
                </p>
                <p className="text-xs text-slate-500 mt-1">
                  Supports .xlsx, .xls, and .csv files exported from Wondersoft POS
                </p>
              </div>

              <div className="pt-2">
                <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-indigo-50 text-indigo-700 border border-indigo-200 text-xs font-semibold">
                  <span>Select File from Computer</span>
                  <ArrowRight className="w-3.5 h-3.5" />
                </span>
              </div>
            </div>
          </div>

          {/* Mapping Specifications preview card */}
          <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-xs text-slate-600">
            <span className="font-bold text-slate-700">
              Automatic Backend Wondersoft Mapping:
            </span>{' '}
            Product Group → <span className="text-indigo-700 font-medium">Department</span>,
            Department → <span className="text-indigo-700 font-medium">Product Name</span>,
            Class → <span className="text-indigo-700 font-medium">Brand</span>, Colour →{' '}
            <span className="text-indigo-700 font-medium">Colour</span>, Size →{' '}
            <span className="text-indigo-700 font-medium">Size</span>, Garment →{' '}
            <span className="text-indigo-700 font-medium">Fabric</span>, EAN Code →{' '}
            <span className="text-indigo-700 font-medium">EAN</span>, MRP →{' '}
            <span className="text-indigo-700 font-medium">Selling Price</span>, Stock →{' '}
            <span className="text-indigo-700 font-medium">Stock Quantity</span>.
          </div>
        </div>
      )}

      {/* Error notification banner */}
      {errorMessage && (
        <div
          id="upload-error-banner"
          className="mt-4 flex items-center gap-2.5 p-3 rounded-lg bg-rose-50 border border-rose-200 text-rose-800 text-xs"
        >
          <AlertCircle className="w-4 h-4 shrink-0 text-rose-600" />
          <span>{errorMessage}</span>
        </div>
      )}
    </div>
  );
}
