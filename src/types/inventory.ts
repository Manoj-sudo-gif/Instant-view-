export type StoreName =
  | 'Salem'
  | 'Karur'
  | 'Mallur'
  | 'Namakkal'
  | 'Thiruvannamalai'
  | 'Kootapalli'
  | 'Kumbakonam'
  | 'GM Fashions Warehouse';

export const ALL_STORES: StoreName[] = [
  'Salem',
  'Karur',
  'Mallur',
  'Namakkal',
  'Thiruvannamalai',
  'Kootapalli',
  'Kumbakonam',
  'GM Fashions Warehouse',
];

// Specific order required for Toon Label Excel Output:
// [Color, Toon Label, Size, Kootapalli, Karur, Salem, Namakkal, Kumbakonam, Thiruvannamalai, Mallur, Size Total]
export const TOON_EXCEL_STORES: StoreName[] = [
  'Kootapalli',
  'Karur',
  'Salem',
  'Namakkal',
  'Kumbakonam',
  'Thiruvannamalai',
  'Mallur',
  'GM Fashions Warehouse',
];

export interface MappedInventoryItem {
  id: string;
  department: string;     // mapped from Wondersoft "Product Group"
  productName: string;    // mapped from Wondersoft "Department"
  brand: string;          // mapped from Wondersoft "Class"
  colour: string;         // mapped from Wondersoft "Colour" / "Color"
  size: string;           // mapped from Wondersoft "Size"
  fabric: string;         // mapped from Wondersoft "Garments" / "Garment"
  ean: string;            // mapped from Wondersoft "EAN Code"
  sellingPrice: number;   // mapped from Wondersoft "MRP"
  stockQuantity: number;  // store-wise quantity (mapped from Wondersoft "Stock")
  totalQuantity: number;  // total stock quantity across all stores
  store: StoreName;       // store location
  toonLabel: string;      // Toon Label (e.g., "6-G028-01")
  rawImagePath?: string;  // Internal only - MUST BE OMITTED in Rough View UI
  storeBreakdown?: Record<string, number>; // Store name (uppercase) -> quantity
}

export interface ToonMatrixRow {
  colour: string;
  toonLabel: string;
  size: string;
  kootapalli: number;
  karur: number;
  salem: number;
  namakkal: number;
  kumbakonam: number;
  thiruvannamalai: number;
  mallur: number;
  gmFashionsWarehouse: number;
  sizeTotal: number;
}

export type SearchType = 'EAN' | 'TOON';

export interface SearchState {
  type: SearchType;
  query: string;
  selectedStores: StoreName[];
  isAllStores: boolean;
  timestamp: number;
}

export interface UploadedFileInfo {
  fileName: string;
  fileSize: number;
  totalRows: number;
  uniqueEANs: number;
  uniqueToonLabels: number;
  uploadedAt: string;
  totalSalesUnits?: number;
  detectedQuantityColumn?: string;
}

export interface MappedSalesItem {
  id: string;
  toonLabel: string;
  colour: string;
  size: string;
  productName: string;
  brand: string;
  ean: string;
  store: StoreName;
  salesQty: number;
  mrp: number;
  billNo?: string;
  date?: string;
  storeBreakdown?: Record<string, number>;
}

export interface SalesMatrixRow {
  colour: string;
  toonLabel: string;
  size: string;
  kootapalli: number;
  karur: number;
  salem: number;
  namakkal: number;
  kumbakonam: number;
  thiruvannamalai: number;
  mallur: number;
  gmFashionsWarehouse: number;
  totalSales: number;
}
