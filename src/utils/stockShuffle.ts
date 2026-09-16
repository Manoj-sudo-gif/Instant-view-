import { ToonMatrixRow } from '../types/inventory';

export type MatrixStoreKey =
  | 'kootapalli'
  | 'karur'
  | 'salem'
  | 'namakkal'
  | 'kumbakonam'
  | 'thiruvannamalai'
  | 'mallur'
  | 'gmFashionsWarehouse';

// Shortcuts exactly as specified:
// Koottapalli na - KP,
// Karur - KR,
// Salem - SLM,
// Namakkal - NML,
// Kumbakonam - KUM,
// Thiruvannamalai - Tir,
// Mallur - MLR,
// GM warhouse - GMW
export const STORE_SHORTCUTS: Record<MatrixStoreKey, string> = {
  kootapalli: 'KP',
  karur: 'KR',
  salem: 'SLM',
  namakkal: 'NML',
  kumbakonam: 'KUM',
  thiruvannamalai: 'Tir',
  mallur: 'MLR',
  gmFashionsWarehouse: 'GMW',
};

export const STORE_FULL_NAMES: Record<MatrixStoreKey, string> = {
  kootapalli: 'Koottappalli',
  karur: 'Karur',
  salem: 'Salem',
  namakkal: 'Namakkal',
  kumbakonam: 'Kumbakonam',
  thiruvannamalai: 'Thiruvannamalai',
  mallur: 'Mallur',
  gmFashionsWarehouse: 'GM Fashions Warehouse',
};

// User's exact preference order for receiving stock when at 0:
// 1st preference Koottappalli (kootapalli fill agurathula first preference ah eduthuko)
// 2nd Salem (first salem fill aganum)
// 3rd Namakkal (apro namakkal fill aganum)
// 4th Karur (apro karur fill aganum)
// 5th Thiruvannamalai (apro thiruvannamalai fill aganum)
// 6th Kumbakonam (apro kumbakonam fill aganum)
// 7th Mallur (apr final ah mallur fill aganum)
export const RECIPIENT_PREFERENCE: MatrixStoreKey[] = [
  'kootapalli',
  'salem',
  'namakkal',
  'karur',
  'thiruvannamalai',
  'kumbakonam',
  'mallur',
];

export interface StoreCellShuffleInfo {
  originalQty: number;
  shuffledQty: number;
  isRedirected: boolean;
  sourceStoreKey?: MatrixStoreKey;
  sourceShortcut?: string;
  sourceStoreName?: string;
  donatedCount: number;
}

export interface ShuffledMatrixRow extends ToonMatrixRow {
  cellInfo: Record<MatrixStoreKey, StoreCellShuffleInfo>;
  totalRedirectedInRow: number;
}

/**
 * Rebalance / Shuffle stock logic for a single matrix row (specific Color + Toon Label + Size).
 * 
 * Rules:
 * 1. Identify stores that have 0 stock.
 * 2. Process zero-stock stores in order of user's recipient preference:
 *    Mallur -> Thiruvannamalai -> Kumbakonam -> Karur -> Salem -> Namakkal -> Koottappalli.
 * 3. Donors are stores with excess stock (current count >= 2).
 *    To keep stocks balanced across all stores, the donor with the HIGHEST current stock donates first.
 *    If tied, GM Fashions Warehouse (GMW) or stores with lower recipient preference donate first.
 * 4. 1 unit is transferred from the donor to the recipient.
 * 5. Size Total remains strictly conserved.
 */
export function shuffleMatrixRow(row: ToonMatrixRow): ShuffledMatrixRow {
  const currentStocks: Record<MatrixStoreKey, number> = {
    mallur: row.mallur,
    thiruvannamalai: row.thiruvannamalai,
    kumbakonam: row.kumbakonam,
    karur: row.karur,
    salem: row.salem,
    namakkal: row.namakkal,
    kootapalli: row.kootapalli,
    gmFashionsWarehouse: row.gmFashionsWarehouse,
  };

  const cellInfo: Record<MatrixStoreKey, StoreCellShuffleInfo> = {
    mallur: { originalQty: row.mallur, shuffledQty: row.mallur, isRedirected: false, donatedCount: 0 },
    thiruvannamalai: { originalQty: row.thiruvannamalai, shuffledQty: row.thiruvannamalai, isRedirected: false, donatedCount: 0 },
    kumbakonam: { originalQty: row.kumbakonam, shuffledQty: row.kumbakonam, isRedirected: false, donatedCount: 0 },
    karur: { originalQty: row.karur, shuffledQty: row.karur, isRedirected: false, donatedCount: 0 },
    salem: { originalQty: row.salem, shuffledQty: row.salem, isRedirected: false, donatedCount: 0 },
    namakkal: { originalQty: row.namakkal, shuffledQty: row.namakkal, isRedirected: false, donatedCount: 0 },
    kootapalli: { originalQty: row.kootapalli, shuffledQty: row.kootapalli, isRedirected: false, donatedCount: 0 },
    gmFashionsWarehouse: { originalQty: row.gmFashionsWarehouse, shuffledQty: row.gmFashionsWarehouse, isRedirected: false, donatedCount: 0 },
  };

  let totalRedirectedInRow = 0;

  const allPossibleDonors: MatrixStoreKey[] = [
    'gmFashionsWarehouse',
    'kootapalli',
    'namakkal',
    'salem',
    'karur',
    'kumbakonam',
    'thiruvannamalai',
    'mallur',
  ];

  // Process 0-stock stores according to recipient preference
  for (const recipientKey of RECIPIENT_PREFERENCE) {
    if (currentStocks[recipientKey] === 0) {
      // Find candidate donors with excess stock (>= 2)
      const candidateDonors = allPossibleDonors
        .filter((k) => k !== recipientKey && currentStocks[k] >= 2)
        .sort((a, b) => {
          // 1. Highest stock first (to keep stocks balanced across all stores)
          if (currentStocks[b] !== currentStocks[a]) {
            return currentStocks[b] - currentStocks[a];
          }
          // 2. Central warehouse (GMW) preference if tied
          if (a === 'gmFashionsWarehouse') return -1;
          if (b === 'gmFashionsWarehouse') return 1;
          // 3. Lower recipient preference stores donate first if tied
          const idxA = RECIPIENT_PREFERENCE.indexOf(a);
          const idxB = RECIPIENT_PREFERENCE.indexOf(b);
          return idxB - idxA;
        });

      if (candidateDonors.length > 0) {
        const donorKey = candidateDonors[0];

        // Transfer 1 stock unit
        currentStocks[donorKey] -= 1;
        currentStocks[recipientKey] += 1;

        cellInfo[recipientKey].shuffledQty = currentStocks[recipientKey];
        cellInfo[recipientKey].isRedirected = true;
        cellInfo[recipientKey].sourceStoreKey = donorKey;
        cellInfo[recipientKey].sourceShortcut = STORE_SHORTCUTS[donorKey];
        cellInfo[recipientKey].sourceStoreName = STORE_FULL_NAMES[donorKey];

        cellInfo[donorKey].shuffledQty = currentStocks[donorKey];
        cellInfo[donorKey].donatedCount += 1;

        totalRedirectedInRow += 1;
      }
    }
  }

  return {
    ...row,
    mallur: currentStocks.mallur,
    thiruvannamalai: currentStocks.thiruvannamalai,
    kumbakonam: currentStocks.kumbakonam,
    karur: currentStocks.karur,
    salem: currentStocks.salem,
    namakkal: currentStocks.namakkal,
    kootapalli: currentStocks.kootapalli,
    gmFashionsWarehouse: currentStocks.gmFashionsWarehouse,
    cellInfo,
    totalRedirectedInRow,
  };
}

/**
 * Shuffles all rows of the Toon Label matrix.
 */
export function shuffleMatrixData(rows: ToonMatrixRow[]): {
  shuffledRows: ShuffledMatrixRow[];
  totalRedirectedUnits: number;
} {
  let totalRedirectedUnits = 0;
  const shuffledRows = rows.map((r) => {
    const res = shuffleMatrixRow(r);
    totalRedirectedUnits += res.totalRedirectedInRow;
    return res;
  });
  return { shuffledRows, totalRedirectedUnits };
}
