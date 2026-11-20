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

export interface BorrowDetail {
  sourceStoreKey: MatrixStoreKey;
  sourceShortcut: string;
  sourceStoreName: string;
  qty: number;
}

export interface StoreCellShuffleInfo {
  originalQty: number;
  shuffledQty: number;
  isRedirected: boolean;
  sourceStoreKey?: MatrixStoreKey;
  sourceShortcut?: string;
  sourceStoreName?: string;
  donatedCount: number;
  borrowedDetails: BorrowDetail[];
  donatedToStores: { storeKey: MatrixStoreKey; shortcut: string; name: string; qty: number }[];
}

export interface ShuffledMatrixRow extends ToonMatrixRow {
  cellInfo: Record<MatrixStoreKey, StoreCellShuffleInfo>;
  totalRedirectedInRow: number;
  originalRow: ToonMatrixRow;
}

/**
 * Rebalance / Shuffle stock logic for a single matrix row (specific Color + Toon Label + Size).
 * 
 * Rules:
 * 1. Identify stores that have 0 stock.
 * 2. Process zero-stock stores in order of user's recipient preference:
 *    Koottappalli -> Salem -> Namakkal -> Karur -> Thiruvannamalai -> Kumbakonam -> Mallur.
 * 3. Donors are preferred first from stores with excess stock (current count >= 2), highest stock first.
 * 4. If no >= 2 donor exists, stores with 1 stock (GM Fashions Warehouse or lower-priority stores)
 *    can lend their remaining 1 unit down to 0 so high-priority 0-stock stores get filled.
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
    mallur: { originalQty: row.mallur, shuffledQty: row.mallur, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    thiruvannamalai: { originalQty: row.thiruvannamalai, shuffledQty: row.thiruvannamalai, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    kumbakonam: { originalQty: row.kumbakonam, shuffledQty: row.kumbakonam, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    karur: { originalQty: row.karur, shuffledQty: row.karur, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    salem: { originalQty: row.salem, shuffledQty: row.salem, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    namakkal: { originalQty: row.namakkal, shuffledQty: row.namakkal, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    kootapalli: { originalQty: row.kootapalli, shuffledQty: row.kootapalli, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    gmFashionsWarehouse: { originalQty: row.gmFashionsWarehouse, shuffledQty: row.gmFashionsWarehouse, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
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
      const recipientPrefIdx = RECIPIENT_PREFERENCE.indexOf(recipientKey);

      // Tier 1: Candidate donors with excess stock (>= 2)
      let candidateDonors = allPossibleDonors
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

      // Tier 2: Allow lending down to 0 when only 1 stock remains in GMW or a lower-priority shop
      if (candidateDonors.length === 0) {
        candidateDonors = allPossibleDonors
          .filter((k) => {
            if (k === recipientKey || currentStocks[k] < 1) return false;
            if (k === 'gmFashionsWarehouse') return true;
            return RECIPIENT_PREFERENCE.indexOf(k) > recipientPrefIdx;
          })
          .sort((a, b) => {
            if (a === 'gmFashionsWarehouse') return -1;
            if (b === 'gmFashionsWarehouse') return 1;
            return RECIPIENT_PREFERENCE.indexOf(b) - RECIPIENT_PREFERENCE.indexOf(a);
          });
      }

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
        cellInfo[recipientKey].borrowedDetails.push({
          sourceStoreKey: donorKey,
          sourceShortcut: STORE_SHORTCUTS[donorKey],
          sourceStoreName: STORE_FULL_NAMES[donorKey],
          qty: 1,
        });

        cellInfo[donorKey].shuffledQty = currentStocks[donorKey];
        cellInfo[donorKey].donatedCount += 1;
        cellInfo[donorKey].donatedToStores.push({
          storeKey: recipientKey,
          shortcut: STORE_SHORTCUTS[recipientKey],
          name: STORE_FULL_NAMES[recipientKey],
          qty: 1,
        });

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
    originalRow: row,
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

export interface ManualTransfer {
  id: string;
  fromStore: MatrixStoreKey;
  toStore: MatrixStoreKey;
  qty: number;
}

/**
 * Applies custom manual store-to-store transfers to a matrix row.
 */
export function applyManualTransfersToMatrixRow(
  row: ToonMatrixRow,
  transfers: ManualTransfer[]
): ShuffledMatrixRow {
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
    mallur: { originalQty: row.mallur, shuffledQty: row.mallur, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    thiruvannamalai: { originalQty: row.thiruvannamalai, shuffledQty: row.thiruvannamalai, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    kumbakonam: { originalQty: row.kumbakonam, shuffledQty: row.kumbakonam, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    karur: { originalQty: row.karur, shuffledQty: row.karur, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    salem: { originalQty: row.salem, shuffledQty: row.salem, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    namakkal: { originalQty: row.namakkal, shuffledQty: row.namakkal, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    kootapalli: { originalQty: row.kootapalli, shuffledQty: row.kootapalli, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
    gmFashionsWarehouse: { originalQty: row.gmFashionsWarehouse, shuffledQty: row.gmFashionsWarehouse, isRedirected: false, donatedCount: 0, borrowedDetails: [], donatedToStores: [] },
  };

  let totalRedirectedInRow = 0;

  for (const t of transfers) {
    if (!t.fromStore || !t.toStore || t.fromStore === t.toStore) continue;
    const qty = Math.max(1, Number(t.qty) || 1);

    currentStocks[t.fromStore] -= qty;
    currentStocks[t.toStore] += qty;

    cellInfo[t.toStore].shuffledQty = currentStocks[t.toStore];
    cellInfo[t.toStore].isRedirected = true;
    cellInfo[t.toStore].sourceStoreKey = t.fromStore;
    cellInfo[t.toStore].sourceShortcut = STORE_SHORTCUTS[t.fromStore];
    cellInfo[t.toStore].sourceStoreName = STORE_FULL_NAMES[t.fromStore];
    cellInfo[t.toStore].borrowedDetails.push({
      sourceStoreKey: t.fromStore,
      sourceShortcut: STORE_SHORTCUTS[t.fromStore],
      sourceStoreName: STORE_FULL_NAMES[t.fromStore],
      qty: qty,
    });

    cellInfo[t.fromStore].shuffledQty = currentStocks[t.fromStore];
    cellInfo[t.fromStore].donatedCount += qty;
    cellInfo[t.fromStore].donatedToStores.push({
      storeKey: t.toStore,
      shortcut: STORE_SHORTCUTS[t.toStore],
      name: STORE_FULL_NAMES[t.toStore],
      qty: qty,
    });

    totalRedirectedInRow += qty;
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
    originalRow: row,
  };
}

/**
 * Extracts all transfer movements from an automated ShuffledMatrixRow.
 */
export function extractTransfersFromShuffledRow(
  shuffledRow: ShuffledMatrixRow
): ManualTransfer[] {
  const transfers: ManualTransfer[] = [];
  const storeKeys: MatrixStoreKey[] = [
    'kootapalli',
    'karur',
    'salem',
    'namakkal',
    'kumbakonam',
    'thiruvannamalai',
    'mallur',
    'gmFashionsWarehouse',
  ];

  for (const recipientKey of storeKeys) {
    const info = shuffledRow.cellInfo[recipientKey];
    if (info && info.borrowedDetails) {
      for (const b of info.borrowedDetails) {
        transfers.push({
          id: `tr_${recipientKey}_${b.sourceStoreKey}_${Math.random().toString(36).substring(2, 8)}`,
          fromStore: b.sourceStoreKey,
          toStore: recipientKey,
          qty: b.qty,
        });
      }
    }
  }

  return transfers;
}

/**
 * Extracts automated transfers for the entire matrix as a map keyed by rowKey.
 */
export function extractAllAutoTransfers(
  shuffledRows: ShuffledMatrixRow[]
): Record<string, ManualTransfer[]> {
  const map: Record<string, ManualTransfer[]> = {};
  shuffledRows.forEach((row) => {
    const key = `${row.colour}__${row.toonLabel}__${row.size}`;
    const transfers = extractTransfersFromShuffledRow(row);
    if (transfers.length > 0) {
      map[key] = transfers;
    }
  });
  return map;
}

/**
 * Applies manual transfers across all rows of the matrix.
 */
export function applyManualMatrixData(
  rows: ToonMatrixRow[],
  manualTransfersMap: Record<string, ManualTransfer[]>
): {
  shuffledRows: ShuffledMatrixRow[];
  totalRedirectedUnits: number;
} {
  let totalRedirectedUnits = 0;
  const shuffledRows = rows.map((r) => {
    const key = `${r.colour}__${r.toonLabel}__${r.size}`;
    const rowTransfers = manualTransfersMap[key] || [];
    const res = applyManualTransfersToMatrixRow(r, rowTransfers);
    totalRedirectedUnits += res.totalRedirectedInRow;
    return res;
  });
  return { shuffledRows, totalRedirectedUnits };
}

