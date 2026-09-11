import XLSX from 'xlsx-js-style';

/**
 * Interface for grouped color blocks
 */
export interface MatrixColorGroup {
  colour: string;
  toonLabel: string;
  startIndex: number; // 0-based data row index
  endIndex: number;   // 0-based data row index (inclusive)
  count: number;
}

/**
 * Visual styling theme for an individual Color group
 */
export interface ColorGroupTheme {
  badgeBg: string;      // Color pill / header cell bg (e.g. 'DBEAFE')
  badgeText: string;    // Bold text color (e.g. '1E3A8A')
  rowBg: string;        // Soft tint for data cells in this group
  rowAltBg: string;     // Slightly alternate tint for even rows in group
  borderColor: string;  // Group divider border color
  totalColBg: string;   // Size total col bg
  totalColText: string; // Size total col text
}

/**
 * Curated list of vibrant yet accessible color group themes
 */
const ROTATING_PALETTES: ColorGroupTheme[] = [
  // 0: Sky Blue
  {
    badgeBg: 'DBEAFE',
    badgeText: '1E40AF',
    rowBg: 'FFFFFF',
    rowAltBg: 'EFF6FF',
    borderColor: '60A5FA',
    totalColBg: 'DBEAFE',
    totalColText: '1E3A8A',
  },
  // 1: Fresh Mint / Sage
  {
    badgeBg: 'DCFCE7',
    badgeText: '166534',
    rowBg: 'FFFFFF',
    rowAltBg: 'F0FDF4',
    borderColor: '4ADE80',
    totalColBg: 'D1FAE5',
    totalColText: '065F46',
  },
  // 2: Warm Amber / Gold
  {
    badgeBg: 'FEF08A',
    badgeText: '854D0E',
    rowBg: 'FFFFFF',
    rowAltBg: 'FEFCE8',
    borderColor: 'FACC15',
    totalColBg: 'FEF3C7',
    totalColText: '92400E',
  },
  // 3: Soft Lavender / Violet
  {
    badgeBg: 'F3E8FF',
    badgeText: '6B21A8',
    rowBg: 'FFFFFF',
    rowAltBg: 'FAF5FF',
    borderColor: 'C084FC',
    totalColBg: 'E9D5FF',
    totalColText: '581C87',
  },
  // 4: Soft Rose / Coral
  {
    badgeBg: 'FFE4E6',
    badgeText: '9F1239',
    rowBg: 'FFFFFF',
    rowAltBg: 'FFF1F2',
    borderColor: 'FB7185',
    totalColBg: 'FCE7F3',
    totalColText: '831843',
  },
  // 5: Cyan Teal
  {
    badgeBg: 'CFFAFE',
    badgeText: '155E75',
    rowBg: 'FFFFFF',
    rowAltBg: 'ECFEFF',
    borderColor: '22D3EE',
    totalColBg: 'CCFBF1',
    totalColText: '115E59',
  },
  // 6: Warm Peach / Orange
  {
    badgeBg: 'FFEDD5',
    badgeText: '9A3412',
    rowBg: 'FFFFFF',
    rowAltBg: 'FFF7ED',
    borderColor: 'FB923C',
    totalColBg: 'FED7AA',
    totalColText: '7C2D12',
  },
  // 7: Soft Indigo
  {
    badgeBg: 'E0E7FF',
    badgeText: '3730A3',
    rowBg: 'FFFFFF',
    rowAltBg: 'EEF2FF',
    borderColor: '818CF8',
    totalColBg: 'C7D2FE',
    totalColText: '312E81',
  },
];

/**
 * Detects color keyword and returns a corresponding vibrant pastel theme
 */
export function getColorGroupTheme(colourName: string, groupIndex: number): ColorGroupTheme {
  const c = (colourName || '').toUpperCase().trim();

  if (/BLUE|NAVY|ROYAL|SKY|DENIM|INDIGO|COBALT|OCEAN/.test(c)) {
    return {
      badgeBg: 'DBEAFE',
      badgeText: '1E40AF',
      rowBg: 'FFFFFF',
      rowAltBg: 'EFF6FF',
      borderColor: '60A5FA',
      totalColBg: 'DBEAFE',
      totalColText: '1E3A8A',
    };
  }

  if (/GREEN|MINT|OLIVE|PISTA|SAGE|BOTTLE|EMERALD|MEHENDI/.test(c)) {
    return {
      badgeBg: 'DCFCE7',
      badgeText: '166534',
      rowBg: 'FFFFFF',
      rowAltBg: 'F0FDF4',
      borderColor: '4ADE80',
      totalColBg: 'D1FAE5',
      totalColText: '065F46',
    };
  }

  if (/YELLOW|GOLD|MUSTARD|LEMON|OCHRE/.test(c)) {
    return {
      badgeBg: 'FEF08A',
      badgeText: '854D0E',
      rowBg: 'FFFFFF',
      rowAltBg: 'FEFCE8',
      borderColor: 'FACC15',
      totalColBg: 'FEF3C7',
      totalColText: '92400E',
    };
  }

  if (/RED|MAROON|WINE|BURGUNDY|CHERRY|CRIMSON/.test(c)) {
    return {
      badgeBg: 'FFE4E6',
      badgeText: '9F1239',
      rowBg: 'FFFFFF',
      rowAltBg: 'FFF1F2',
      borderColor: 'FB7185',
      totalColBg: 'FCE7F3',
      totalColText: '831843',
    };
  }

  if (/ORANGE|PEACH|CORAL|RUST|APRICOT/.test(c)) {
    return {
      badgeBg: 'FFEDD5',
      badgeText: '9A3412',
      rowBg: 'FFFFFF',
      rowAltBg: 'FFF7ED',
      borderColor: 'FB923C',
      totalColBg: 'FED7AA',
      totalColText: '7C2D12',
    };
  }

  if (/PURPLE|VIOLET|LAVENDER|LILAC|PLUM|MAUVE/.test(c)) {
    return {
      badgeBg: 'F3E8FF',
      badgeText: '6B21A8',
      rowBg: 'FFFFFF',
      rowAltBg: 'FAF5FF',
      borderColor: 'C084FC',
      totalColBg: 'E9D5FF',
      totalColText: '581C87',
    };
  }

  if (/PINK|ROSE|MAGENTA|FUCHSIA/.test(c)) {
    return {
      badgeBg: 'FCE7F3',
      badgeText: '9D174D',
      rowBg: 'FFFFFF',
      rowAltBg: 'FDF2F8',
      borderColor: 'F472B6',
      totalColBg: 'FBCFE8',
      totalColText: '831843',
    };
  }

  if (/BLACK|CHARCOAL|GREY|GRAY|MELANGE|CARBON/.test(c)) {
    return {
      badgeBg: 'E2E8F0',
      badgeText: '0F172A',
      rowBg: 'FFFFFF',
      rowAltBg: 'F1F5F9',
      borderColor: '94A3B8',
      totalColBg: 'CBD5E1',
      totalColText: '1E293B',
    };
  }

  if (/WHITE|CREAM|BEIGE|IVORY|KHAKI|TAN|BROWN|FAWN|CAMEL/.test(c)) {
    return {
      badgeBg: 'FEF3C7',
      badgeText: '78350F',
      rowBg: 'FFFFFF',
      rowAltBg: 'FFFBEB',
      borderColor: 'FCD34D',
      totalColBg: 'FDE68A',
      totalColText: '78350F',
    };
  }

  // Fallback rotating vibrant pastel themes
  return ROTATING_PALETTES[groupIndex % ROTATING_PALETTES.length];
}

/**
 * Common color palette for professional GM Fashions Excel exports
 */
export const EXCEL_COLORS = {
  // Stock Matrix Theme (Indigo / Deep Slate)
  stockHeaderBg: '1E293B', // Deep Slate
  stockHeaderText: 'FFFFFF',
  stockSubHeaderBg: '334155',
  stockTotalBg: 'FEF3C7', // Amber highlight
  stockTotalText: '92400E',
  stockRowEven: 'FFFFFF',
  stockRowOdd: 'F8FAFC',
  stockQtyPositiveBg: 'DCFCE7', // Soft emerald green highlight
  stockQtyPositiveText: '166534',
  stockTotalColBg: 'EEF2FF', // Soft Indigo for Size Total column
  stockTotalColText: '3730A3',

  // Sales Matrix Theme (Emerald / Forest Green)
  salesHeaderBg: '064E3B', // Deep Emerald Green
  salesHeaderText: 'FFFFFF',
  salesSubHeaderBg: '047857',
  salesTotalBg: 'FEF08A', // Rich Yellow/Gold
  salesTotalText: '854D0E',
  salesRowEven: 'FFFFFF',
  salesRowOdd: 'F0FDF4', // Very light mint
  salesQtyPositiveBg: 'D1FAE5', // Mint highlight for positive sales
  salesQtyPositiveText: '065F46',
  salesTotalColBg: 'ECFDF5', // Soft Mint for Total Sales column
  salesTotalColText: '047857',

  // Borders
  borderLight: 'E2E8F0',
  borderDark: '94A3B8',
  borderTotal: 'CBD5E1',

  // Text
  textMuted: '94A3B8',
  textRegular: '1E293B',
  textBold: '0F172A',
};

const BORDER_THIN = {
  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
  bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
  right: { style: 'thin', color: { rgb: 'CBD5E1' } },
};

const BORDER_TOTAL = {
  top: { style: 'medium', color: { rgb: '475569' } },
  bottom: { style: 'double', color: { rgb: '0F172A' } },
  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
  right: { style: 'thin', color: { rgb: 'CBD5E1' } },
};

/**
 * Auto-calculates optimal column widths based on headers and data length
 */
export function calculateColumnWidths(
  headers: string[],
  rows: Record<string, unknown>[]
): { wch: number }[] {
  return headers.map((header) => {
    let maxLen = header.length;
    for (let i = 0; i < Math.min(rows.length, 100); i++) {
      const val = rows[i]?.[header];
      if (val !== null && val !== undefined) {
        const str = String(val);
        if (str.length > maxLen) {
          maxLen = Math.min(str.length, 45); // cap at 45
        }
      }
    }

    // Give spacious comfortable padding for key columns
    if (header === 'Color') return { wch: Math.max(maxLen + 4, 16) };
    if (header === 'Toon Label') return { wch: Math.max(maxLen + 4, 18) };
    if (header === 'Size') return { wch: Math.max(maxLen + 4, 10) };
    if (header.includes('Total')) return { wch: Math.max(maxLen + 4, 14) };

    return { wch: Math.max(maxLen + 4, 14) };
  });
}

/**
 * Styles a Matrix Worksheet with grouped, merged Color blocks and vibrant themes
 * - Vertically merges Color cell across all sizes for that color group (e.g. Blue appears once spanning 10 rows)
 * - Vertically merges Toon Label cell across that group
 * - Applies a distinct, vibrant pastel background theme for each color group
 * - Adds clear solid separator borders between groups
 * - Highlights positive store quantities with vivid emerald/mint badges
 * - Highlights the Size Total column and bottom TOTAL summary row
 */
export function styleMatrixWorksheet(
  worksheet: XLSX.WorkSheet,
  headers: string[],
  rowCount: number,
  theme: 'stock' | 'sales',
  groups: MatrixColorGroup[] = []
) {
  const isStock = theme === 'stock';
  const headerBg = isStock ? EXCEL_COLORS.stockHeaderBg : EXCEL_COLORS.salesHeaderBg;
  const headerText = isStock ? EXCEL_COLORS.stockHeaderText : EXCEL_COLORS.salesHeaderText;
  const posQtyBg = isStock ? EXCEL_COLORS.stockQtyPositiveBg : EXCEL_COLORS.salesQtyPositiveBg;
  const posQtyText = isStock ? EXCEL_COLORS.stockQtyPositiveText : EXCEL_COLORS.salesQtyPositiveText;
  const summaryRowBg = isStock ? EXCEL_COLORS.stockTotalBg : EXCEL_COLORS.salesTotalBg;
  const summaryRowText = isStock ? EXCEL_COLORS.stockTotalText : EXCEL_COLORS.salesTotalText;

  const totalCols = headers.length;
  const lastColIdx = totalCols - 1;

  // Set row heights
  worksheet['!rows'] = [
    { hpt: 28 }, // Header row height
  ];

  // Configure SheetJS cell merges for Color (col 0) and Toon Label (col 1)
  const merges: XLSX.Range[] = [];

  // Map each data row (0-based) to its group and group index
  const rowGroupMap = new Map<
    number,
    { group: MatrixColorGroup; groupIndex: number; theme: ColorGroupTheme }
  >();

  groups.forEach((g, gIdx) => {
    const groupTheme = getColorGroupTheme(g.colour, gIdx);
    const startRow = g.startIndex + 1; // 1-based row index in sheet (row 0 is header)
    const endRow = g.endIndex + 1;

    // Record merge ranges if group has multiple rows
    if (endRow > startRow) {
      // Column 0: Color merge
      merges.push({
        s: { r: startRow, c: 0 },
        e: { r: endRow, c: 0 },
      });
      // Column 1: Toon Label merge
      merges.push({
        s: { r: startRow, c: 1 },
        e: { r: endRow, c: 1 },
      });
    }

    for (let rIdx = g.startIndex; rIdx <= g.endIndex; rIdx++) {
      rowGroupMap.set(rIdx, { group: g, groupIndex: gIdx, theme: groupTheme });
    }
  });

  if (merges.length > 0) {
    worksheet['!merges'] = merges;
  }

  // Iterate all cells in the sheet to apply exquisite styles
  for (let r = 0; r <= rowCount; r++) {
    const isHeader = r === 0;
    const isSummaryRow = r === rowCount && rowCount > 1;
    const dataRowIdx = r - 1;
    const groupInfo = rowGroupMap.get(dataRowIdx);
    const isFirstInGroup = groupInfo ? dataRowIdx === groupInfo.group.startIndex : false;
    const isLastInGroup = groupInfo ? dataRowIdx === groupInfo.group.endIndex : false;
    const currentTheme = groupInfo?.theme || ROTATING_PALETTES[0];

    for (let c = 0; c < totalCols; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      let cell = worksheet[cellRef];

      // Ensure empty cells inside merged areas exist so borders & backgrounds don't glitch
      if (!cell) {
        cell = { t: 's', v: '' };
        worksheet[cellRef] = cell;
      }

      const isColorCol = c === 0;
      const isToonCol = c === 1;
      const isSizeCol = c === 2;
      const isStoreCol = c >= 3 && c < lastColIdx;
      const isTotalCol = c === lastColIdx;
      const numVal = typeof cell.v === 'number' ? cell.v : parseFloat(String(cell.v));
      const hasValue = !isNaN(numVal) && numVal > 0;

      // Group divider borders
      const borderTop = isFirstInGroup
        ? { style: 'medium', color: { rgb: currentTheme.borderColor || '475569' } }
        : BORDER_THIN.top;
      const borderBottom = isLastInGroup
        ? { style: 'medium', color: { rgb: currentTheme.borderColor || '475569' } }
        : BORDER_THIN.bottom;

      if (isHeader) {
        // TOP HEADER ROW
        cell.s = {
          fill: { fgColor: { rgb: headerBg } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: headerText },
          },
          alignment: {
            horizontal: isColorCol || isToonCol ? 'left' : isStoreCol || isTotalCol || isSizeCol ? 'center' : 'left',
            vertical: 'center',
            wrapText: true,
          },
          border: BORDER_THIN,
        };
      } else if (isSummaryRow) {
        // BOTTOM TOTAL SUMMARY ROW
        cell.s = {
          fill: { fgColor: { rgb: summaryRowBg } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: summaryRowText },
          },
          alignment: {
            horizontal: c === 0 ? 'left' : isStoreCol || isTotalCol ? 'center' : 'center',
            vertical: 'center',
          },
          border: BORDER_TOTAL,
        };
      } else if (isColorCol) {
        // COLOR CELL (Merged block for color group)
        cell.s = {
          fill: { fgColor: { rgb: currentTheme.badgeBg } },
          font: {
            name: 'Arial',
            sz: 12,
            bold: true,
            color: { rgb: currentTheme.badgeText },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true,
          },
          border: {
            top: borderTop,
            bottom: borderBottom,
            left: { style: 'medium', color: { rgb: currentTheme.borderColor || '475569' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } },
          },
        };
      } else if (isToonCol) {
        // TOON LABEL CELL (Merged block for Toon Label)
        cell.s = {
          fill: { fgColor: { rgb: isStock ? 'EEF2FF' : 'ECFDF5' } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: isStock ? '3730A3' : '047857' },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true,
          },
          border: {
            top: borderTop,
            bottom: borderBottom,
            left: BORDER_THIN.left,
            right: { style: 'thin', color: { rgb: 'CBD5E1' } },
          },
        };
      } else if (isSizeCol) {
        // SIZE CELL
        const isEven = dataRowIdx % 2 === 0;
        cell.s = {
          fill: { fgColor: { rgb: isEven ? currentTheme.rowBg : currentTheme.rowAltBg } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: '0F172A' },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
          border: {
            top: borderTop,
            bottom: borderBottom,
            left: BORDER_THIN.left,
            right: BORDER_THIN.right,
          },
        };
      } else if (isTotalCol) {
        // SIZE TOTAL / TOTAL SALES COLUMN
        cell.s = {
          fill: { fgColor: { rgb: currentTheme.totalColBg } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: currentTheme.totalColText },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
          border: {
            top: borderTop,
            bottom: borderBottom,
            left: BORDER_THIN.left,
            right: { style: 'medium', color: { rgb: currentTheme.borderColor || '475569' } },
          },
        };
      } else if (isStoreCol) {
        // STORE DATA CELLS
        const isEven = dataRowIdx % 2 === 0;
        let cellBg = isEven ? currentTheme.rowBg : currentTheme.rowAltBg;
        let fontColor = EXCEL_COLORS.textRegular;
        let isBold = false;

        if (hasValue) {
          // Store cell with positive stock/sales -> Highlighted badge!
          cellBg = posQtyBg;
          fontColor = posQtyText;
          isBold = true;
        } else {
          // Zero value store cell -> subtle muted gray
          fontColor = EXCEL_COLORS.textMuted;
        }

        cell.s = {
          fill: { fgColor: { rgb: cellBg } },
          font: {
            name: 'Arial',
            sz: 10,
            bold: isBold,
            color: { rgb: fontColor },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
          border: {
            top: borderTop,
            bottom: borderBottom,
            left: BORDER_THIN.left,
            right: BORDER_THIN.right,
          },
        };
      }
    }
  }
}

/**
 * Styles a Detailed List Worksheet (EAN Stock List or Sales Table List)
 */
export function styleDetailedWorksheet(
  worksheet: XLSX.WorkSheet,
  headers: string[],
  rowCount: number,
  theme: 'stock' | 'sales',
  highlightColIndex?: number
) {
  const isStock = theme === 'stock';
  const headerBg = isStock ? EXCEL_COLORS.stockHeaderBg : EXCEL_COLORS.salesHeaderBg;
  const headerText = isStock ? EXCEL_COLORS.stockHeaderText : EXCEL_COLORS.salesHeaderText;
  const highlightBg = isStock ? EXCEL_COLORS.stockQtyPositiveBg : EXCEL_COLORS.salesQtyPositiveBg;
  const highlightText = isStock ? EXCEL_COLORS.stockQtyPositiveText : EXCEL_COLORS.salesQtyPositiveText;

  const totalCols = headers.length;

  worksheet['!rows'] = [
    { hpt: 26 }, // Header row height
  ];

  for (let r = 0; r <= rowCount; r++) {
    const isHeader = r === 0;

    for (let c = 0; c < totalCols; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = worksheet[cellRef];
      if (!cell) continue;

      const isHighlightCol = c === highlightColIndex;

      if (isHeader) {
        cell.s = {
          fill: { fgColor: { rgb: headerBg } },
          font: {
            name: 'Arial',
            sz: 11,
            bold: true,
            color: { rgb: headerText },
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true,
          },
          border: BORDER_THIN,
        };
      } else {
        const isEven = r % 2 === 0;
        let cellBg = isEven ? 'FFFFFF' : isStock ? 'F8FAFC' : 'F0FDF4';
        let fontColor = EXCEL_COLORS.textRegular;
        let isBold = false;

        if (isHighlightCol) {
          cellBg = highlightBg;
          fontColor = highlightText;
          isBold = true;
        }

        cell.s = {
          fill: { fgColor: { rgb: cellBg } },
          font: {
            name: 'Arial',
            sz: 10,
            bold: isBold,
            color: { rgb: fontColor },
          },
          alignment: {
            horizontal: isHighlightCol ? 'center' : 'left',
            vertical: 'center',
          },
          border: BORDER_THIN,
        };
      }
    }
  }
}

export { XLSX };
