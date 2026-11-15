import ExcelJS from 'exceljs';

export const EXCEL_COLORS = {
  crimsonDark: 'FFC00000',
  crimsonBg: 'FFFFF1F2',
  gold: 'FFFFC000',
  goldLight: 'FFFFF9C4',
  goldSoft: 'FFFFFDE7',
  yellowText: 'FFFDE047',
  navyDark: 'FF0F172A',
  slate800: 'FF1E293B',
  slate700: 'FF334155',
  slate100: 'FFF1F5F9',
  slate50: 'FFF8FAFC',
  slateMuted: 'FF94A3B8',
  amberHeader: 'FFB45309',
  amberSubBg: 'FFFEF3C7',
  amberText: 'FF78350F',
  indigoHeader: 'FF3730A3',
  indigoSubBg: 'FFE0E7FF',
  indigoText: 'FF1E1B4B',
  purpleHeader: 'FF581C87',
  purpleSubBg: 'FFF3E8FF',
  purpleText: 'FF4C1D95',
  emeraldHeader: 'FF065F46',
  emeraldSubBg: 'FFD1FAE5',
  emeraldSoft: 'FFECFDF5',
  emeraldText: 'FF047857',
  skyHeader: 'FF0369A1',
  skySubBg: 'FFE0F2FE',
  roseHeader: 'FFBE123C',
  roseText: 'FFBE123C',
  white: 'FFFFFFFF',
  black: 'FF000000',
};

export const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
};

export const DARK_THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF1E293B' } },
  left: { style: 'thin', color: { argb: 'FF1E293B' } },
  bottom: { style: 'thin', color: { argb: 'FF1E293B' } },
  right: { style: 'thin', color: { argb: 'FF1E293B' } },
};

export const MEDIUM_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'medium', color: { argb: 'FF0F172A' } },
  left: { style: 'medium', color: { argb: 'FF0F172A' } },
  bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
  right: { style: 'medium', color: { argb: 'FF0F172A' } },
};

export const NUM_FMT_INT = '#,##0';
export const NUM_FMT_IDR = '"Rp "#,##0';
export const NUM_FMT_PCT = '0.0%';

export function colToLetter(col: number): string {
  let temp = 0;
  let letter = '';
  let c = col;
  while (c > 0) {
    temp = (c - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    c = (c - temp - 1) / 26;
  }
  return letter;
}

export function solidFill(argb: string): ExcelJS.Fill {
  return {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb },
  };
}

/**
 * Adds the unified official 3-tier corporate banner at top of any worksheet.
 */
export function addUnifiedCorporateHeader(
  ws: ExcelJS.Worksheet,
  colSpan: number,
  moduleTitle: string,
  metaSubtitle: string,
  accentHex: string = EXCEL_COLORS.crimsonDark
): number {
  const safeSpan = Math.max(2, colSpan);

  // Row 1: Crimson Company Brand Banner
  const r1 = ws.addRow(['REPORT BAZAR MATAHARI — TANJUNG · PT MDS RETAILING TBK']);
  r1.height = 28;
  ws.mergeCells(r1.number, 1, r1.number, safeSpan);
  for (let c = 1; c <= safeSpan; c++) {
    const cell = r1.getCell(c);
    cell.fill = solidFill(accentHex);
    cell.font = {
      name: 'Calibri',
      size: 14,
      bold: true,
      color: { argb: EXCEL_COLORS.white },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = MEDIUM_BORDER;
  }

  // Row 2: Module Title Banner (Gold)
  const r2 = ws.addRow([moduleTitle]);
  r2.height = 22;
  ws.mergeCells(r2.number, 1, r2.number, safeSpan);
  for (let c = 1; c <= safeSpan; c++) {
    const cell = r2.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.font = {
      name: 'Calibri',
      size: 11.5,
      bold: true,
      color: { argb: EXCEL_COLORS.black },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = DARK_THIN_BORDER;
  }

  // Row 3: Periode & Metadata Bar (Navy)
  const stamp = new Date().toLocaleDateString('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });
  const r3 = ws.addRow([`${metaSubtitle}   |   Diekspor: ${stamp}`]);
  r3.height = 20;
  ws.mergeCells(r3.number, 1, r3.number, safeSpan);
  for (let c = 1; c <= safeSpan; c++) {
    const cell = r3.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.navyDark);
    cell.font = {
      name: 'Calibri',
      size: 9.5,
      bold: true,
      color: { argb: EXCEL_COLORS.yellowText },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = DARK_THIN_BORDER;
  }

  return r3.number;
}

export function addSectionTitleBar(
  ws: ExcelJS.Worksheet,
  title: string,
  colSpan: number,
  bgArgb: string = EXCEL_COLORS.slate800,
  fgArgb: string = EXCEL_COLORS.yellowText
): ExcelJS.Row {
  const row = ws.addRow([title]);
  row.height = 23;
  ws.mergeCells(row.number, 1, row.number, colSpan);
  for (let c = 1; c <= colSpan; c++) {
    const cell = row.getCell(c);
    cell.fill = solidFill(bgArgb);
    cell.font = { name: 'Calibri', size: 10.5, bold: true, color: { argb: fgArgb } };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = DARK_THIN_BORDER;
  }
  return row;
}

export interface KpiBannerItem {
  label: string;
  value: string | number;
  numFmt?: string;
  formula?: string;
  result?: number;
  bgArgb?: string;
  fgArgb?: string;
}

/**
 * Adds a clean 2-row KPI card strip across up to colSpan columns.
 */
export function addKpiStrip(
  ws: ExcelJS.Worksheet,
  items: KpiBannerItem[],
  colSpan: number
) {
  if (items.length === 0) return;
  const colsPerCard = Math.max(1, Math.floor(colSpan / items.length));

  const labelRow = ws.addRow(new Array(colSpan).fill(''));
  labelRow.height = 20;
  const valRow = ws.addRow(new Array(colSpan).fill(''));
  valRow.height = 24;

  items.forEach((item, idx) => {
    const startCol = idx * colsPerCard + 1;
    const endCol =
      idx === items.length - 1
        ? colSpan
        : Math.min(colSpan, startCol + colsPerCard - 1);
    if (startCol > colSpan) return;

    if (endCol > startCol) {
      ws.mergeCells(labelRow.number, startCol, labelRow.number, endCol);
      ws.mergeCells(valRow.number, startCol, valRow.number, endCol);
    }

    const lCell = labelRow.getCell(startCol);
    lCell.value = item.label;
    const vCell = valRow.getCell(startCol);
    if (item.formula) {
      vCell.value = {
        formula: item.formula,
        result: item.result ?? (typeof item.value === 'number' ? item.value : 0),
      };
    } else {
      vCell.value = item.value;
    }
    if (item.numFmt) {
      vCell.numFmt = item.numFmt;
    }

    const bg = item.bgArgb || EXCEL_COLORS.slate100;
    const fg = item.fgArgb || EXCEL_COLORS.navyDark;

    for (let c = startCol; c <= endCol; c++) {
      const lc = labelRow.getCell(c);
      lc.fill = solidFill(bg);
      lc.font = { name: 'Calibri', size: 8.5, bold: true, color: { argb: EXCEL_COLORS.slate700 } };
      lc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      lc.border = {
        top: { style: 'thin', color: { argb: 'FF94A3B8' } },
        left: { style: 'thin', color: { argb: 'FF94A3B8' } },
        right: { style: 'thin', color: { argb: 'FF94A3B8' } },
      };

      const vc = valRow.getCell(c);
      vc.fill = solidFill(bg);
      vc.font = { name: 'Calibri', size: 11, bold: true, color: { argb: fg } };
      vc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      vc.border = {
        bottom: { style: 'medium', color: { argb: 'FF475569' } },
        left: { style: 'thin', color: { argb: 'FF94A3B8' } },
        right: { style: 'thin', color: { argb: 'FF94A3B8' } },
      };
    }
  });
}

/**
 * Applies Freeze Panes compatible with both Microsoft Excel and WPS Spreadsheets.
 * Always freezes AFTER the bottom row of a merged table header (`ySplit`) and sets
 * explicit `topLeftCell` + `activeCell` so WPS/Excel scroll panes stay aligned.
 */
export function applySheetFreeze(
  ws: ExcelJS.Worksheet,
  xSplit: number,
  ySplit: number,
  options?: { showGridLines?: boolean; setPrintTitles?: boolean }
) {
  const safeX = Math.max(0, Math.floor(xSplit));
  const safeY = Math.max(0, Math.floor(ySplit));
  const showGridLines = options?.showGridLines !== false;

  if (safeX === 0 && safeY === 0) {
    ws.views = [{ state: 'normal', showGridLines }];
    return;
  }

  const topLeft = `${colToLetter(safeX + 1)}${safeY + 1}`;
  ws.views = [
    {
      state: 'frozen',
      xSplit: safeX,
      ySplit: safeY,
      topLeftCell: topLeft,
      activeCell: topLeft,
      showGridLines,
    },
  ];

  if (options?.setPrintTitles !== false && safeY > 0 && safeY <= 12 && ws.pageSetup) {
    ws.pageSetup.printTitlesRow = `1:${safeY}`;
  }
}

export function triggerWorkbookDownload(workbook: ExcelJS.Workbook, fileName: string) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  });
}

export function computeMultilineRowHeight(texts: string[], minHeight = 22, charsPerLine = 42, lineHeight = 14.5): number {
  let maxLines = 1;
  for (const t of texts) {
    if (!t) continue;
    const explicitLines = String(t).split('\n');
    let count = 0;
    for (const seg of explicitLines) {
      count += Math.max(1, Math.ceil(seg.length / charsPerLine));
    }
    if (count > maxLines) maxLines = count;
  }
  return Math.max(minHeight, maxLines * lineHeight + 6);
}
