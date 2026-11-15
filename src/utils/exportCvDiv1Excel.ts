/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * Streamlined Excel (.xlsx) export utility for:
 * 1) Sheet "Report CV": Harian Bulan Aktif + Per Week Bulan Aktif + Rekap 12 Bulan (2026)
 * 2) Sheet "Data CV": Data CV Per Monthly & Per Week (Inventory Stock Awal/Akhir, Promo, Reason & Action Plan)
 */

import ExcelJS from 'exceljs';
import {
  CvDailyRecord,
  CvInventoryBrandKey,
  CvInventoryOpsEntry,
  CV_BRAND_CATALOG,
  CV_DIVISI_DEFAULT,
  calcCvDaySummary,
  calcCvWeeklySummary,
  calcCvMonthlySummary,
  calcCvAnnualSummary,
  calcCvLiveSalesQtyMap,
  createDefaultCvOpsEntry,
  getCvOpsEntryId,
  loadCvInventoryOps,
  MONTH_NAMES_ID,
} from '../data/cvDiv1Data';
import { ANNUAL_2026_WEEKS, getIndoDayName, isWeekendDate } from '../data/initialData';
import {
  EXCEL_COLORS,
  THIN_BORDER,
  DARK_THIN_BORDER,
  MEDIUM_BORDER,
  NUM_FMT_INT,
  NUM_FMT_IDR,
  NUM_FMT_PCT,
  colToLetter,
  solidFill,
  addUnifiedCorporateHeader,
  addSectionTitleBar,
  addKpiStrip,
  applySheetFreeze,
  triggerWorkbookDownload,
  computeMultilineRowHeight,
} from './excelStyles';
import {
  ExcelChartBarItem,
  ExcelChartSeriesConfig,
  embedExcelSalesDiagramPair,
} from './excelCharts';

const CV_CHART_CONFIG: ExcelChartSeriesConfig = {
  seriesALabel: '6201 Graphis',
  seriesAColor: '#4F46E5',
  seriesALegendColor: '#C7D2FE',
  seriesBLabel: '6244 Simplicity',
  seriesBColor: '#F59E0B',
  seriesBLegendColor: '#FDE68A',
  badgeText: 'DIAGRAM SALES CV DIVISI 1 (OMZET RUPIAH)',
  badgeBg: '#7C3AED',
  cardSubtitle: 'Grafik Batang Dept 6201 Graphis vs Dept 6244 Simplicity & Garis Tren',
};

function getSafeSheetName(workbook: ExcelJS.Workbook, preferred: string): string {
  const cleaned = preferred.replace(/[\\/?*\[\]:]/g, '').slice(0, 31);
  if (!workbook.getWorksheet(cleaned)) return cleaned;
  for (let idx = 2; idx < 20; idx++) {
    const suffix = ` (${idx})`;
    const candidate = cleaned.slice(0, 31 - suffix.length) + suffix;
    if (!workbook.getWorksheet(candidate)) return candidate;
  }
  return `${cleaned.slice(0, 26)}_${Date.now().toString().slice(-4)}`;
}

function styleDualHeader(
  topRow: ExcelJS.Row,
  subRow: ExcelJS.Row,
  colCount: number,
  colorFn: (col: number) => { topBg: string; topFg: string; subBg: string; subFg: string }
) {
  for (let c = 1; c <= colCount; c++) {
    const { topBg, topFg, subBg, subFg } = colorFn(c);
    const tc = topRow.getCell(c);
    tc.fill = solidFill(topBg);
    tc.font = { name: 'Calibri', size: 10, bold: true, color: { argb: topFg } };
    tc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    tc.border = DARK_THIN_BORDER;

    const sc = subRow.getCell(c);
    sc.fill = solidFill(subBg);
    sc.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: subFg } };
    sc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    sc.border = DARK_THIN_BORDER;
  }
}

export function appendCvDiv1SheetsToWorkbook(
  workbook: ExcelJS.Workbook,
  records: CvDailyRecord[],
  activeWeekNumber: number = 45,
  activeMonthKey: string = '11',
  year: number = 2026,
  inventoryOps?: CvInventoryOpsEntry[],
  options?: {
    includePrefix?: boolean;
    compact?: boolean;
    includeSales?: boolean;
    includeDataCv?: boolean;
    includeCharts?: boolean;
    salesSheetName?: string;
    dataCvSheetName?: string;
  }
): void {
  const monthName = MONTH_NAMES_ID[activeMonthKey] || `Bulan ${activeMonthKey}`;
  const daysInMonth = new Date(year, parseInt(activeMonthKey, 10) || 11, 0).getDate();
  const monthlyData = calcCvMonthlySummary(records, activeMonthKey, year);
  const annualData = calcCvAnnualSummary(records, year);
  const opsEntries =
    inventoryOps && inventoryOps.length > 0 ? inventoryOps : loadCvInventoryOps();
  const opsById = new Map<string, CvInventoryOpsEntry>();
  opsEntries.forEach((e) => opsById.set(e.id, e));
  const liveSalesQty = calcCvLiveSalesQtyMap(records);
  const opsBrands: CvInventoryBrandKey[] = ['GRAPHIS', 'SIMPLICITY'];

  const compact = options?.compact !== false;
  const includeSales = options?.includeSales !== false;
  const includeDataCv = options?.includeDataCv !== false;
  const includeCharts = options?.includeCharts !== false;
  const pfx = options?.includePrefix && !compact ? 'CV ' : '';

  const recMap = new Map<string, CvDailyRecord>();
  records.forEach((r) => recMap.set((r.tanggal || '').trim(), r));

  const monthWeeks = ANNUAL_2026_WEEKS.filter(
    (w) =>
      w.monthKey === activeMonthKey ||
      w.slashDates.some((d) => d.split('/')[1] === activeMonthKey)
  );

  const weeklySalesValMap = new Map<string, number>();
  const monthlySalesValMap = new Map<string, number>();
  records.forEach((r) => {
    const gSales = (Number(r.graphis?.pagi?.sales) || 0) + (Number(r.graphis?.siang?.sales) || 0);
    const sSales = (Number(r.simplicity?.pagi?.sales) || 0) + (Number(r.simplicity?.siang?.sales) || 0);
    weeklySalesValMap.set(`${r.weekNumber}:GRAPHIS`, (weeklySalesValMap.get(`${r.weekNumber}:GRAPHIS`) || 0) + gSales);
    weeklySalesValMap.set(`${r.weekNumber}:SIMPLICITY`, (weeklySalesValMap.get(`${r.weekNumber}:SIMPLICITY`) || 0) + sSales);
    monthlySalesValMap.set(`${r.monthKey}:GRAPHIS`, (monthlySalesValMap.get(`${r.monthKey}:GRAPHIS`) || 0) + gSales);
    monthlySalesValMap.set(`${r.monthKey}:SIMPLICITY`, (monthlySalesValMap.get(`${r.monthKey}:SIMPLICITY`) || 0) + sSales);
  });

  if (includeSales) {
    /* =========================================================================
       COMPACT "Report CV" (Harian + Per Week + 12 Bulan dalam 1 Sheet — 19 Kolom)
       ========================================================================= */
    const wsSales = workbook.addWorksheet(
      getSafeSheetName(
        workbook,
        options?.salesSheetName || (compact ? 'Report CV' : `${pfx}Report (${monthName.slice(0, 3)})`)
      ),
      {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      }
    );

    const dWidths = [
      13.5, 11.5, 9.5, 19,
      17, 10.5, 10.5, 11.5,
      17, 10.5, 10.5, 11.5,
      19, 11.5, 12, 14.5,
      16, 12, 32,
    ];
    dWidths.forEach((w, i) => {
      wsSales.getColumn(i + 1).width = w;
    });

    addUnifiedCorporateHeader(
      wsSales,
      19,
      'REPORT KATEGORI CV — DIVISI 1 (6201 GRAPHIS & 6244 SIMPLICITY)',
      `Bulan Aktif: ${monthName.toUpperCase()} ${year} (Tgl 1–${daysInMonth}) · Penjualan Harian, Per Week & Rekap 12 Bulan`,
      EXCEL_COLORS.purpleHeader
    );

    addKpiStrip(
      wsSales,
      [
        { label: '6201 - GRAPHIS', value: monthlyData.graphis.sales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.indigoSubBg, fgArgb: EXCEL_COLORS.indigoText },
        { label: '6244 - SIMPLICITY', value: monthlyData.simplicity.sales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.amberSubBg, fgArgb: EXCEL_COLORS.amberText },
        { label: 'TOTAL SALES CV DIV. 1', value: monthlyData.total.sales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.goldLight, fgArgb: EXCEL_COLORS.black },
        { label: 'TOTAL QTY / NOTA CV', value: `${monthlyData.total.qty.toLocaleString('id-ID')} pcs | ${monthlyData.total.trx.toLocaleString('id-ID')} nota`, bgArgb: EXCEL_COLORS.purpleSubBg, fgArgb: EXCEL_COLORS.purpleText },
        { label: 'PENCAPAIAN TARGET', value: monthlyData.targetMonthly > 0 ? monthlyData.total.sales / monthlyData.targetMonthly : 0, numFmt: NUM_FMT_PCT, bgArgb: EXCEL_COLORS.emeraldSoft, fgArgb: EXCEL_COLORS.emeraldText },
      ],
      19
    );

    const sp1 = wsSales.addRow([]);
    sp1.height = 6;

    // BAGIAN A: HARIAN CV
    addSectionTitleBar(
      wsSales,
      `BAGIAN A: PENJUALAN HARIAN CV — BULAN ${monthName.toUpperCase()} ${year}`,
      19,
      EXCEL_COLORS.slate800,
      EXCEL_COLORS.yellowText
    );

    const dTop = wsSales.addRow([
      'Tanggal', 'Hari', 'Week', 'Petugas (Pagi/Siang)',
      'DEPT 6201 — BRAND: GRAPHIS', '', '', '',
      'DEPT 6244 — BRAND: SIMPLICITY', '', '', '',
      'TOTAL CV DIVISI 1', '', '', '',
      'Target & Pencapaian', '',
      'Reason / Event',
    ]);
    dTop.height = 24;

    const dSub = wsSales.addRow([
      '', '', '', '',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Total Sales', 'Total Qty', 'Total Nota', 'Price Point',
      'Target (Rp)', 'Ach %',
      '',
    ]);
    dSub.height = 25;

    [1, 2, 3, 4, 19].forEach((c) => wsSales.mergeCells(dTop.number, c, dSub.number, c));
    wsSales.mergeCells(dTop.number, 5, dTop.number, 8);
    wsSales.mergeCells(dTop.number, 9, dTop.number, 12);
    wsSales.mergeCells(dTop.number, 13, dTop.number, 16);
    wsSales.mergeCells(dTop.number, 17, dTop.number, 18);

    styleDualHeader(dTop, dSub, 19, (c) => {
      if (c >= 5 && c <= 8) return { topBg: EXCEL_COLORS.indigoHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.indigoSubBg, subFg: EXCEL_COLORS.indigoText };
      if (c >= 9 && c <= 12) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 13 && c <= 16) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 17 && c <= 18) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
      return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleHeader, subFg: EXCEL_COLORS.white };
    });
    applySheetFreeze(wsSales, 3, dSub.number);

    const firstDRow = dSub.number + 1;
    let activeDays = 0;
    const cvDailyBars: ExcelChartBarItem[] = [];

    for (let dNum = 1; dNum <= daysInMonth; dNum++) {
      const dd = String(dNum).padStart(2, '0');
      const tanggal = `${dd}/${activeMonthKey}/${year}`;
      const dayName = getIndoDayName(tanggal);
      const wknd = isWeekendDate(tanggal);
      const schedW = ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(tanggal));
      const rec = recMap.get(tanggal);
      const sum = rec ? calcCvDaySummary(rec) : null;

      const gSales = sum ? sum.graphis.sales : 0;
      const gQty = sum ? sum.graphis.qty : 0;
      const gNota = sum ? sum.graphis.trx : 0;
      const sSales = sum ? sum.simplicity.sales : 0;
      const sQty = sum ? sum.simplicity.qty : 0;
      const sNota = sum ? sum.simplicity.trx : 0;
      const totSales = gSales + sSales;
      if (totSales > 0) activeDays++;

      cvDailyBars.push({
        label: dd,
        subLabel: dayName.slice(0, 3),
        seriesAVal: gSales,
        seriesBVal: sSales,
        totalVal: totSales,
        isWeekend: wknd,
      });

      const targetDaily = sum?.targetDaily || 500000;
      const pStaff = rec?.petugasPagi?.trim();
      const sStaff = rec?.petugasSiang?.trim();
      const staffLabel = pStaff || sStaff ? `${pStaff || '-'} / ${sStaff || '-'}` : '-';
      const reason = (rec?.reasonEvent || '').trim() || '-';

      const r = wsSales.addRow([
        tanggal, dayName, `W-${rec?.weekNumber || schedW?.weekNumber || 0}`, staffLabel,
        gSales, gQty, gNota, null,
        sSales, sQty, sNota, null,
        null, null, null, null,
        targetDaily, null, reason,
      ]);
      const n = r.number;
      r.height = computeMultilineRowHeight([reason], 20, 28, 13.5);

      r.getCell(8).value = { formula: `IF(M${n}>0,E${n}/M${n},0)`, result: totSales > 0 ? gSales / totSales : 0 };
      r.getCell(12).value = { formula: `IF(M${n}>0,I${n}/M${n},0)`, result: totSales > 0 ? sSales / totSales : 0 };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: totSales };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: gQty + sQty };
      r.getCell(15).value = { formula: `G${n}+K${n}`, result: gNota + sNota };
      r.getCell(16).value = { formula: `IF(N${n}>0,ROUND(M${n}/N${n},0),0)` };
      r.getCell(18).value = { formula: `IF(Q${n}>0,M${n}/Q${n},0)` };

      const isZero = totSales === 0;
      for (let c = 1; c <= 19; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: c <= 3 || c === 13 || c === 15,
          color: { argb: wknd && c <= 2 ? EXCEL_COLORS.roseText : isZero && c >= 5 && c <= 18 ? EXCEL_COLORS.slateMuted : EXCEL_COLORS.navyDark },
        };
        if (c === 13) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (c >= 14 && c <= 16) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
        else if (wknd) cell.fill = solidFill(EXCEL_COLORS.crimsonBg);
        else if (dNum % 2 === 0) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([8, 12, 18].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if ([5, 6, 7, 9, 10, 11, 13, 14, 15, 16, 17].includes(c)) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c === 19 || c === 4) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    }

    const lastDRow = firstDRow + daysInMonth - 1;
    const totDRow = wsSales.addRow([
      `TOTAL ${monthName.toUpperCase()} (${activeDays}/${daysInMonth} HR)`,
      '', '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, '',
    ]);
    totDRow.height = 24;
    const dn = totDRow.number;
    wsSales.mergeCells(dn, 1, dn, 4);

    [5, 6, 7, 9, 10, 11, 13, 14, 15, 17].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totDRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstDRow}:${colL}${lastDRow})` };
    });
    totDRow.getCell(8).value = { formula: `IF(M${dn}>0,E${dn}/M${dn},0)` };
    totDRow.getCell(12).value = { formula: `IF(M${dn}>0,I${dn}/M${dn},0)` };
    totDRow.getCell(16).value = { formula: `IF(N${dn}>0,ROUND(M${dn}/N${dn},0),0)` };
    totDRow.getCell(18).value = { formula: `IF(Q${dn}>0,M${dn}/Q${dn},0)` };

    for (let c = 1; c <= 19; c++) {
      const cell = totDRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if ([8, 12, 18].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 5 && c <= 17) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    const avgDRow = wsSales.addRow([
      `RATA-RATA / HARI AKTIF (${activeDays} HR)`,
      '', '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, '',
    ]);
    avgDRow.height = 21;
    const adn = avgDRow.number;
    wsSales.mergeCells(adn, 1, adn, 4);
    [5, 6, 7, 9, 10, 11, 13, 14, 15, 17].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      avgDRow.getCell(cIdx).value = {
        formula: `IF(COUNTIF($M$${firstDRow}:$M$${lastDRow},">0")>0,ROUND(${colL}${dn}/COUNTIF($M$${firstDRow}:$M$${lastDRow},">0"),0),0)`,
      };
    });
    avgDRow.getCell(8).value = { formula: `H${dn}` };
    avgDRow.getCell(12).value = { formula: `L${dn}` };
    avgDRow.getCell(16).value = { formula: `P${dn}` };
    avgDRow.getCell(18).value = { formula: `R${dn}` };
    for (let c = 1; c <= 19; c++) {
      const cell = avgDRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.navyDark);
      cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: c >= 13 ? EXCEL_COLORS.yellowText : EXCEL_COLORS.white } };
      cell.border = DARK_THIN_BORDER;
      if ([8, 12, 18].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 5 && c <= 17) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && cvDailyBars.length > 0) {
      embedExcelSalesDiagramPair(workbook, wsSales, {
        salesTitle: `DIAGRAM SALES HARIAN CV DIVISI 1 — ${monthName.toUpperCase()} ${year}`,
        salesSubtitle: 'Grafik Batang Harian Dept 6201 Graphis vs Dept 6244 Simplicity & Garis Tren Penjualan',
        growthTitle: `DIAGRAM GROWTH & KONTRIBUSI HARIAN CV — ${monthName.toUpperCase()} ${year}`,
        growthSubtitle: 'Diagram Pertumbuhan Harian (Growth %) & Porsi Kontribusi Tiap Tanggal terhadap Total Bulan',
        bars: cvDailyBars,
        config: CV_CHART_CONFIG,
        rightFilterBadgeText: `${monthName.toUpperCase()} ${year} (${daysInMonth} HR)`,
      });
    }

    const spB = wsSales.addRow([]);
    spB.height = 10;

    // BAGIAN B: REKAP PER WEEK (BULAN AKTIF) — ALIGNED KE 19 KOLOM
    addSectionTitleBar(
      wsSales,
      `BAGIAN B: REKAP PER WEEK CV — BULAN ${monthName.toUpperCase()} ${year}`,
      19,
      EXCEL_COLORS.indigoHeader,
      EXCEL_COLORS.white
    );

    const wTop = wsSales.addRow([
      'Week', 'Periode Tanggal', '', 'Bulan',
      'DEPT 6201 — GRAPHIS', '', '', '',
      'DEPT 6244 — SIMPLICITY', '', '', '',
      'TOTAL WEEKLY CV', '', '', '',
      'Target Weekly', 'Ach %', 'Kontribusi %',
    ]);
    wTop.height = 24;
    const wSub = wsSales.addRow([
      '', '', '', '',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Total Sales', 'Total Qty', 'Total Nota', 'Price Point',
      '', '', '',
    ]);
    wSub.height = 25;

    [1, 4, 17, 18, 19].forEach((c) => wsSales.mergeCells(wTop.number, c, wSub.number, c));
    wsSales.mergeCells(wTop.number, 2, wSub.number, 3);
    wsSales.mergeCells(wTop.number, 5, wTop.number, 8);
    wsSales.mergeCells(wTop.number, 9, wTop.number, 12);
    wsSales.mergeCells(wTop.number, 13, wTop.number, 16);

    styleDualHeader(wTop, wSub, 19, (c) => {
      if (c >= 5 && c <= 8) return { topBg: EXCEL_COLORS.indigoHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.indigoSubBg, subFg: EXCEL_COLORS.indigoText };
      if (c >= 9 && c <= 12) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 13 && c <= 16) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 17) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldHeader, subFg: EXCEL_COLORS.white };
      return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleHeader, subFg: EXCEL_COLORS.white };
    });

    const firstPwRow = wSub.number + 1;
    const totPwRowNum = firstPwRow + monthWeeks.length;
    const cvWeeklyBars: ExcelChartBarItem[] = [];

    monthWeeks.forEach((w, idx) => {
      const wsSum = calcCvWeeklySummary(records, w.weekNumber);
      const isAct = w.weekNumber === activeWeekNumber;
      cvWeeklyBars.push({
        label: `WEEK ${w.weekNumber}`,
        subLabel: wsSum.periodeText.slice(0, 11),
        seriesAVal: wsSum.graphis.sales,
        seriesBVal: wsSum.simplicity.sales,
        totalVal: wsSum.total.sales,
      });
      const r = wsSales.addRow([
        `WEEK ${w.weekNumber}`, wsSum.periodeText, '', wsSum.monthName,
        wsSum.graphis.sales, wsSum.graphis.qty, wsSum.graphis.trx, null,
        wsSum.simplicity.sales, wsSum.simplicity.qty, wsSum.simplicity.trx, null,
        null, null, null, null,
        wsSum.targetWeekly, null, null,
      ]);
      const n = r.number;
      r.height = 21;
      wsSales.mergeCells(n, 2, n, 3);
      r.getCell(8).value = { formula: `IF(M${n}>0,E${n}/M${n},0)` };
      r.getCell(12).value = { formula: `IF(M${n}>0,I${n}/M${n},0)` };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: wsSum.total.sales };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: wsSum.total.qty };
      r.getCell(15).value = { formula: `G${n}+K${n}`, result: wsSum.total.trx };
      r.getCell(16).value = { formula: `IF(N${n}>0,ROUND(M${n}/N${n},0),0)` };
      r.getCell(18).value = { formula: `IF(Q${n}>0,M${n}/Q${n},0)` };
      r.getCell(19).value = { formula: `IF($M$${totPwRowNum}>0,M${n}/$M$${totPwRowNum},0)` };

      for (let c = 1; c <= 19; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = { name: 'Calibri', size: 9.5, bold: isAct || c === 1 || c === 13 || c === 15 };
        if (isAct && c <= 4) cell.fill = solidFill(EXCEL_COLORS.purpleSubBg);
        else if (c === 13) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([8, 12, 18, 19].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c >= 5 && c <= 17) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c === 2) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    });

    const lastPwRow = Math.max(firstPwRow, firstPwRow + monthWeeks.length - 1);
    const totPwRow = wsSales.addRow([
      `TOTAL PER WEEK (${monthName.toUpperCase()})`, '', '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, null,
    ]);
    totPwRow.height = 23;
    const pwn = totPwRow.number;
    wsSales.mergeCells(pwn, 1, pwn, 4);
    [5, 6, 7, 9, 10, 11, 13, 14, 15, 17].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totPwRow.getCell(cIdx).value = monthWeeks.length > 0 ? { formula: `SUM(${colL}${firstPwRow}:${colL}${lastPwRow})` } : 0;
    });
    totPwRow.getCell(8).value = { formula: `IF(M${pwn}>0,E${pwn}/M${pwn},0)` };
    totPwRow.getCell(12).value = { formula: `IF(M${pwn}>0,I${pwn}/M${pwn},0)` };
    totPwRow.getCell(16).value = { formula: `IF(N${pwn}>0,ROUND(M${pwn}/N${pwn},0),0)` };
    totPwRow.getCell(18).value = { formula: `IF(Q${pwn}>0,M${pwn}/Q${pwn},0)` };
    totPwRow.getCell(19).value = monthWeeks.length > 0 ? { formula: `SUM(S${firstPwRow}:S${lastPwRow})` } : 0;

    for (let c = 1; c <= 19; c++) {
      const cell = totPwRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if ([8, 12, 18, 19].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 5 && c <= 17) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && cvWeeklyBars.length > 0) {
      embedExcelSalesDiagramPair(workbook, wsSales, {
        salesTitle: `DIAGRAM SALES PER WEEK CV DIVISI 1 — ${monthName.toUpperCase()} ${year}`,
        salesSubtitle: 'Perbandingan Omzet Mingguan Dept 6201 Graphis vs Dept 6244 Simplicity & Garis Tren',
        growthTitle: `DIAGRAM GROWTH (WOW) & KONTRIBUSI PER WEEK CV — ${monthName.toUpperCase()} ${year}`,
        growthSubtitle: 'Diagram Pertumbuhan Mingguan (WoW %) & Porsi Kontribusi Tiap Week terhadap Bulan',
        bars: cvWeeklyBars,
        config: CV_CHART_CONFIG,
        rightFilterBadgeText: `${cvWeeklyBars.length} WEEKLY (${monthName.toUpperCase()})`,
      });
    }

    const spC = wsSales.addRow([]);
    spC.height = 10;

    // BAGIAN C: 12 BULAN KALENDER (JAN–DES 2026) — ALIGNED KE 19 KOLOM
    addSectionTitleBar(
      wsSales,
      `BAGIAN C: REKAPITULASI 12 BULAN KALENDER CV DIVISI 1 (JANUARI – DESEMBER ${year})`,
      19,
      EXCEL_COLORS.purpleHeader,
      EXCEL_COLORS.white
    );

    const mTop = wsSales.addRow([
      'Bulan', '', 'Hari', 'Avg/Hari',
      'DEPT 6201 — GRAPHIS', '', '', '',
      'DEPT 6244 — SIMPLICITY', '', '', '',
      'TOTAL CV DIVISI 1', '', '', '',
      'Target', 'Ach %', 'Kontribusi %',
    ]);
    mTop.height = 24;
    const mSub = wsSales.addRow([
      '', '', '', '',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Sales', 'Qty', 'Nota', 'Share %',
      'Total Sales', 'Total Qty', 'Total Nota', 'Price Point',
      '', '', '',
    ]);
    mSub.height = 25;

    [3, 4, 17, 18, 19].forEach((c) => wsSales.mergeCells(mTop.number, c, mSub.number, c));
    wsSales.mergeCells(mTop.number, 1, mSub.number, 2);
    wsSales.mergeCells(mTop.number, 5, mTop.number, 8);
    wsSales.mergeCells(mTop.number, 9, mTop.number, 12);
    wsSales.mergeCells(mTop.number, 13, mTop.number, 16);

    styleDualHeader(mTop, mSub, 19, (c) => {
      if (c >= 5 && c <= 8) return { topBg: EXCEL_COLORS.indigoHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.indigoSubBg, subFg: EXCEL_COLORS.indigoText };
      if (c >= 9 && c <= 12) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 13 && c <= 16) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 17) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldHeader, subFg: EXCEL_COLORS.white };
      return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleHeader, subFg: EXCEL_COLORS.white };
    });

    const firstMRow = mSub.number + 1;
    const totalMRowNum = firstMRow + 12;
    const cvMonthly12Bars: ExcelChartBarItem[] = [];

    annualData.monthlySummaries.forEach((m, idx) => {
      const isSel = m.monthKey === activeMonthKey;
      cvMonthly12Bars.push({
        label: m.monthName.slice(0, 3),
        subLabel: String(year),
        seriesAVal: m.graphis.sales,
        seriesBVal: m.simplicity.sales,
        totalVal: m.total.sales,
      });
      const r = wsSales.addRow([
        `${m.monthKey}. ${m.monthName}`, '', m.daysCount, null,
        m.graphis.sales, m.graphis.qty, m.graphis.trx, null,
        m.simplicity.sales, m.simplicity.qty, m.simplicity.trx, null,
        null, null, null, null,
        m.targetMonthly, null, null,
      ]);
      const n = r.number;
      r.height = 21;
      wsSales.mergeCells(n, 1, n, 2);

      r.getCell(4).value = { formula: `IF(C${n}>0,ROUND(M${n}/C${n},0),0)` };
      r.getCell(8).value = { formula: `IF(M${n}>0,E${n}/M${n},0)` };
      r.getCell(12).value = { formula: `IF(M${n}>0,I${n}/M${n},0)` };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: m.total.sales };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: m.total.qty };
      r.getCell(15).value = { formula: `G${n}+K${n}`, result: m.total.trx };
      r.getCell(16).value = { formula: `IF(N${n}>0,ROUND(M${n}/N${n},0),0)` };
      r.getCell(18).value = { formula: `IF(Q${n}>0,M${n}/Q${n},0)` };
      r.getCell(19).value = { formula: `IF($M$${totalMRowNum}>0,M${n}/$M$${totalMRowNum},0)` };

      for (let c = 1; c <= 19; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = { name: 'Calibri', size: 9.5, bold: isSel || c === 1 || c === 13 || c === 15 };
        if (isSel && c <= 3) cell.fill = solidFill(EXCEL_COLORS.purpleSubBg);
        else if (c === 13) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (c >= 14 && c <= 16) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
        else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([8, 12, 18, 19].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c >= 3 && c <= 17) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        }
      }
    });

    const lastMRow = firstMRow + 11;
    const totMRow = wsSales.addRow([
      `GRAND TOTAL 1 TAHUN (${year})`, '', null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, null,
    ]);
    totMRow.height = 24;
    const mn = totMRow.number;
    wsSales.mergeCells(mn, 1, mn, 2);
    [3, 5, 6, 7, 9, 10, 11, 13, 14, 15, 17].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totMRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstMRow}:${colL}${lastMRow})` };
    });
    totMRow.getCell(4).value = { formula: `IF(C${mn}>0,ROUND(M${mn}/C${mn},0),0)` };
    totMRow.getCell(8).value = { formula: `IF(M${mn}>0,E${mn}/M${mn},0)` };
    totMRow.getCell(12).value = { formula: `IF(M${mn}>0,I${mn}/M${mn},0)` };
    totMRow.getCell(16).value = { formula: `IF(N${mn}>0,ROUND(M${mn}/N${mn},0),0)` };
    totMRow.getCell(18).value = { formula: `IF(Q${mn}>0,M${mn}/Q${mn},0)` };
    totMRow.getCell(19).value = { formula: `SUM(S${firstMRow}:S${lastMRow})` };

    for (let c = 1; c <= 19; c++) {
      const cell = totMRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if ([8, 12, 18, 19].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 3 && c <= 17) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && cvMonthly12Bars.length > 0) {
      embedExcelSalesDiagramPair(workbook, wsSales, {
        salesTitle: `DIAGRAM SALES TAHUNAN CV DIVISI 1 — 12 BULAN (${year})`,
        salesSubtitle: 'Grafik Batang Tahunan Dept 6201 Graphis vs Dept 6244 Simplicity & Garis Tren (Jan–Des)',
        growthTitle: `DIAGRAM GROWTH (MOM) & KONTRIBUSI 12 BULAN CV (${year})`,
        growthSubtitle: 'Diagram Pertumbuhan Bulanan (MoM %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahun',
        bars: cvMonthly12Bars,
        config: CV_CHART_CONFIG,
        rightFilterBadgeText: `12 BULAN (JAN – DES ${year})`,
      });
    }
  }

  if (includeDataCv) {
    /* =========================================================================
       COMPACT "Data CV" (Monthly + Weekly Inventory, Promo, Reason & Action Plan)
       13 Aligned Columns
       ========================================================================= */
    const wsDataCv = workbook.addWorksheet(
      getSafeSheetName(
        workbook,
        options?.dataCvSheetName || 'Data CV'
      ),
      {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      }
    );

    const opsWidths = [16.5, 22, 13.5, 21, 13.5, 18.5, 13.5, 18.5, 13.5, 17, 26, 30, 30];
    opsWidths.forEach((w, i) => {
      wsDataCv.getColumn(i + 1).width = w;
    });

    addUnifiedCorporateHeader(
      wsDataCv,
      13,
      'DATA CV — MONITORING INVENTORY, PROMO, REASON/KENDALA & ACTION PLAN',
      `Divisi 1 (Dept 6201 Graphis & Dept 6244 Simplicity) · Stock Akhir = Stock Awal + Barang Datang - Sales Qty - Barang Keluar`,
      EXCEL_COLORS.indigoHeader
    );

    const sp4 = wsDataCv.addRow([]);
    sp4.height = 6;

    // BAGIAN A: DATA CV PER MONTHLY (12 BULAN)
    addSectionTitleBar(
      wsDataCv,
      `BAGIAN A: DATA CV PER MONTHLY (12 BULAN ${year})`,
      13,
      EXCEL_COLORS.purpleHeader,
      EXCEL_COLORS.white
    );

    const omTop = wsDataCv.addRow([
      'Bulan', 'Keterangan', 'Divisi', 'Brand (Dept)',
      'MONITORING INVENTORY (PCS)', '', '', '', '',
      'Sales Value (Rp)', 'Promo', 'Reason / Kendala', 'Action Plan',
    ]);
    omTop.height = 24;
    const omSub = wsDataCv.addRow([
      '', '', '', '',
      'Stock Awal', 'Barang Datang (+)', 'Sales Qty (-)', 'Barang Keluar (-)', 'Stock Akhir',
      '', '', '', '',
    ]);
    omSub.height = 26;

    [1, 2, 3, 4, 10, 11, 12, 13].forEach((c) => wsDataCv.mergeCells(omTop.number, c, omSub.number, c));
    wsDataCv.mergeCells(omTop.number, 5, omTop.number, 9);

    styleDualHeader(omTop, omSub, 13, (c) => {
      if (c >= 5 && c <= 9) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
      if (c === 10) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.gold, subFg: EXCEL_COLORS.black };
      return { topBg: EXCEL_COLORS.slate800, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.slate800, subFg: EXCEL_COLORS.yellowText };
    });
    applySheetFreeze(wsDataCv, 4, omSub.number);

    const monthKeys = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    const firstOmRow = omSub.number + 1;
    monthKeys.forEach((mKey, mIdx) => {
      opsBrands.forEach((brand) => {
        const id = getCvOpsEntryId('monthly', mKey, brand);
        const entry = opsById.get(id) || createDefaultCvOpsEntry('monthly', mKey, brand);
        const liveQty = liveSalesQty.monthly.get(mKey)?.[brand] || 0;
        const effSalesQty = entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined ? Number(entry.salesQtyOverride) || 0 : liveQty;
        const stockAwal = Number(entry.stockAwal) || 0;
        const barangDatang = Number(entry.barangDatang) || 0;
        const barangKeluar = Number(entry.barangKeluar) || 0;
        const stockAkhir = stockAwal + barangDatang - effSalesQty - barangKeluar;
        const salesVal = monthlySalesValMap.get(`${mKey}:${brand}`) || 0;
        const promoText = (entry.promo || '').trim() || '-';
        const reasonText = (entry.reasonKendala || '').trim() || '-';
        const actionText = (entry.actionPlan || '').trim() || '-';

        const r = wsDataCv.addRow([
          `${MONTH_NAMES_ID[mKey] || mKey} ${year}`,
          `Bulan ${mKey}`,
          entry.divisi || CV_DIVISI_DEFAULT,
          `${CV_BRAND_CATALOG[brand].label} (${CV_BRAND_CATALOG[brand].deptCode})`,
          stockAwal, barangDatang, effSalesQty, barangKeluar, null,
          salesVal, promoText, reasonText, actionText,
        ]);
        const n = r.number;
        r.height = computeMultilineRowHeight([promoText, reasonText, actionText], 21, 26, 13.5);
        r.getCell(9).value = { formula: `E${n}+F${n}-G${n}-H${n}`, result: stockAkhir };

        for (let c = 1; c <= 13; c++) {
          const cell = r.getCell(c);
          cell.border = THIN_BORDER;
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: c === 1 || c === 4 || c === 9 || c === 10,
            color: { argb: c === 4 ? (brand === 'GRAPHIS' ? EXCEL_COLORS.indigoHeader : EXCEL_COLORS.amberHeader) : c === 9 && stockAkhir < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark },
          };
          const isSelMonth = mKey === activeMonthKey;
          if (c === 9) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
          else if (isSelMonth && c <= 4) cell.fill = solidFill(EXCEL_COLORS.indigoSubBg);
          else if (mIdx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

          if (c >= 5 && c <= 10) {
            cell.numFmt = NUM_FMT_INT;
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else if (c === 2 || c === 3) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
          }
        }
      });
    });
    const lastOmRow = Math.max(firstOmRow, wsDataCv.rowCount);
    const totOmRow = wsDataCv.addRow([`TOTAL DATA CV MONTHLY (${year})`, '', '', '', null, null, null, null, null, null, '', '', '']);
    totOmRow.height = 23;
    wsDataCv.mergeCells(totOmRow.number, 1, totOmRow.number, 4);
    [5, 6, 7, 8, 9, 10].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totOmRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstOmRow}:${colL}${lastOmRow})` };
    });
    for (let c = 1; c <= 13; c++) {
      const cell = totOmRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if (c >= 5 && c <= 10) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    const sp5 = wsDataCv.addRow([]);
    sp5.height = 10;

    // BAGIAN B: DATA CV PER WEEK
    // Include active month weeks + any weeks with non-zero inventory/promo/notes so table stays clean & complete
    const targetWeeks = ANNUAL_2026_WEEKS.filter((w) => {
      const inMonth =
        w.monthKey === activeMonthKey ||
        w.slashDates.some((d) => d.split('/')[1] === activeMonthKey);
      const hasData = opsBrands.some((brand) => {
        const entry = opsById.get(getCvOpsEntryId('weekly', w.weekNumber, brand));
        const liveQty = liveSalesQty.weekly.get(w.weekNumber)?.[brand] || 0;
        const val = weeklySalesValMap.get(`${w.weekNumber}:${brand}`) || 0;
        return (
          liveQty > 0 ||
          val > 0 ||
          (entry &&
            ((entry.stockAwal || 0) > 0 ||
              (entry.barangDatang || 0) > 0 ||
              (entry.barangKeluar || 0) > 0 ||
              (entry.promo && entry.promo.trim() !== '') ||
              (entry.reasonKendala && entry.reasonKendala.trim() !== '') ||
              (entry.actionPlan && entry.actionPlan.trim() !== '')))
        );
      });
      return !compact || inMonth || hasData;
    });
    const weeksToRender = targetWeeks.length > 0 ? targetWeeks : ANNUAL_2026_WEEKS;

    addSectionTitleBar(
      wsDataCv,
      `BAGIAN B: DATA CV PER WEEK — BULAN ${monthName.toUpperCase()} & MINGGU AKTIF (${weeksToRender.length} WEEK)`,
      13,
      EXCEL_COLORS.indigoHeader,
      EXCEL_COLORS.white
    );

    const owTop = wsDataCv.addRow([
      'Week', 'Periode Tanggal', 'Divisi', 'Brand (Dept)',
      'MONITORING INVENTORY (PCS)', '', '', '', '',
      'Sales Value (Rp)', 'Promo', 'Reason / Kendala', 'Action Plan',
    ]);
    owTop.height = 24;
    const owSub = wsDataCv.addRow([
      '', '', '', '',
      'Stock Awal', 'Barang Datang (+)', 'Sales Qty (-)', 'Barang Keluar (-)', 'Stock Akhir',
      '', '', '', '',
    ]);
    owSub.height = 26;

    [1, 2, 3, 4, 10, 11, 12, 13].forEach((c) => wsDataCv.mergeCells(owTop.number, c, owSub.number, c));
    wsDataCv.mergeCells(owTop.number, 5, owTop.number, 9);

    styleDualHeader(owTop, owSub, 13, (c) => {
      if (c >= 5 && c <= 9) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
      if (c === 10) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.gold, subFg: EXCEL_COLORS.black };
      return { topBg: EXCEL_COLORS.slate800, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.slate800, subFg: EXCEL_COLORS.yellowText };
    });

    const firstOwRow = owSub.number + 1;
    weeksToRender.forEach((w, wIdx) => {
      opsBrands.forEach((brand) => {
        const id = getCvOpsEntryId('weekly', w.weekNumber, brand);
        const entry = opsById.get(id) || createDefaultCvOpsEntry('weekly', w.weekNumber, brand);
        const liveQty = liveSalesQty.weekly.get(w.weekNumber)?.[brand] || 0;
        const effSalesQty = entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined ? Number(entry.salesQtyOverride) || 0 : liveQty;
        const stockAwal = Number(entry.stockAwal) || 0;
        const barangDatang = Number(entry.barangDatang) || 0;
        const barangKeluar = Number(entry.barangKeluar) || 0;
        const stockAkhir = stockAwal + barangDatang - effSalesQty - barangKeluar;
        const salesVal = weeklySalesValMap.get(`${w.weekNumber}:${brand}`) || 0;
        const promoText = (entry.promo || '').trim() || '-';
        const reasonText = (entry.reasonKendala || '').trim() || '-';
        const actionText = (entry.actionPlan || '').trim() || '-';

        const r = wsDataCv.addRow([
          `W-${w.weekNumber}`, w.periodeText, entry.divisi || CV_DIVISI_DEFAULT,
          `${CV_BRAND_CATALOG[brand].label} (${CV_BRAND_CATALOG[brand].deptCode})`,
          stockAwal, barangDatang, effSalesQty, barangKeluar, null,
          salesVal, promoText, reasonText, actionText,
        ]);
        const n = r.number;
        r.height = computeMultilineRowHeight([promoText, reasonText, actionText], 21, 26, 13.5);
        r.getCell(9).value = { formula: `E${n}+F${n}-G${n}-H${n}`, result: stockAkhir };

        for (let c = 1; c <= 13; c++) {
          const cell = r.getCell(c);
          cell.border = THIN_BORDER;
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: c === 1 || c === 4 || c === 9 || c === 10,
            color: { argb: c === 4 ? (brand === 'GRAPHIS' ? EXCEL_COLORS.indigoHeader : EXCEL_COLORS.amberHeader) : c === 9 && stockAkhir < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark },
          };
          if (c === 9) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
          else if (wIdx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

          if (c >= 5 && c <= 10) {
            cell.numFmt = NUM_FMT_INT;
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
          } else if (c === 1 || c === 3) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
          } else {
            cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
          }
        }
      });
    });

    const lastOwRow = Math.max(firstOwRow, wsDataCv.rowCount);
    const totOwRow = wsDataCv.addRow(['TOTAL DATA CV WEEKLY', '', '', '', null, null, null, null, null, null, '', '', '']);
    totOwRow.height = 23;
    wsDataCv.mergeCells(totOwRow.number, 1, totOwRow.number, 4);
    [5, 6, 7, 8, 9, 10].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totOwRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstOwRow}:${colL}${lastOwRow})` };
    });
    for (let c = 1; c <= 13; c++) {
      const cell = totOwRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if (c >= 5 && c <= 10) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  }
}

export async function exportCvDiv1ToExcel(
  records: CvDailyRecord[],
  activeWeekNumber: number = 45,
  activeMonthKey: string = '11',
  year: number = 2026,
  inventoryOps?: CvInventoryOpsEntry[],
  includeCharts: boolean = true
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'REPORT BAZAR MATAHARI — Tanjung (PT MDS Retailing Tbk)';
  workbook.lastModifiedBy = 'PT MDS Retailing Tbk - CV Divisi 1';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  appendCvDiv1SheetsToWorkbook(
    workbook,
    records,
    activeWeekNumber,
    activeMonthKey,
    year,
    inventoryOps,
    { compact: true, includeSales: true, includeDataCv: true, includeCharts }
  );

  const dateStr = new Date().toISOString().slice(0, 10);
  await triggerWorkbookDownload(
    workbook,
    `Report_CV_&_Data_CV_Matahari_Tanjung_${dateStr}.xlsx`
  );
}
