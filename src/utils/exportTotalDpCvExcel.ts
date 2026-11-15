import ExcelJS from 'exceljs';
import {
  WeeklySalesTable,
  ANNUAL_2026_WEEKS,
  getIndoDayName,
  isWeekendDate,
} from '../data/initialData';
import { CvDailyRecord, MONTH_NAMES_ID } from '../data/cvDiv1Data';
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
  ExcelTrafficChartItem,
  embedExcelSalesDiagramPair,
  embedExcelTrafficConversionDiagram,
} from './excelCharts';

const DP_CV_CHART_CONFIG: ExcelChartSeriesConfig = {
  seriesALabel: '1. Report DP',
  seriesAColor: '#F59E0B',
  seriesALegendColor: '#FDE68A',
  seriesBLabel: '2. Report CV (Div. 1)',
  seriesBColor: '#7C3AED',
  seriesBLegendColor: '#DDD6FE',
  badgeText: 'DIAGRAM KONSOLIDASI DP + CV (OMZET RUPIAH)',
  badgeBg: '#059669',
  cardSubtitle: 'Grafik Batang 1. Report DP vs 2. Report CV & Garis Tren Grand Total',
};

export interface ConsolidatedCalculationData {
  dpDateMap: Map<
    string,
    {
      sales: number;
      qty: number;
      trx: number;
      pagiTraffic: number;
      siangTraffic: number;
      trafficTotal: number;
      pagiStaff: string;
      siangStaff: string;
      reasons: string[];
      weekNumber: number;
    }
  >;
  cvDateMap: Map<
    string,
    {
      sales: number;
      qty: number;
      trx: number;
      reasonEvent: string;
      weekNumber: number;
    }
  >;
}

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

export function buildConsolidatedMaps(
  weeklyTables: WeeklySalesTable[],
  cvRecords: CvDailyRecord[]
): ConsolidatedCalculationData {
  const dpDateMap = new Map<
    string,
    {
      sales: number;
      qty: number;
      trx: number;
      pagiTraffic: number;
      siangTraffic: number;
      trafficTotal: number;
      pagiStaff: string;
      siangStaff: string;
      reasons: string[];
      weekNumber: number;
    }
  >();

  weeklyTables.forEach((table) => {
    if (!Array.isArray(table.days)) return;
    table.days.forEach((day) => {
      const tgl = (day.tanggal || '').trim();
      if (!tgl) return;

      const pSales = Number(day.pagi?.sales) || 0;
      const sSales = Number(day.siang?.sales) || 0;
      const pQty = Number(day.pagi?.qty) || 0;
      const sQty = Number(day.siang?.qty) || 0;
      const pTrx = Number(day.pagi?.trx) || 0;
      const sTrx = Number(day.siang?.trx) || 0;
      const pTraffic = Number(day.pagi?.traffic) || 0;
      const sTraffic = Number(day.siang?.traffic) || 0;

      const reasons: string[] = [];
      if (day.reasonPagi && day.reasonPagi.trim()) reasons.push(`Pagi: ${day.reasonPagi.trim()}`);
      if (day.reasonSiang && day.reasonSiang.trim()) reasons.push(`Siang: ${day.reasonSiang.trim()}`);

      const entry = {
        sales: pSales + sSales,
        qty: pQty + sQty,
        trx: pTrx + sTrx,
        pagiTraffic: pTraffic,
        siangTraffic: sTraffic,
        trafficTotal: pTraffic + sTraffic,
        pagiStaff: (day.pagi?.nama || '').trim(),
        siangStaff: (day.siang?.nama || '').trim(),
        reasons,
        weekNumber: table.weekNumber,
      };

      const prev = dpDateMap.get(tgl);
      if (
        !prev ||
        entry.sales > 0 ||
        entry.trafficTotal > 0 ||
        (prev.sales === 0 && prev.trafficTotal === 0)
      ) {
        dpDateMap.set(tgl, entry);
      }
    });
  });

  const cvDateMap = new Map<
    string,
    {
      sales: number;
      qty: number;
      trx: number;
      reasonEvent: string;
      weekNumber: number;
    }
  >();

  cvRecords.forEach((r) => {
    const tgl = (r.tanggal || '').trim();
    if (!tgl) return;
    const gSales = (Number(r.graphis?.pagi?.sales) || 0) + (Number(r.graphis?.siang?.sales) || 0);
    const gQty = (Number(r.graphis?.pagi?.qty) || 0) + (Number(r.graphis?.siang?.qty) || 0);
    const gTrx = (Number(r.graphis?.pagi?.trx) || 0) + (Number(r.graphis?.siang?.trx) || 0);

    const sSales = (Number(r.simplicity?.pagi?.sales) || 0) + (Number(r.simplicity?.siang?.sales) || 0);
    const sQty = (Number(r.simplicity?.pagi?.qty) || 0) + (Number(r.simplicity?.siang?.qty) || 0);
    const sTrx = (Number(r.simplicity?.pagi?.trx) || 0) + (Number(r.simplicity?.siang?.trx) || 0);

    cvDateMap.set(tgl, {
      sales: gSales + sSales,
      qty: gQty + sQty,
      trx: gTrx + sTrx,
      reasonEvent: (r.reasonEvent || '').trim(),
      weekNumber: r.weekNumber,
    });
  });

  return { dpDateMap, cvDateMap };
}

function applyDualHeaderGroupColors(
  topRow: ExcelJS.Row,
  subRow: ExcelJS.Row,
  colCount: number,
  groupColorFn: (col: number) => { topBg: string; topFg: string; subBg: string; subFg: string }
) {
  for (let c = 1; c <= colCount; c++) {
    const { topBg, topFg, subBg, subFg } = groupColorFn(c);
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

/**
 * Renders Consolidated DP + CV + Traffic & Conversion Rate into either:
 * - compact=true  -> 1 unified sheet "Total DP+CV" (Daily + Kasir Traffic + Weekly + 12 Bulan)
 * - compact=false -> 4 detailed separate sheets (Harian, Weekly, 12 Bulan, Traffic & Conv Rate)
 */
export function appendTotalDpCvSheetsToWorkbook(
  workbook: ExcelJS.Workbook,
  weeklyTables: WeeklySalesTable[],
  cvRecords: CvDailyRecord[],
  selectedMonthKey: string = '11',
  year: number = 2026,
  options?: {
    weeklyScope?: 'month' | 'all';
    includePrefix?: boolean;
    compact?: boolean;
    includeCharts?: boolean;
    sheetName?: string;
  }
) {
  const monthName = MONTH_NAMES_ID[selectedMonthKey] || `Bulan ${selectedMonthKey}`;
  const daysInMonth = new Date(year, parseInt(selectedMonthKey, 10) || 11, 0).getDate();
  const maps = buildConsolidatedMaps(weeklyTables, cvRecords);
  const weeklyScope = options?.weeklyScope || 'month';
  const compact = options?.compact !== false;
  const includeCharts = options?.includeCharts !== false;
  const pfx = options?.includePrefix ? 'DP+CV ' : '';

  interface DailyCalcRow {
    dayNumber: number;
    tanggal: string;
    dayName: string;
    isWeekend: boolean;
    weekNumber: number;
    dpSales: number;
    dpQty: number;
    dpTrx: number;
    cvSales: number;
    cvQty: number;
    cvNota: number;
    totalSales: number;
    totalQty: number;
    totalTrxNota: number;
    pricePoint: number;
    trafficPagi: number;
    trafficSiang: number;
    trafficTotal: number;
    convRateRatio: number;
    salesPerVisitor: number;
    pagiStaff: string;
    siangStaff: string;
    reasonsText: string;
  }

  const dailyRows: DailyCalcRow[] = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dd = String(d).padStart(2, '0');
    const tanggal = `${dd}/${selectedMonthKey}/${year}`;
    const dayName = getIndoDayName(tanggal);
    const wknd = isWeekendDate(tanggal);
    const dp = maps.dpDateMap.get(tanggal);
    const cv = maps.cvDateMap.get(tanggal);
    const schedW = ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(tanggal));

    const dpSales = dp?.sales || 0;
    const dpQty = dp?.qty || 0;
    const dpTrx = dp?.trx || 0;
    const cvSales = cv?.sales || 0;
    const cvQty = cv?.qty || 0;
    const cvNota = cv?.trx || 0;

    const totalSales = dpSales + cvSales;
    const totalQty = dpQty + cvQty;
    const totalTrxNota = dpTrx + cvNota;
    const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;

    const trafficPagi = dp?.pagiTraffic || 0;
    const trafficSiang = dp?.siangTraffic || 0;
    const trafficTotal = trafficPagi + trafficSiang;
    const convRateRatio = trafficTotal > 0 ? totalTrxNota / trafficTotal : 0;
    const salesPerVisitor = trafficTotal > 0 ? Math.round(totalSales / trafficTotal) : 0;

    const reasons = [...(dp?.reasons || [])];
    if (cv?.reasonEvent) reasons.push(`CV: ${cv.reasonEvent}`);

    dailyRows.push({
      dayNumber: d,
      tanggal,
      dayName,
      isWeekend: wknd,
      weekNumber: dp?.weekNumber || cv?.weekNumber || schedW?.weekNumber || 0,
      dpSales,
      dpQty,
      dpTrx,
      cvSales,
      cvQty,
      cvNota,
      totalSales,
      totalQty,
      totalTrxNota,
      pricePoint,
      trafficPagi,
      trafficSiang,
      trafficTotal,
      convRateRatio,
      salesPerVisitor,
      pagiStaff: dp?.pagiStaff || '-',
      siangStaff: dp?.siangStaff || '-',
      reasonsText: reasons.join(' | ') || '-',
    });
  }

  const mTotDpSales = dailyRows.reduce((s, r) => s + r.dpSales, 0);
  const mTotDpQty = dailyRows.reduce((s, r) => s + r.dpQty, 0);
  const mTotDpTrx = dailyRows.reduce((s, r) => s + r.dpTrx, 0);
  const mTotCvSales = dailyRows.reduce((s, r) => s + r.cvSales, 0);
  const mTotCvQty = dailyRows.reduce((s, r) => s + r.cvQty, 0);
  const mTotCvNota = dailyRows.reduce((s, r) => s + r.cvNota, 0);
  const mTotGrandSales = mTotDpSales + mTotCvSales;
  const mTotGrandQty = mTotDpQty + mTotCvQty;
  const mTotGrandTrxNota = mTotDpTrx + mTotCvNota;
  const mTotTrafficPagi = dailyRows.reduce((s, r) => s + r.trafficPagi, 0);
  const mTotTrafficSiang = dailyRows.reduce((s, r) => s + r.trafficSiang, 0);
  const mTotTraffic = mTotTrafficPagi + mTotTrafficSiang;
  const mActiveDays = dailyRows.filter((r) => r.totalSales > 0).length;

  const weekCandidates =
    weeklyScope === 'all'
      ? ANNUAL_2026_WEEKS
      : ANNUAL_2026_WEEKS.filter(
          (w) =>
            w.monthKey === selectedMonthKey ||
            w.slashDates.some((dt) => dt.split('/')[1] === selectedMonthKey)
        );

  const weekComputed = weekCandidates.map((w) => {
    let wDpSales = 0, wDpQty = 0, wDpTrx = 0;
    let wCvSales = 0, wCvQty = 0, wCvNota = 0;
    let wTrafPagi = 0, wTrafSiang = 0;

    w.slashDates.forEach((dt) => {
      const dp = maps.dpDateMap.get(dt);
      if (dp) {
        wDpSales += dp.sales;
        wDpQty += dp.qty;
        wDpTrx += dp.trx;
        wTrafPagi += dp.pagiTraffic;
        wTrafSiang += dp.siangTraffic;
      }
      const cv = maps.cvDateMap.get(dt);
      if (cv) {
        wCvSales += cv.sales;
        wCvQty += cv.qty;
        wCvNota += cv.trx;
      }
    });
    const wTotalSales = wDpSales + wCvSales;
    const wTotalQty = wDpQty + wCvQty;
    const wTotalTrxNota = wDpTrx + wCvNota;
    const wTotalTraf = wTrafPagi + wTrafSiang;
    return {
      w,
      wDpSales,
      wDpQty,
      wDpTrx,
      wCvSales,
      wCvQty,
      wCvNota,
      wTotalSales,
      wTotalQty,
      wTotalTrxNota,
      wTrafPagi,
      wTrafSiang,
      wTotalTraf,
      daysLen: Math.max(1, w.slashDates.length),
    };
  });
  const allWeeksGrandSales = weekComputed.reduce((s, item) => s + item.wTotalSales, 0);

  const monthKeys = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
  const monthlyStats = monthKeys.map((mKey) => {
    const mName = MONTH_NAMES_ID[mKey] || `Bulan ${mKey}`;
    const dCount = new Date(year, parseInt(mKey, 10), 0).getDate();
    let mDpSales = 0, mDpQty = 0, mDpTrx = 0;
    let mCvSales = 0, mCvQty = 0, mCvNota = 0;
    let mTrafPagi = 0, mTrafSiang = 0;
    for (let day = 1; day <= dCount; day++) {
      const tgl = `${String(day).padStart(2, '0')}/${mKey}/${year}`;
      const dp = maps.dpDateMap.get(tgl);
      if (dp) {
        mDpSales += dp.sales;
        mDpQty += dp.qty;
        mDpTrx += dp.trx;
        mTrafPagi += dp.pagiTraffic;
        mTrafSiang += dp.siangTraffic;
      }
      const cv = maps.cvDateMap.get(tgl);
      if (cv) {
        mCvSales += cv.sales;
        mCvQty += cv.qty;
        mCvNota += cv.trx;
      }
    }
    return {
      mKey,
      mName,
      dCount,
      mDpSales,
      mDpQty,
      mDpTrx,
      mCvSales,
      mCvQty,
      mCvNota,
      mTotalSales: mDpSales + mCvSales,
      mTrafPagi,
      mTrafSiang,
      mTrafTotal: mTrafPagi + mTrafSiang,
    };
  });
  const grandYearSales = monthlyStats.reduce((s, m) => s + m.mTotalSales, 0);

  const dailySalesBars: ExcelChartBarItem[] = dailyRows.map((d) => ({
    label: String(d.dayNumber).padStart(2, '0'),
    subLabel: d.dayName.slice(0, 3),
    seriesAVal: d.dpSales,
    seriesBVal: d.cvSales,
    totalVal: d.totalSales,
    isWeekend: d.isWeekend,
  }));

  const dailyTrafficBars: ExcelTrafficChartItem[] = dailyRows.map((d) => ({
    label: String(d.dayNumber).padStart(2, '0'),
    subLabel: d.dayName.slice(0, 3),
    trafficPagi: d.trafficPagi,
    trafficSiang: d.trafficSiang,
    totalTraffic: d.trafficTotal,
    totalTrxNota: d.totalTrxNota,
    convRatePct: d.convRateRatio * 100,
    isWeekend: d.isWeekend,
  }));

  const weeklySalesBars: ExcelChartBarItem[] = weekComputed.map((wk) => ({
    label: `WEEK ${wk.w.weekNumber}`,
    subLabel: wk.w.periodeText.slice(0, 11),
    seriesAVal: wk.wDpSales,
    seriesBVal: wk.wCvSales,
    totalVal: wk.wTotalSales,
  }));

  const monthly12Bars: ExcelChartBarItem[] = monthlyStats.map((m) => ({
    label: m.mName.slice(0, 3),
    subLabel: String(year),
    seriesAVal: m.mDpSales,
    seriesBVal: m.mCvSales,
    totalVal: m.mTotalSales,
  }));

  if (compact) {
    /* =========================================================================
       COMPACT SINGLE SHEET: "Total DP+CV" (Harian + Traffic Kasir + Weekly + 12 Bulan)
       23 Aligned Columns
       ========================================================================= */
    const ws = workbook.addWorksheet(
      getSafeSheetName(workbook, options?.sheetName || 'Total DP+CV'),
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

    const compactWidths = [
      14, 12, 9.5,
      17, 10.5, 10.5, 11,
      17, 10.5, 10.5, 11,
      18.5, 11.5, 15, 14,
      15, 11.5, 15, 11.5, 12.5, 13.5, 14, 30,
    ];
    compactWidths.forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });

    addUnifiedCorporateHeader(
      ws,
      23,
      'TOTAL KONSOLIDASI (REPORT DP + REPORT CV + TRAFFIC KASIR & CONVERSION RATE)',
      `Bulan Aktif: ${monthName.toUpperCase()} ${year} · Harian, Weekly & Kalender 12 Bulan dalam 1 Sheet Ringkas`
    );

    addKpiStrip(
      ws,
      [
        { label: `1. REPORT DP (${monthName.slice(0, 3).toUpperCase()})`, value: mTotDpSales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.amberSubBg, fgArgb: EXCEL_COLORS.amberText },
        { label: `2. REPORT CV (${monthName.slice(0, 3).toUpperCase()})`, value: mTotCvSales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.purpleSubBg, fgArgb: EXCEL_COLORS.purpleText },
        { label: 'GRAND TOTAL (DP + CV)', value: mTotGrandSales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.goldLight, fgArgb: EXCEL_COLORS.black },
        { label: 'TOTAL TRX+NOTA', value: mTotGrandTrxNota, numFmt: NUM_FMT_INT, bgArgb: EXCEL_COLORS.emeraldSoft, fgArgb: EXCEL_COLORS.emeraldText },
        { label: 'TOTAL TRAFFIC (ORG)', value: mTotTraffic, numFmt: NUM_FMT_INT, bgArgb: EXCEL_COLORS.skySubBg, fgArgb: EXCEL_COLORS.skyHeader },
        { label: 'CONVERSION RATE %', value: mTotTraffic > 0 ? mTotGrandTrxNota / mTotTraffic : 0, numFmt: NUM_FMT_PCT, bgArgb: EXCEL_COLORS.indigoSubBg, fgArgb: EXCEL_COLORS.indigoText },
      ],
      23
    );

    const spA = ws.addRow([]);
    spA.height = 6;

    // BAGIAN A: HARIAN + TRAFFIC KASIR PAGI/SIANG
    addSectionTitleBar(
      ws,
      `BAGIAN A: TOTAL HARIAN (DP + CV) & TRAFFIC KASIR PAGI / SIANG — ${monthName.toUpperCase()} ${year}`,
      23,
      EXCEL_COLORS.slate800,
      EXCEL_COLORS.yellowText
    );

    const hTop = ws.addRow([
      'Tanggal', 'Hari', 'Week',
      '1. REPORT DP', '', '', '',
      '2. REPORT CV (DIV. 1)', '', '', '',
      '3. GRAND TOTAL (DP + CV)', '', '', '',
      '4. TRAFFIC KASIR & TOTAL CONVERSION RATE', '', '', '', '', '', '',
      'Catatan / Event',
    ]);
    hTop.height = 24;

    const hSub = ws.addRow([
      '', '', '',
      'Sales DP', 'Qty DP', 'Trx DP', 'Share %',
      'Sales CV', 'Qty CV', 'Nota CV', 'Share %',
      'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point',
      'Kasir Pagi', 'Traf Pagi', 'Kasir Siang', 'Traf Siang', 'Total Traf', 'Conv. Rate %', 'Sales / Org',
      '',
    ]);
    hSub.height = 25;

    [1, 2, 3, 23].forEach((c) => ws.mergeCells(hTop.number, c, hSub.number, c));
    ws.mergeCells(hTop.number, 4, hTop.number, 7);
    ws.mergeCells(hTop.number, 8, hTop.number, 11);
    ws.mergeCells(hTop.number, 12, hTop.number, 15);
    ws.mergeCells(hTop.number, 16, hTop.number, 22);

    applyDualHeaderGroupColors(hTop, hSub, 23, (c) => {
      if (c >= 4 && c <= 7) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 8 && c <= 11) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
      if (c >= 12 && c <= 15) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 16 && c <= 22) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
      return { topBg: EXCEL_COLORS.crimsonDark, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.crimsonDark, subFg: EXCEL_COLORS.white };
    });
    applySheetFreeze(ws, 3, hSub.number);

    const firstHRow = hSub.number + 1;
    dailyRows.forEach((d) => {
      const r = ws.addRow([
        d.tanggal, d.dayName, d.weekNumber > 0 ? `W-${d.weekNumber}` : '-',
        d.dpSales, d.dpQty, d.dpTrx, null,
        d.cvSales, d.cvQty, d.cvNota, null,
        null, null, null, null,
        d.pagiStaff, d.trafficPagi, d.siangStaff, d.trafficSiang, null, null, null,
        d.reasonsText,
      ]);
      const n = r.number;
      r.height = computeMultilineRowHeight([d.reasonsText], 20, 28, 13.5);

      r.getCell(7).value = { formula: `IF(L${n}>0,D${n}/L${n},0)`, result: d.totalSales > 0 ? d.dpSales / d.totalSales : 0 };
      r.getCell(11).value = { formula: `IF(L${n}>0,H${n}/L${n},0)`, result: d.totalSales > 0 ? d.cvSales / d.totalSales : 0 };
      r.getCell(12).value = { formula: `D${n}+H${n}`, result: d.totalSales };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: d.totalQty };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: d.totalTrxNota };
      r.getCell(15).value = { formula: `IF(M${n}>0,ROUND(L${n}/M${n},0),0)`, result: d.pricePoint };
      r.getCell(20).value = { formula: `Q${n}+S${n}`, result: d.trafficTotal };
      r.getCell(21).value = { formula: `IF(T${n}>0,N${n}/T${n},0)`, result: d.convRateRatio };
      r.getCell(22).value = { formula: `IF(T${n}>0,ROUND(L${n}/T${n},0),0)`, result: d.salesPerVisitor };

      const isEmpty = d.totalSales === 0 && d.trafficTotal === 0;
      for (let c = 1; c <= 23; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: c <= 3 || c === 12 || c === 14 || c === 20 || c === 21,
          color: {
            argb:
              d.isWeekend && c <= 2
                ? EXCEL_COLORS.roseText
                : isEmpty && c >= 4 && c <= 22
                  ? EXCEL_COLORS.slateMuted
                  : c === 21 && d.convRateRatio >= 0.45
                    ? EXCEL_COLORS.emeraldText
                    : EXCEL_COLORS.navyDark,
          },
        };
        if (c === 12) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (c >= 13 && c <= 15) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
        else if (c === 20 || c === 21) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
        else if (d.isWeekend) cell.fill = solidFill(EXCEL_COLORS.crimsonBg);
        else if (d.dayNumber % 2 === 0) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([7, 11, 21].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if ([4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 17, 19, 20, 22].includes(c)) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c === 23) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    });

    const lastHRow = Math.max(firstHRow, firstHRow + dailyRows.length - 1);
    const totHRow = ws.addRow([
      `TOTAL ${monthName.toUpperCase()} (${mActiveDays}/${daysInMonth} HARI)`,
      '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      '', null, '', null, null, null, null,
      '',
    ]);
    totHRow.height = 24;
    const tn = totHRow.number;
    ws.mergeCells(tn, 1, tn, 3);
    [4, 5, 6, 8, 9, 10, 12, 13, 14, 17, 19, 20].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totHRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstHRow}:${colL}${lastHRow})` };
    });
    totHRow.getCell(7).value = { formula: `IF(L${tn}>0,D${tn}/L${tn},0)` };
    totHRow.getCell(11).value = { formula: `IF(L${tn}>0,H${tn}/L${tn},0)` };
    totHRow.getCell(15).value = { formula: `IF(M${tn}>0,ROUND(L${tn}/M${tn},0),0)` };
    totHRow.getCell(21).value = { formula: `IF(T${tn}>0,N${tn}/T${tn},0)` };
    totHRow.getCell(22).value = { formula: `IF(T${tn}>0,ROUND(L${tn}/T${tn},0),0)` };

    for (let c = 1; c <= 23; c++) {
      const cell = totHRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: EXCEL_COLORS.black } };
      cell.border = MEDIUM_BORDER;
      if ([7, 11, 21].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 17, 19, 20, 22].includes(c)) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    const avgHRow = ws.addRow([
      `RATA-RATA / HARI AKTIF (${mActiveDays} HR)`,
      '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      '', null, '', null, null, null, null, '',
    ]);
    avgHRow.height = 21;
    const an = avgHRow.number;
    ws.mergeCells(an, 1, an, 3);
    [4, 5, 6, 8, 9, 10, 12, 13, 14, 17, 19, 20].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      avgHRow.getCell(cIdx).value = {
        formula: `IF(COUNTIF($L$${firstHRow}:$L$${lastHRow},">0")>0,ROUND(${colL}${tn}/COUNTIF($L$${firstHRow}:$L$${lastHRow},">0"),0),0)`,
      };
    });
    avgHRow.getCell(7).value = { formula: `G${tn}` };
    avgHRow.getCell(11).value = { formula: `K${tn}` };
    avgHRow.getCell(15).value = { formula: `O${tn}` };
    avgHRow.getCell(21).value = { formula: `U${tn}` };
    avgHRow.getCell(22).value = { formula: `V${tn}` };

    for (let c = 1; c <= 23; c++) {
      const cell = avgHRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.navyDark);
      cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: c >= 12 ? EXCEL_COLORS.yellowText : EXCEL_COLORS.white } };
      cell.border = DARK_THIN_BORDER;
      if ([7, 11, 21].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 17, 19, 20, 22].includes(c)) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && dailySalesBars.length > 0) {
      embedExcelSalesDiagramPair(workbook, ws, {
        salesTitle: `DIAGRAM SALES HARIAN (DP + CV) — ${monthName.toUpperCase()} ${year}`,
        salesSubtitle: 'Grafik Batang Harian 1. Report DP vs 2. Report CV & Garis Tren Grand Total',
        growthTitle: `DIAGRAM GROWTH & KONTRIBUSI HARIAN (DP + CV) — ${monthName.toUpperCase()} ${year}`,
        growthSubtitle: 'Diagram Pertumbuhan Harian (Growth %) & Porsi Kontribusi Tiap Tanggal terhadap Total Bulan',
        bars: dailySalesBars,
        config: DP_CV_CHART_CONFIG,
        rightFilterBadgeText: `${monthName.toUpperCase()} ${year} (${daysInMonth} HR)`,
      });
      embedExcelTrafficConversionDiagram(workbook, ws, {
        title: `DIAGRAM TRAFFIC KASIR (PAGI & SIANG) & CONVERSION RATE — ${monthName.toUpperCase()} ${year}`,
        subtitle: 'Jumlah Pengunjung Shift Pagi + Siang vs Persentase Conversion Rate (Total Trx+Nota / Traffic)',
        items: dailyTrafficBars,
        rightFilterBadgeText: `TOTAL TRAFFIC: ${mTotTraffic.toLocaleString('id-ID')} ORG`,
      });
    }

    const spB = ws.addRow([]);
    spB.height = 10;

    // BAGIAN B: WEEKLY KONSOLIDASI (DP + CV + TRAFFIC)
    addSectionTitleBar(
      ws,
      `BAGIAN B: TOTAL WEEKLY KONSOLIDASI (DP + CV + TRAFFIC) — ${weeklyScope === 'all' ? 'TAHUN 2026' : `BULAN ${monthName.toUpperCase()} ${year}`}`,
      23,
      EXCEL_COLORS.indigoHeader,
      EXCEL_COLORS.white
    );

    const wTop = ws.addRow([
      'Week', 'Periode Tanggal', '',
      '1. REPORT DP', '', '', '',
      '2. REPORT CV', '', '', '',
      '3. GRAND TOTAL WEEKLY (DP + CV)', '', '', '',
      'Rata2/Hari', 'Traf Pagi', 'Bulan', 'Traf Siang', 'Total Traf', 'Conv. Rate %', 'Kontribusi %', 'WoW Growth %',
    ]);
    wTop.height = 24;
    const wSub = ws.addRow([
      '', '', '',
      'Sales DP', 'Qty DP', 'Trx DP', 'Share %',
      'Sales CV', 'Qty CV', 'Nota CV', 'Share %',
      'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point',
      '', '', '', '', '', '', '', '',
    ]);
    wSub.height = 25;

    [1, 16, 17, 18, 19, 20, 21, 22, 23].forEach((c) => ws.mergeCells(wTop.number, c, wSub.number, c));
    ws.mergeCells(wTop.number, 2, wSub.number, 3);
    ws.mergeCells(wTop.number, 4, wTop.number, 7);
    ws.mergeCells(wTop.number, 8, wTop.number, 11);
    ws.mergeCells(wTop.number, 12, wTop.number, 15);

    applyDualHeaderGroupColors(wTop, wSub, 23, (c) => {
      if (c >= 4 && c <= 7) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 8 && c <= 11) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
      if (c >= 12 && c <= 15) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 16 && c <= 21) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldHeader, subFg: EXCEL_COLORS.white };
      if (c >= 22) return { topBg: EXCEL_COLORS.navyDark, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.navyDark, subFg: EXCEL_COLORS.yellowText };
      return { topBg: EXCEL_COLORS.crimsonDark, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.crimsonDark, subFg: EXCEL_COLORS.white };
    });

    const firstWRow = wSub.number + 1;
    const totWRowNum = firstWRow + weekComputed.length;

    weekComputed.forEach((wk, idx) => {
      const prevSales = idx > 0 ? weekComputed[idx - 1].wTotalSales : 0;
      const wowRatio = idx > 0 && prevSales > 0 ? (wk.wTotalSales - prevSales) / prevSales : 0;

      const r = ws.addRow([
        `WEEK ${wk.w.weekNumber}`, wk.w.periodeText, '',
        wk.wDpSales, wk.wDpQty, wk.wDpTrx, null,
        wk.wCvSales, wk.wCvQty, wk.wCvNota, null,
        null, null, null, null,
        null, wk.wTrafPagi, wk.w.monthName, wk.wTrafSiang, null, null, null, null,
      ]);
      const n = r.number;
      r.height = 21;
      ws.mergeCells(n, 2, n, 3);

      r.getCell(7).value = { formula: `IF(L${n}>0,D${n}/L${n},0)` };
      r.getCell(11).value = { formula: `IF(L${n}>0,H${n}/L${n},0)` };
      r.getCell(12).value = { formula: `D${n}+H${n}`, result: wk.wTotalSales };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: wk.wTotalQty };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: wk.wTotalTrxNota };
      r.getCell(15).value = { formula: `IF(M${n}>0,ROUND(L${n}/M${n},0),0)` };
      r.getCell(16).value = { formula: `ROUND(L${n}/${wk.daysLen},0)` };
      r.getCell(20).value = { formula: `Q${n}+S${n}`, result: wk.wTotalTraf };
      r.getCell(21).value = { formula: `IF(T${n}>0,N${n}/T${n},0)` };
      r.getCell(22).value = { formula: `IF($L$${totWRowNum}>0,L${n}/$L$${totWRowNum},0)`, result: allWeeksGrandSales > 0 ? wk.wTotalSales / allWeeksGrandSales : 0 };
      r.getCell(23).value = idx === 0 ? 0 : { formula: `IF(L${n - 1}>0,(L${n}-L${n - 1})/L${n - 1},0)`, result: wowRatio };

      for (let c = 1; c <= 23; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: c === 1 || c === 12 || c === 14 || c >= 20,
          color: { argb: c === 23 ? (wowRatio > 0 ? EXCEL_COLORS.emeraldText : wowRatio < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark) : EXCEL_COLORS.navyDark },
        };
        if (c === 12) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (c >= 13 && c <= 15) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
        else if (c === 20 || c === 21) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
        else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([7, 11, 21, 22, 23].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if ([4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 19, 20].includes(c)) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c === 2) {
          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    });

    const lastWRow = Math.max(firstWRow, firstWRow + weekComputed.length - 1);
    const totWRow = ws.addRow([
      `TOTAL WEEKLY (${weekComputed.length} MINGGU)`, '', '',
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, '-', null, null, null, null, '-',
    ]);
    totWRow.height = 24;
    const wn = totWRow.number;
    ws.mergeCells(wn, 1, wn, 3);
    [4, 5, 6, 8, 9, 10, 12, 13, 14, 17, 19, 20].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totWRow.getCell(cIdx).value = weekComputed.length > 0 ? { formula: `SUM(${colL}${firstWRow}:${colL}${lastWRow})` } : 0;
    });
    totWRow.getCell(7).value = { formula: `IF(L${wn}>0,D${wn}/L${wn},0)` };
    totWRow.getCell(11).value = { formula: `IF(L${wn}>0,H${wn}/L${wn},0)` };
    totWRow.getCell(15).value = { formula: `IF(M${wn}>0,ROUND(L${wn}/M${wn},0),0)` };
    totWRow.getCell(16).value = { formula: `ROUND(L${wn}/${Math.max(1, weekComputed.length * 7)},0)` };
    totWRow.getCell(21).value = { formula: `IF(T${wn}>0,N${wn}/T${wn},0)` };
    totWRow.getCell(22).value = weekComputed.length > 0 ? { formula: `SUM(V${firstWRow}:V${lastWRow})` } : 0;

    for (let c = 1; c <= 23; c++) {
      const cell = totWRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if ([7, 11, 21, 22].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 19, 20].includes(c)) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && weeklySalesBars.length > 0) {
      embedExcelSalesDiagramPair(workbook, ws, {
        salesTitle: `DIAGRAM SALES WEEKLY (DP + CV) — ${weeklyScope === 'all' ? `TAHUN ${year}` : `${monthName.toUpperCase()} ${year}`}`,
        salesSubtitle: 'Perbandingan Omzet Mingguan 1. Report DP vs 2. Report CV & Garis Tren Grand Total',
        growthTitle: `DIAGRAM GROWTH (WOW) & KONTRIBUSI WEEKLY (DP + CV)`,
        growthSubtitle: 'Diagram Pertumbuhan Mingguan (WoW %) & Porsi Kontribusi Tiap Week terhadap Periode',
        bars: weeklySalesBars,
        config: DP_CV_CHART_CONFIG,
        rightFilterBadgeText: `${weeklySalesBars.length} WEEKLY`,
      });
    }

    const spC = ws.addRow([]);
    spC.height = 10;

    // BAGIAN C: 12 BULAN KALENDER (2026) + TARGET BENCHMARK
    addSectionTitleBar(
      ws,
      `BAGIAN C: TOTAL KONSOLIDASI 12 BULAN KALENDER (JANUARI – DESEMBER ${year}) · TARGET TAHUNAN RP 540.000.000`,
      23,
      EXCEL_COLORS.emeraldHeader,
      EXCEL_COLORS.white
    );

    const aTop = ws.addRow([
      'Bulan', '', 'Hari',
      '1. REPORT DP', '', '', '',
      '2. REPORT CV', '', '', '',
      '3. GRAND TOTAL BAZAR (DP + CV)', '', '', '',
      'Rata2/Hari', 'Traf Pagi', 'Sales/Org', 'Traf Siang', 'Total Traf', 'Conv. Rate %', 'Kontribusi %', 'MoM Growth %',
    ]);
    aTop.height = 24;
    const aSub = ws.addRow([
      '', '', '',
      'Sales DP', 'Qty DP', 'Trx DP', 'Share %',
      'Sales CV', 'Qty CV', 'Nota CV', 'Share %',
      'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point',
      '', '', '', '', '', '', '', '',
    ]);
    aSub.height = 25;

    [3, 16, 17, 18, 19, 20, 21, 22, 23].forEach((c) => ws.mergeCells(aTop.number, c, aSub.number, c));
    ws.mergeCells(aTop.number, 1, aSub.number, 2);
    ws.mergeCells(aTop.number, 4, aTop.number, 7);
    ws.mergeCells(aTop.number, 8, aTop.number, 11);
    ws.mergeCells(aTop.number, 12, aTop.number, 15);

    applyDualHeaderGroupColors(aTop, aSub, 23, (c) => {
      if (c >= 4 && c <= 7) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
      if (c >= 8 && c <= 11) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
      if (c >= 12 && c <= 15) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
      if (c >= 16 && c <= 21) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldHeader, subFg: EXCEL_COLORS.white };
      if (c >= 22) return { topBg: EXCEL_COLORS.navyDark, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.navyDark, subFg: EXCEL_COLORS.yellowText };
      return { topBg: EXCEL_COLORS.crimsonDark, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.crimsonDark, subFg: EXCEL_COLORS.white };
    });

    const firstARow = aSub.number + 1;
    const totalARowNum = firstARow + 12;

    monthlyStats.forEach((m, idx) => {
      const prevS = idx > 0 ? monthlyStats[idx - 1].mTotalSales : 0;
      const momRatio = idx > 0 && prevS > 0 ? (m.mTotalSales - prevS) / prevS : 0;
      const isSel = m.mKey === selectedMonthKey;

      const r = ws.addRow([
        `${m.mKey}. ${m.mName}`, '', m.dCount,
        m.mDpSales, m.mDpQty, m.mDpTrx, null,
        m.mCvSales, m.mCvQty, m.mCvNota, null,
        null, null, null, null,
        null, m.mTrafPagi, null, m.mTrafSiang, null, null, null, null,
      ]);
      const n = r.number;
      r.height = 21;
      ws.mergeCells(n, 1, n, 2);

      r.getCell(7).value = { formula: `IF(L${n}>0,D${n}/L${n},0)` };
      r.getCell(11).value = { formula: `IF(L${n}>0,H${n}/L${n},0)` };
      r.getCell(12).value = { formula: `D${n}+H${n}`, result: m.mTotalSales };
      r.getCell(13).value = { formula: `E${n}+I${n}`, result: m.mDpQty + m.mCvQty };
      r.getCell(14).value = { formula: `F${n}+J${n}`, result: m.mDpTrx + m.mCvNota };
      r.getCell(15).value = { formula: `IF(M${n}>0,ROUND(L${n}/M${n},0),0)` };
      r.getCell(16).value = { formula: `IF(C${n}>0,ROUND(L${n}/C${n},0),0)` };
      r.getCell(18).value = { formula: `IF(T${n}>0,ROUND(L${n}/T${n},0),0)` };
      r.getCell(20).value = { formula: `Q${n}+S${n}`, result: m.mTrafTotal };
      r.getCell(21).value = { formula: `IF(T${n}>0,N${n}/T${n},0)` };
      r.getCell(22).value = { formula: `IF($L$${totalARowNum}>0,L${n}/$L$${totalARowNum},0)`, result: grandYearSales > 0 ? m.mTotalSales / grandYearSales : 0 };
      r.getCell(23).value = idx === 0 ? 0 : { formula: `IF(L${n - 1}>0,(L${n}-L${n - 1})/L${n - 1},0)`, result: momRatio };

      for (let c = 1; c <= 23; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: isSel || c === 1 || c === 12 || c === 14 || c >= 20,
          color: {
            argb:
              c === 23
                ? momRatio > 0 ? EXCEL_COLORS.emeraldText : momRatio < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark
                : isSel && c <= 2 ? EXCEL_COLORS.indigoHeader : EXCEL_COLORS.navyDark,
          },
        };
        if (isSel && c <= 3) cell.fill = solidFill(EXCEL_COLORS.indigoSubBg);
        else if (c === 12) cell.fill = solidFill(EXCEL_COLORS.goldLight);
        else if (c >= 13 && c <= 15) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
        else if (c === 20 || c === 21) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
        else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

        if ([7, 11, 21, 22, 23].includes(c)) {
          cell.numFmt = NUM_FMT_PCT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if ([3, 4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20].includes(c)) {
          cell.numFmt = NUM_FMT_INT;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        }
      }
    });

    const lastARow = firstARow + 11;
    const totARow = ws.addRow([
      `GRAND TOTAL 1 TAHUN (${year})`, '', null,
      null, null, null, null,
      null, null, null, null,
      null, null, null, null,
      null, null, null, null, null, null, null, '-',
    ]);
    totARow.height = 25;
    const anTot = totARow.number;
    ws.mergeCells(anTot, 1, anTot, 2);
    [3, 4, 5, 6, 8, 9, 10, 12, 13, 14, 17, 19, 20].forEach((cIdx) => {
      const colL = colToLetter(cIdx);
      totARow.getCell(cIdx).value = { formula: `SUM(${colL}${firstARow}:${colL}${lastARow})` };
    });
    totARow.getCell(7).value = { formula: `IF(L${anTot}>0,D${anTot}/L${anTot},0)` };
    totARow.getCell(11).value = { formula: `IF(L${anTot}>0,H${anTot}/L${anTot},0)` };
    totARow.getCell(15).value = { formula: `IF(M${anTot}>0,ROUND(L${anTot}/M${anTot},0),0)` };
    totARow.getCell(16).value = { formula: `IF(C${anTot}>0,ROUND(L${anTot}/C${anTot},0),0)` };
    totARow.getCell(18).value = { formula: `IF(T${anTot}>0,ROUND(L${anTot}/T${anTot},0),0)` };
    totARow.getCell(21).value = { formula: `IF(T${anTot}>0,N${anTot}/T${anTot},0)` };
    totARow.getCell(22).value = { formula: `SUM(V${firstARow}:V${lastARow})` };

    for (let c = 1; c <= 23; c++) {
      const cell = totARow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.gold);
      cell.font = { name: 'Calibri', size: 10, bold: true };
      cell.border = MEDIUM_BORDER;
      if ([7, 11, 21, 22].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 3 && c <= 20) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    const targetRow = ws.addRow([
      'TARGET ANNUAL BAZAR (RP 540 JT)', '', 365,
      null, null, null, null, null, null, null, null,
      540000000, 'Pencapaian % ->', null, null, null,
      null, null, null, null, null, null, '-',
    ]);
    targetRow.height = 22;
    const tRn = targetRow.number;
    ws.mergeCells(tRn, 1, tRn, 2);
    ws.mergeCells(tRn, 13, tRn, 15);
    targetRow.getCell(16).value = { formula: `ROUND(L${tRn}/365,0)` };
    targetRow.getCell(22).value = { formula: `IF(L${tRn}>0,L${anTot}/L${tRn},0)` };
    for (let c = 1; c <= 23; c++) {
      const cell = targetRow.getCell(c);
      cell.fill = solidFill(EXCEL_COLORS.navyDark);
      cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: c === 12 || c === 22 ? EXCEL_COLORS.yellowText : EXCEL_COLORS.white } };
      cell.border = DARK_THIN_BORDER;
      if (c === 22) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([3, 12, 16].includes(c)) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    if (includeCharts && monthly12Bars.length > 0) {
      embedExcelSalesDiagramPair(workbook, ws, {
        salesTitle: `DIAGRAM SALES TAHUNAN (DP + CV) — 12 BULAN (${year})`,
        salesSubtitle: 'Grafik Batang Tahunan 1. Report DP vs 2. Report CV & Garis Tren (Januari – Desember)',
        growthTitle: `DIAGRAM GROWTH (MOM) & KONTRIBUSI 12 BULAN (DP + CV) (${year})`,
        growthSubtitle: 'Diagram Pertumbuhan Bulanan (MoM %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahunan',
        bars: monthly12Bars,
        config: DP_CV_CHART_CONFIG,
        rightFilterBadgeText: `12 BULAN (JAN – DES ${year})`,
      });
    }

    return;
  }

  /* =========================================================================
     DETAILED MODE (4 SEPARATE WORKSHEETS) — IF compact=false
     ========================================================================= */
  const wsHarian = workbook.addWorksheet(
    getSafeSheetName(workbook, `${pfx}Total Harian (${monthName.slice(0, 3)})`),
    {
      views: [{ showGridLines: true }],
      pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    }
  );
  [13.5, 11.5, 9.5, 17, 10.5, 10.5, 11, 17, 10.5, 10.5, 11, 18.5, 11.5, 15, 14, 11.5, 11.5, 13, 13.5, 15, 34].forEach((w, i) => {
    wsHarian.getColumn(i + 1).width = w;
  });
  addUnifiedCorporateHeader(wsHarian, 21, 'TOTAL KONSOLIDASI HARIAN (REPORT DP + REPORT CV + CONVERSION RATE)', `Periode: ${monthName.toUpperCase()} ${year}`);
  addKpiStrip(
    wsHarian,
    [
      { label: '1. REPORT DP', value: mTotDpSales, numFmt: NUM_IDR_FMT_SAFE(NUM_FMT_IDR), bgArgb: EXCEL_COLORS.amberSubBg, fgArgb: EXCEL_COLORS.amberText },
      { label: '2. REPORT CV', value: mTotCvSales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.purpleSubBg, fgArgb: EXCEL_COLORS.purpleText },
      { label: 'GRAND TOTAL (DP + CV)', value: mTotGrandSales, numFmt: NUM_FMT_IDR, bgArgb: EXCEL_COLORS.goldLight, fgArgb: EXCEL_COLORS.black },
      { label: 'TOTAL TRX+NOTA', value: mTotGrandTrxNota, numFmt: NUM_FMT_INT, bgArgb: EXCEL_COLORS.emeraldSoft, fgArgb: EXCEL_COLORS.emeraldText },
      { label: 'TOTAL TRAFFIC', value: mTotTraffic, numFmt: NUM_FMT_INT, bgArgb: EXCEL_COLORS.skySubBg, fgArgb: EXCEL_COLORS.skyHeader },
      { label: 'CONV. RATE %', value: mTotTraffic > 0 ? mTotGrandTrxNota / mTotTraffic : 0, numFmt: NUM_FMT_PCT, bgArgb: EXCEL_COLORS.indigoSubBg, fgArgb: EXCEL_COLORS.indigoText },
    ],
    21
  );
  const sp1 = wsHarian.addRow([]);
  sp1.height = 6;
  const hTop = wsHarian.addRow(['Tanggal', 'Hari', 'Week', '1. REPORT DP', '', '', '', '2. REPORT CV', '', '', '', '3. GRAND TOTAL', '', '', '', '4. TRAFFIC & CONV. RATE', '', '', '', '', 'Catatan / Event']);
  hTop.height = 24;
  const hSub = wsHarian.addRow(['', '', '', 'Sales DP', 'Qty DP', 'Trx DP', 'Share %', 'Sales CV', 'Qty CV', 'Nota CV', 'Share %', 'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point', 'Traf Pagi', 'Traf Siang', 'Total Traffic', 'Conv. Rate %', 'Sales / Org', '']);
  hSub.height = 25;
  [1, 2, 3, 21].forEach((c) => wsHarian.mergeCells(hTop.number, c, hSub.number, c));
  wsHarian.mergeCells(hTop.number, 4, hTop.number, 7);
  wsHarian.mergeCells(hTop.number, 8, hTop.number, 11);
  wsHarian.mergeCells(hTop.number, 12, hTop.number, 15);
  wsHarian.mergeCells(hTop.number, 16, hTop.number, 20);
  applyDualHeaderGroupColors(hTop, hSub, 21, (c) => {
    if (c >= 4 && c <= 7) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
    if (c >= 8 && c <= 11) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
    if (c >= 12 && c <= 15) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
    if (c >= 16 && c <= 20) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
    return { topBg: EXCEL_COLORS.crimsonDark, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.crimsonDark, subFg: EXCEL_COLORS.white };
  });
  applySheetFreeze(wsHarian, 3, hSub.number);
  const firstHRow = hSub.number + 1;
  dailyRows.forEach((d) => {
    const r = wsHarian.addRow([d.tanggal, d.dayName, d.weekNumber > 0 ? `W-${d.weekNumber}` : '-', d.dpSales, d.dpQty, d.dpTrx, null, d.cvSales, d.cvQty, d.cvNota, null, null, null, null, null, d.trafficPagi, d.trafficSiang, null, null, null, d.reasonsText]);
    const n = r.number;
    r.height = computeMultilineRowHeight([d.reasonsText], 20, 32, 13.5);
    r.getCell(7).value = { formula: `IF(L${n}>0,D${n}/L${n},0)`, result: d.totalSales > 0 ? d.dpSales / d.totalSales : 0 };
    r.getCell(11).value = { formula: `IF(L${n}>0,H${n}/L${n},0)`, result: d.totalSales > 0 ? d.cvSales / d.totalSales : 0 };
    r.getCell(12).value = { formula: `D${n}+H${n}`, result: d.totalSales };
    r.getCell(13).value = { formula: `E${n}+I${n}`, result: d.totalQty };
    r.getCell(14).value = { formula: `F${n}+J${n}`, result: d.totalTrxNota };
    r.getCell(15).value = { formula: `IF(M${n}>0,ROUND(L${n}/M${n},0),0)`, result: d.pricePoint };
    r.getCell(18).value = { formula: `P${n}+Q${n}`, result: d.trafficTotal };
    r.getCell(19).value = { formula: `IF(R${n}>0,N${n}/R${n},0)`, result: d.convRateRatio };
    r.getCell(20).value = { formula: `IF(R${n}>0,ROUND(L${n}/R${n},0),0)`, result: d.salesPerVisitor };
    const isEmpty = d.totalSales === 0 && d.trafficTotal === 0;
    for (let c = 1; c <= 21; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: c <= 3 || c === 12 || c === 14 || c === 18 || c === 19,
        color: {
          argb:
            d.isWeekend && c <= 2
              ? EXCEL_COLORS.roseText
              : isEmpty && c >= 4 && c <= 20
                ? EXCEL_COLORS.slateMuted
                : c === 19 && d.convRateRatio >= 0.45
                  ? EXCEL_COLORS.emeraldText
                  : EXCEL_COLORS.navyDark,
        },
      };
      if (c === 12) cell.fill = solidFill(EXCEL_COLORS.goldLight);
      else if (c >= 13 && c <= 15) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
      else if (c === 18 || c === 19) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
      else if (d.isWeekend) cell.fill = solidFill(EXCEL_COLORS.crimsonBg);
      else if (d.dayNumber % 2 === 0) cell.fill = solidFill(EXCEL_COLORS.slate50);

      if ([7, 11, 19].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 4 && c <= 20) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 21) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });
  const lastHRow = Math.max(firstHRow, firstHRow + dailyRows.length - 1);
  const totHRow = wsHarian.addRow([`TOTAL ${monthName.toUpperCase()} (${mActiveDays}/${daysInMonth} HARI)`, '', '', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, '']);
  totHRow.height = 24;
  const tn = totHRow.number;
  wsHarian.mergeCells(tn, 1, tn, 3);
  [4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18].forEach((cIdx) => {
    const colL = colToLetter(cIdx);
    totHRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstHRow}:${colL}${lastHRow})` };
  });
  totHRow.getCell(7).value = { formula: `IF(L${tn}>0,D${tn}/L${tn},0)` };
  totHRow.getCell(11).value = { formula: `IF(L${tn}>0,H${tn}/L${tn},0)` };
  totHRow.getCell(15).value = { formula: `IF(M${tn}>0,ROUND(L${tn}/M${tn},0),0)` };
  totHRow.getCell(19).value = { formula: `IF(R${tn}>0,N${tn}/R${tn},0)` };
  totHRow.getCell(20).value = { formula: `IF(R${tn}>0,ROUND(L${tn}/R${tn},0),0)` };
  for (let c = 1; c <= 21; c++) {
    const cell = totHRow.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: EXCEL_COLORS.black } };
    cell.border = MEDIUM_BORDER;
    if ([7, 11, 19].includes(c)) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c >= 4 && c <= 20) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  const avgHRowD = wsHarian.addRow([`RATA-RATA / HARI AKTIF (${mActiveDays} HR)`, '', '', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, '']);
  avgHRowD.height = 21;
  const anD = avgHRowD.number;
  wsHarian.mergeCells(anD, 1, anD, 3);
  [4, 5, 6, 8, 9, 10, 12, 13, 14, 16, 17, 18].forEach((cIdx) => {
    const colL = colToLetter(cIdx);
    avgHRowD.getCell(cIdx).value = {
      formula: `IF(COUNTIF($L$${firstHRow}:$L$${lastHRow},">0")>0,ROUND(${colL}${tn}/COUNTIF($L$${firstHRow}:$L$${lastHRow},">0"),0),0)`,
    };
  });
  avgHRowD.getCell(7).value = { formula: `G${tn}` };
  avgHRowD.getCell(11).value = { formula: `K${tn}` };
  avgHRowD.getCell(15).value = { formula: `O${tn}` };
  avgHRowD.getCell(19).value = { formula: `S${tn}` };
  avgHRowD.getCell(20).value = { formula: `T${tn}` };
  for (let c = 1; c <= 21; c++) {
    const cell = avgHRowD.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.navyDark);
    cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: c >= 12 ? EXCEL_COLORS.yellowText : EXCEL_COLORS.white } };
    cell.border = DARK_THIN_BORDER;
    if ([7, 11, 19].includes(c)) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c >= 4 && c <= 20) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  if (includeCharts && dailySalesBars.length > 0) {
    embedExcelSalesDiagramPair(workbook, wsHarian, {
      salesTitle: `DIAGRAM SALES HARIAN (DP + CV) — ${monthName.toUpperCase()} ${year}`,
      salesSubtitle: 'Grafik Batang Harian 1. Report DP vs 2. Report CV & Garis Tren Grand Total',
      growthTitle: `DIAGRAM GROWTH & KONTRIBUSI HARIAN (DP + CV) — ${monthName.toUpperCase()} ${year}`,
      growthSubtitle: 'Diagram Pertumbuhan Harian (Growth %) & Kontribusi per Tanggal',
      bars: dailySalesBars,
      config: DP_CV_CHART_CONFIG,
      rightFilterBadgeText: `${monthName.toUpperCase()} ${year}`,
    });
  }

  // Detailed Sheet 2: Total Weekly (DP+CV)
  const wsWeekly = workbook.addWorksheet(getSafeSheetName(workbook, `${pfx}Total Weekly (DP+CV)`), {
    views: [{ showGridLines: true }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  [12, 24, 13, 17, 10.5, 10.5, 11, 17, 10.5, 10.5, 11, 18.5, 11.5, 15, 14, 14.5, 11.5, 11.5, 13, 13.5, 13, 13].forEach((w, i) => (wsWeekly.getColumn(i + 1).width = w));
  addUnifiedCorporateHeader(wsWeekly, 22, 'TOTAL KONSOLIDASI WEEKLY (DP + CV)', `${monthName.toUpperCase()} ${year}`);
  const wTopD = wsWeekly.addRow(['Week', 'Periode', 'Bulan', '1. REPORT DP', '', '', '', '2. REPORT CV', '', '', '', '3. GRAND TOTAL', '', '', '', '', '4. TRAFFIC', '', '', '', 'Kontribusi %', 'WoW Growth %']);
  wTopD.height = 24;
  const wSubD = wsWeekly.addRow(['', '', '', 'Sales DP', 'Qty DP', 'Trx DP', 'Share %', 'Sales CV', 'Qty CV', 'Nota CV', 'Share %', 'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point', 'Rata2/Hari', 'Traf Pagi', 'Traf Siang', 'Total Traffic', 'Conv. Rate %', '', '']);
  wSubD.height = 25;
  [1, 2, 3, 21, 22].forEach((c) => wsWeekly.mergeCells(wTopD.number, c, wSubD.number, c));
  wsWeekly.mergeCells(wTopD.number, 4, wTopD.number, 7);
  wsWeekly.mergeCells(wTopD.number, 8, wTopD.number, 11);
  wsWeekly.mergeCells(wTopD.number, 12, wTopD.number, 16);
  wsWeekly.mergeCells(wTopD.number, 17, wTopD.number, 20);
  applyDualHeaderGroupColors(wTopD, wSubD, 22, (c) => {
    if (c >= 4 && c <= 7) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
    if (c >= 8 && c <= 11) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
    if (c >= 12 && c <= 16) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
    if (c >= 17 && c <= 20) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
    return { topBg: EXCEL_COLORS.navyDark, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.navyDark, subFg: EXCEL_COLORS.yellowText };
  });
  applySheetFreeze(wsWeekly, 3, wSubD.number);
  const firstWRowD = wSubD.number + 1;
  const totWRowDNum = firstWRowD + weekComputed.length;
  weekComputed.forEach((wk, idx) => {
    const prevSales = idx > 0 ? weekComputed[idx - 1].wTotalSales : 0;
    const wowRatio = idx > 0 && prevSales > 0 ? (wk.wTotalSales - prevSales) / prevSales : 0;
    const ppVal = wk.wTotalQty > 0 ? Math.round(wk.wTotalSales / wk.wTotalQty) : 0;
    const avgVal = Math.round(wk.wTotalSales / Math.max(1, wk.daysLen));
    const r = wsWeekly.addRow([
      `WEEK ${wk.w.weekNumber}`, wk.w.periodeText, wk.w.monthName,
      wk.wDpSales, wk.wDpQty, wk.wDpTrx, null,
      wk.wCvSales, wk.wCvQty, wk.wCvNota, null,
      null, null, null, null, null,
      wk.wTrafPagi, wk.wTrafSiang, null, null, null, null,
    ]);
    const n = r.number;
    r.height = 21;
    r.getCell(7).value = { formula: `IF(L${n}>0,D${n}/L${n},0)` };
    r.getCell(11).value = { formula: `IF(L${n}>0,H${n}/L${n},0)` };
    r.getCell(12).value = { formula: `D${n}+H${n}`, result: wk.wTotalSales };
    r.getCell(13).value = { formula: `E${n}+I${n}`, result: wk.wTotalQty };
    r.getCell(14).value = { formula: `F${n}+J${n}`, result: wk.wTotalTrxNota };
    r.getCell(15).value = { formula: `IF(M${n}>0,ROUND(L${n}/M${n},0),0)`, result: ppVal };
    r.getCell(16).value = { formula: `ROUND(L${n}/${wk.daysLen},0)`, result: avgVal };
    r.getCell(19).value = { formula: `Q${n}+R${n}`, result: wk.wTotalTraf };
    r.getCell(20).value = { formula: `IF(S${n}>0,N${n}/S${n},0)` };
    r.getCell(21).value = { formula: `IF($L$${totWRowDNum}>0,L${n}/$L$${totWRowDNum},0)`, result: allWeeksGrandSales > 0 ? wk.wTotalSales / allWeeksGrandSales : 0 };
    r.getCell(22).value = idx === 0 ? 0 : { formula: `IF(L${n - 1}>0,(L${n}-L${n - 1})/L${n - 1},0)`, result: wowRatio };
    for (let c = 1; c <= 22; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: c === 1 || c === 12 || c === 14 || c >= 19,
        color: { argb: c === 22 ? (wowRatio > 0 ? EXCEL_COLORS.emeraldText : wowRatio < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark) : EXCEL_COLORS.navyDark },
      };
      if (c === 12) cell.fill = solidFill(EXCEL_COLORS.goldLight);
      else if (c >= 13 && c <= 16) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
      else if (c === 19 || c === 20) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
      else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

      if ([7, 11, 20, 21, 22].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 4 && c <= 19) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 2) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });
  const lastWRowD = Math.max(firstWRowD, firstWRowD + weekComputed.length - 1);
  const totWRowD = wsWeekly.addRow([`TOTAL WEEKLY (${weekComputed.length} MINGGU)`, '', '-', null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, '-']);
  totWRowD.height = 24;
  const wnD = totWRowD.number;
  wsWeekly.mergeCells(wnD, 1, wnD, 2);
  [4, 5, 6, 8, 9, 10, 12, 13, 14, 17, 18, 19].forEach((cIdx) => {
    const colL = colToLetter(cIdx);
    totWRowD.getCell(cIdx).value = weekComputed.length > 0 ? { formula: `SUM(${colL}${firstWRowD}:${colL}${lastWRowD})` } : 0;
  });
  totWRowD.getCell(7).value = { formula: `IF(L${wnD}>0,D${wnD}/L${wnD},0)` };
  totWRowD.getCell(11).value = { formula: `IF(L${wnD}>0,H${wnD}/L${wnD},0)` };
  totWRowD.getCell(15).value = { formula: `IF(M${wnD}>0,ROUND(L${wnD}/M${wnD},0),0)` };
  totWRowD.getCell(16).value = { formula: `ROUND(L${wnD}/${Math.max(1, weekComputed.length * 7)},0)` };
  totWRowD.getCell(20).value = { formula: `IF(S${wnD}>0,N${wnD}/S${wnD},0)` };
  totWRowD.getCell(21).value = weekComputed.length > 0 ? { formula: `SUM(U${firstWRowD}:U${lastWRowD})` } : 0;
  for (let c = 1; c <= 22; c++) {
    const cell = totWRowD.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = MEDIUM_BORDER;
    if ([7, 11, 20, 21].includes(c)) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c >= 4 && c <= 19) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  if (includeCharts && weeklySalesBars.length > 0) {
    embedExcelSalesDiagramPair(workbook, wsWeekly, {
      salesTitle: `DIAGRAM SALES WEEKLY (DP + CV) — ${monthName.toUpperCase()} ${year}`,
      salesSubtitle: 'Perbandingan Omzet Mingguan 1. Report DP vs 2. Report CV & Garis Tren Grand Total',
      growthTitle: `DIAGRAM GROWTH (WOW) & KONTRIBUSI WEEKLY (DP + CV)`,
      growthSubtitle: 'Diagram Pertumbuhan Mingguan (WoW %) & Kontribusi Tiap Week',
      bars: weeklySalesBars,
      config: DP_CV_CHART_CONFIG,
      rightFilterBadgeText: `${weeklySalesBars.length} WEEKLY`,
    });
  }

  // Detailed Sheet 3: Total 12 Bulan
  const wsAnnual = workbook.addWorksheet(getSafeSheetName(workbook, `${pfx}Total 12 Bulan (${year})`), {
    views: [{ showGridLines: true }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  [16.5, 10, 17, 10.5, 10.5, 11, 17, 10.5, 10.5, 11, 18.5, 11.5, 15, 14, 14.5, 11.5, 11.5, 13, 13.5, 14.5, 13, 13].forEach((w, i) => (wsAnnual.getColumn(i + 1).width = w));
  addUnifiedCorporateHeader(wsAnnual, 22, `TOTAL KONSOLIDASI 12 BULAN (${year})`, `Januari s/d Desember ${year} · Target Tahunan Rp 540.000.000`);
  const aTopD = wsAnnual.addRow(['Bulan', 'Hari', '1. REPORT DP', '', '', '', '2. REPORT CV', '', '', '', '3. GRAND TOTAL', '', '', '', '', '4. TRAFFIC', '', '', '', '', 'MoM Growth %', 'Kontribusi %']);
  aTopD.height = 24;
  const aSubD = wsAnnual.addRow(['', '', 'Sales DP', 'Qty DP', 'Trx DP', 'Share %', 'Sales CV', 'Qty CV', 'Nota CV', 'Share %', 'Grand Sales', 'Total Qty', 'Total Trx+Nota', 'Price Point', 'Rata2/Hari', 'Traf Pagi', 'Traf Siang', 'Total Traffic', 'Conv. Rate %', 'Sales / Org', '', '']);
  aSubD.height = 25;
  [1, 2, 21, 22].forEach((c) => wsAnnual.mergeCells(aTopD.number, c, aSubD.number, c));
  wsAnnual.mergeCells(aTopD.number, 3, aTopD.number, 6);
  wsAnnual.mergeCells(aTopD.number, 7, aTopD.number, 10);
  wsAnnual.mergeCells(aTopD.number, 11, aTopD.number, 15);
  wsAnnual.mergeCells(aTopD.number, 16, aTopD.number, 20);
  applyDualHeaderGroupColors(aTopD, aSubD, 22, (c) => {
    if (c >= 3 && c <= 6) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
    if (c >= 7 && c <= 10) return { topBg: EXCEL_COLORS.purpleHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.purpleSubBg, subFg: EXCEL_COLORS.purpleText };
    if (c >= 11 && c <= 15) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
    if (c >= 16 && c <= 20) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
    return { topBg: EXCEL_COLORS.navyDark, topFg: EXCEL_COLORS.yellowText, subBg: EXCEL_COLORS.navyDark, subFg: EXCEL_COLORS.yellowText };
  });
  applySheetFreeze(wsAnnual, 2, aSubD.number);
  const firstARowD = aSubD.number + 1;
  const totARowDNum = firstARowD + 12;
  monthlyStats.forEach((m, idx) => {
    const prevS = idx > 0 ? monthlyStats[idx - 1].mTotalSales : 0;
    const momRatio = idx > 0 && prevS > 0 ? (m.mTotalSales - prevS) / prevS : 0;
    const isSel = m.mKey === selectedMonthKey;
    const r = wsAnnual.addRow([
      `${m.mKey}. ${m.mName}`, m.dCount,
      m.mDpSales, m.mDpQty, m.mDpTrx, null,
      m.mCvSales, m.mCvQty, m.mCvNota, null,
      null, null, null, null, null,
      m.mTrafPagi, m.mTrafSiang, null, null, null,
      null, null,
    ]);
    const n = r.number;
    r.height = 21;
    r.getCell(6).value = { formula: `IF(K${n}>0,C${n}/K${n},0)` };
    r.getCell(10).value = { formula: `IF(K${n}>0,G${n}/K${n},0)` };
    r.getCell(11).value = { formula: `C${n}+G${n}`, result: m.mTotalSales };
    r.getCell(12).value = { formula: `D${n}+H${n}`, result: m.mDpQty + m.mCvQty };
    r.getCell(13).value = { formula: `E${n}+I${n}`, result: m.mDpTrx + m.mCvNota };
    r.getCell(14).value = { formula: `IF(L${n}>0,ROUND(K${n}/L${n},0),0)` };
    r.getCell(15).value = { formula: `IF(B${n}>0,ROUND(K${n}/B${n},0),0)` };
    r.getCell(18).value = { formula: `P${n}+Q${n}`, result: m.mTrafTotal };
    r.getCell(19).value = { formula: `IF(R${n}>0,M${n}/R${n},0)` };
    r.getCell(20).value = { formula: `IF(R${n}>0,ROUND(K${n}/R${n},0),0)` };
    r.getCell(21).value = idx === 0 ? 0 : { formula: `IF(K${n - 1}>0,(K${n}-K${n - 1})/K${n - 1},0)`, result: momRatio };
    r.getCell(22).value = { formula: `IF($K$${totARowDNum}>0,K${n}/$K$${totARowDNum},0)`, result: grandYearSales > 0 ? m.mTotalSales / grandYearSales : 0 };
    for (let c = 1; c <= 22; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: isSel || c === 1 || c === 11 || c === 13 || c >= 18,
        color: {
          argb:
            c === 21
              ? momRatio > 0 ? EXCEL_COLORS.emeraldText : momRatio < 0 ? EXCEL_COLORS.roseText : EXCEL_COLORS.navyDark
              : isSel && c <= 2 ? EXCEL_COLORS.indigoHeader : EXCEL_COLORS.navyDark,
        },
      };
      if (isSel && c <= 2) cell.fill = solidFill(EXCEL_COLORS.indigoSubBg);
      else if (c === 11) cell.fill = solidFill(EXCEL_COLORS.goldLight);
      else if (c >= 12 && c <= 15) cell.fill = solidFill(EXCEL_COLORS.goldSoft);
      else if (c === 18 || c === 19) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
      else if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);

      if ([6, 10, 19, 21, 22].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c >= 2 && c <= 20) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      }
    }
  });
  const lastARowD = firstARowD + 11;
  const totARowD = wsAnnual.addRow([`GRAND TOTAL 1 TAHUN (${year})`, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, '-', null]);
  totARowD.height = 25;
  const anTotD = totARowD.number;
  [2, 3, 4, 5, 7, 8, 9, 11, 12, 13, 16, 17, 18].forEach((cIdx) => {
    const colL = colToLetter(cIdx);
    totARowD.getCell(cIdx).value = { formula: `SUM(${colL}${firstARowD}:${colL}${lastARowD})` };
  });
  totARowD.getCell(6).value = { formula: `IF(K${anTotD}>0,C${anTotD}/K${anTotD},0)` };
  totARowD.getCell(10).value = { formula: `IF(K${anTotD}>0,G${anTotD}/K${anTotD},0)` };
  totARowD.getCell(14).value = { formula: `IF(L${anTotD}>0,ROUND(K${anTotD}/L${anTotD},0),0)` };
  totARowD.getCell(15).value = { formula: `IF(B${anTotD}>0,ROUND(K${anTotD}/B${anTotD},0),0)` };
  totARowD.getCell(19).value = { formula: `IF(R${anTotD}>0,M${anTotD}/R${anTotD},0)` };
  totARowD.getCell(20).value = { formula: `IF(R${anTotD}>0,ROUND(K${anTotD}/R${anTotD},0),0)` };
  totARowD.getCell(22).value = { formula: `SUM(V${firstARowD}:V${lastARowD})` };
  for (let c = 1; c <= 22; c++) {
    const cell = totARowD.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = MEDIUM_BORDER;
    if ([6, 10, 19, 22].includes(c)) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c >= 2 && c <= 20) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  const targetRowD = wsAnnual.addRow(['TARGET ANNUAL (RP 540 JT)', 365, null, null, null, null, null, null, null, null, 540000000, 'Pencapaian % ->', null, null, null, null, null, null, null, null, '-', null]);
  targetRowD.height = 22;
  const tRnD = targetRowD.number;
  wsAnnual.mergeCells(tRnD, 12, tRnD, 14);
  targetRowD.getCell(15).value = { formula: `ROUND(K${tRnD}/365,0)` };
  targetRowD.getCell(22).value = { formula: `IF(K${tRnD}>0,K${anTotD}/K${tRnD},0)` };
  for (let c = 1; c <= 22; c++) {
    const cell = targetRowD.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.navyDark);
    cell.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: c === 11 || c === 22 ? EXCEL_COLORS.yellowText : EXCEL_COLORS.white } };
    cell.border = DARK_THIN_BORDER;
    if (c === 22) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if ([2, 11, 15].includes(c)) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  if (includeCharts && monthly12Bars.length > 0) {
    embedExcelSalesDiagramPair(workbook, wsAnnual, {
      salesTitle: `DIAGRAM SALES TAHUNAN (DP + CV) — 12 BULAN (${year})`,
      salesSubtitle: 'Grafik Batang Tahunan 1. Report DP vs 2. Report CV & Garis Tren (Jan–Des)',
      growthTitle: `DIAGRAM GROWTH (MOM) & KONTRIBUSI 12 BULAN (DP + CV) (${year})`,
      growthSubtitle: 'Diagram Pertumbuhan Bulanan (MoM %) & Kontribusi per Bulan',
      bars: monthly12Bars,
      config: DP_CV_CHART_CONFIG,
      rightFilterBadgeText: `12 BULAN (${year})`,
    });
  }

  // Detailed Sheet 4: Traffic & Conversion Rate
  const wsTraffic = workbook.addWorksheet(getSafeSheetName(workbook, `${pfx}Traffic & Conv Rate`), {
    views: [{ showGridLines: true }],
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  [13.5, 11.5, 9.5, 15, 12.5, 12, 15, 12.5, 12, 13.5, 11, 11, 15, 14, 16.5, 18.5, 15].forEach((w, i) => (wsTraffic.getColumn(i + 1).width = w));
  addUnifiedCorporateHeader(wsTraffic, 17, 'TRAFFIC KASIR (PAGI & SIANG) & TOTAL CONVERSION RATE', `${monthName.toUpperCase()} ${year}`);
  const tTop = wsTraffic.addRow(['Tanggal', 'Hari', 'Week', 'SHIFT PAGI', '', '', 'SHIFT SIANG', '', '', 'TOTAL TRAFFIC & CONVERSION RATE', '', '', '', '', '', 'KONVERSI BELANJA', '']);
  tTop.height = 24;
  const tSub = wsTraffic.addRow(['', '', '', 'Kasir Pagi', 'Traf Pagi', 'Share Traf %', 'Kasir Siang', 'Traf Siang', 'Share Traf %', 'Total Traffic', 'Trx DP', 'Nota CV', 'Total Trx+Nota', 'Conv. Rate %', 'Kategori', 'Grand Sales (Rp)', 'Sales / Org']);
  tSub.height = 25;
  [1, 2, 3].forEach((c) => wsTraffic.mergeCells(tTop.number, c, tSub.number, c));
  wsTraffic.mergeCells(tTop.number, 4, tTop.number, 6);
  wsTraffic.mergeCells(tTop.number, 7, tTop.number, 9);
  wsTraffic.mergeCells(tTop.number, 10, tTop.number, 15);
  wsTraffic.mergeCells(tTop.number, 16, tTop.number, 17);
  applyDualHeaderGroupColors(tTop, tSub, 17, (c) => {
    if (c >= 4 && c <= 6) return { topBg: EXCEL_COLORS.amberHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.amberSubBg, subFg: EXCEL_COLORS.amberText };
    if (c >= 7 && c <= 9) return { topBg: EXCEL_COLORS.indigoHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.indigoSubBg, subFg: EXCEL_COLORS.indigoText };
    if (c >= 10 && c <= 15) return { topBg: EXCEL_COLORS.emeraldHeader, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.emeraldSubBg, subFg: EXCEL_COLORS.emeraldText };
    if (c >= 16) return { topBg: EXCEL_COLORS.gold, topFg: EXCEL_COLORS.black, subBg: EXCEL_COLORS.goldLight, subFg: EXCEL_COLORS.navyDark };
    return { topBg: EXCEL_COLORS.crimsonDark, topFg: EXCEL_COLORS.white, subBg: EXCEL_COLORS.crimsonDark, subFg: EXCEL_COLORS.white };
  });
  applySheetFreeze(wsTraffic, 3, tSub.number);
  const firstTRow = tSub.number + 1;
  dailyRows.forEach((d) => {
    const statusLabel = d.trafficTotal === 0 ? '-' : d.convRateRatio >= 0.45 ? 'Tinggi' : d.convRateRatio >= 0.25 ? 'Normal' : 'Rendah';
    const r = wsTraffic.addRow([
      d.tanggal, d.dayName, d.weekNumber > 0 ? `W-${d.weekNumber}` : '-',
      d.pagiStaff, d.trafficPagi, null,
      d.siangStaff, d.trafficSiang, null,
      null, d.dpTrx, d.cvNota, null, null, null,
      d.totalSales, null,
    ]);
    const n = r.number;
    r.height = 20;
    r.getCell(6).value = { formula: `IF(J${n}>0,E${n}/J${n},0)` };
    r.getCell(9).value = { formula: `IF(J${n}>0,H${n}/J${n},0)` };
    r.getCell(10).value = { formula: `E${n}+H${n}`, result: d.trafficTotal };
    r.getCell(13).value = { formula: `K${n}+L${n}`, result: d.totalTrxNota };
    r.getCell(14).value = { formula: `IF(J${n}>0,M${n}/J${n},0)`, result: d.convRateRatio };
    r.getCell(15).value = { formula: `IF(J${n}=0,"-",IF(N${n}>=0.45,"Tinggi",IF(N${n}>=0.25,"Normal","Rendah")))`, result: statusLabel };
    r.getCell(17).value = { formula: `IF(J${n}>0,ROUND(P${n}/J${n},0),0)`, result: d.salesPerVisitor };
    const isEmpty = d.totalSales === 0 && d.trafficTotal === 0;
    for (let c = 1; c <= 17; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: c <= 3 || c === 10 || c === 13 || c === 14 || c === 16,
        color: {
          argb:
            d.isWeekend && c <= 2
              ? EXCEL_COLORS.roseText
              : isEmpty && c >= 5
                ? EXCEL_COLORS.slateMuted
                : c === 14 && d.convRateRatio >= 0.45
                  ? EXCEL_COLORS.emeraldText
                  : EXCEL_COLORS.navyDark,
        },
      };
      if (c === 10 || c === 14) cell.fill = solidFill(EXCEL_COLORS.emeraldSoft);
      else if (c === 16) cell.fill = solidFill(EXCEL_COLORS.goldLight);
      else if (d.isWeekend) cell.fill = solidFill(EXCEL_COLORS.crimsonBg);
      else if (d.dayNumber % 2 === 0) cell.fill = solidFill(EXCEL_COLORS.slate50);

      if ([6, 9, 14].includes(c)) {
        cell.numFmt = NUM_FMT_PCT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([5, 8, 10, 11, 12, 13, 16, 17].includes(c)) {
        cell.numFmt = NUM_FMT_INT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });
  const lastTRow = Math.max(firstTRow, firstTRow + dailyRows.length - 1);
  const totTRow = wsTraffic.addRow([`TOTAL ${monthName.toUpperCase()}`, '', '', '-', null, null, '-', null, null, null, null, null, null, null, '-', null, null]);
  totTRow.height = 24;
  const trn = totTRow.number;
  wsTraffic.mergeCells(trn, 1, trn, 3);
  [5, 8, 10, 11, 12, 13, 16].forEach((cIdx) => {
    const colL = colToLetter(cIdx);
    totTRow.getCell(cIdx).value = { formula: `SUM(${colL}${firstTRow}:${colL}${lastTRow})` };
  });
  totTRow.getCell(6).value = { formula: `IF(J${trn}>0,E${trn}/J${trn},0)` };
  totTRow.getCell(9).value = { formula: `IF(J${trn}>0,H${trn}/J${trn},0)` };
  totTRow.getCell(14).value = { formula: `IF(J${trn}>0,M${trn}/J${trn},0)` };
  totTRow.getCell(17).value = { formula: `IF(J${trn}>0,ROUND(P${trn}/J${trn},0),0)` };
  for (let c = 1; c <= 17; c++) {
    const cell = totTRow.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.font = { name: 'Calibri', size: 10, bold: true };
    cell.border = MEDIUM_BORDER;
    if ([6, 9, 14].includes(c)) {
      cell.numFmt = NUM_FMT_PCT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if ([5, 8, 10, 11, 12, 13, 16, 17].includes(c)) {
      cell.numFmt = NUM_FMT_INT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }
  if (includeCharts && dailyTrafficBars.length > 0) {
    embedExcelTrafficConversionDiagram(workbook, wsTraffic, {
      title: `DIAGRAM TRAFFIC KASIR (PAGI & SIANG) & CONVERSION RATE — ${monthName.toUpperCase()} ${year}`,
      subtitle: 'Jumlah Pengunjung Shift Pagi + Siang vs Persentase Conversion Rate (Total Trx+Nota / Traffic)',
      items: dailyTrafficBars,
      rightFilterBadgeText: `TOTAL TRAFFIC: ${mTotTraffic.toLocaleString('id-ID')} ORG`,
    });
  }
}

function NUM_IDR_FMT_SAFE(fmt: string): string {
  return fmt || '#,##0';
}

export async function exportTotalDpCvToExcel(
  weeklyTables: WeeklySalesTable[],
  cvRecords: CvDailyRecord[],
  selectedMonthKey: string = '11',
  year: number = 2026,
  weeklyScope: 'month' | 'all' = 'month',
  compact: boolean = true,
  includeCharts: boolean = true
) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'REPORT BAZAR MATAHARI — Tanjung (PT MDS Retailing Tbk)';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  appendTotalDpCvSheetsToWorkbook(workbook, weeklyTables, cvRecords, selectedMonthKey, year, {
    weeklyScope,
    includePrefix: false,
    compact,
    includeCharts,
  });

  await triggerWorkbookDownload(
    workbook,
    `Total_Report_DP_CV_Matahari_${selectedMonthKey}_${year}.xlsx`
  );
}
