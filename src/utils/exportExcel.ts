import ExcelJS from 'exceljs';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  calculateDayMetrics,
  calculateWeekMetrics,
  getWeeklyPeriodeText,
  getIndoDayName,
  getWeeklyDayCycleText,
  isWeekendDate,
  getTableMonthKey,
  getTableMonthName,
  getTableWeekendInfo,
  createWeeklyTableFromAnnualSchedule,
  formatReviewBulletItem,
  getReviewDocWeekNumber,
  getReviewDocMonthKey,
  getReviewDocWeeklyLabel,
  createReviewDocForWeek,
  ANNUAL_2026_WEEKS,
  buildAll12MonthsCalendarSummary,
} from '../data/initialData';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  loadCvDiv1Records,
  loadCvInventoryOps,
  MONTH_NAMES_ID,
} from '../data/cvDiv1Data';
import { appendCvDiv1SheetsToWorkbook } from './exportCvDiv1Excel';
import {
  appendTotalDpCvSheetsToWorkbook,
  buildConsolidatedMaps,
} from './exportTotalDpCvExcel';
import {
  EXCEL_COLORS,
  NUM_FMT_IDR,
  solidFill,
  addUnifiedCorporateHeader,
  addSectionTitleBar,
  addKpiStrip,
  applySheetFreeze,
} from './excelStyles';

export interface ExportExcelOptions {
  includeWeeklyMatrix: boolean;
  includeDailyTable: boolean;
  includeMonthlyTable: boolean;
  includeReasonEvaluation: boolean;
  includeStaffAndReview: boolean;
  includeChartDiagrams?: boolean;
  includeExecutiveCover?: boolean;
  includeCvReport?: boolean;
  includeDataCv?: boolean;
  includeTotalConsolidated?: boolean;
  compactMode?: boolean;
  cvRecords?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  activeCvWeekNumber?: number;
  sheetLayout?:
    | 'SEPARATE_SHEETS'
    | 'COMBINED_AND_SEPARATE'
    | 'SINGLE_COMBINED_SHEET';
  selectedWeekIds?: string[]; // If empty or undefined, uses all tables passed
  selectedMonthKey?: string; // 'ALL' or '01'..'12'
  selectedWeekendFilter?: string; // 'ALL' | 'WEEKEND_1'..'WEEKEND_5' | 'w-1'..'w-53' | table.id
  dayCategoryFilter?: 'ALL_DAYS' | 'WEEKEND_ONLY' | 'WEEKDAY_ONLY';
  allReviewDocs?: ReviewMonitoringDoc[];
  selectedReviewWeekFilter?: string; // 'FOLLOW_WEEKLY' | 'ALL' | 'w-1'..'w-53' | doc.id
  customFileName?: string;
}

export const DEFAULT_EXPORT_EXCEL_OPTIONS: ExportExcelOptions = {
  includeWeeklyMatrix: true,
  includeDailyTable: true,
  includeMonthlyTable: true,
  includeReasonEvaluation: true,
  includeStaffAndReview: true,
  includeChartDiagrams: true,
  includeExecutiveCover: false,
  includeCvReport: false,
  includeDataCv: false,
  includeTotalConsolidated: false,
  compactMode: false,
  sheetLayout: 'SEPARATE_SHEETS',
  selectedWeekIds: [],
  selectedMonthKey: 'ALL',
  selectedWeekendFilter: 'ALL',
  dayCategoryFilter: 'ALL_DAYS',
  selectedReviewWeekFilter: 'FOLLOW_WEEKLY',
};

const INDO_MONTHS: Record<string, string> = {
  '01': 'Januari',
  '02': 'Februari',
  '03': 'Maret',
  '04': 'April',
  '05': 'Mei',
  '06': 'Juni',
  '07': 'Juli',
  '08': 'Agustus',
  '09': 'September',
  '10': 'Oktober',
  '11': 'November',
  '12': 'Desember',
};

function colToLetter(col: number): string {
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

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF1E293B' } },
  left: { style: 'thin', color: { argb: 'FF1E293B' } },
  bottom: { style: 'thin', color: { argb: 'FF1E293B' } },
  right: { style: 'thin', color: { argb: 'FF1E293B' } },
};

const MEDIUM_OUTER_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'medium', color: { argb: 'FF0F172A' } },
  left: { style: 'medium', color: { argb: 'FF0F172A' } },
  bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
  right: { style: 'medium', color: { argb: 'FF0F172A' } },
};

const NUM_FMT_INTEGER = '#,##0';
const NUM_FMT_PERCENT = '0.0%';

function addOfficialReportHeader(ws: ExcelJS.Worksheet, colSpan: number) {
  const rTitle = ws.addRow(['REPORT BAZAR MATAHARI']);
  rTitle.height = 28;
  ws.mergeCells(rTitle.number, 1, rTitle.number, colSpan);
  for (let c = 1; c <= colSpan; c++) {
    const cell = rTitle.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFC00000' },
    };
    cell.font = {
      name: 'Calibri',
      size: 15,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = MEDIUM_OUTER_BORDER;
  }

  const rSub = ws.addRow(['Tanjung']);
  rSub.height = 22;
  ws.mergeCells(rSub.number, 1, rSub.number, colSpan);
  for (let c = 1; c <= colSpan; c++) {
    const cell = rSub.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFC000' },
    };
    cell.font = {
      name: 'Calibri',
      size: 12,
      bold: true,
      color: { argb: 'FF000000' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = THIN_BORDER;
  }

  const rCompany = ws.addRow(['PT MDS Retailing Tbk']);
  rCompany.height = 20;
  ws.mergeCells(rCompany.number, 1, rCompany.number, colSpan);
  for (let c = 1; c <= colSpan; c++) {
    const cell = rCompany.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFDE047' },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = MEDIUM_OUTER_BORDER;
  }

  ws.addRow([]);
}

function addSectionBanner(
  ws: ExcelJS.Worksheet,
  title: string,
  colSpan: number,
  bgArgb: string = 'FF0F172A',
  textArgb: string = 'FFFDE047'
) {
  const row = ws.addRow([title]);
  row.height = 28;
  ws.mergeCells(row.number, 1, row.number, colSpan);
  for (let c = 1; c <= colSpan; c++) {
    const cell = row.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: bgArgb },
    };
    cell.font = {
      name: 'Calibri',
      size: 12,
      bold: true,
      color: { argb: textArgb },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    cell.border = MEDIUM_OUTER_BORDER;
  }
  return row;
}

interface ChartBarItem {
  label: string;
  subLabel?: string;
  pagiSales: number;
  siangSales: number;
  totalSales: number;
  pagiQty?: number;
  siangQty?: number;
  totalQty?: number;
  pagiTrx?: number;
  siangTrx?: number;
  totalTrx?: number;
  pricePoint?: number;
  bigSize?: number;
  isWeekend?: boolean;
}

const MONTH_META_LIST = [
  { key: '01', name: 'Januari', short: 'Jan', quarter: 'Q1', daysCount: 31 },
  { key: '02', name: 'Februari', short: 'Feb', quarter: 'Q1', daysCount: 28 },
  { key: '03', name: 'Maret', short: 'Mar', quarter: 'Q1', daysCount: 31 },
  { key: '04', name: 'April', short: 'Apr', quarter: 'Q2', daysCount: 30 },
  { key: '05', name: 'Mei', short: 'Mei', quarter: 'Q2', daysCount: 31 },
  { key: '06', name: 'Juni', short: 'Jun', quarter: 'Q2', daysCount: 30 },
  { key: '07', name: 'Juli', short: 'Jul', quarter: 'Q3', daysCount: 31 },
  { key: '08', name: 'Agustus', short: 'Agu', quarter: 'Q3', daysCount: 31 },
  { key: '09', name: 'September', short: 'Sep', quarter: 'Q3', daysCount: 30 },
  { key: '10', name: 'Oktober', short: 'Okt', quarter: 'Q4', daysCount: 31 },
  { key: '11', name: 'November', short: 'Nov', quarter: 'Q4', daysCount: 30 },
  { key: '12', name: 'Desember', short: 'Des', quarter: 'Q4', daysCount: 31 },
] as const;

export function isWeeklyTableActive(t: WeeklySalesTable): boolean {
  if (!t || !Array.isArray(t.days)) return false;
  if (calculateWeekMetrics(t).weeklySales > 0) return true;
  return t.days.some(
    (d) =>
      (d.pagi?.sales ?? 0) > 0 ||
      (d.siang?.sales ?? 0) > 0 ||
      (d.pagi?.qty ?? 0) > 0 ||
      (d.siang?.qty ?? 0) > 0 ||
      (d.pagi?.trx ?? 0) > 0 ||
      (d.siang?.trx ?? 0) > 0 ||
      (d.pagi?.traffic ?? 0) > 0 ||
      (d.siang?.traffic ?? 0) > 0 ||
      (d.pagi?.nama && d.pagi.nama.trim() !== '') ||
      (d.siang?.nama && d.siang.nama.trim() !== '') ||
      (d.reasonPagi && d.reasonPagi.trim() !== '') ||
      (d.reasonSiang && d.reasonSiang.trim() !== '')
  );
}

export function hasReviewDocContent(doc?: ReviewMonitoringDoc | null): boolean {
  if (!doc || !Array.isArray(doc.sections)) return false;
  return doc.sections.some((sec) => {
    const hasDirect =
      Array.isArray(sec.directItems) &&
      sec.directItems.some((it) => typeof it === 'string' && it.trim() !== '');
    const hasSub =
      Array.isArray(sec.subRows) &&
      sec.subRows.some(
        (sr) =>
          Array.isArray(sr.items) &&
          sr.items.some((it) => typeof it === 'string' && it.trim() !== '')
      );
    return hasDirect || hasSub;
  });
}

function getDefaultCalendarMonthKey2026(): string {
  const now = new Date();
  if (now.getFullYear() === 2026) {
    return String(now.getMonth() + 1).padStart(2, '0');
  }
  return '11';
}

export function resolveLatestActiveMonthKey(
  explicitMonthKey?: string,
  allTables?: WeeklySalesTable[],
  cvRecords?: CvDailyRecord[],
  cvOps?: CvInventoryOpsEntry[]
): string {
  if (
    explicitMonthKey &&
    explicitMonthKey !== 'ALL' &&
    MONTH_META_LIST.some((m) => m.key === explicitMonthKey)
  ) {
    return explicitMonthKey;
  }

  let maxFilledMonth: string | null = null;
  const considerMonth = (mKey?: string) => {
    if (!mKey) return;
    const norm = String(mKey).padStart(2, '0');
    if (!MONTH_META_LIST.some((m) => m.key === norm)) return;
    if (!maxFilledMonth || norm > maxFilledMonth) {
      maxFilledMonth = norm;
    }
  };

  (allTables || []).forEach((t) => {
    (t.days || []).forEach((d) => {
      const hasVal =
        (d.pagi?.sales ?? 0) > 0 ||
        (d.siang?.sales ?? 0) > 0 ||
        (d.pagi?.qty ?? 0) > 0 ||
        (d.siang?.qty ?? 0) > 0 ||
        (d.pagi?.trx ?? 0) > 0 ||
        (d.siang?.trx ?? 0) > 0 ||
        (d.pagi?.traffic ?? 0) > 0 ||
        (d.siang?.traffic ?? 0) > 0 ||
        (d.pagi?.nama && d.pagi.nama.trim() !== '') ||
        (d.siang?.nama && d.siang.nama.trim() !== '') ||
        (d.reasonPagi && d.reasonPagi.trim() !== '') ||
        (d.reasonSiang && d.reasonSiang.trim() !== '');
      if (!hasVal) return;
      const parts = (d.tanggal || '').trim().split('/');
      if (parts.length === 3 && parts[2] === '2026') {
        considerMonth(parts[1]);
      } else if (parts.length !== 3) {
        considerMonth(getTableMonthKey(t));
      }
    });
  });

  (cvRecords || []).forEach((r) => {
    const hasVal =
      (Number(r.graphis?.pagi?.sales) || 0) > 0 ||
      (Number(r.graphis?.siang?.sales) || 0) > 0 ||
      (Number(r.simplicity?.pagi?.sales) || 0) > 0 ||
      (Number(r.simplicity?.siang?.sales) || 0) > 0 ||
      (Number(r.graphis?.pagi?.qty) || 0) > 0 ||
      (Number(r.graphis?.siang?.qty) || 0) > 0 ||
      (Number(r.simplicity?.pagi?.qty) || 0) > 0 ||
      (Number(r.simplicity?.siang?.qty) || 0) > 0 ||
      (Number(r.graphis?.pagi?.trx) || 0) > 0 ||
      (Number(r.graphis?.siang?.trx) || 0) > 0 ||
      (Number(r.simplicity?.pagi?.trx) || 0) > 0 ||
      (Number(r.simplicity?.siang?.trx) || 0) > 0 ||
      (r.reasonEvent && r.reasonEvent.trim() !== '') ||
      (r.petugasPagi && r.petugasPagi.trim() !== '') ||
      (r.petugasSiang && r.petugasSiang.trim() !== '');
    if (!hasVal) return;
    const parts = (r.tanggal || '').trim().split('/');
    if (parts.length === 3 && parts[2] === '2026') {
      considerMonth(parts[1]);
    } else if (r.monthKey) {
      considerMonth(r.monthKey);
    }
  });

  (cvOps || []).forEach((op) => {
    const hasVal =
      (Number(op.stockAwal) || 0) > 0 ||
      (Number(op.barangDatang) || 0) > 0 ||
      (Number(op.barangKeluar) || 0) > 0 ||
      (op.salesQtyOverride !== null &&
        op.salesQtyOverride !== undefined &&
        Number(op.salesQtyOverride) > 0) ||
      (op.promo && op.promo.trim() !== '') ||
      (op.reasonKendala && op.reasonKendala.trim() !== '') ||
      (op.actionPlan && op.actionPlan.trim() !== '');
    if (!hasVal) return;
    if (op.periodType === 'monthly' && op.monthKey) {
      considerMonth(op.monthKey);
    } else if (op.periodType === 'weekly' && op.weekNumber) {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === op.weekNumber);
      if (sched) considerMonth(sched.monthKey);
    }
  });

  if (maxFilledMonth) {
    return maxFilledMonth;
  }

  return getDefaultCalendarMonthKey2026();
}

export function resolveLatestActiveWeekNumber(
  explicitWeek?: number,
  tables?: WeeklySalesTable[],
  cvRecords?: CvDailyRecord[],
  cvOps?: CvInventoryOpsEntry[],
  activeMonthKey?: string
): number {
  if (explicitWeek && explicitWeek >= 1 && explicitWeek <= 53) {
    return explicitWeek;
  }

  let maxWk = 0;
  (tables || []).forEach((t) => {
    if (isWeeklyTableActive(t) && t.weekNumber > maxWk) {
      maxWk = t.weekNumber;
    }
  });

  (cvRecords || []).forEach((r) => {
    const hasVal =
      (Number(r.graphis?.pagi?.sales) || 0) > 0 ||
      (Number(r.graphis?.siang?.sales) || 0) > 0 ||
      (Number(r.simplicity?.pagi?.sales) || 0) > 0 ||
      (Number(r.simplicity?.siang?.sales) || 0) > 0;
    if (hasVal && r.weekNumber > maxWk) {
      maxWk = r.weekNumber;
    }
  });

  (cvOps || []).forEach((op) => {
    if (op.periodType !== 'weekly' || !op.weekNumber) return;
    const hasVal =
      (Number(op.stockAwal) || 0) > 0 ||
      (Number(op.barangDatang) || 0) > 0 ||
      (Number(op.barangKeluar) || 0) > 0 ||
      (op.promo && op.promo.trim() !== '') ||
      (op.reasonKendala && op.reasonKendala.trim() !== '') ||
      (op.actionPlan && op.actionPlan.trim() !== '');
    if (hasVal && op.weekNumber > maxWk) {
      maxWk = op.weekNumber;
    }
  });

  if (maxWk >= 1 && maxWk <= 53) {
    return maxWk;
  }

  const mKey = activeMonthKey || getDefaultCalendarMonthKey2026();
  const monthWeeks = ANNUAL_2026_WEEKS.filter((w) => w.monthKey === mKey);
  if (monthWeeks.length > 0) {
    return monthWeeks[0].weekNumber;
  }
  return tables && tables.length > 0 ? tables[tables.length - 1].weekNumber : 45;
}

function drawRoundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number
) {
  const radius = Math.max(0, Math.min(r, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + w - radius, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
  ctx.lineTo(x + w, y + h - radius);
  ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
  ctx.lineTo(x + radius, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function drawSalesBarTrendCard(
  ctx: CanvasRenderingContext2D,
  cardX: number,
  cardY: number,
  cardW: number,
  cardH: number,
  badgeText: string,
  badgeBg: string,
  titleText: string,
  summaryRightText: string,
  bars: ChartBarItem[],
  showContributionPct: boolean = false
) {
  // Inner White Card with rounded border (matching PhotoStyleSalesDiagram in app)
  ctx.fillStyle = '#FFFFFF';
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, 8);
  ctx.fill();
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.2;
  ctx.stroke();

  // Top-Left Orange Badge: [DIAGRAM SALES (OMZET RUPIAH)]
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  const bText = badgeText || 'DIAGRAM SALES (OMZET RUPIAH)';
  const badgeW = Math.max(178, ctx.measureText(bText).width + 20);
  ctx.fillStyle = badgeBg || '#D97706';
  drawRoundedRect(ctx, cardX + 14, cardY + 12, badgeW, 20, 4);
  ctx.fill();
  ctx.fillStyle = '#FFFFFF';
  ctx.textAlign = 'left';
  ctx.fillText(bText, cardX + 22, cardY + 25.5);

  // Subtitle below badge: "Grafik Batang Shift Pagi vs Siang & Garis Tren"
  ctx.textAlign = 'left';
  ctx.fillStyle = '#0F172A';
  ctx.font = 'bold 11px Inter, Calibri, sans-serif';
  ctx.fillText(titleText, cardX + 14, cardY + 47);

  // Top-Right Green Pill Badge: [Total: Rp 33.522.206]
  ctx.font = 'bold 10.5px Inter, Calibri, sans-serif';
  const sumW = ctx.measureText(summaryRightText).width + 22;
  const pillX = cardX + cardW - 14 - sumW;
  const pillY = cardY + 11;
  ctx.fillStyle = '#ECFDF5';
  drawRoundedRect(ctx, pillX, pillY, sumW, 23, 5);
  ctx.fill();
  ctx.strokeStyle = '#A7F3D0';
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.textAlign = 'center';
  ctx.fillStyle = '#047857';
  ctx.fillText(summaryRightText, pillX + sumW / 2, pillY + 15.5);

  // Plot area (Y-Axis + Dashed Gridlines + Stacked Vertical Bars + Green Trendline + Value Labels)
  const plotX = cardX + 52;
  const plotY = cardY + 72;
  const plotW = cardW - 68;
  const plotH = cardH - 112;

  const rawMax = Math.max(0, ...bars.map((b) => b.totalSales));
  const maxVal = rawMax > 0 ? rawMax : 2000000;
  const sumAllBars = bars.reduce((s, b) => s + b.totalSales, 0);

  // Dashed Y-Axis Gridlines (100%, 50%, 0%) matching app ("8.0jt", "4.0jt", "0")
  [1, 0.5, 0].forEach((ratio) => {
    const gy = plotY + plotH * (1 - ratio);
    const v = Math.round(maxVal * ratio);
    ctx.save();
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1;
    if (ratio > 0) {
      ctx.setLineDash([4, 4]);
    }
    ctx.beginPath();
    ctx.moveTo(plotX, gy);
    ctx.lineTo(plotX + plotW, gy);
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = '#64748B';
    ctx.font = 'bold 10px Inter, Calibri, sans-serif';
    ctx.textAlign = 'right';
    const labelTxt =
      v >= 1000000
        ? `${(v / 1000000).toFixed(1)}jt`
        : v > 0
          ? `${Math.round(v / 1000)}rb`
          : '0';
    ctx.fillText(labelTxt, plotX - 10, gy + 3.5);
  });

  const count = Math.max(1, bars.length);
  const isCompactMany = count > 14;
  const slotW = plotW / count;
  const barW = isCompactMany
    ? Math.min(22, Math.max(10, slotW * 0.58))
    : Math.min(38, Math.max(16, slotW * 0.42));
  const trendPoints: Array<{ x: number; y: number; bar: ChartBarItem }> = [];

  // 1) Draw Vertical Stacked Bars (Top = Indigo #4F46E5 Shift Siang, Bottom = Amber #F59E0B Shift Pagi)
  bars.forEach((b, idx) => {
    const cx = plotX + slotW * (idx + 0.5);
    const hRatio =
      maxVal > 0
        ? Math.max(b.totalSales > 0 ? 0.06 : 0.02, b.totalSales / maxVal)
        : 0.02;
    const totalH = Math.min(plotH, hRatio * plotH);
    const pagiShare = b.totalSales > 0 ? b.pagiSales / b.totalSales : 0;
    const siangShare = b.totalSales > 0 ? 1 - pagiShare : 0;
    const siangH = totalH * siangShare;
    const pagiH = totalH * pagiShare;

    const byTop = plotY + plotH - totalH;
    trendPoints.push({ x: cx, y: byTop, bar: b });

    ctx.save();
    drawRoundedRect(ctx, cx - barW / 2, byTop, barW, totalH, 2);
    ctx.clip();

    if (b.totalSales > 0) {
      // Top segment: Shift Siang (Indigo #4F46E5)
      if (siangH > 0) {
        ctx.fillStyle = '#4F46E5';
        ctx.fillRect(cx - barW / 2, byTop, barW, siangH + 0.5);
      }
      // Bottom segment: Shift Pagi (Amber #F59E0B)
      if (pagiH > 0) {
        ctx.fillStyle = '#F59E0B';
        ctx.fillRect(cx - barW / 2, byTop + siangH, barW, pagiH + 0.5);
      }
    } else {
      ctx.fillStyle = '#E2E8F0';
      ctx.fillRect(cx - barW / 2, byTop, barW, totalH);
    }
    ctx.restore();

    // X-Axis Labels below bar (e.g. "Sen (H1)" + "28/09")
    ctx.fillStyle = b.isWeekend ? '#BE123C' : '#334155';
    ctx.font = isCompactMany
      ? 'bold 8px Inter, Calibri, sans-serif'
      : 'bold 9.5px Inter, Calibri, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(b.label, cx, plotY + plotH + 15);
    if (b.subLabel) {
      ctx.fillStyle = '#64748B';
      ctx.font = isCompactMany
        ? '7.5px Inter, Calibri, sans-serif'
        : '9px Inter, Calibri, sans-serif';
      ctx.fillText(b.subLabel, cx, plotY + plotH + 26);
    }
  });

  // 2) Draw Green Trend Line (#10B981) + White-Centered Green Circles (#059669)
  if (trendPoints.length > 1) {
    ctx.strokeStyle = '#10B981';
    ctx.lineWidth = 2.2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    trendPoints.forEach((pt, i) => {
      if (i === 0) ctx.moveTo(pt.x, pt.y);
      else ctx.lineTo(pt.x, pt.y);
    });
    ctx.stroke();

    trendPoints.forEach((pt) => {
      ctx.fillStyle = '#FFFFFF';
      ctx.strokeStyle = '#059669';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(pt.x, pt.y, isCompactMany ? 3 : 3.6, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
    });
  }

  // 3) Draw Compact Value Text right above each bar top ("3.5jt", "2.8jt", etc. matching app)
  trendPoints.forEach((pt) => {
    const b = pt.bar;
    const compactNominal =
      b.totalSales >= 1000000
        ? `${(b.totalSales / 1000000).toFixed(1)}jt`
        : b.totalSales > 0
          ? `${Math.round(b.totalSales / 1000)}rb`
          : '0';
    const labelText =
      showContributionPct && sumAllBars > 0 && b.totalSales > 0 && !isCompactMany
        ? `${compactNominal} (${Math.round((b.totalSales / sumAllBars) * 100)}%)`
        : compactNominal;

    ctx.font = isCompactMany
      ? 'bold 8px Inter, Calibri, sans-serif'
      : 'bold 9.5px Inter, Calibri, sans-serif';
    ctx.fillStyle = '#0F172A';
    ctx.textAlign = 'center';
    ctx.fillText(labelText, pt.x, Math.max(cardY + 56, pt.y - 7));
  });
}

function generateSalesDiagramBase64(
  chartTitle: string,
  chartSubtitle: string,
  bars: ChartBarItem[],
  totalPagi: number,
  totalSiang: number,
  rightFilterBadgeText?: string
): string | null {
  if (typeof document === 'undefined') return null;
  const dpr = 2; // 2x Retina resolution for ultra-sharp text & bars in Excel/WPS
  const isWide = bars.length > 14;
  const logicalW = isWide ? Math.max(980, bars.length * 32 + 90) : 760;
  const logicalH = 314;
  const canvas = document.createElement('canvas');
  canvas.width = logicalW * dpr;
  canvas.height = logicalH * dpr;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.scale(dpr, dpr);

  const width = logicalW;
  const height = logicalH;

  const grandTotal = totalPagi + totalSiang;
  const pagiPct = grandTotal > 0 ? Math.round((totalPagi / grandTotal) * 100) : 0;
  const siangPct = grandTotal > 0 ? 100 - pagiPct : 0;

  // Outer Card Background (#F1F5F9 matching PhotoStyleSalesDiagram in app)
  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(0.75, 0.75, width - 1.5, height - 1.5);

  // Top Dark Navy Header Banner (#0F172A, 3 lines matching app)
  const headerH = 62;
  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, width, headerH);

  // Line 1 Left: Yellow Bold Title
  ctx.textAlign = 'left';
  ctx.fillStyle = '#FBBF24';
  ctx.font = 'bold 12px Inter, Calibri, sans-serif';
  ctx.fillText(chartTitle.toUpperCase(), 16, 20);

  // Line 1 Right: Optional Filter Weekly / Period Badge (matching top-right dropdown in app)
  if (rightFilterBadgeText) {
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const fbW = ctx.measureText(rightFilterBadgeText).width + 18;
    const fbX = width - 14 - fbW;
    const fbY = 7;
    ctx.fillStyle = '#1E293B';
    drawRoundedRect(ctx, fbX, fbY, fbW, 19, 4);
    ctx.fill();
    ctx.strokeStyle = '#FBBF24';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#FCD34D';
    ctx.textAlign = 'center';
    ctx.fillText(rightFilterBadgeText, fbX + fbW / 2, fbY + 13);
    ctx.textAlign = 'left';
  }

  // Line 2: Light Slate Subtitle
  ctx.fillStyle = '#CBD5E1';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillText(
    chartSubtitle ||
      'Diagram Batang Vertikal & Garis Tren Penjualan Shift Pagi vs Shift Siang',
    16,
    36
  );

  // Line 3: Legend inside Dark Navy Banner
  const pagText = `Shift Pagi: Rp ${totalPagi.toLocaleString('id-ID')} (${pagiPct}%)`;
  const siaText = `Shift Siang: Rp ${totalSiang.toLocaleString('id-ID')} (${siangPct}%)`;
  const trnText = 'Garis Tren Total Sales';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';

  // Shift Pagi Legend
  ctx.fillStyle = '#F59E0B';
  ctx.fillRect(16, 44, 9, 9);
  ctx.fillStyle = '#FDE68A';
  ctx.fillText(pagText, 29, 52);

  // Shift Siang Legend
  const siaX = 240;
  ctx.fillStyle = '#4F46E5';
  ctx.fillRect(siaX, 44, 9, 9);
  ctx.fillStyle = '#C7D2FE';
  ctx.fillText(siaText, siaX + 13, 52);

  // Garis Tren Legend
  const trnX = 485;
  ctx.strokeStyle = '#10B981';
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  ctx.moveTo(trnX, 48.5);
  ctx.lineTo(trnX + 15, 48.5);
  ctx.stroke();
  ctx.fillStyle = '#6EE7B7';
  ctx.fillText(trnText, trnX + 20, 52);

  drawSalesBarTrendCard(
    ctx,
    12,
    headerH + 10,
    width - 24,
    height - headerH - 20,
    'DIAGRAM SALES (OMZET RUPIAH)',
    '#D97706',
    'Grafik Batang Shift Pagi vs Siang & Garis Tren',
    `Total: Rp ${grandTotal.toLocaleString('id-ID')}`,
    bars,
    false
  );
  return canvas.toDataURL('image/png');
}

function generateGrowthAndContributionDiagramBase64(
  chartTitle: string,
  chartSubtitle: string,
  bars: ChartBarItem[],
  rightFilterBadgeText?: string
): string | null {
  if (typeof document === 'undefined' || bars.length === 0) return null;
  const dpr = 2;
  const isWide = bars.length > 14;
  const logicalW = isWide ? Math.max(980, bars.length * 34 + 90) : 760;
  // When >14 bars (e.g. 31 days in Daily sheet), stack the 2 charts vertically so each bar has plenty of width;
  // when <=14 bars (Weekly 7 days, Monthly weeks, or 12 Months), place Growth & Kontribusi side-by-side!
  const logicalH = isWide ? 540 : 326;
  const canvas = document.createElement('canvas');
  canvas.width = logicalW * dpr;
  canvas.height = logicalH * dpr;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.scale(dpr, dpr);

  const width = logicalW;
  const height = logicalH;

  const grandTotalSales = bars.reduce((s, b) => s + (b.totalSales || 0), 0);

  // Enrich bars with Growth % and Contribution %
  const enriched = bars.map((b, idx) => {
    const contribPct =
      grandTotalSales > 0 ? (b.totalSales / grandTotalSales) * 100 : 0;
    const prevSales = idx > 0 ? bars[idx - 1].totalSales : null;
    let growthPct = 0;
    let growthStatus: 'UP' | 'DOWN' | 'FLAT' | 'BASE' | 'NEW' = 'FLAT';

    if (prevSales === null) {
      growthPct = 0;
      growthStatus = 'BASE';
    } else if (prevSales <= 0 && b.totalSales <= 0) {
      growthPct = 0;
      growthStatus = 'FLAT';
    } else if (prevSales <= 0 && b.totalSales > 0) {
      growthPct = 100;
      growthStatus = 'NEW';
    } else {
      growthPct = ((b.totalSales - prevSales) / prevSales) * 100;
      if (Math.abs(growthPct) < 0.05) {
        growthPct = 0;
        growthStatus = 'FLAT';
      } else {
        growthStatus = growthPct > 0 ? 'UP' : 'DOWN';
      }
    }
    return {
      ...b,
      contribPct,
      growthPct,
      growthStatus,
    };
  });

  // Outer Card Background (#F1F5F9)
  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(0.75, 0.75, width - 1.5, height - 1.5);

  // Top Dark Navy Header Banner (#0F172A)
  const headerH = 62;
  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, width, headerH);

  // Line 1 Left: Emerald Pill + Yellow Bold Title
  ctx.fillStyle = '#10B981';
  drawRoundedRect(ctx, 16, 7, 130, 17, 4);
  ctx.fill();
  ctx.fillStyle = '#022C22';
  ctx.font = 'bold 9px Inter, Calibri, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('GROWTH & KONTRIBUSI', 16 + 65, 18.5);

  ctx.textAlign = 'left';
  ctx.fillStyle = '#FBBF24';
  ctx.font = 'bold 11.5px Inter, Calibri, sans-serif';
  ctx.fillText(chartTitle.toUpperCase(), 154, 19.5);

  if (rightFilterBadgeText) {
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const fbW = ctx.measureText(rightFilterBadgeText).width + 18;
    const fbX = width - 14 - fbW;
    const fbY = 7;
    ctx.fillStyle = '#1E293B';
    drawRoundedRect(ctx, fbX, fbY, fbW, 19, 4);
    ctx.fill();
    ctx.strokeStyle = '#FBBF24';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#FCD34D';
    ctx.textAlign = 'center';
    ctx.fillText(rightFilterBadgeText, fbX + fbW / 2, fbY + 13);
    ctx.textAlign = 'left';
  }

  // Line 2: Subtitle
  ctx.fillStyle = '#CBD5E1';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillText(
    chartSubtitle ||
      'Visualisasi Diagram Batang Pertumbuhan (Growth %) & Porsi Kontribusi Sales (%)',
    16,
    36
  );

  // Line 3: Legend inside Dark Navy Banner
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillStyle = '#10B981';
  ctx.fillRect(16, 44, 9, 9);
  ctx.fillStyle = '#6EE7B7';
  ctx.fillText('Growth Naik (▲ +%)', 29, 52);

  ctx.fillStyle = '#F43F5E';
  ctx.fillRect(160, 44, 9, 9);
  ctx.fillStyle = '#FDA4AF';
  ctx.fillText('Growth Turun (▼ -%)', 173, 52);

  ctx.fillStyle = '#0284C7';
  ctx.fillRect(310, 44, 9, 9);
  ctx.fillStyle = '#7DD3FC';
  ctx.fillText('Kontribusi terhadap Total (%)', 323, 52);

  const topContrib = [...enriched].sort((a, b) => b.contribPct - a.contribPct)[0];
  if (topContrib && topContrib.totalSales > 0) {
    ctx.fillStyle = '#FDE68A';
    ctx.fillText(
      `Kontribusi Tertinggi: ${topContrib.label} (${topContrib.contribPct.toFixed(1)}%)`,
      500,
      52
    );
  }

  // Helper to draw a sub-card (either GROWTH or KONTRIBUSI)
  const drawSubChart = (
    cardX: number,
    cardY: number,
    cardW: number,
    cardH: number,
    type: 'GROWTH' | 'KONTRIBUSI'
  ) => {
    ctx.fillStyle = '#FFFFFF';
    drawRoundedRect(ctx, cardX, cardY, cardW, cardH, 8);
    ctx.fill();
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1.2;
    ctx.stroke();

    const isGrowth = type === 'GROWTH';
    const badgeTxt = isGrowth
      ? 'DIAGRAM BATANG GROWTH (%)'
      : 'DIAGRAM BATANG KONTRIBUSI (%)';
    const badgeColor = isGrowth ? '#059669' : '#0284C7';
    const subTitleTxt = isGrowth
      ? 'Grafik Batang Pertumbuhan Sales (Naik ▲ / Turun ▼)'
      : 'Grafik Batang Porsi Kontribusi terhadap Total Sales (100%)';

    ctx.font = 'bold 9px Inter, Calibri, sans-serif';
    const bW = Math.max(150, ctx.measureText(badgeTxt).width + 16);
    ctx.fillStyle = badgeColor;
    drawRoundedRect(ctx, cardX + 12, cardY + 10, bW, 18, 4);
    ctx.fill();
    ctx.fillStyle = '#FFFFFF';
    ctx.textAlign = 'left';
    ctx.fillText(badgeTxt, cardX + 20, cardY + 22);

    ctx.fillStyle = '#0F172A';
    ctx.font = 'bold 10px Inter, Calibri, sans-serif';
    ctx.fillText(subTitleTxt, cardX + 12, cardY + 42);

    // Top-Right Summary Pill
    const topGrowthItem = [...enriched]
      .filter((it) => it.growthStatus === 'UP' || it.growthStatus === 'NEW')
      .sort((a, b) => b.growthPct - a.growthPct)[0];
    const rightTxt = isGrowth
      ? topGrowthItem
        ? `Tertinggi: ${topGrowthItem.label} (+${topGrowthItem.growthPct.toFixed(1)}%)`
        : `Total: Rp ${grandTotalSales.toLocaleString('id-ID')}`
      : `Total 100% (Rp ${grandTotalSales.toLocaleString('id-ID')})`;

    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const rW = ctx.measureText(rightTxt).width + 16;
    const rX = cardX + cardW - 12 - rW;
    const rY = cardY + 10;
    ctx.fillStyle = isGrowth ? '#ECFDF5' : '#F0F9FF';
    drawRoundedRect(ctx, rX, rY, rW, 20, 4);
    ctx.fill();
    ctx.strokeStyle = isGrowth ? '#A7F3D0' : '#BAE6FD';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = isGrowth ? '#047857' : '#0369A1';
    ctx.textAlign = 'center';
    ctx.fillText(rightTxt, rX + rW / 2, rY + 13.5);

    const plotX = cardX + 44;
    const plotY = cardY + 62;
    const plotW = cardW - 56;
    const plotH = cardH - 96;

    const maxScale = isGrowth
      ? Math.max(
          15,
          ...enriched.map((it) => Math.min(200, Math.abs(it.growthPct)))
        )
      : Math.max(10, ...enriched.map((it) => it.contribPct));

    // Gridlines (100%, 50%, 0%)
    [1, 0.5, 0].forEach((ratio) => {
      const gy = plotY + plotH * (1 - ratio);
      const val = Math.round(maxScale * ratio);
      ctx.save();
      ctx.strokeStyle = '#CBD5E1';
      ctx.lineWidth = 1;
      if (ratio > 0) ctx.setLineDash([4, 4]);
      ctx.beginPath();
      ctx.moveTo(plotX, gy);
      ctx.lineTo(plotX + plotW, gy);
      ctx.stroke();
      ctx.restore();

      ctx.fillStyle = '#64748B';
      ctx.font = 'bold 9px Inter, Calibri, sans-serif';
      ctx.textAlign = 'right';
      const yLbl = isGrowth ? (val === 0 ? '0%' : `±${val}%`) : `${val}%`;
      ctx.fillText(yLbl, plotX - 6, gy + 3);
    });

    // Average line for Contribution chart
    if (!isGrowth && enriched.length > 0 && grandTotalSales > 0) {
      const avgPct = 100 / enriched.length;
      if (avgPct <= maxScale) {
        const avgY = plotY + plotH * (1 - avgPct / maxScale);
        ctx.save();
        ctx.strokeStyle = '#F59E0B';
        ctx.lineWidth = 1.2;
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.moveTo(plotX, avgY);
        ctx.lineTo(plotX + plotW, avgY);
        ctx.stroke();
        ctx.restore();
      }
    }

    const count = Math.max(1, enriched.length);
    const isCompactMany = count > 10;
    const slotW = plotW / count;
    const barW = isCompactMany
      ? Math.min(20, Math.max(9, slotW * 0.56))
      : Math.min(32, Math.max(14, slotW * 0.44));

    const pts: Array<{
      x: number;
      y: number;
      item: (typeof enriched)[number];
    }> = [];

    enriched.forEach((it, idx) => {
      const cx = plotX + slotW * (idx + 0.5);
      const val = isGrowth
        ? Math.min(maxScale, Math.abs(it.growthPct))
        : it.contribPct;
      const minRatio = isGrowth
        ? it.growthStatus === 'UP' ||
          it.growthStatus === 'DOWN' ||
          it.growthStatus === 'NEW'
          ? 0.08
          : 0.03
        : it.totalSales > 0
          ? 0.06
          : 0.02;
      const hRatio = maxScale > 0 ? Math.max(minRatio, val / maxScale) : 0.03;
      const barH = Math.min(plotH, hRatio * plotH);
      const byTop = plotY + plotH - barH;
      pts.push({ x: cx, y: byTop, item: it });

      ctx.save();
      drawRoundedRect(ctx, cx - barW / 2, byTop, barW, barH, 2.5);
      ctx.clip();

      if (isGrowth) {
        if (it.growthStatus === 'UP' || it.growthStatus === 'NEW') {
          ctx.fillStyle = '#10B981';
        } else if (it.growthStatus === 'DOWN') {
          ctx.fillStyle = '#F43F5E';
        } else if (it.growthStatus === 'BASE') {
          ctx.fillStyle = '#F59E0B';
        } else {
          ctx.fillStyle = '#CBD5E1';
        }
        ctx.fillRect(cx - barW / 2, byTop, barW, barH);
      } else {
        if (it.totalSales > 0) {
          const pShare = it.totalSales > 0 ? it.pagiSales / it.totalSales : 0;
          const sShare = 1 - pShare;
          ctx.fillStyle = '#0284C7';
          ctx.fillRect(cx - barW / 2, byTop, barW, barH * sShare + 0.5);
          ctx.fillStyle = '#38BDF8';
          ctx.fillRect(
            cx - barW / 2,
            byTop + barH * sShare,
            barW,
            barH * pShare + 0.5
          );
        } else {
          ctx.fillStyle = '#E2E8F0';
          ctx.fillRect(cx - barW / 2, byTop, barW, barH);
        }
      }
      ctx.restore();

      // X-Axis Labels
      ctx.fillStyle = it.isWeekend ? '#BE123C' : '#334155';
      ctx.font = isCompactMany
        ? 'bold 7.5px Inter, Calibri, sans-serif'
        : 'bold 9px Inter, Calibri, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(it.label, cx, plotY + plotH + 14);
      if (it.subLabel && !isCompactMany) {
        ctx.fillStyle = '#64748B';
        ctx.font = '8px Inter, Calibri, sans-serif';
        ctx.fillText(it.subLabel, cx, plotY + plotH + 24);
      }
    });

    // Trendline
    if (pts.length > 1) {
      ctx.save();
      ctx.strokeStyle = isGrowth ? '#0D9488' : '#0284C7';
      ctx.lineWidth = 2;
      if (isGrowth) ctx.setLineDash([4, 3]);
      ctx.beginPath();
      pts.forEach((p, i) => {
        if (i === 0) ctx.moveTo(p.x, p.y);
        else ctx.lineTo(p.x, p.y);
      });
      ctx.stroke();
      ctx.restore();

      pts.forEach((p) => {
        ctx.fillStyle = '#FFFFFF';
        ctx.strokeStyle = isGrowth
          ? p.item.growthStatus === 'DOWN'
            ? '#E11D48'
            : p.item.growthStatus === 'UP' || p.item.growthStatus === 'NEW'
              ? '#059669'
              : '#64748B'
          : '#0369A1';
        ctx.lineWidth = 1.8;
        ctx.beginPath();
        ctx.arc(p.x, p.y, isCompactMany ? 2.6 : 3.2, 0, Math.PI * 2);
        ctx.fill();
        ctx.stroke();
      });
    }

    // Top Value Labels above bars
    pts.forEach((p) => {
      const it = p.item;
      let valTxt = '';
      if (isGrowth) {
        if (it.growthStatus === 'BASE') valTxt = 'Basis';
        else if (it.growthStatus === 'UP' || it.growthStatus === 'NEW')
          valTxt = `+${it.growthPct.toFixed(1)}%`;
        else if (it.growthStatus === 'DOWN')
          valTxt = `${it.growthPct.toFixed(1)}%`;
        else valTxt = '0%';
      } else {
        valTxt = `${it.contribPct.toFixed(1)}%`;
      }

      ctx.font = isCompactMany
        ? 'bold 7.5px Inter, Calibri, sans-serif'
        : 'bold 8.5px Inter, Calibri, sans-serif';
      ctx.fillStyle = isGrowth
        ? it.growthStatus === 'DOWN'
          ? '#BE123C'
          : it.growthStatus === 'UP' || it.growthStatus === 'NEW'
            ? '#047857'
            : '#475569'
        : '#0369A1';
      ctx.textAlign = 'center';
      ctx.fillText(valTxt, p.x, Math.max(cardY + 50, p.y - 6));
    });
  };

  if (isWide) {
    const eachH = Math.floor((height - headerH - 28) / 2);
    drawSubChart(12, headerH + 8, width - 24, eachH, 'GROWTH');
    drawSubChart(12, headerH + 16 + eachH, width - 24, eachH, 'KONTRIBUSI');
  } else {
    const eachW = Math.floor((width - 32) / 2);
    const subH = height - headerH - 18;
    drawSubChart(12, headerH + 9, eachW, subH, 'GROWTH');
    drawSubChart(20 + eachW, headerH + 9, eachW, subH, 'KONTRIBUSI');
  }

  return canvas.toDataURL('image/png');
}

function renderGrowthAndContributionDiagramCard(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  blockTitle: string,
  blockSubtitle: string,
  items: ChartBarItem[],
  rightFilterBadgeText?: string
) {
  if (items.length === 0) return;

  const base64Img = generateGrowthAndContributionDiagramBase64(
    blockTitle,
    blockSubtitle,
    items,
    rightFilterBadgeText
  );

  if (base64Img) {
    const isWide = items.length > 14;
    const imgRowsCount = isWide ? 18 : 11;
    const startRowIdx = ws.lastRow ? ws.lastRow.number + 1 : 1;
    for (let r = 0; r < imgRowsCount; r++) {
      const row = ws.addRow([]);
      row.height = 25;
    }

    const imageId = workbook.addImage({
      base64: base64Img,
      extension: 'png',
    });
    ws.addImage(imageId, {
      tl: { col: 1.02, row: startRowIdx - 1 + 0.04 },
      ext: { width: isWide ? 820 : 620, height: isWide ? 430 : 260 },
      editAs: 'oneCell',
    });
  }
}

function renderPhotoStyleDiagramCard(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  blockTitle: string,
  blockSubtitle: string,
  items: ChartBarItem[],
  _colSpan: number = 10,
  rightFilterBadgeText?: string
) {
  if (items.length === 0) return;

  const totalPagi = items.reduce((s, it) => s + it.pagiSales, 0);
  const totalSiang = items.reduce((s, it) => s + it.siangSales, 0);

  const base64Img = generateSalesDiagramBase64(
    blockTitle,
    blockSubtitle,
    items,
    totalPagi,
    totalSiang,
    rightFilterBadgeText
  );

  if (base64Img) {
    const isWide = items.length > 14;
    const imgRowsCount = isWide ? 11 : 10;
    const startRowIdx = ws.lastRow ? ws.lastRow.number + 1 : 1;
    for (let r = 0; r < imgRowsCount; r++) {
      const row = ws.addRow([]);
      row.height = 25;
    }

    const imageId = workbook.addImage({
      base64: base64Img,
      extension: 'png',
    });
    ws.addImage(imageId, {
      tl: { col: 1.02, row: startRowIdx - 1 + 0.04 },
      ext: { width: isWide ? 820 : 620, height: isWide ? 265 : 250 },
      editAs: 'oneCell',
    });
  }
}

function embedSalesDiagramIntoSheet(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  chartTitle: string,
  chartSubtitle: string,
  bars: ChartBarItem[],
  _totalPagi: number,
  _totalSiang: number,
  _weekdaySales: number,
  _weekendSales: number,
  mode: 'SALES_ONLY' | 'MONTHLY_WITH_CONTRIBUTION' = 'SALES_ONLY',
  monthly12Bars?: ChartBarItem[],
  rightFilterBadgeText?: string
) {
  const topSpacerRow = ws.addRow([]);
  topSpacerRow.height = 8;

  if (mode === 'SALES_ONLY') {
    renderPhotoStyleDiagramCard(
      workbook,
      ws,
      chartTitle,
      chartSubtitle,
      bars,
      10,
      rightFilterBadgeText
    );

    const gcSpacer = ws.addRow([]);
    gcSpacer.height = 8;
    renderGrowthAndContributionDiagramCard(
      workbook,
      ws,
      chartTitle.replace(/^DIAGRAM SALES/, 'DIAGRAM GROWTH & KONTRIBUSI'),
      'Diagram Batang Pertumbuhan (Growth %) & Porsi Kontribusi terhadap Total Sales (100%)',
      bars,
      rightFilterBadgeText
    );
  } else {
    renderPhotoStyleDiagramCard(
      workbook,
      ws,
      chartTitle,
      chartSubtitle,
      bars,
      10,
      rightFilterBadgeText
    );

    const gcWeeklySpacer = ws.addRow([]);
    gcWeeklySpacer.height = 8;
    renderGrowthAndContributionDiagramCard(
      workbook,
      ws,
      chartTitle.replace(/^DIAGRAM SALES/, 'DIAGRAM GROWTH & KONTRIBUSI'),
      'Diagram Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Tiap Weekly terhadap Bulan (100%)',
      bars,
      rightFilterBadgeText
    );

    if (monthly12Bars && monthly12Bars.length > 0) {
      const midSpacer = ws.addRow([]);
      midSpacer.height = 10;
      renderPhotoStyleDiagramCard(
        workbook,
        ws,
        'DIAGRAM SALES TAHUNAN — 12 BULAN (JANUARI - DESEMBER 2026)',
        'Diagram Batang Vertikal 12 Bulan & Garis Tren Penjualan Shift Pagi vs Shift Siang',
        monthly12Bars,
        13,
        '12 BULAN (JAN – DES 2026)'
      );

      const gcMonthlySpacer = ws.addRow([]);
      gcMonthlySpacer.height = 8;
      renderGrowthAndContributionDiagramCard(
        workbook,
        ws,
        'DIAGRAM GROWTH & KONTRIBUSI TAHUNAN — 12 BULAN (JANUARI - DESEMBER 2026)',
        'Diagram Batang Pertumbuhan Bulanan (MoM Growth %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahunan (100%)',
        monthly12Bars,
        '12 BULAN (JAN – DES 2026)'
      );
    }
  }

  const bottomSpacerRow = ws.addRow([]);
  bottomSpacerRow.height = 8;
}

/* ============================================================================
   SAME-SHEET & HIDDEN DATABASE LOOKUP FOR "TOTAL WEEKLY 1 ▼" .. "TOTAL WEEKLY 53 ▼"
   ============================================================================ */
function buildWeeklyLookupRowValues(
  wk: number,
  allTables: WeeklySalesTable[],
  withFilterArrow: boolean
): (string | number)[] {
  const existing = allTables.find((t) => t.weekNumber === wk);
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
  const tbl =
    existing ||
    (sched ? createWeeklyTableFromAnnualSchedule(sched) : allTables[0]);

  const arrowSuffix = withFilterArrow ? ' ▼' : '';
  const keyLabel = `TOTAL WEEKLY ${wk}${arrowSuffix}`;
  const nextWk = wk < 53 ? wk + 1 : 53;
  const prevWk = wk > 1 ? wk - 1 : 1;
  const nextKeyLabel = `TOTAL WEEKLY ${nextWk} ▼`;
  const prevKeyLabel = `TOTAL WEEKLY ${prevWk} ▼`;
  const periodeLabel = getWeeklyPeriodeText(tbl);
  const dayCycleLabel = getWeeklyDayCycleText(tbl);

  // Col 1 (Z): KeyLabel
  // Col 2 (AA): NextSequentialKeyLabel (so lower tables automatically follow +1 sequence of table above!)
  // Col 3 (AB): PrevSequentialKeyLabel (so upper tables automatically follow -1 sequence of table below!)
  // Col 4 (AC): Periode Banner
  // Col 5 (AD): Right Merged Banner
  const rowVals: (string | number)[] = [
    keyLabel,
    nextKeyLabel,
    prevKeyLabel,
    `Periode : ${periodeLabel} (${dayCycleLabel})`,
    `TOTAL WEEKLY ${wk}\n(${dayCycleLabel})`,
  ];

  // Cols 6..75 (dbBase = 6 + i * 10 for i = 0..6)
  for (let i = 0; i < 7; i++) {
    const d = tbl.days[i];
    const tgl = d?.tanggal || '';
    const dName = tgl ? getIndoDayName(tgl) : `Hari ${i + 1}`;
    rowVals.push(
      tgl,
      `Hari ${i + 1} - ${dName}`,
      d?.pagi?.nama || '',
      d?.siang?.nama || '',
      d?.pagi?.sales ?? 0,
      d?.siang?.sales ?? 0,
      d?.pagi?.qty ?? 0,
      d?.siang?.qty ?? 0,
      d?.pagi?.trx ?? 0,
      d?.siang?.trx ?? 0
    );
  }

  // Cols 76..89 (ReasonPagi & ReasonSiang for i = 0..6)
  for (let i = 0; i < 7; i++) {
    const d = tbl.days[i];
    rowVals.push(d?.reasonPagi || '-', d?.reasonSiang || '-');
  }

  return rowVals;
}

/**
 * Populates hidden columns Z..DJ (Columns 26..114) on the SAME worksheet
 * so that WPS Office Android & Microsoft Excel natively display the dropdown
 * arrow ($Z$2:$Z$54) and evaluate VLOOKUP($A$6, $Z$2:$DJ$107, ...) instantly.
 */
function populateWorksheetWeeklyLookupTable(
  ws: ExcelJS.Worksheet,
  allTables: WeeklySalesTable[]
) {
  const startCol = 26; // Column Z
  const totalCols = 89; // Z (26) through DJ (114)

  for (let c = startCol; c < startCol + totalCols; c++) {
    const col = ws.getColumn(c);
    col.hidden = true;
    col.width = 18;
  }

  // Rows 2..54: "TOTAL WEEKLY 1 ▼" .. "TOTAL WEEKLY 53 ▼" (used by Data Validation $Z$2:$Z$54)
  for (let wk = 1; wk <= 53; wk++) {
    const rowNum = 1 + wk; // 2..54
    const row = ws.getRow(rowNum);
    const vals = buildWeeklyLookupRowValues(wk, allTables, true);
    for (let idx = 0; idx < vals.length; idx++) {
      row.getCell(startCol + idx).value = vals[idx];
    }
  }

  // Rows 55..107: "TOTAL WEEKLY 1" .. "TOTAL WEEKLY 53" (fallback without ▼ symbol)
  for (let wk = 1; wk <= 53; wk++) {
    const rowNum = 54 + wk; // 55..107
    const row = ws.getRow(rowNum);
    const vals = buildWeeklyLookupRowValues(wk, allTables, false);
    for (let idx = 0; idx < vals.length; idx++) {
      row.getCell(startCol + idx).value = vals[idx];
    }
  }
}

function buildHiddenWeeklyDatabaseSheet(
  workbook: ExcelJS.Workbook,
  allTables: WeeklySalesTable[]
): ExcelJS.Worksheet {
  const wsDb = workbook.addWorksheet('_Database_Weekly', {
    state: 'hidden',
  });

  const headers: string[] = [
    'Key_Total_Weekly',
    'Next_Sequential_Weekly',
    'Prev_Sequential_Weekly',
    'Periode_Banner',
    'Right_Header',
  ];
  for (let i = 1; i <= 7; i++) {
    headers.push(
      `D${i}_Tanggal`,
      `D${i}_Hari`,
      `D${i}_NamaPagi`,
      `D${i}_NamaSiang`,
      `D${i}_SalesPagi`,
      `D${i}_SalesSiang`,
      `D${i}_QtyPagi`,
      `D${i}_QtySiang`,
      `D${i}_TrxPagi`,
      `D${i}_TrxSiang`
    );
  }
  for (let i = 1; i <= 7; i++) {
    headers.push(`D${i}_ReasonPagi`, `D${i}_ReasonSiang`);
  }
  wsDb.addRow(headers);

  for (let wk = 1; wk <= 53; wk++) {
    wsDb.addRow(buildWeeklyLookupRowValues(wk, allTables, true));
  }
  for (let wk = 1; wk <= 53; wk++) {
    wsDb.addRow(buildWeeklyLookupRowValues(wk, allTables, false));
  }

  return wsDb;
}

/* ============================================================================
   SECTION RENDERER 1: SALES TABLE WEEKLY (MATRIKS MINGGUAN WPS TERFORMULA)
   ============================================================================ */
function renderWeeklyMatrixSection(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  tables: WeeklySalesTable[],
  showDivZero: boolean,
  includeReasonRows: boolean = true,
  includeChartDiagrams: boolean = true
): number {
  addSectionBanner(
    ws,
    'BAGIAN 1: SALES TABLE WEEKLY (KLIK TANDA FILTER ▼ PADA TOTAL WEEKLY UNTUK PILIH TOTAL WEEKLY 1 S/D 53)',
    16,
    'FF0F172A',
    'FFFDE047'
  );
  const topBannerSpacer = ws.addRow([]);
  topBannerSpacer.height = 10;

  let prevTableTitleRowNum: number | null = null;
  const renderedTitleRows: { rowNum: number; activeKey: string }[] = [];

  for (const table of tables) {
    const weekMetrics = calculateWeekMetrics(table);
    const dayMetricsList = table.days.map((d) => calculateDayMetrics(d));
    const periodeLabel = getWeeklyPeriodeText(table);
    const dayCycleLabel = getWeeklyDayCycleText(table);
    const numDays = table.days.length;
    const totalColIdx = 1 + numDays * 2 + 1; // 16 (Col P when 7 days)
    const totalColLetter = colToLetter(totalColIdx);
    const lastShiftColLetter = colToLetter(1 + numDays * 2); // O
    const activeWeeklyKey = `TOTAL WEEKLY ${table.weekNumber} ▼`;
    const activeRightBanner = `TOTAL WEEKLY ${table.weekNumber}\n(${dayCycleLabel})`;

    // Title Row: Col 1 is "TOTAL WEEKLY 40 ▼" with Dropdown ($Z$2:$Z$54), Col 2..16 is dynamic Periode
    const titleRowVals = new Array(totalColIdx).fill('');
    titleRowVals[0] = activeWeeklyKey;
    titleRowVals[1] = `Periode : ${periodeLabel} (${dayCycleLabel})`;
    const titleRow = ws.addRow(titleRowVals);
    titleRow.height = 26;
    const tRowNum = titleRow.number;

    if (totalColIdx >= 2) {
      ws.mergeCells(tRowNum, 2, tRowNum, totalColIdx);
    }

    // Col 1: Interactive Filter Dropdown Cell "TOTAL WEEKLY 40 ▼"
    // - Lower tables link to the table above (+1 via Col 2 in $Z$2:$DJ$107)
    // - After loop, the top table is also linked to the table below (-1 via Col 3 in $Z$2:$DJ$107)
    //   so changing EITHER the top OR lower TOTAL WEEKLY updates the other in sequence!
    const titleCellA = titleRow.getCell(1);
    if (prevTableTitleRowNum === null) {
      titleCellA.value = activeWeeklyKey;
    } else {
      titleCellA.value = {
        formula: `IFERROR(VLOOKUP($A$${prevTableTitleRowNum},$Z$2:$DJ$107,2,FALSE),"${activeWeeklyKey}")`,
        result: activeWeeklyKey,
      };
    }
    prevTableTitleRowNum = tRowNum;
    renderedTitleRows.push({ rowNum: tRowNum, activeKey: activeWeeklyKey });

    titleCellA.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' },
    };
    titleCellA.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFDE047' },
    };
    titleCellA.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    titleCellA.border = {
      top: { style: 'medium', color: { argb: 'FFFFC000' } },
      left: { style: 'medium', color: { argb: 'FFFFC000' } },
      bottom: { style: 'medium', color: { argb: 'FFFFC000' } },
      right: { style: 'medium', color: { argb: 'FFFFC000' } },
    };
    titleCellA.dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['$Z$2:$Z$54'],
      showErrorMessage: false,
      promptTitle: 'Filter Total Weekly (1 s/d 53)',
      prompt: 'Pilih TOTAL WEEKLY 1 s/d 53. Tabel di atas maupun di bawah otomatis mengikuti urutan.',
    };

    // Col 2..16: Dynamic Periode lookup based on Col 1 ($A$<tRowNum>) -> Col 4 in $Z$2:$DJ$107
    titleRow.getCell(2).value = {
      formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,4,FALSE),"Periode : ${periodeLabel} (${dayCycleLabel})")`,
      result: `Periode : ${periodeLabel} (${dayCycleLabel})`,
    };
    for (let c = 2; c <= totalColIdx; c++) {
      const cell = titleRow.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E293B' },
      };
      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      cell.border = THIN_BORDER;
    }

    const rowHari = ws.addRow(['Hari (Weekly)']);
    rowHari.height = 24;
    const rowTanggal = ws.addRow(['Tanggal']);
    rowTanggal.height = 21;
    const rowShift = ws.addRow(['Shift']);
    rowShift.height = 20;
    const rowNama = ws.addRow(['Nama']);
    rowNama.height = 21;
    const rowSales = ws.addRow(['Sales']);
    rowSales.height = 21;
    const rowQty = ws.addRow(['Qty']);
    rowQty.height = 21;
    const rowTrx = ws.addRow(['Trx']);
    rowTrx.height = 21;

    [rowHari, rowTanggal, rowShift, rowNama, rowSales, rowQty, rowTrx].forEach(
      (r) => {
        const c1 = r.getCell(1);
        c1.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFC00000' },
        };
        c1.font = {
          name: 'Calibri',
          size: 10.5,
          bold: true,
          color: { argb: 'FFFFFFFF' },
        };
        c1.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        c1.border = THIN_BORDER;
      }
    );

    ws.mergeCells(rowHari.number, totalColIdx, rowTrx.number, totalColIdx);
    for (let rNum = rowHari.number; rNum <= rowTrx.number; rNum++) {
      const cell = ws.getRow(rNum).getCell(totalColIdx);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FF000000' },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true,
      };
      cell.border = THIN_BORDER;
    }
    rowHari.getCell(totalColIdx).value = {
      formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,5,FALSE),"${activeRightBanner}")`,
      result: activeRightBanner,
    };

    for (let i = 0; i < numDays; i++) {
      const day = table.days[i];
      const colPagi = 2 + i * 2;
      const colSiang = colPagi + 1;
      const dName = getIndoDayName(day.tanggal);
      const isWknd = isWeekendDate(day.tanggal);
      const dbBase = 6 + i * 10;

      ws.mergeCells(rowHari.number, colPagi, rowHari.number, colSiang);
      rowHari.getCell(colPagi).value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 1},FALSE),"Hari ${i + 1} - ${dName}")`,
        result: `Hari ${i + 1} - ${dName}`,
      };
      [colPagi, colSiang].forEach((cIdx) => {
        const cell = rowHari.getCell(cIdx);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isWknd ? 'FFFFE4E6' : 'FFE0E7FF' },
        };
        cell.font = {
          name: 'Calibri',
          size: 10,
          bold: true,
          color: { argb: isWknd ? 'FF9F1239' : 'FF1E1B4B' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = THIN_BORDER;
      });

      ws.mergeCells(rowTanggal.number, colPagi, rowTanggal.number, colSiang);
      rowTanggal.getCell(colPagi).value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase},FALSE),"${day.tanggal}")`,
        result: day.tanggal,
      };
      [colPagi, colSiang].forEach((cIdx) => {
        const cell = rowTanggal.getCell(cIdx);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isWknd ? 'FFFFF1F2' : 'FFF8FAFC' },
        };
        cell.font = {
          name: 'Calibri',
          size: 10,
          bold: true,
          color: { argb: 'FF0F172A' },
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = THIN_BORDER;
      });

      const cellShiftPagi = rowShift.getCell(colPagi);
      cellShiftPagi.value = 'Pagi';
      cellShiftPagi.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFBEB' },
      };
      cellShiftPagi.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: 'FF92400E' },
      };
      cellShiftPagi.alignment = { vertical: 'middle', horizontal: 'center' };
      cellShiftPagi.border = THIN_BORDER;

      const cellShiftSiang = rowShift.getCell(colSiang);
      cellShiftSiang.value = 'Siang';
      cellShiftSiang.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEEF2FF' },
      };
      cellShiftSiang.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: 'FF3730A3' },
      };
      cellShiftSiang.alignment = { vertical: 'middle', horizontal: 'center' };
      cellShiftSiang.border = THIN_BORDER;

      const cellNamaPagi = rowNama.getCell(colPagi);
      cellNamaPagi.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 2},FALSE),"")`,
        result: day.pagi.nama || '',
      };
      cellNamaPagi.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
      };
      cellNamaPagi.alignment = { vertical: 'middle', horizontal: 'center' };
      cellNamaPagi.border = THIN_BORDER;

      const cellNamaSiang = rowNama.getCell(colSiang);
      cellNamaSiang.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 3},FALSE),"")`,
        result: day.siang.nama || '',
      };
      cellNamaSiang.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
      };
      cellNamaSiang.alignment = { vertical: 'middle', horizontal: 'center' };
      cellNamaSiang.border = THIN_BORDER;

      const cellSalesPagi = rowSales.getCell(colPagi);
      cellSalesPagi.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 4},FALSE),0)`,
        result: day.pagi.sales ?? 0,
      };
      cellSalesPagi.numFmt = NUM_FMT_INTEGER;
      cellSalesPagi.font = { name: 'Calibri', size: 10 };
      cellSalesPagi.alignment = { vertical: 'middle', horizontal: 'right' };
      cellSalesPagi.border = THIN_BORDER;

      const cellSalesSiang = rowSales.getCell(colSiang);
      cellSalesSiang.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 5},FALSE),0)`,
        result: day.siang.sales ?? 0,
      };
      cellSalesSiang.numFmt = NUM_FMT_INTEGER;
      cellSalesSiang.font = { name: 'Calibri', size: 10 };
      cellSalesSiang.alignment = { vertical: 'middle', horizontal: 'right' };
      cellSalesSiang.border = THIN_BORDER;

      const cellQtyPagi = rowQty.getCell(colPagi);
      cellQtyPagi.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 6},FALSE),0)`,
        result: day.pagi.qty ?? 0,
      };
      cellQtyPagi.numFmt = NUM_FMT_INTEGER;
      cellQtyPagi.font = { name: 'Calibri', size: 10 };
      cellQtyPagi.alignment = { vertical: 'middle', horizontal: 'right' };
      cellQtyPagi.border = THIN_BORDER;

      const cellQtySiang = rowQty.getCell(colSiang);
      cellQtySiang.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 7},FALSE),0)`,
        result: day.siang.qty ?? 0,
      };
      cellQtySiang.numFmt = NUM_FMT_INTEGER;
      cellQtySiang.font = { name: 'Calibri', size: 10 };
      cellQtySiang.alignment = { vertical: 'middle', horizontal: 'right' };
      cellQtySiang.border = THIN_BORDER;

      const cellTrxPagi = rowTrx.getCell(colPagi);
      cellTrxPagi.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 8},FALSE),0)`,
        result: day.pagi.trx ?? 0,
      };
      cellTrxPagi.numFmt = NUM_FMT_INTEGER;
      cellTrxPagi.font = { name: 'Calibri', size: 10 };
      cellTrxPagi.alignment = { vertical: 'middle', horizontal: 'right' };
      cellTrxPagi.border = THIN_BORDER;

      const cellTrxSiang = rowTrx.getCell(colSiang);
      cellTrxSiang.value = {
        formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${dbBase + 9},FALSE),0)`,
        result: day.siang.trx ?? 0,
      };
      cellTrxSiang.numFmt = NUM_FMT_INTEGER;
      cellTrxSiang.font = { name: 'Calibri', size: 10 };
      cellTrxSiang.alignment = { vertical: 'middle', horizontal: 'right' };
      cellTrxSiang.border = THIN_BORDER;
    }

    const rowTotalSales = ws.addRow(['Total Sales']);
    const rowTotalQty = ws.addRow(['Total Qty']);
    const rowTotalTrx = ws.addRow(['Total Trx']);
    const rowPricePoint = ws.addRow(['Price Point']);
    const rowBigSize = ws.addRow(['Big Size']);

    [
      rowTotalSales,
      rowTotalQty,
      rowTotalTrx,
      rowPricePoint,
      rowBigSize,
    ].forEach((r) => {
      r.height = 22;
      const c1 = r.getCell(1);
      c1.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      c1.font = {
        name: 'Calibri',
        size: 10.5,
        bold: true,
        color: { argb: 'FF000000' },
      };
      c1.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      c1.border = THIN_BORDER;
    });

    for (let i = 0; i < numDays; i++) {
      const m = dayMetricsList[i];
      const colPagi = 2 + i * 2;
      const colSiang = colPagi + 1;
      const letPagi = colToLetter(colPagi);
      const letSiang = colToLetter(colSiang);

      ws.mergeCells(
        rowTotalSales.number,
        colPagi,
        rowTotalSales.number,
        colSiang
      );
      rowTotalSales.getCell(colPagi).value = {
        formula: `${letPagi}${rowSales.number}+${letSiang}${rowSales.number}`,
        result: m.totalSales,
      };

      ws.mergeCells(rowTotalQty.number, colPagi, rowTotalQty.number, colSiang);
      rowTotalQty.getCell(colPagi).value = {
        formula: `${letPagi}${rowQty.number}+${letSiang}${rowQty.number}`,
        result: m.totalQty,
      };

      ws.mergeCells(rowTotalTrx.number, colPagi, rowTotalTrx.number, colSiang);
      rowTotalTrx.getCell(colPagi).value = {
        formula: `${letPagi}${rowTrx.number}+${letSiang}${rowTrx.number}`,
        result: m.totalTrx,
      };

      ws.mergeCells(
        rowPricePoint.number,
        colPagi,
        rowPricePoint.number,
        colSiang
      );
      rowPricePoint.getCell(colPagi).value = {
        formula: showDivZero
          ? `ROUND(${letPagi}${rowTotalSales.number}/${letPagi}${rowTotalQty.number},0)`
          : `IF(${letPagi}${rowTotalQty.number}>0,ROUND(${letPagi}${rowTotalSales.number}/${letPagi}${rowTotalQty.number},0),0)`,
        result: m.pricePoint ?? 0,
      };

      ws.mergeCells(rowBigSize.number, colPagi, rowBigSize.number, colSiang);
      rowBigSize.getCell(colPagi).value = {
        formula: showDivZero
          ? `ROUND(${letPagi}${rowTotalSales.number}/${letPagi}${rowTotalTrx.number},0)`
          : `IF(${letPagi}${rowTotalTrx.number}>0,ROUND(${letPagi}${rowTotalSales.number}/${letPagi}${rowTotalTrx.number},0),0)`,
        result: m.bigSize ?? 0,
      };

      [
        { r: rowTotalSales, bg: 'FFFFF9C4' },
        { r: rowTotalQty, bg: 'FFFFFDE7' },
        { r: rowTotalTrx, bg: 'FFFFFDE7' },
        { r: rowPricePoint, bg: 'FFF8FAFC' },
        { r: rowBigSize, bg: 'FFF8FAFC' },
      ].forEach(({ r, bg }) => {
        [colPagi, colSiang].forEach((cIdx) => {
          const cell = r.getCell(cIdx);
          cell.numFmt = NUM_FMT_INTEGER;
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: bg },
          };
          cell.font = {
            name: 'Calibri',
            size: 10,
            bold: true,
            color: { argb: 'FF0F172A' },
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = THIN_BORDER;
        });
      });
    }

    const cellWeeklySales = rowTotalSales.getCell(totalColIdx);
    cellWeeklySales.value = {
      formula: `SUM(B${rowSales.number}:${lastShiftColLetter}${rowSales.number})`,
      result: weekMetrics.weeklySales,
    };

    const cellWeeklyQty = rowTotalQty.getCell(totalColIdx);
    cellWeeklyQty.value = {
      formula: `SUM(B${rowQty.number}:${lastShiftColLetter}${rowQty.number})`,
      result: weekMetrics.weeklyQty,
    };

    const cellWeeklyTrx = rowTotalTrx.getCell(totalColIdx);
    cellWeeklyTrx.value = {
      formula: `SUM(B${rowTrx.number}:${lastShiftColLetter}${rowTrx.number})`,
      result: weekMetrics.weeklyTrx,
    };

    const cellWeeklyPP = rowPricePoint.getCell(totalColIdx);
    cellWeeklyPP.value = {
      formula: `IF(${totalColLetter}${rowTotalQty.number}>0,ROUND(${totalColLetter}${rowTotalSales.number}/${totalColLetter}${rowTotalQty.number},0),0)`,
      result: weekMetrics.weeklyPricePoint ?? 0,
    };

    const cellWeeklyBS = rowBigSize.getCell(totalColIdx);
    cellWeeklyBS.value = {
      formula: `IF(${totalColLetter}${rowTotalTrx.number}>0,ROUND(${totalColLetter}${rowTotalSales.number}/${totalColLetter}${rowTotalTrx.number},0),0)`,
      result: weekMetrics.weeklyBigSize ?? 0,
    };

    [
      cellWeeklySales,
      cellWeeklyQty,
      cellWeeklyTrx,
      cellWeeklyPP,
      cellWeeklyBS,
    ].forEach((cell) => {
      cell.numFmt = NUM_FMT_INTEGER;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FF000000' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = THIN_BORDER;
    });

    if (includeReasonRows) {
      const rowReasonPagi = ws.addRow(['Reason / Event\nShift Pagi']);
      rowReasonPagi.height = 48;
      const rowReasonSiang = ws.addRow(['Reason / Event\nShift Siang']);
      rowReasonSiang.height = 48;

      [rowReasonPagi, rowReasonSiang].forEach((r) => {
        const c1 = r.getCell(1);
        c1.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFC000' },
        };
        c1.font = {
          name: 'Calibri',
          size: 10,
          bold: true,
          color: { argb: 'FF000000' },
        };
        c1.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
          indent: 1,
        };
        c1.border = THIN_BORDER;
      });

      for (let i = 0; i < numDays; i++) {
        const day = table.days[i];
        const colPagi = 2 + i * 2;
        const colSiang = colPagi + 1;
        const reasonPagiColIdx = 76 + i * 2;
        const reasonSiangColIdx = 77 + i * 2;

        ws.mergeCells(
          rowReasonPagi.number,
          colPagi,
          rowReasonPagi.number,
          colSiang
        );
        rowReasonPagi.getCell(colPagi).value = {
          formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${reasonPagiColIdx},FALSE),"${(day.reasonPagi || '-').replace(/"/g, '""')}")`,
          result: day.reasonPagi || '-',
        };
        [colPagi, colSiang].forEach((cIdx) => {
          const cell = rowReasonPagi.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFBEB' },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            color: { argb: 'FF1E293B' },
          };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'left',
            wrapText: true,
          };
          cell.border = THIN_BORDER;
        });

        ws.mergeCells(
          rowReasonSiang.number,
          colPagi,
          rowReasonSiang.number,
          colSiang
        );
        rowReasonSiang.getCell(colPagi).value = {
          formula: `IFERROR(VLOOKUP($A$${tRowNum},$Z$2:$DJ$107,${reasonSiangColIdx},FALSE),"${(day.reasonSiang || '-').replace(/"/g, '""')}")`,
          result: day.reasonSiang || '-',
        };
        [colPagi, colSiang].forEach((cIdx) => {
          const cell = rowReasonSiang.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEEF2FF' },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            color: { argb: 'FF1E293B' },
          };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'left',
            wrapText: true,
          };
          cell.border = THIN_BORDER;
        });
      }

      const rpSummary = rowReasonPagi.getCell(totalColIdx);
      rpSummary.value = 'Evaluasi Shift Pagi';
      rpSummary.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFEF3C7' },
      };
      rpSummary.font = {
        name: 'Calibri',
        size: 9.5,
        italic: true,
        bold: true,
        color: { argb: 'FF92400E' },
      };
      rpSummary.alignment = { vertical: 'middle', horizontal: 'center' };
      rpSummary.border = THIN_BORDER;

      const rsSummary = rowReasonSiang.getCell(totalColIdx);
      rsSummary.value = 'Evaluasi Shift Siang';
      rsSummary.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E7FF' },
      };
      rsSummary.font = {
        name: 'Calibri',
        size: 9.5,
        italic: true,
        bold: true,
        color: { argb: 'FF3730A3' },
      };
      rsSummary.alignment = { vertical: 'middle', horizontal: 'center' };
      rsSummary.border = THIN_BORDER;
    }

    if (includeChartDiagrams) {
      const diagSpacer = ws.addRow([]);
      diagSpacer.height = 8;

      const rHariNum = rowHari.number;
      const rTanggalNum = rowTanggal.number;
      const rSalesNum = rowSales.number;
      const rTotSalesNum = rowTotalSales.number;
      const wkTotalRef = `${totalColLetter}${rTotSalesNum}`;

      let wPagiInit = 0;
      let wSiangInit = 0;
      let maxShiftInit = 0;
      let maxDayTotalInit = 0;
      const pagiColsExprParts: string[] = [];
      const siangColsExprParts: string[] = [];

      for (let i = 0; i < numDays; i++) {
        const d = table.days[i];
        const ps = d.pagi.sales ?? 0;
        const ss = d.siang.sales ?? 0;
        const ts = ps + ss;
        wPagiInit += ps;
        wSiangInit += ss;
        if (ps > maxShiftInit) maxShiftInit = ps;
        if (ss > maxShiftInit) maxShiftInit = ss;
        if (ts > maxDayTotalInit) maxDayTotalInit = ts;

        const cPagi = 2 + i * 2;
        const cSiang = cPagi + 1;
        pagiColsExprParts.push(`${colToLetter(cPagi)}${rSalesNum}`);
        siangColsExprParts.push(`${colToLetter(cSiang)}${rSalesNum}`);
      }

      const wTotalInit = wPagiInit + wSiangInit;
      const pagiPctInit =
        wTotalInit > 0 ? Math.round((wPagiInit / wTotalInit) * 100) : 0;
      const siangPctInit =
        wTotalInit > 0 ? Math.round((wSiangInit / wTotalInit) * 100) : 0;
      const pagiSumFormulaExpr = `(${pagiColsExprParts.join('+')})`;
      const siangSumFormulaExpr = `(${siangColsExprParts.join('+')})`;

      // ROW 1: Dark Navy Header Banner (Dynamic Title + Dynamic Filter Badge)
      const diagHeaderRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagHeaderRow.height = 24;
      const dhNum = diagHeaderRow.number;
      ws.mergeCells(dhNum, 2, dhNum, 13);
      ws.mergeCells(dhNum, 14, dhNum, totalColIdx);

      const cleanTitleInit = `DIAGRAM SALES WEEKLY — TOTAL WEEKLY ${table.weekNumber} (${periodeLabel.toUpperCase()} (${dayCycleLabel.toUpperCase()}))`;
      diagHeaderRow.getCell(2).value = {
        formula: `"DIAGRAM SALES WEEKLY — " & SUBSTITUTE(A${tRowNum}," ▼","") & " (" & UPPER(MID(B${tRowNum},11,100)) & ")"`,
        result: cleanTitleInit,
      };
      diagHeaderRow.getCell(14).value = {
        formula: `A${tRowNum} & " (" & MID(B${tRowNum},11,24) & ")"`,
        result: `${activeWeeklyKey} (${periodeLabel})`,
      };

      for (let c = 2; c <= totalColIdx; c++) {
        const cell = diagHeaderRow.getCell(c);
        const isBadge = c >= 14;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isBadge ? 'FF451A03' : 'FF0F172A' },
        };
        cell.font = {
          name: 'Calibri',
          size: isBadge ? 9.5 : 10.5,
          bold: true,
          color: { argb: 'FFFDE047' },
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: isBadge ? 'center' : 'left',
          indent: isBadge ? 0 : 1,
        };
        cell.border = THIN_BORDER;
      }

      // ROW 2: Dark Navy Subtitle & Live Shift Pagi / Shift Siang / Garis Tren Legend
      const diagLegendRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagLegendRow.height = 21;
      const dlNum = diagLegendRow.number;
      ws.mergeCells(dlNum, 2, dlNum, 5);
      ws.mergeCells(dlNum, 6, dlNum, 9);
      ws.mergeCells(dlNum, 10, dlNum, 13);
      ws.mergeCells(dlNum, 14, dlNum, totalColIdx);

      diagLegendRow.getCell(2).value =
        'Diagram Batang Vertikal 7 Hari & Garis Tren Shift Pagi vs Siang';
      diagLegendRow.getCell(6).value = {
        formula: `"■ Shift Pagi: Rp " & TEXT(${pagiSumFormulaExpr},"#,##0") & " (" & TEXT(IF(${wkTotalRef}>0,${pagiSumFormulaExpr}/${wkTotalRef},0),"0%") & ")"`,
        result: `■ Shift Pagi: Rp ${wPagiInit.toLocaleString('id-ID')} (${pagiPctInit}%)`,
      };
      diagLegendRow.getCell(10).value = {
        formula: `"■ Shift Siang: Rp " & TEXT(${siangSumFormulaExpr},"#,##0") & " (" & TEXT(IF(${wkTotalRef}>0,${siangSumFormulaExpr}/${wkTotalRef},0),"0%") & ")"`,
        result: `■ Shift Siang: Rp ${wSiangInit.toLocaleString('id-ID')} (${siangPctInit}%)`,
      };
      diagLegendRow.getCell(14).value = '━●━ Garis Tren Total Sales Harian';

      for (let c = 2; c <= totalColIdx; c++) {
        const cell = diagLegendRow.getCell(c);
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF0F172A' },
        };
        const fgColor =
          c >= 6 && c <= 9
            ? 'FFFDE68A'
            : c >= 10 && c <= 13
              ? 'FFC7D2FE'
              : c >= 14
                ? 'FF6EE7B7'
                : 'FFCBD5E1';
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: c >= 6,
          color: { argb: fgColor },
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: c >= 14 ? 'center' : 'left',
          indent: c < 14 ? 1 : 0,
        };
        cell.border = THIN_BORDER;
      }

      // ROW 3: White Card Sub-Header (Orange Badge + Card Title + Emerald Total Omzet Pill)
      const diagSubRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagSubRow.height = 22;
      const dsNum = diagSubRow.number;
      ws.mergeCells(dsNum, 2, dsNum, 5);
      ws.mergeCells(dsNum, 6, dsNum, 13);
      ws.mergeCells(dsNum, 14, dsNum, totalColIdx);

      diagSubRow.getCell(2).value = 'DIAGRAM SALES (OMZET RUPIAH)';
      diagSubRow.getCell(6).value =
        'Grafik Batang Shift Pagi (Oranye) vs Shift Siang (Biru) & Garis Tren (Otomatis Mengikuti Filter Weekly)';
      diagSubRow.getCell(14).value = {
        formula: `"Total: Rp " & TEXT(${wkTotalRef},"#,##0")`,
        result: `Total: Rp ${wTotalInit.toLocaleString('id-ID')}`,
      };

      for (let c = 2; c <= totalColIdx; c++) {
        const cell = diagSubRow.getCell(c);
        const isOrangeBadge = c >= 2 && c <= 5;
        const isGreenBadge = c >= 14;
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: isOrangeBadge
              ? 'FFD97706'
              : isGreenBadge
                ? 'FFECFDF5'
                : 'FFFFFFFF',
          },
        };
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: true,
          color: {
            argb: isOrangeBadge
              ? 'FFFFFFFF'
              : isGreenBadge
                ? 'FF047857'
                : 'FF0F172A',
          },
        };
        cell.alignment = {
          vertical: 'middle',
          horizontal: isOrangeBadge || isGreenBadge ? 'center' : 'left',
          indent: isOrangeBadge || isGreenBadge ? 0 : 1,
        };
        cell.border = THIN_BORDER;
      }

      // ROW 4: Dynamic Top Data Labels & Green Trend Nodes ("● 3.5jt", "● 2.8jt", etc.)
      const diagTopValRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagTopValRow.height = 20;
      const dtvNum = diagTopValRow.number;

      for (let i = 0; i < numDays; i++) {
        const cPagi = 2 + i * 2;
        const cSiang = cPagi + 1;
        const letPagi = colToLetter(cPagi);
        const dayTotRef = `${letPagi}${rTotSalesNum}`;
        const dTotVal = dayMetricsList[i].totalSales;
        const dTotText =
          dTotVal >= 1000000
            ? `● ${(dTotVal / 1000000).toFixed(1)}jt`
            : dTotVal > 0
              ? `● ${Math.round(dTotVal / 1000)}rb`
              : '● 0';

        ws.mergeCells(dtvNum, cPagi, dtvNum, cSiang);
        diagTopValRow.getCell(cPagi).value = {
          formula: `IF(${dayTotRef}>=1000000,"● " & TEXT(${dayTotRef}/1000000,"0.0") & "jt",IF(${dayTotRef}>0,"● " & TEXT(${dayTotRef}/1000,"0") & "rb","● 0"))`,
          result: dTotText,
        };

        [cPagi, cSiang].forEach((cIdx) => {
          const cell = diagTopValRow.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: true,
            color: { argb: 'FF047857' },
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'dotted', color: { argb: 'FFCBD5E1' } },
          };
        });
      }

      const peakInitText =
        maxDayTotalInit >= 1000000
          ? `Puncak: ${(maxDayTotalInit / 1000000).toFixed(1)}jt`
          : `Puncak: ${maxDayTotalInit.toLocaleString('id-ID')}`;
      const peakCell = diagTopValRow.getCell(totalColIdx);
      peakCell.value = {
        formula: `"Puncak: " & IF(MAX(B${rTotSalesNum}:${lastShiftColLetter}${rTotSalesNum})>=1000000,TEXT(MAX(B${rTotSalesNum}:${lastShiftColLetter}${rTotSalesNum})/1000000,"0.0")&"jt",TEXT(MAX(B${rTotSalesNum}:${lastShiftColLetter}${rTotSalesNum}),"#,##0"))`,
        result: peakInitText,
      };
      peakCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8FAFC' },
      };
      peakCell.font = {
        name: 'Calibri',
        size: 9,
        bold: true,
        color: { argb: 'FF334155' },
      };
      peakCell.alignment = { vertical: 'middle', horizontal: 'center' };
      peakCell.border = THIN_BORDER;

      // ROWS 5..9: 5 Vertical Bar Levels (100%, 80%, 60%, 40%, 20% Scale)
      // Left sub-col = Shift Pagi Bar (Amber #F59E0B), Right sub-col = Shift Siang Bar (Indigo #4F46E5)
      const barLevels = [
        { full: 0.85, half: 0.72, scaleFactor: 1.0 },
        { full: 0.65, half: 0.52, scaleFactor: 0.75 },
        { full: 0.45, half: 0.32, scaleFactor: 0.5 },
        { full: 0.25, half: 0.14, scaleFactor: 0.25 },
        { full: 0.08, half: 0.001, scaleFactor: 0.0 },
      ];
      const maxShiftRangeExpr = `MAX($B$${rSalesNum}:$${lastShiftColLetter}$${rSalesNum})`;

      barLevels.forEach((lvl, lvlIdx) => {
        const barRow = ws.addRow(new Array(totalColIdx).fill(''));
        barRow.height = 17;

        for (let i = 0; i < numDays; i++) {
          const d = table.days[i];
          const ps = d.pagi.sales ?? 0;
          const ss = d.siang.sales ?? 0;
          const cPagi = 2 + i * 2;
          const cSiang = cPagi + 1;
          const letPagi = colToLetter(cPagi);
          const letSiang = colToLetter(cSiang);

          const pRatio = maxShiftInit > 0 ? ps / maxShiftInit : 0;
          const sRatio = maxShiftInit > 0 ? ss / maxShiftInit : 0;
          const pInitBar =
            pRatio >= lvl.full
              ? '██████████'
              : pRatio >= lvl.half
                ? '▄▄▄▄▄▄▄▄▄▄'
                : '';
          const sInitBar =
            sRatio >= lvl.full
              ? '██████████'
              : sRatio >= lvl.half
                ? '▄▄▄▄▄▄▄▄▄▄'
                : '';

          const cellP = barRow.getCell(cPagi);
          cellP.value = {
            formula: `IF(${maxShiftRangeExpr}>0,IF(${letPagi}$${rSalesNum}/${maxShiftRangeExpr}>=${lvl.full},"██████████",IF(${letPagi}$${rSalesNum}/${maxShiftRangeExpr}>=${lvl.half},"▄▄▄▄▄▄▄▄▄▄","")),"")`,
            result: pInitBar,
          };
          cellP.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          cellP.font = {
            name: 'Calibri',
            size: 10,
            bold: true,
            color: { argb: 'FFF59E0B' },
          };
          cellP.alignment = { vertical: 'bottom', horizontal: 'center' };
          cellP.border = {
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'dotted', color: { argb: 'FFF1F5F9' } },
            top: { style: 'dotted', color: { argb: 'FFE2E8F0' } },
            bottom:
              lvlIdx === barLevels.length - 1
                ? { style: 'thin', color: { argb: 'FF94A3B8' } }
                : { style: 'dotted', color: { argb: 'FFE2E8F0' } },
          };

          const cellS = barRow.getCell(cSiang);
          cellS.value = {
            formula: `IF(${maxShiftRangeExpr}>0,IF(${letSiang}$${rSalesNum}/${maxShiftRangeExpr}>=${lvl.full},"██████████",IF(${letSiang}$${rSalesNum}/${maxShiftRangeExpr}>=${lvl.half},"▄▄▄▄▄▄▄▄▄▄","")),"")`,
            result: sInitBar,
          };
          cellS.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          cellS.font = {
            name: 'Calibri',
            size: 10,
            bold: true,
            color: { argb: 'FF4F46E5' },
          };
          cellS.alignment = { vertical: 'bottom', horizontal: 'center' };
          cellS.border = {
            left: { style: 'dotted', color: { argb: 'FFF1F5F9' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            top: { style: 'dotted', color: { argb: 'FFE2E8F0' } },
            bottom:
              lvlIdx === barLevels.length - 1
                ? { style: 'thin', color: { argb: 'FF94A3B8' } }
                : { style: 'dotted', color: { argb: 'FFE2E8F0' } },
          };
        }

        // Dynamic Y-Axis Scale in Col P
        const scaleCell = barRow.getCell(totalColIdx);
        const scaleValInit = Math.round(maxShiftInit * lvl.scaleFactor);
        const scaleTextInit =
          lvl.scaleFactor === 0
            ? '0'
            : scaleValInit >= 1000000
              ? `${(scaleValInit / 1000000).toFixed(1)}jt`
              : scaleValInit.toLocaleString('id-ID');
        scaleCell.value =
          lvl.scaleFactor === 0
            ? '0'
            : {
                formula: `IF(${maxShiftRangeExpr}>=1000000,TEXT(${maxShiftRangeExpr}*${lvl.scaleFactor}/1000000,"0.0")&"jt",TEXT(ROUND(${maxShiftRangeExpr}*${lvl.scaleFactor},0),"#,##0"))`,
                result: scaleTextInit,
              };
        scaleCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8FAFC' },
        };
        scaleCell.font = {
          name: 'Calibri',
          size: 8.5,
          bold: true,
          color: { argb: 'FF64748B' },
        };
        scaleCell.alignment = { vertical: 'middle', horizontal: 'center' };
        scaleCell.border = THIN_BORDER;
      });

      // ROW 10: Shift Pagi vs Shift Siang Nominal Values directly under each bar
      const diagShiftValRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagShiftValRow.height = 19;
      for (let i = 0; i < numDays; i++) {
        const d = table.days[i];
        const cPagi = 2 + i * 2;
        const cSiang = cPagi + 1;
        const letPagi = colToLetter(cPagi);
        const letSiang = colToLetter(cSiang);

        const cellP = diagShiftValRow.getCell(cPagi);
        cellP.value = {
          formula: `${letPagi}${rSalesNum}`,
          result: d.pagi.sales ?? 0,
        };
        cellP.numFmt = NUM_FMT_INTEGER;
        cellP.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFEF3C7' },
        };
        cellP.font = {
          name: 'Calibri',
          size: 9,
          bold: true,
          color: { argb: 'FF92400E' },
        };
        cellP.alignment = { vertical: 'middle', horizontal: 'right' };
        cellP.border = THIN_BORDER;

        const cellS = diagShiftValRow.getCell(cSiang);
        cellS.value = {
          formula: `${letSiang}${rSalesNum}`,
          result: d.siang.sales ?? 0,
        };
        cellS.numFmt = NUM_FMT_INTEGER;
        cellS.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFE0E7FF' },
        };
        cellS.font = {
          name: 'Calibri',
          size: 9,
          bold: true,
          color: { argb: 'FF3730A3' },
        };
        cellS.alignment = { vertical: 'middle', horizontal: 'right' };
        cellS.border = THIN_BORDER;
      }

      const totBarFooterCell = diagShiftValRow.getCell(totalColIdx);
      totBarFooterCell.value = {
        formula: wkTotalRef,
        result: wTotalInit,
      };
      totBarFooterCell.numFmt = NUM_FMT_INTEGER;
      totBarFooterCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      totBarFooterCell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FF000000' },
      };
      totBarFooterCell.alignment = { vertical: 'middle', horizontal: 'center' };
      totBarFooterCell.border = THIN_BORDER;

      // ROW 11: Dynamic X-Axis Labels (Day Name + Date from rowHari & rowTanggal)
      const diagAxisRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagAxisRow.height = 22;
      const daxNum = diagAxisRow.number;

      for (let i = 0; i < numDays; i++) {
        const d = table.days[i];
        const cPagi = 2 + i * 2;
        const cSiang = cPagi + 1;
        const letPagi = colToLetter(cPagi);
        const isWk = isWeekendDate(d.tanggal);
        const dShort = getIndoDayName(d.tanggal).slice(0, 3);
        const dtShort = (d.tanggal || '').slice(0, 5);

        ws.mergeCells(daxNum, cPagi, daxNum, cSiang);
        diagAxisRow.getCell(cPagi).value = {
          formula: `MID(${letPagi}${rHariNum},10,3) & " (H${i + 1}) • " & LEFT(${letPagi}${rTanggalNum},5)`,
          result: `${dShort} (H${i + 1}) • ${dtShort}`,
        };

        [cPagi, cSiang].forEach((cIdx) => {
          const cell = diagAxisRow.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: isWk ? 'FFFFE4E6' : 'FFF1F5F9' },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: true,
            color: { argb: isWk ? 'FF9F1239' : 'FF0F172A' },
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = THIN_BORDER;
        });
      }

      const axisRightCell = diagAxisRow.getCell(totalColIdx);
      axisRightCell.value = {
        formula: `SUBSTITUTE(A${tRowNum}," ▼","")`,
        result: `TOTAL WEEKLY ${table.weekNumber}`,
      };
      axisRightCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F172A' },
      };
      axisRightCell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FFFDE047' },
      };
      axisRightCell.alignment = { vertical: 'middle', horizontal: 'center' };
      axisRightCell.border = THIN_BORDER;

      // ROW 12 & 13: LIVE FORMULA KONTRIBUSI (%) & GROWTH HARIAN (%) DENGAN BAR VISUAL
      const diagContribRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagContribRow.height = 21;
      const dcrNum = diagContribRow.number;
      diagContribRow.getCell(1).value = 'Kontribusi (%)';
      diagContribRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0284C7' },
      };
      diagContribRow.getCell(1).font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      diagContribRow.getCell(1).alignment = {
        vertical: 'middle',
        horizontal: 'left',
        indent: 1,
      };
      diagContribRow.getCell(1).border = THIN_BORDER;

      const diagGrowthRow = ws.addRow(new Array(totalColIdx).fill(''));
      diagGrowthRow.height = 21;
      const dgrNum = diagGrowthRow.number;
      diagGrowthRow.getCell(1).value = 'Growth Harian (%)';
      diagGrowthRow.getCell(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF059669' },
      };
      diagGrowthRow.getCell(1).font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      diagGrowthRow.getCell(1).alignment = {
        vertical: 'middle',
        horizontal: 'left',
        indent: 1,
      };
      diagGrowthRow.getCell(1).border = THIN_BORDER;

      for (let i = 0; i < numDays; i++) {
        const cPagi = 2 + i * 2;
        const cSiang = cPagi + 1;
        const letPagi = colToLetter(cPagi);
        const dayTotRef = `${letPagi}${rTotSalesNum}`;
        const dTotVal = dayMetricsList[i].totalSales;
        const contribPctInit =
          wTotalInit > 0 ? Math.round((dTotVal / wTotalInit) * 1000) / 10 : 0;
        const barBlocksCount = Math.min(
          8,
          Math.max(0, Math.round(contribPctInit / 4))
        );
        const contribTextInit = `${'█'.repeat(barBlocksCount)} ${contribPctInit.toFixed(1)}%`.trim();

        ws.mergeCells(dcrNum, cPagi, dcrNum, cSiang);
        diagContribRow.getCell(cPagi).value = {
          formula: `IF(${wkTotalRef}>0,REPT("█",MIN(8,ROUND((${dayTotRef}/${wkTotalRef})*25,0))) & " " & TEXT(${dayTotRef}/${wkTotalRef},"0.0%"),"0.0%")`,
          result: contribTextInit,
        };
        [cPagi, cSiang].forEach((cIdx) => {
          const cell = diagContribRow.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF0F9FF' },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: true,
            color: { argb: 'FF0369A1' },
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = THIN_BORDER;
        });

        // Growth Harian (DoD vs previous day in the same week)
        ws.mergeCells(dgrNum, cPagi, dgrNum, cSiang);
        if (i === 0) {
          diagGrowthRow.getCell(cPagi).value = 'Basis (H1)';
        } else {
          const prevCPagi = 2 + (i - 1) * 2;
          const prevLetPagi = colToLetter(prevCPagi);
          const prevDayTotRef = `${prevLetPagi}${rTotSalesNum}`;
          const prevVal = dayMetricsList[i - 1].totalSales;
          const gPctInit =
            prevVal > 0 ? ((dTotVal - prevVal) / prevVal) * 100 : dTotVal > 0 ? 100 : 0;
          const gTextInit =
            prevVal <= 0 && dTotVal <= 0
              ? '▬ 0.0%'
              : gPctInit > 0
                ? `▲ +${gPctInit.toFixed(1)}%`
                : gPctInit < 0
                  ? `▼ ${gPctInit.toFixed(1)}%`
                  : '▬ 0.0%';

          diagGrowthRow.getCell(cPagi).value = {
            formula: `IF(AND(${prevDayTotRef}<=0,${dayTotRef}<=0),"▬ 0.0%",IF(${prevDayTotRef}<=0,"▲ +100.0%",IF(${dayTotRef}>${prevDayTotRef},"▲ +" & TEXT((${dayTotRef}-${prevDayTotRef})/${prevDayTotRef},"0.0%"),IF(${dayTotRef}<${prevDayTotRef},"▼ " & TEXT((${dayTotRef}-${prevDayTotRef})/${prevDayTotRef},"0.0%"),"▬ 0.0%"))))`,
            result: gTextInit,
          };
        }

        const prevValForColor = i > 0 ? dayMetricsList[i - 1].totalSales : null;
        const isUpInit =
          prevValForColor !== null && dTotVal > prevValForColor && dTotVal > 0;
        const isDownInit =
          prevValForColor !== null && dTotVal < prevValForColor && prevValForColor > 0;

        [cPagi, cSiang].forEach((cIdx) => {
          const cell = diagGrowthRow.getCell(cIdx);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
              argb:
                i === 0
                  ? 'FFFEF3C7'
                  : isUpInit
                    ? 'FFECFDF5'
                    : isDownInit
                      ? 'FFFFF1F2'
                      : 'FFF8FAFC',
            },
          };
          cell.font = {
            name: 'Calibri',
            size: 9.5,
            bold: true,
            color: {
              argb:
                i === 0
                  ? 'FF92400E'
                  : isUpInit
                    ? 'FF047857'
                    : isDownInit
                      ? 'FFBE123C'
                      : 'FF475569',
            },
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.border = THIN_BORDER;
        });
      }

      const contribRightCell = diagContribRow.getCell(totalColIdx);
      contribRightCell.value = '100.0%';
      contribRightCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0F2FE' },
      };
      contribRightCell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FF0369A1' },
      };
      contribRightCell.alignment = { vertical: 'middle', horizontal: 'center' };
      contribRightCell.border = THIN_BORDER;

      const growthRightCell = diagGrowthRow.getCell(totalColIdx);
      growthRightCell.value = 'Tren H1→H7';
      growthRightCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFECFDF5' },
      };
      growthRightCell.font = {
        name: 'Calibri',
        size: 9.5,
        bold: true,
        color: { argb: 'FF047857' },
      };
      growthRightCell.alignment = { vertical: 'middle', horizontal: 'center' };
      growthRightCell.border = THIN_BORDER;

      // Also embed visual high-DPI Bar Chart card for Growth & Kontribusi of this Weekly table
      const weeklyDayChartItems: ChartBarItem[] = table.days.map((d, idx) => {
        const ps = d.pagi.sales ?? 0;
        const ss = d.siang.sales ?? 0;
        return {
          label: `${getIndoDayName(d.tanggal).slice(0, 3)} (H${idx + 1})`,
          subLabel: (d.tanggal || '').slice(0, 5),
          pagiSales: ps,
          siangSales: ss,
          totalSales: ps + ss,
          isWeekend: isWeekendDate(d.tanggal),
        };
      });

      const gcChartSpacer = ws.addRow([]);
      gcChartSpacer.height = 6;
      renderGrowthAndContributionDiagramCard(
        workbook,
        ws,
        `DIAGRAM GROWTH & KONTRIBUSI WEEKLY ${table.weekNumber} (${periodeLabel.toUpperCase()})`,
        'Diagram Batang Pertumbuhan Harian (Growth %) & Porsi Kontribusi 7 Hari terhadap Total Weekly (100%)',
        weeklyDayChartItems,
        `WEEKLY ${table.weekNumber}`
      );

      const afterDiagSpacer = ws.addRow([]);
      afterDiagSpacer.height = 10;
    } else {
      const compactGapRow = ws.addRow([]);
      compactGapRow.height = 10;
    }
  }

  // Note: We intentionally keep the top table cell (first.rowNum) as a plain string value
  // ("TOTAL WEEKLY X ▼") so there is never a circular reference in WPS Office or Excel.
  // Lower tables (2nd, 3rd, etc.) already link one-way (+1 week) to the table above them
  // and also have their own Data Validation dropdown ($Z$2:$Z$54).
  return renderedTitleRows[0]?.rowNum ?? 7;
}

/* ============================================================================
   SECTION RENDERER 2: SALES TABLE DAILY (PERSIS FORMAT DI APLIKASI - 16 KOLOM & HEADER 2 BARIS)
   ============================================================================ */
function renderDailyTableSection(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  tables: WeeklySalesTable[],
  includeChartDiagrams: boolean = true,
  enableSheetAutoFilter: boolean = false,
  selectedMonthKey?: string,
  allTablesPool?: WeeklySalesTable[],
  resolvedActiveMonthKey?: string
): number {
  const sourceTables =
    allTablesPool && allTablesPool.length > 0
      ? [...allTablesPool, ...tables]
      : tables;

  const activeMonthKey =
    resolvedActiveMonthKey ||
    resolveLatestActiveMonthKey(selectedMonthKey, sourceTables);
  const activeMonthMeta =
    MONTH_META_LIST.find((m) => m.key === activeMonthKey) || MONTH_META_LIST[10];

  // Build map of DD/MM/YYYY -> { dayData, weekNumber, weekTitle } from all tables
  const dateMap = new Map<
    string,
    {
      day: WeeklySalesTable['days'][number];
      weekNumber: number;
      weekTitle: string;
    }
  >();
  sourceTables.forEach((t) => {
    t.days.forEach((d) => {
      if (d.tanggal) {
        const existing = dateMap.get(d.tanggal);
        const curSales = (d.pagi.sales ?? 0) + (d.siang.sales ?? 0);
        const curTraf = (Number(d.pagi.traffic) || 0) + (Number(d.siang.traffic) || 0);
        const curExtra =
          (d.pagi.qty ?? 0) +
          (d.siang.qty ?? 0) +
          (d.pagi.trx ?? 0) +
          (d.siang.trx ?? 0) +
          ((d.reasonPagi || '').trim() ? 1 : 0) +
          ((d.reasonSiang || '').trim() ? 1 : 0);
        const curScore = curSales * 100 + curTraf * 10 + curExtra;

        const exSales = existing
          ? (existing.day.pagi.sales ?? 0) + (existing.day.siang.sales ?? 0)
          : -1;
        const exTraf = existing
          ? (Number(existing.day.pagi.traffic) || 0) +
            (Number(existing.day.siang.traffic) || 0)
          : -1;
        const exExtra = existing
          ? (existing.day.pagi.qty ?? 0) +
            (existing.day.siang.qty ?? 0) +
            (existing.day.pagi.trx ?? 0) +
            (existing.day.siang.trx ?? 0) +
            ((existing.day.reasonPagi || '').trim() ? 1 : 0) +
            ((existing.day.reasonSiang || '').trim() ? 1 : 0)
          : -1;
        const exScore = existing ? exSales * 100 + exTraf * 10 + exExtra : -1;

        if (!existing || curScore >= exScore) {
          dateMap.set(d.tanggal, {
            day: d,
            weekNumber: t.weekNumber,
            weekTitle: t.title,
          });
        }
      }
    });
  });

  // Always build full calendar month from Tanggal 1 to Tanggal 30/31 matching DailySalesSpreadsheetView in the app
  interface DailyRowRecord {
    dayNum: number | string;
    tanggal: string;
    hari: string;
    isWeekend: boolean;
    weeklyLabel: string;
    pagiNama: string;
    pagiSales: number;
    pagiQty: number;
    pagiTrx: number;
    siangNama: string;
    siangSales: number;
    siangQty: number;
    siangTrx: number;
    totalSales: number;
    totalQty: number;
    totalTrx: number;
    pricePoint: number;
    bigSize: number;
    pagiTraffic: number;
    siangTraffic: number;
    totalTraffic: number;
    convRateRatio: number;
    reasonPagi: string;
    reasonSiang: string;
  }

  const dailyRecords: DailyRowRecord[] = [];

  for (let dNum = 1; dNum <= activeMonthMeta.daysCount; dNum++) {
    const dd = String(dNum).padStart(2, '0');
    const fullDate = `${dd}/${activeMonthMeta.key}/2026`;
    const dName = getIndoDayName(fullDate);
    const wknd = isWeekendDate(fullDate);
    const matched = dateMap.get(fullDate);
    const schedWeek = ANNUAL_2026_WEEKS.find((w) =>
      w.slashDates.includes(fullDate)
    );
    const weekLabel = matched
      ? `W${matched.weekNumber}`
      : schedWeek
        ? `W${schedWeek.weekNumber}`
        : '-';

    const pSales = matched ? matched.day.pagi.sales ?? 0 : 0;
    const pQty = matched ? matched.day.pagi.qty ?? 0 : 0;
    const pTrx = matched ? matched.day.pagi.trx ?? 0 : 0;
    const sSales = matched ? matched.day.siang.sales ?? 0 : 0;
    const sQty = matched ? matched.day.siang.qty ?? 0 : 0;
    const sTrx = matched ? matched.day.siang.trx ?? 0 : 0;
    const tSales = pSales + sSales;
    const tQty = pQty + sQty;
    const tTrx = pTrx + sTrx;
    const pp = tQty > 0 ? Math.round(tSales / tQty) : 0;
    const bs = tTrx > 0 ? Math.round(tSales / tTrx) : 0;
    const pTraffic = matched ? Number(matched.day.pagi.traffic) || 0 : 0;
    const sTraffic = matched ? Number(matched.day.siang.traffic) || 0 : 0;
    const tTraffic = pTraffic + sTraffic;
    const convRatio = tTraffic > 0 ? tTrx / tTraffic : 0;

    const rPagi = matched ? (matched.day.reasonPagi || '').trim() : '';
    const rSiang = matched ? (matched.day.reasonSiang || '').trim() : '';

    dailyRecords.push({
      dayNum: dNum,
      tanggal: fullDate,
      hari: dName,
      isWeekend: wknd,
      weeklyLabel: weekLabel,
      pagiNama: matched?.day.pagi.nama || '-',
      pagiSales: pSales,
      pagiQty: pQty,
      pagiTrx: pTrx,
      siangNama: matched?.day.siang.nama || '-',
      siangSales: sSales,
      siangQty: sQty,
      siangTrx: sTrx,
      totalSales: tSales,
      totalQty: tQty,
      totalTrx: tTrx,
      pricePoint: pp,
      bigSize: bs,
      pagiTraffic: pTraffic,
      siangTraffic: sTraffic,
      totalTraffic: tTraffic,
      convRateRatio: convRatio,
      reasonPagi: rPagi || '-',
      reasonSiang: rSiang || '-',
    });
  }

  const periodBannerTitle = `TABEL SALES HARIAN (DAILY) DP & TRAFFIC — BULAN ${activeMonthMeta.name.toUpperCase()} 2026 (TANGGAL 1 S/D ${activeMonthMeta.daysCount} ${activeMonthMeta.name.toUpperCase()})`;

  addSectionBanner(ws, periodBannerTitle, 23, 'FF0F172A', 'FFFDE047');

  // 1) EXACT 2-ROW GROUPED TABLE HEADER WITH TRAFFIC PAGI/SIANG + CONV RATE (COLS 18..21) & REASON PAGI/SIANG (COLS 22..23)
  const dHeaderTop = ws.addRow([
    'Tgl',
    'Tanggal',
    'Hari',
    'Weekly',
    'Shift Pagi',
    '',
    '',
    '',
    'Shift Siang',
    '',
    '',
    '',
    'Total Harian (Pagi + Siang)',
    '',
    '',
    '',
    '',
    'Traffic & Conv. Rate (DP)',
    '',
    '',
    '',
    'Reason / Event Pagi',
    'Reason / Event Siang',
  ]);
  dHeaderTop.height = 24;

  const dHeaderSub = ws.addRow([
    '',
    '',
    '',
    '',
    'Petugas',
    'Sales (Rp)',
    'Qty',
    'Trx',
    'Petugas',
    'Sales (Rp)',
    'Qty',
    'Trx',
    'Sales (Rp)',
    'Qty',
    'Trx',
    'Price Point',
    'Big Size',
    'Traf Pagi',
    'Traf Siang',
    'Total Traf',
    'Conv. Rate',
    '',
    '',
  ]);
  dHeaderSub.height = 25;

  const topNum = dHeaderTop.number;
  const subNum = dHeaderSub.number;

  // Merge rowSpan=2 for Cols 1..4 & 22..23
  [1, 2, 3, 4, 22, 23].forEach((cIdx) => {
    ws.mergeCells(topNum, cIdx, subNum, cIdx);
  });
  ws.mergeCells(topNum, 5, topNum, 8);
  ws.mergeCells(topNum, 9, topNum, 12);
  ws.mergeCells(topNum, 13, topNum, 17);
  ws.mergeCells(topNum, 18, topNum, 21);

  for (let c = 1; c <= 23; c++) {
    const topCell = dHeaderTop.getCell(c);
    const subCell = dHeaderSub.getCell(c);

    const topBg =
      c >= 5 && c <= 8
        ? 'FFB45309'
        : c >= 9 && c <= 12
          ? 'FF3730A3'
          : c >= 13 && c <= 17
            ? 'FFFFC000'
            : c >= 18 && c <= 21
              ? 'FF065F46'
              : 'FFC00000';
    const topFg = c >= 13 && c <= 17 ? 'FF000000' : 'FFFFFFFF';

    topCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: topBg },
    };
    topCell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: topFg },
    };
    topCell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    topCell.border = THIN_BORDER;

    const subBg =
      c >= 5 && c <= 8
        ? 'FFFEF3C7'
        : c >= 9 && c <= 12
          ? 'FFE0E7FF'
          : c === 13
            ? 'FFFDE047'
            : c >= 14 && c <= 17
              ? 'FFFEF9C3'
              : c >= 18 && c <= 21
                ? 'FFD1FAE5'
                : topBg;
    const subFg =
      c >= 5 && c <= 8
        ? 'FF78350F'
        : c >= 9 && c <= 12
          ? 'FF1E1B4B'
          : c >= 13 && c <= 17
            ? 'FF0F172A'
            : c >= 18 && c <= 21
              ? 'FF065F46'
              : 'FFFFFFFF';

    subCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: subBg },
    };
    subCell.font = {
      name: 'Calibri',
      size: 9.5,
      bold: true,
      color: { argb: subFg },
    };
    subCell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    subCell.border = THIN_BORDER;
  }

  const firstDailyDataRow = subNum + 1;

  dailyRecords.forEach((rec) => {
    const r = ws.addRow([
      rec.dayNum,
      rec.tanggal,
      rec.hari,
      rec.weeklyLabel,
      rec.pagiNama,
      rec.pagiSales,
      rec.pagiQty,
      rec.pagiTrx,
      rec.siangNama,
      rec.siangSales,
      rec.siangQty,
      rec.siangTrx,
      null,
      null,
      null,
      null,
      null,
      rec.pagiTraffic,
      rec.siangTraffic,
      null,
      null,
      rec.reasonPagi,
      rec.reasonSiang,
    ]);
    const hasLongReason =
      rec.reasonPagi.length > 45 ||
      rec.reasonSiang.length > 45 ||
      rec.reasonPagi.includes('\n') ||
      rec.reasonSiang.includes('\n');
    r.height = hasLongReason ? 30 : 21;
    const rNum = r.number;

    // Live Excel formulas for Total Harian (Cols 13..17) & Traffic / Conv Rate (Cols 20..21)
    r.getCell(13).value = {
      formula: `F${rNum}+J${rNum}`,
      result: rec.totalSales,
    };
    r.getCell(14).value = {
      formula: `G${rNum}+K${rNum}`,
      result: rec.totalQty,
    };
    r.getCell(15).value = {
      formula: `H${rNum}+L${rNum}`,
      result: rec.totalTrx,
    };
    r.getCell(16).value = {
      formula: `IF(N${rNum}>0,ROUND(M${rNum}/N${rNum},0),0)`,
      result: rec.pricePoint,
    };
    r.getCell(17).value = {
      formula: `IF(O${rNum}>0,ROUND(M${rNum}/O${rNum},0),0)`,
      result: rec.bigSize,
    };
    r.getCell(20).value = {
      formula: `R${rNum}+S${rNum}`,
      result: rec.totalTraffic,
    };
    r.getCell(21).value = {
      formula: `IF(T${rNum}>0,O${rNum}/T${rNum},0)`,
      result: rec.convRateRatio,
    };

    for (let c = 1; c <= 23; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      const isZeroDay = rec.totalSales === 0 && rec.totalTraffic === 0;
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold:
          c === 1 ||
          c === 2 ||
          c === 3 ||
          c === 4 ||
          (c >= 13 && c <= 17) ||
          c === 20 ||
          c === 21,
        color: {
          argb:
            rec.isWeekend && (c === 1 || c === 3)
              ? 'FFBE123C'
              : c === 4
                ? 'FF3730A3'
                : isZeroDay && c >= 5 && c <= 21
                  ? 'FF94A3B8'
                  : 'FF0F172A',
        },
      };

      if (c === 13) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFF9C4' },
        };
      } else if (c >= 14 && c <= 17) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFDE7' },
        };
      } else if (c === 20 || c === 21) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFECFDF5' },
        };
      } else if (rec.isWeekend) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFF1F2' },
        };
      }

      if (c === 21) {
        cell.numFmt = NUM_FMT_PERCENT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20].includes(c)) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 22 || c === 23) {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });

  const lastDailyDataRow = Math.max(
    firstDailyDataRow,
    firstDailyDataRow + dailyRecords.length - 1
  );

  if (enableSheetAutoFilter && dailyRecords.length > 0) {
    ws.autoFilter = {
      from: { row: subNum, column: 1 },
      to: { row: lastDailyDataRow, column: 23 },
    };
  }

  if (dailyRecords.length > 0) {
    const activeDaysCount = dailyRecords.filter((d) => d.totalSales > 0).length;
    const sumPagiSales = dailyRecords.reduce((s, d) => s + d.pagiSales, 0);
    const sumPagiQty = dailyRecords.reduce((s, d) => s + d.pagiQty, 0);
    const sumPagiTrx = dailyRecords.reduce((s, d) => s + d.pagiTrx, 0);
    const sumSiangSales = dailyRecords.reduce((s, d) => s + d.siangSales, 0);
    const sumSiangQty = dailyRecords.reduce((s, d) => s + d.siangQty, 0);
    const sumSiangTrx = dailyRecords.reduce((s, d) => s + d.siangTrx, 0);
    const sumPagiTraffic = dailyRecords.reduce((s, d) => s + d.pagiTraffic, 0);
    const sumSiangTraffic = dailyRecords.reduce((s, d) => s + d.siangTraffic, 0);
    const sumTotalTraffic = sumPagiTraffic + sumSiangTraffic;
    const sumTotalSales = sumPagiSales + sumSiangSales;
    const sumTotalQty = sumPagiQty + sumSiangQty;
    const sumTotalTrx = sumPagiTrx + sumSiangTrx;
    const grandPP = sumTotalQty > 0 ? Math.round(sumTotalSales / sumTotalQty) : 0;
    const grandBS = sumTotalTrx > 0 ? Math.round(sumTotalSales / sumTotalTrx) : 0;
    const grandConvRatio = sumTotalTraffic > 0 ? sumTotalTrx / sumTotalTraffic : 0;

    const grandRow = ws.addRow([
      `GRAND TOTAL HARIAN (${activeDaysCount}/${dailyRecords.length} HARI AKTIF — ${activeMonthMeta.name.toUpperCase()} 2026)`,
      '',
      '',
      '',
      '',
      null,
      null,
      null,
      '',
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      '',
      '',
    ]);
    grandRow.height = 25;
    const gNum = grandRow.number;
    ws.mergeCells(gNum, 1, gNum, 5);
    ws.mergeCells(gNum, 22, gNum, 23);

    const colSumsMap: Record<number, number> = {
      6: sumPagiSales,
      7: sumPagiQty,
      8: sumPagiTrx,
      10: sumSiangSales,
      11: sumSiangQty,
      12: sumSiangTrx,
      13: sumTotalSales,
      14: sumTotalQty,
      15: sumTotalTrx,
      18: sumPagiTraffic,
      19: sumSiangTraffic,
      20: sumTotalTraffic,
    };

    [6, 7, 8, 10, 11, 12, 13, 14, 15, 18, 19, 20].forEach((cIdx) => {
      const letCol = colToLetter(cIdx);
      grandRow.getCell(cIdx).value = {
        formula: `SUM(${letCol}${firstDailyDataRow}:${letCol}${lastDailyDataRow})`,
        result: colSumsMap[cIdx] ?? 0,
      };
    });
    grandRow.getCell(16).value = {
      formula: `IF(N${gNum}>0,ROUND(M${gNum}/N${gNum},0),0)`,
      result: grandPP,
    };
    grandRow.getCell(17).value = {
      formula: `IF(O${gNum}>0,ROUND(M${gNum}/O${gNum},0),0)`,
      result: grandBS,
    };
    grandRow.getCell(21).value = {
      formula: `IF(T${gNum}>0,O${gNum}/T${gNum},0)`,
      result: grandConvRatio,
    };

    for (let c = 1; c <= 23; c++) {
      const cell = grandRow.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      cell.font = {
        name: 'Calibri',
        size: 10.5,
        bold: true,
        color: { argb: 'FF000000' },
      };
      cell.border = MEDIUM_OUTER_BORDER;
      if (c === 21) {
        cell.numFmt = NUM_FMT_PERCENT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20].includes(c)) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }

    // RATA-RATA HARIAN & RATA-RATA TIKET PER HARI (RINGKASAN BULANAN OTOMATIS)
    const avgDivisor = Math.max(1, activeDaysCount);
    const avgTrxPerActiveDay =
      activeDaysCount > 0
        ? Number((sumTotalTrx / activeDaysCount).toFixed(1))
        : 0;
    const avgTrxPerCalDay =
      dailyRecords.length > 0
        ? Number((sumTotalTrx / dailyRecords.length).toFixed(1))
        : 0;

    const avgDailyRow = ws.addRow([
      `RATA-RATA PER HARI AKTIF (${activeDaysCount} HARI TERISI — ${activeMonthMeta.name.toUpperCase()} 2026)`,
      '',
      '',
      '',
      '',
      null,
      null,
      null,
      '',
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      `Kumulatif Sales: Rp ${sumTotalSales.toLocaleString('id-ID')} | Total Transaksi: ${sumTotalTrx} trx | Rata-rata Trx/Hari Aktif: ${avgTrxPerActiveDay} trx/hr (Kalender: ${avgTrxPerCalDay} trx/hr)`,
      '',
    ]);
    avgDailyRow.height = 23;
    const avgDNum = avgDailyRow.number;
    ws.mergeCells(avgDNum, 1, avgDNum, 5);
    ws.mergeCells(avgDNum, 22, avgDNum, 23);

    [6, 7, 8, 10, 11, 12, 13, 14, 15, 18, 19, 20].forEach((cIdx) => {
      const letCol = colToLetter(cIdx);
      const rawSum = colSumsMap[cIdx] ?? 0;
      avgDailyRow.getCell(cIdx).value = {
        formula: `IF(COUNTIF($M$${firstDailyDataRow}:$M$${lastDailyDataRow},">0")>0,ROUND(${letCol}${gNum}/COUNTIF($M$${firstDailyDataRow}:$M$${lastDailyDataRow},">0"),0),0)`,
        result: activeDaysCount > 0 ? Math.round(rawSum / avgDivisor) : 0,
      };
    });
    avgDailyRow.getCell(16).value = {
      formula: `P${gNum}`,
      result: grandPP,
    };
    avgDailyRow.getCell(17).value = {
      formula: `Q${gNum}`,
      result: grandBS,
    };
    avgDailyRow.getCell(21).value = {
      formula: `U${gNum}`,
      result: grandConvRatio,
    };

    for (let c = 1; c <= 23; c++) {
      const cell = avgDailyRow.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F172A' },
      };
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: c >= 13 && c <= 23 ? 'FFFDE047' : 'FFFFFFFF' },
      };
      cell.border = THIN_BORDER;
      if (c === 21) {
        cell.numFmt = NUM_FMT_PERCENT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if ([6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20].includes(c)) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 22 || c === 23) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  }

  // 2) EMBED DIAGRAM SALES HARIAN AT THE BOTTOM BELOW THE TABLE (SESUAI PERMINTAAN & TAMPILAN APLIKASI)
  if (includeChartDiagrams && dailyRecords.length > 0) {
    let dPagi = 0;
    let dSiang = 0;
    let dWkday = 0;
    let dWknd = 0;
    const dailyBars: ChartBarItem[] = dailyRecords.map((rec) => {
      dPagi += rec.pagiSales;
      dSiang += rec.siangSales;
      if (rec.isWeekend) dWknd += rec.totalSales;
      else dWkday += rec.totalSales;
      return {
        label: String(rec.dayNum).padStart(2, '0'),
        subLabel: rec.hari.slice(0, 3),
        pagiSales: rec.pagiSales,
        siangSales: rec.siangSales,
        totalSales: rec.totalSales,
        isWeekend: rec.isWeekend,
      };
    });

    embedSalesDiagramIntoSheet(
      workbook,
      ws,
      `DIAGRAM SALES HARIAN — BULAN ${activeMonthMeta.name.toUpperCase()} 2026 (TANGGAL 1 - ${activeMonthMeta.daysCount})`,
      `Diagram Batang Vertikal ${dailyBars.length} Hari & Garis Tren Penjualan Shift Pagi vs Shift Siang`,
      dailyBars,
      dPagi,
      dSiang,
      dWkday,
      dWknd,
      'SALES_ONLY',
      undefined,
      `1–${activeMonthMeta.daysCount} ${activeMonthMeta.name} 2026`
    );
  }

  ws.addRow([]);
  return subNum;
}

/* ============================================================================
   SECTION RENDERER 3: SALES TABLE BULANAN & REKAP 12 BULAN (PERSIS FORMAT DI APLIKASI - 15 KOLOM)
   ============================================================================ */
function renderMonthlyTableSection(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  tables: WeeklySalesTable[],
  includeChartDiagrams: boolean = true,
  enableSheetAutoFilter: boolean = false,
  selectedMonthKey?: string,
  allTablesPool?: WeeklySalesTable[],
  resolvedActiveMonthKey?: string
): number {
  addSectionBanner(
    ws,
    'TABEL SALES BULANAN (MONTHLY) — REKAPITULASI PER WEEKLY & KALENDER 12 BULAN (TGL 1–30/31) DENGAN KONTRIBUSI & GROWTH',
    16,
    'FF0F172A',
    'FFFDE047'
  );

  const calendarSourceTables =
    allTablesPool && allTablesPool.length > 0 ? allTablesPool : tables;
  const resolvedMonth =
    resolvedActiveMonthKey ||
    resolveLatestActiveMonthKey(selectedMonthKey, calendarSourceTables);

  const byWeekNum = new Map<number, WeeklySalesTable>();
  tables.forEach((t) => byWeekNum.set(t.weekNumber, t));

  let weeklyRowsToRender: WeeklySalesTable[] = [];
  if (tables.length === 1) {
    weeklyRowsToRender = [...tables];
  } else if (selectedMonthKey && selectedMonthKey !== 'ALL') {
    const candidateWeeks = ANNUAL_2026_WEEKS.filter(
      (w) => w.monthKey === selectedMonthKey
    );
    weeklyRowsToRender = candidateWeeks.map(
      (w) =>
        byWeekNum.get(w.weekNumber) ||
        createWeeklyTableFromAnnualSchedule(w)
    );
  } else {
    const sorted = [...tables].sort((a, b) => a.weekNumber - b.weekNumber);
    if (sorted.length > 12 && !sorted.some(isWeeklyTableActive)) {
      // Avoid dumping 53 blank weeks when no DP weekly data has been entered yet
      const activeMonthTables = sorted.filter(
        (t) => getTableMonthKey(t) === resolvedMonth
      );
      weeklyRowsToRender =
        activeMonthTables.length > 0 ? activeMonthTables : sorted.slice(-4);
    } else {
      weeklyRowsToRender = sorted;
    }
  }

  const all12MonthsSummary = buildAll12MonthsCalendarSummary(
    calendarSourceTables,
    2026
  );
  const all12MonthsCal = all12MonthsSummary.months;

  // 1) TABLE 1: MATRIKS REKAPITULASI PER WEEKLY (16 COLUMNS WITH KONTRIBUSI (%) & GROWTH (%))
  // Col 1: Periode Weekly
  // Col 2: Rentang Tanggal
  // Col 3: Hari Aktif
  // Col 4..6: Akumulasi Shift Pagi (Sales Pagi | Qty | Trx)
  // Col 7..9: Akumulasi Shift Siang (Sales Siang | Qty | Trx)
  // Col 10..14: Grand Total Mingguan (Total Sales (Rp) | Total Qty | Total Trx | Price Point | Big Size)
  // Col 15: Kontribusi (%)
  // Col 16: Growth (%)
  const mHeaderTop = ws.addRow([
    'Periode Weekly',
    'Rentang Tanggal',
    'Hari Aktif',
    'Akumulasi Shift Pagi',
    '',
    '',
    'Akumulasi Shift Siang',
    '',
    '',
    'Grand Total Mingguan (Weekly Total)',
    '',
    '',
    '',
    '',
    'Kontribusi (%)',
    'Growth (%)',
  ]);
  mHeaderTop.height = 24;

  const mHeaderSub = ws.addRow([
    '',
    '',
    '',
    'Sales Pagi',
    'Qty',
    'Trx',
    'Sales Siang',
    'Qty',
    'Trx',
    'Total Sales (Rp)',
    'Total Qty',
    'Total Trx',
    'Price Point',
    'Big Size',
    '',
    '',
  ]);
  mHeaderSub.height = 25;

  const mTopNum = mHeaderTop.number;
  const mSubNum = mHeaderSub.number;

  [1, 2, 3, 15, 16].forEach((cIdx) => {
    ws.mergeCells(mTopNum, cIdx, mSubNum, cIdx);
  });
  ws.mergeCells(mTopNum, 4, mTopNum, 6);
  ws.mergeCells(mTopNum, 7, mTopNum, 9);
  ws.mergeCells(mTopNum, 10, mTopNum, 14);

  for (let c = 1; c <= 16; c++) {
    const topCell = mHeaderTop.getCell(c);
    const subCell = mHeaderSub.getCell(c);

    const topBg =
      c >= 4 && c <= 6
        ? 'FFB45309'
        : c >= 7 && c <= 9
          ? 'FF3730A3'
          : c >= 10 && c <= 14
            ? 'FFFFC000'
            : c === 15
              ? 'FF0F172A'
              : c === 16
                ? 'FF065F46'
                : 'FFC00000';
    const topFg =
      c >= 10 && c <= 14
        ? 'FF000000'
        : c === 15 || c === 16
          ? 'FFFDE047'
          : 'FFFFFFFF';

    topCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: topBg },
    };
    topCell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: topFg },
    };
    topCell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    topCell.border = THIN_BORDER;

    const subBg =
      c >= 4 && c <= 6
        ? 'FFFEF3C7'
        : c >= 7 && c <= 9
          ? 'FFE0E7FF'
          : c === 10
            ? 'FFFDE047'
            : c >= 11 && c <= 14
              ? 'FFFEF9C3'
              : topBg;
    const subFg =
      c >= 4 && c <= 6
        ? 'FF78350F'
        : c >= 7 && c <= 9
          ? 'FF1E1B4B'
          : c >= 10 && c <= 14
            ? 'FF0F172A'
            : topFg;

    subCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: subBg },
    };
    subCell.font = {
      name: 'Calibri',
      size: 9.5,
      bold: true,
      color: { argb: subFg },
    };
    subCell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true,
    };
    subCell.border = THIN_BORDER;
  }

  const firstWeekSummaryRow = mSubNum + 1;
  const totalGrandSalesAllWeeks = weeklyRowsToRender.reduce(
    (s, t) => s + calculateWeekMetrics(t).weeklySales,
    0
  );
  const grandTotalWeekRowNum = firstWeekSummaryRow + weeklyRowsToRender.length;

  weeklyRowsToRender.forEach((table, idx) => {
    const wm = calculateWeekMetrics(table);
    let pSales = 0;
    let pQty = 0;
    let pTrx = 0;
    let sSales = 0;
    let sQty = 0;
    let sTrx = 0;
    let activeDays = 0;

    table.days.forEach((d) => {
      const ps = d.pagi.sales ?? 0;
      const ss = d.siang.sales ?? 0;
      pSales += ps;
      pQty += d.pagi.qty ?? 0;
      pTrx += d.pagi.trx ?? 0;
      sSales += ss;
      sQty += d.siang.qty ?? 0;
      sTrx += d.siang.trx ?? 0;
      if (ps + ss > 0) activeDays += 1;
    });

    const prevSales =
      idx > 0 ? calculateWeekMetrics(weeklyRowsToRender[idx - 1]).weeklySales : null;
    const growthRatio =
      prevSales !== null && prevSales > 0
        ? (wm.weeklySales - prevSales) / prevSales
        : 0;

    const wkInfo = getTableWeekendInfo(table);
    const r = ws.addRow([
      `TOTAL WEEKLY ${table.weekNumber}`,
      `${getWeeklyPeriodeText(table)} (${wkInfo.shortBadge} - ${getTableMonthName(table)})`,
      `${activeDays}/${table.days.length} Hari`,
      pSales,
      pQty,
      pTrx,
      sSales,
      sQty,
      sTrx,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
    ]);
    r.height = 22;
    const rNum = r.number;

    // Formulas for Grand Total Mingguan (Cols 10..14), Kontribusi (Col 15), & Growth (Col 16)
    r.getCell(10).value = {
      formula: `D${rNum}+G${rNum}`,
      result: wm.weeklySales,
    };
    r.getCell(11).value = {
      formula: `E${rNum}+H${rNum}`,
      result: wm.weeklyQty,
    };
    r.getCell(12).value = {
      formula: `F${rNum}+I${rNum}`,
      result: wm.weeklyTrx,
    };
    r.getCell(13).value = {
      formula: `IF(K${rNum}>0,ROUND(J${rNum}/K${rNum},0),0)`,
      result: wm.weeklyPricePoint ?? 0,
    };
    r.getCell(14).value = {
      formula: `IF(L${rNum}>0,ROUND(J${rNum}/L${rNum},0),0)`,
      result: wm.weeklyBigSize ?? 0,
    };
    r.getCell(15).value = {
      formula: `IF($J$${grandTotalWeekRowNum}>0,J${rNum}/$J$${grandTotalWeekRowNum},0)`,
      result:
        totalGrandSalesAllWeeks > 0
          ? wm.weeklySales / totalGrandSalesAllWeeks
          : 0,
    };
    if (idx === 0) {
      r.getCell(16).value = 0;
    } else {
      const prevRowNum = rNum - 1;
      r.getCell(16).value = {
        formula: `IF(J${prevRowNum}>0,(J${rNum}-J${prevRowNum})/J${prevRowNum},0)`,
        result: growthRatio,
      };
    }

    for (let c = 1; c <= 16; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold:
          c === 1 ||
          c === 4 ||
          c === 7 ||
          (c >= 10 && c <= 16),
        color: {
          argb:
            c === 4
              ? 'FF92400E'
              : c === 7
                ? 'FF3730A3'
                : c === 16
                  ? growthRatio > 0
                    ? 'FF047857'
                    : growthRatio < 0
                      ? 'FFBE123C'
                      : 'FF0F172A'
                  : 'FF0F172A',
        },
      };
      if (c === 10) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFF9C4' },
        };
      } else if (c >= 11 && c <= 14) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFDE7' },
        };
      } else if (c === 16) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb:
              growthRatio > 0
                ? 'FFECFDF5'
                : growthRatio < 0
                  ? 'FFFFF1F2'
                  : 'FFF8FAFC',
          },
        };
      }
      if (c >= 4 && c <= 14) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 15 || c === 16) {
        cell.numFmt = NUM_FMT_PERCENT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 2) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });

  const lastWeekSummaryRow = Math.max(
    firstWeekSummaryRow,
    firstWeekSummaryRow + weeklyRowsToRender.length - 1
  );

  if (enableSheetAutoFilter && weeklyRowsToRender.length > 0) {
    ws.autoFilter = {
      from: { row: mSubNum, column: 1 },
      to: { row: lastWeekSummaryRow, column: 16 },
    };
  }

  const totalActiveDaysWeeks = weeklyRowsToRender.reduce(
    (s, t) =>
      s +
      t.days.filter((d) => (d.pagi.sales ?? 0) + (d.siang.sales ?? 0) > 0)
        .length,
    0
  );
  const totalCalendarDaysWeeks = weeklyRowsToRender.reduce(
    (s, t) => s + t.days.length,
    0
  );

  // GRAND TOTAL BULANAN Row (matching app)
  const mTotalRow = ws.addRow([
    `GRAND TOTAL BULANAN (${weeklyRowsToRender.length} MINGGU)`,
    '',
    `${totalActiveDaysWeeks}/${totalCalendarDaysWeeks} Hari`,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    '-',
  ]);
  mTotalRow.height = 25;
  const mtNum = mTotalRow.number;
  ws.mergeCells(mtNum, 1, mtNum, 2);

  [4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((cIdx) => {
    const letCol = colToLetter(cIdx);
    mTotalRow.getCell(cIdx).value = {
      formula: `SUM(${letCol}${firstWeekSummaryRow}:${letCol}${lastWeekSummaryRow})`,
    };
  });
  mTotalRow.getCell(13).value = {
    formula: `IF(K${mtNum}>0,ROUND(J${mtNum}/K${mtNum},0),0)`,
  };
  mTotalRow.getCell(14).value = {
    formula: `IF(L${mtNum}>0,ROUND(J${mtNum}/L${mtNum},0),0)`,
  };
  mTotalRow.getCell(15).value = {
    formula: `SUM(O${firstWeekSummaryRow}:O${lastWeekSummaryRow})`,
    result: 1,
  };

  for (let c = 1; c <= 16; c++) {
    const cell = mTotalRow.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFC000' },
    };
    cell.font = { name: 'Calibri', size: 10.5, bold: true };
    cell.border = MEDIUM_OUTER_BORDER;
    if (c >= 4 && c <= 14) {
      cell.numFmt = NUM_FMT_INTEGER;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c === 15) {
      cell.numFmt = NUM_FMT_PERCENT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }

  // RATA-RATA PER MINGGU (AVERAGE WEEKLY) Row (matching MonthlySalesSpreadsheetView in app!)
  const avgWeekRow = ws.addRow([
    'RATA-RATA PER MINGGU (AVERAGE WEEKLY)',
    '',
    'Per Minggu',
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    `${weeklyRowsToRender.length > 0 ? Math.round(100 / weeklyRowsToRender.length) : 0}% / Wk`,
    '-',
  ]);
  avgWeekRow.height = 22;
  const avgNum = avgWeekRow.number;
  ws.mergeCells(avgNum, 1, avgNum, 2);

  [4, 5, 6, 7, 8, 9, 10, 11, 12].forEach((cIdx) => {
    const letCol = colToLetter(cIdx);
    avgWeekRow.getCell(cIdx).value = {
      formula: `IF(${weeklyRowsToRender.length}>0,ROUND(${letCol}${mtNum}/${weeklyRowsToRender.length},0),0)`,
    };
  });
  avgWeekRow.getCell(13).value = {
    formula: `M${mtNum}`,
  };
  avgWeekRow.getCell(14).value = {
    formula: `N${mtNum}`,
  };

  for (let c = 1; c <= 16; c++) {
    const cell = avgWeekRow.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: c === 10 || c === 15 || c === 16 ? 'FFFDE047' : 'FFFFFFFF' },
    };
    cell.border = THIN_BORDER;
    if (c >= 4 && c <= 14) {
      cell.numFmt = NUM_FMT_INTEGER;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }

  ws.addRow([]);

  // 2) TABLE 2: REKAPITULASI KALENDER 12 BULAN (TANGGAL 1 - 30/31 TIAP BULAN — TAHUN 2026) MATCHING APP (12 COLUMNS WITH KONTRIBUSI & GROWTH)
  addSectionBanner(
    ws,
    'REKAPITULASI KALENDER 12 BULAN (TANGGAL 1 - 30/31 PER BULAN — TAHUN 2026) DENGAN KONTRIBUSI & GROWTH',
    12,
    'FF0F172A',
    'FFFDE047'
  );

  const m12HeaderRow = ws.addRow([
    'Bulan',
    'Rentang Tanggal (1 - Akhir Bulan)',
    'Hari Terisi',
    'Sales Shift Pagi',
    'Sales Shift Siang',
    'Total Sales Bulanan (Tgl 1-30/31)',
    'Total Qty',
    'Total Trx',
    'Price Point',
    'Big Size',
    'Kontribusi (%)',
    'Growth (%)',
  ]);
  m12HeaderRow.height = 25;
  for (let c = 1; c <= 12; c++) {
    const cell = m12HeaderRow.getCell(c);
    const bg =
      c === 4
        ? 'FFB45309'
        : c === 5
          ? 'FF3730A3'
          : c === 6
            ? 'FFFFC000'
            : c === 11
              ? 'FF0F172A'
              : c === 12
                ? 'FF065F46'
                : 'FF1E293B';
    const fg =
      c === 6 ? 'FF000000' : c === 11 || c === 12 ? 'FFFDE047' : 'FFFFFFFF';
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: bg },
    };
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: fg },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = THIN_BORDER;
  }

  const firstM12Row = m12HeaderRow.number + 1;
  const m12TotalRowNum = firstM12Row + 12;
  const grandTotal12Sales = all12MonthsSummary.grandTotals.totalSales;

  all12MonthsCal.forEach((mItem, idx) => {
    const pS = mItem.pagiSales;
    const sS = mItem.siangSales;
    const qT = mItem.totalQty;
    const tT = mItem.totalTrx;
    const prevMonthSales = idx > 0 ? all12MonthsCal[idx - 1].totalSales : null;
    const mGrowthRatio =
      prevMonthSales !== null && prevMonthSales > 0
        ? (pS + sS - prevMonthSales) / prevMonthSales
        : 0;

    const r = ws.addRow([
      `${mItem.monthName} 2026`,
      mItem.slashDateRange,
      `${mItem.activeDays}/${mItem.lastDay} Hari`,
      pS,
      sS,
      null,
      qT,
      tT,
      null,
      null,
      null,
      null,
    ]);
    r.height = 21;
    const rNum = r.number;
    r.getCell(6).value = { formula: `D${rNum}+E${rNum}`, result: pS + sS };
    r.getCell(9).value = {
      formula: `IF(G${rNum}>0,ROUND(F${rNum}/G${rNum},0),0)`,
      result: qT > 0 ? Math.round((pS + sS) / qT) : 0,
    };
    r.getCell(10).value = {
      formula: `IF(H${rNum}>0,ROUND(F${rNum}/H${rNum},0),0)`,
      result: tT > 0 ? Math.round((pS + sS) / tT) : 0,
    };
    r.getCell(11).value = {
      formula: `IF($F$${m12TotalRowNum}>0,F${rNum}/$F$${m12TotalRowNum},0)`,
      result: grandTotal12Sales > 0 ? (pS + sS) / grandTotal12Sales : 0,
    };
    if (idx === 0) {
      r.getCell(12).value = 0;
    } else {
      const prevRowNum = rNum - 1;
      r.getCell(12).value = {
        formula: `IF(F${prevRowNum}>0,(F${rNum}-F${prevRowNum})/F${prevRowNum},0)`,
        result: mGrowthRatio,
      };
    }

    for (let c = 1; c <= 12; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: c === 1 || c === 6 || c === 11 || c === 12,
        color: {
          argb:
            c === 4
              ? 'FF92400E'
              : c === 5
                ? 'FF3730A3'
                : c === 12
                  ? mGrowthRatio > 0
                    ? 'FF047857'
                    : mGrowthRatio < 0
                      ? 'FFBE123C'
                      : 'FF0F172A'
                  : 'FF0F172A',
        },
      };
      if (c === 6) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFF9C4' },
        };
      }
      if (c >= 4 && c <= 10) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 11 || c === 12) {
        cell.numFmt = NUM_FMT_PERCENT;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    }
  });

  const lastM12Row = firstM12Row + 11;
  const totalActiveDays12 = all12MonthsSummary.grandTotals.activeDays;
  const m12Total = ws.addRow([
    'GRAND TOTAL TAHUNAN (12 BULAN — TGL 1 S/D 30/31)',
    '',
    `${totalActiveDays12}/${all12MonthsSummary.grandTotals.totalCalendarDays} Hari`,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    null,
    '-',
  ]);
  m12Total.height = 25;
  const m12TNum = m12Total.number;
  ws.mergeCells(m12TNum, 1, m12TNum, 2);

  [4, 5, 6, 7, 8].forEach((cIdx) => {
    const letCol = colToLetter(cIdx);
    m12Total.getCell(cIdx).value = {
      formula: `SUM(${letCol}${firstM12Row}:${letCol}${lastM12Row})`,
    };
  });
  m12Total.getCell(9).value = {
    formula: `IF(G${m12TNum}>0,ROUND(F${m12TNum}/G${m12TNum},0),0)`,
  };
  m12Total.getCell(10).value = {
    formula: `IF(H${m12TNum}>0,ROUND(F${m12TNum}/H${m12TNum},0),0)`,
  };
  m12Total.getCell(11).value = {
    formula: `SUM(K${firstM12Row}:K${lastM12Row})`,
    result: 1,
  };

  for (let c = 1; c <= 12; c++) {
    const cell = m12Total.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFC000' },
    };
    cell.font = { name: 'Calibri', size: 10.5, bold: true };
    cell.border = MEDIUM_OUTER_BORDER;
    if (c >= 4 && c <= 10) {
      cell.numFmt = NUM_FMT_INTEGER;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else if (c === 11) {
      cell.numFmt = NUM_FMT_PERCENT;
      cell.alignment = { vertical: 'middle', horizontal: 'right' };
    } else {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }
  }

  // 3) EMBED DIAGRAM SALES BULANAN AT THE BOTTOM BELOW THE TABLES
  if (includeChartDiagrams) {
    let mPagi = 0;
    let mSiang = 0;
    let mWkday = 0;
    let mWknd = 0;
    const weeklyBars: ChartBarItem[] = weeklyRowsToRender.map((t) => {
      let ps = 0;
      let ss = 0;
      t.days.forEach((d) => {
        const p = d.pagi.sales ?? 0;
        const s = d.siang.sales ?? 0;
        ps += p;
        ss += s;
        if (isWeekendDate(d.tanggal)) mWknd += p + s;
        else mWkday += p + s;
      });
      mPagi += ps;
      mSiang += ss;
      return {
        label: `W${t.weekNumber}`,
        subLabel: getWeeklyPeriodeText(t).slice(0, 6),
        pagiSales: ps,
        siangSales: ss,
        totalSales: ps + ss,
      };
    });

    const monthly12Bars: ChartBarItem[] = all12MonthsCal.map((mItem) => ({
      label: mItem.monthName.slice(0, 3),
      subLabel: `1–${mItem.lastDay}`,
      pagiSales: mItem.pagiSales,
      siangSales: mItem.siangSales,
      totalSales: mItem.totalSales,
    }));

    embedSalesDiagramIntoSheet(
      workbook,
      ws,
      'DIAGRAM SALES BULANAN — REKAPITULASI PER PERIODE WEEKLY',
      `Diagram Batang Vertikal ${weeklyRowsToRender.length} Periode Weekly & Garis Tren Penjualan Shift Pagi vs Shift Siang`,
      weeklyBars,
      mPagi,
      mSiang,
      mWkday,
      mWknd,
      'MONTHLY_WITH_CONTRIBUTION',
      monthly12Bars,
      `${weeklyRowsToRender.length} PERIODE WEEKLY`
    );
  }

  ws.addRow([]);
  return mSubNum;
}

/* ============================================================================
   SECTION RENDERER 4: EVALUASI REASON & EVENT SHIFT PAGI vs SIANG (PERSIS FORMAT DI APLIKASI - 10 KOLOM PER WEEKLY)
   ============================================================================ */
function renderReasonEvaluationSection(
  ws: ExcelJS.Worksheet,
  tables: WeeklySalesTable[],
  resolvedActiveMonthKey?: string
): number {
  const orderedTables = (() => {
    const base = [...tables];
    if (base.length > 12 && !base.some(isWeeklyTableActive)) {
      const mKey =
        resolvedActiveMonthKey || resolveLatestActiveMonthKey('ALL', base);
      const monthMatched = base.filter((t) => getTableMonthKey(t) === mKey);
      if (monthMatched.length > 0) {
        return monthMatched.sort((a, b) => b.weekNumber - a.weekNumber);
      }
      return base.slice(-4).sort((a, b) => b.weekNumber - a.weekNumber);
    }
    return base.sort((a, b) => b.weekNumber - a.weekNumber);
  })();

  addSectionBanner(
    ws,
    'EVALUASI REASON / EVENT — PERBANDINGAN DETAIL SHIFT PAGI vs SHIFT SIANG',
    10,
    'FF78350F',
    'FFFDE047'
  );

  // Calculate overall Shift Pagi vs Shift Siang comparison summary (matching top comparison cards in ReasonEventEvaluationView)
  let grandPagiSales = 0;
  let grandPagiQty = 0;
  let grandPagiTrx = 0;
  let grandSiangSales = 0;
  let grandSiangQty = 0;
  let grandSiangTrx = 0;

  orderedTables.forEach((t) => {
    t.days.forEach((d) => {
      grandPagiSales += d.pagi.sales ?? 0;
      grandPagiQty += d.pagi.qty ?? 0;
      grandPagiTrx += d.pagi.trx ?? 0;
      grandSiangSales += d.siang.sales ?? 0;
      grandSiangQty += d.siang.qty ?? 0;
      grandSiangTrx += d.siang.trx ?? 0;
    });
  });
  const grandAllSales = grandPagiSales + grandSiangSales;
  const pagiShare =
    grandAllSales > 0 ? Math.round((grandPagiSales / grandAllSales) * 100) : 0;
  const siangShare = grandAllSales > 0 ? 100 - pagiShare : 0;

  const summaryBanner = ws.addRow([
    'RINGKASAN SHIFT PAGI',
    `Rp ${grandPagiSales.toLocaleString('id-ID')} (${pagiShare}%)`,
    `Qty: ${grandPagiQty.toLocaleString('id-ID')} pcs | Trx: ${grandPagiTrx.toLocaleString('id-ID')}`,
    '',
    '',
    'RINGKASAN SHIFT SIANG',
    `Rp ${grandSiangSales.toLocaleString('id-ID')} (${siangShare}%)`,
    `Qty: ${grandSiangQty.toLocaleString('id-ID')} pcs | Trx: ${grandSiangTrx.toLocaleString('id-ID')}`,
    'TOTAL GABUNGAN',
    `Rp ${grandAllSales.toLocaleString('id-ID')}`,
  ]);
  summaryBanner.height = 24;
  ws.mergeCells(summaryBanner.number, 3, summaryBanner.number, 5);
  for (let c = 1; c <= 10; c++) {
    const cell = summaryBanner.getCell(c);
    cell.border = THIN_BORDER;
    const isPagi = c >= 1 && c <= 5;
    const isSiang = c >= 6 && c <= 8;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: isPagi ? 'FFFEF3C7' : isSiang ? 'FFE0E7FF' : 'FFFFC000',
      },
    };
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: {
        argb: isPagi ? 'FF78350F' : isSiang ? 'FF1E1B4B' : 'FF000000',
      },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  const gapTop = ws.addRow([]);
  gapTop.height = 8;

  let firstReasonHeaderRowNum: number | null = null;
  for (const t of orderedTables) {
    const periodeLabel = getWeeklyPeriodeText(t);
    const wkInfo = getTableWeekendInfo(t);
    const mName = getTableMonthName(t);

    // Per-Weekly Dark Slate Banner matching ReasonEventEvaluationView in the app (with filter sign ▼ & Weekly 1..53 dropdown)
    const activeReasonKey = `TOTAL WEEKLY ${t.weekNumber} ▼`;
    const wkBanner = ws.addRow([
      activeReasonKey,
      `Periode: ${periodeLabel} (${wkInfo.shortBadge} - ${mName} 2026)`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
    wkBanner.height = 24;
    ws.mergeCells(wkBanner.number, 2, wkBanner.number, 10);
    for (let c = 1; c <= 10; c++) {
      const cell = wkBanner.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: c === 1 ? 'FF1E293B' : 'FF0F172A' },
      };
      cell.font = {
        name: 'Calibri',
        size: 10.5,
        bold: true,
        color: { argb: 'FFFDE047' },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: c === 1 ? 'center' : 'left',
        indent: c > 1 ? 1 : 0,
      };
      cell.border =
        c === 1
          ? {
              top: { style: 'medium', color: { argb: 'FFFFC000' } },
              left: { style: 'medium', color: { argb: 'FFFFC000' } },
              bottom: { style: 'medium', color: { argb: 'FFFFC000' } },
              right: { style: 'medium', color: { argb: 'FFFFC000' } },
            }
          : THIN_BORDER;
    }
    wkBanner.getCell(1).dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: ['$Z$2:$Z$54'],
      showErrorMessage: false,
      promptTitle: 'Filter Pilih Weekly (1 s/d 53)',
      prompt: 'Pilih TOTAL WEEKLY 1 s/d 53.',
    };

    // Exact 10 Columns Header matching ReasonEventEvaluationView in the app
    const rHeader = ws.addRow([
      'No',
      'Hari & Tanggal',
      'Petugas Pagi',
      'Sales / Qty / Trx Pagi',
      'Reason / Event Shift Pagi',
      'Petugas Siang',
      'Sales / Qty / Trx Siang',
      'Reason / Event Shift Siang',
      'Total Sales Harian',
      'Selisih & Dominasi Shift',
    ]);
    rHeader.height = 26;
    if (firstReasonHeaderRowNum === null) {
      firstReasonHeaderRowNum = rHeader.number;
    }
    for (let c = 1; c <= 10; c++) {
      const cell = rHeader.getCell(c);
      const bg =
        c === 1 || c === 2
          ? 'FFC00000'
          : c >= 3 && c <= 5
            ? 'FFB45309'
            : c >= 6 && c <= 8
              ? 'FF3730A3'
              : c === 9
                ? 'FFFFC000'
                : 'FF1E293B';
      const fg = c === 9 ? 'FF000000' : 'FFFFFFFF';
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bg },
      };
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: fg },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true,
      };
      cell.border = THIN_BORDER;
    }

    let wPagiSales = 0;
    let wPagiQty = 0;
    let wPagiTrx = 0;
    let wSiangSales = 0;
    let wSiangQty = 0;
    let wSiangTrx = 0;

    t.days.forEach((d, idx) => {
      const dm = calculateDayMetrics(d);
      const dName = getIndoDayName(d.tanggal);
      const wknd = isWeekendDate(d.tanggal);
      const pSales = d.pagi.sales ?? 0;
      const pQty = d.pagi.qty ?? 0;
      const pTrx = d.pagi.trx ?? 0;
      const sSales = d.siang.sales ?? 0;
      const sQty = d.siang.qty ?? 0;
      const sTrx = d.siang.trx ?? 0;

      wPagiSales += pSales;
      wPagiQty += pQty;
      wPagiTrx += pTrx;
      wSiangSales += sSales;
      wSiangQty += sQty;
      wSiangTrx += sTrx;

      const pShare =
        dm.totalSales > 0 ? Math.round((pSales / dm.totalSales) * 100) : 0;
      const sShare = dm.totalSales > 0 ? 100 - pShare : 0;
      const diff = Math.abs(sSales - pSales);
      const domLabel =
        dm.totalSales === 0
          ? '-'
          : sSales > pSales
            ? `Siang Unggul (+Rp ${diff.toLocaleString('id-ID')})`
            : pSales > sSales
              ? `Pagi Unggul (+Rp ${diff.toLocaleString('id-ID')})`
              : 'Seimbang';

      const r = ws.addRow([
        idx + 1,
        `Hari ${idx + 1} - ${dName}${wknd ? ' (WEEKEND)' : ''}\n${d.tanggal}`,
        d.pagi.nama || '-',
        `Rp ${pSales.toLocaleString('id-ID')}\nQty: ${pQty} | Trx: ${pTrx} (${pShare}%)`,
        d.reasonPagi || '-',
        d.siang.nama || '-',
        `Rp ${sSales.toLocaleString('id-ID')}\nQty: ${sQty} | Trx: ${sTrx} (${sShare}%)`,
        d.reasonSiang || '-',
        dm.totalSales,
        domLabel,
      ]);
      r.height = 34;

      for (let c = 1; c <= 10; c++) {
        const cell = r.getCell(c);
        cell.border = THIN_BORDER;
        cell.font = {
          name: 'Calibri',
          size: 9.5,
          bold: c <= 4 || c === 6 || c === 7 || c >= 9,
          color: {
            argb:
              wknd && c === 2
                ? 'FFBE123C'
                : c === 3 || c === 4
                  ? 'FF78350F'
                  : c === 6 || c === 7
                    ? 'FF1E1B4B'
                    : c === 10
                      ? sSales > pSales
                        ? 'FF3730A3'
                        : pSales > sSales
                          ? 'FF92400E'
                          : 'FF334155'
                      : 'FF0F172A',
          },
        };

        if (c === 5) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFBEB' },
          };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'left',
            wrapText: true,
          };
        } else if (c === 8) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEEF2FF' },
          };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'left',
            wrapText: true,
          };
        } else if (c === 9) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFF9C4' },
          };
          cell.numFmt = NUM_FMT_INTEGER;
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else if (c === 4 || c === 7) {
          cell.alignment = {
            vertical: 'middle',
            horizontal: 'right',
            wrapText: true,
          };
        } else {
          if (wknd && c <= 2) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFF1F2' },
            };
          }
          cell.alignment = {
            vertical: 'middle',
            horizontal: 'center',
            wrapText: true,
          };
        }
      }
    });

    const wTotalSales = wPagiSales + wSiangSales;
    const wDiff = Math.abs(wSiangSales - wPagiSales);
    const wDom =
      wTotalSales === 0
        ? '-'
        : wSiangSales >= wPagiSales
          ? `Shift Siang Dominan (+Rp ${wDiff.toLocaleString('id-ID')})`
          : `Shift Pagi Dominan (+Rp ${wDiff.toLocaleString('id-ID')})`;

    const wFoot = ws.addRow([
      `TOTAL ${t.title}`,
      '',
      '',
      `Rp ${wPagiSales.toLocaleString('id-ID')}\nQty: ${wPagiQty} | Trx: ${wPagiTrx}`,
      `Total Sales Shift Pagi (${wTotalSales > 0 ? Math.round((wPagiSales / wTotalSales) * 100) : 0}%)`,
      '',
      `Rp ${wSiangSales.toLocaleString('id-ID')}\nQty: ${wSiangQty} | Trx: ${wSiangTrx}`,
      `Total Sales Shift Siang (${wTotalSales > 0 ? Math.round((wSiangSales / wTotalSales) * 100) : 0}%)`,
      wTotalSales,
      wDom,
    ]);
    wFoot.height = 28;
    ws.mergeCells(wFoot.number, 1, wFoot.number, 3);

    for (let c = 1; c <= 10; c++) {
      const cell = wFoot.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC000' },
      };
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: 'FF000000' },
      };
      cell.border = MEDIUM_OUTER_BORDER;
      if (c === 9) {
        cell.numFmt = NUM_FMT_INTEGER;
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (c === 4 || c === 7) {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'right',
          wrapText: true,
        };
      } else {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true,
        };
      }
    }

    const wkGap = ws.addRow([]);
    wkGap.height = 10;
  }

  ws.addRow([]);
  return firstReasonHeaderRowNum ?? 9;
}

/* ============================================================================
   SECTION RENDERER 5: REVIEW MONITORING BAZAR MATAHARI TANJUNG (DENGAN FILTER WEEKLY)
   ============================================================================ */
function resolveReviewDocsForExport(
  tables: WeeklySalesTable[],
  primaryReviewDoc: ReviewMonitoringDoc,
  allReviewDocs?: ReviewMonitoringDoc[],
  selectedReviewWeekFilter?: string,
  selectedMonthKey?: string
): {
  docsToRender: ReviewMonitoringDoc[];
  weeklyFilterSummary: string;
} {
  const pool: ReviewMonitoringDoc[] =
    allReviewDocs && allReviewDocs.length > 0
      ? [...allReviewDocs]
      : [primaryReviewDoc];

  // Ensure primaryReviewDoc is in pool
  if (primaryReviewDoc && !pool.some((d) => d.id === primaryReviewDoc.id)) {
    pool.unshift(primaryReviewDoc);
  }

  const filterMode = selectedReviewWeekFilter || 'FOLLOW_WEEKLY';

  // 1) Explicit Weekly number filter like 'w-41' or 'weekly-41'
  if (
    filterMode.startsWith('w-') ||
    filterMode.startsWith('weekly-') ||
    /^\d+$/.test(filterMode)
  ) {
    const wkNum = Number(filterMode.replace(/[^0-9]/g, ''));
    if (wkNum >= 1 && wkNum <= 53) {
      const existing = pool.find((d) => getReviewDocWeekNumber(d) === wkNum);
      const targetDoc =
        existing || createReviewDocForWeek(wkNum, primaryReviewDoc);
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum);
      const label = sched
        ? `WEEKLY ${wkNum} (${sched.periodeText}) — Bulan ${sched.monthName} 2026`
        : getReviewDocWeeklyLabel(targetDoc);
      return {
        docsToRender: [targetDoc],
        weeklyFilterSummary: label,
      };
    }
  }

  // 2) Specific Review Doc ID
  const matchedById = pool.find((d) => d.id === filterMode);
  if (matchedById) {
    return {
      docsToRender: [matchedById],
      weeklyFilterSummary: getReviewDocWeeklyLabel(matchedById),
    };
  }

  // 3) 'ALL' Review Docs (optionally filtered by selectedMonthKey)
  if (filterMode === 'ALL') {
    const docMap = new Map<number, ReviewMonitoringDoc>();
    for (const d of pool) {
      const wk = getReviewDocWeekNumber(d);
      if (wk !== null) docMap.set(wk, d);
    }
    for (const t of tables) {
      if (!docMap.has(t.weekNumber)) {
        docMap.set(
          t.weekNumber,
          createReviewDocForWeek(t.weekNumber, primaryReviewDoc)
        );
      }
    }
    let list = Array.from(docMap.values());
    if (selectedMonthKey && selectedMonthKey !== 'ALL') {
      const monthFiltered = list.filter(
        (d) => getReviewDocMonthKey(d) === selectedMonthKey
      );
      if (monthFiltered.length > 0) {
        list = monthFiltered;
      }
    }
    list.sort(
      (a, b) =>
        (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
    );
    const monthLabel =
      selectedMonthKey && selectedMonthKey !== 'ALL'
        ? `Bulan ${INDO_MONTHS[selectedMonthKey] || selectedMonthKey} 2026`
        : 'Semua Periode Weekly';
    return {
      docsToRender: list,
      weeklyFilterSummary: `${monthLabel} (${list.length} Laporan Review Weekly: ${list
        .map((d) => {
          const w = getReviewDocWeekNumber(d);
          return w ? `W${w}` : d.periode;
        })
        .join(', ')})`,
    };
  }

  // 4) 'FOLLOW_WEEKLY' — Follow filtered Weekly tables or active primaryReviewDoc
  if (tables.length === 1) {
    const singleWeekNum = tables[0].weekNumber;
    const matchedWeekDoc = pool.find(
      (d) => getReviewDocWeekNumber(d) === singleWeekNum
    );
    const docToUse =
      matchedWeekDoc ||
      (getReviewDocWeekNumber(primaryReviewDoc) === singleWeekNum
        ? primaryReviewDoc
        : createReviewDocForWeek(singleWeekNum, primaryReviewDoc));
    return {
      docsToRender: [docToUse],
      weeklyFilterSummary: `WEEKLY ${singleWeekNum} (${getWeeklyPeriodeText(tables[0])}) — Bulan ${getTableMonthName(tables[0])} 2026`,
    };
  }

  // Multiple tables selected: ensure every table in `tables` (plus any filled review docs in `pool`) is represented
  const sortedTables = [...tables].sort((a, b) => a.weekNumber - b.weekNumber);
  const docByWeekMap = new Map<number, ReviewMonitoringDoc>();
  sortedTables.forEach((t) => {
    const existing = pool.find((d) => getReviewDocWeekNumber(d) === t.weekNumber);
    docByWeekMap.set(
      t.weekNumber,
      existing || createReviewDocForWeek(t.weekNumber, primaryReviewDoc)
    );
  });
  pool.forEach((d) => {
    const wk = getReviewDocWeekNumber(d);
    if (wk !== null && hasReviewDocContent(d) && !docByWeekMap.has(wk)) {
      docByWeekMap.set(wk, d);
    }
  });
  const finalDocs = Array.from(docByWeekMap.values()).sort(
    (a, b) =>
      (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
  );

  const weeksSummary = finalDocs
    .map((d) => {
      const w = getReviewDocWeekNumber(d);
      return w ? `WEEKLY ${w}` : d.periode;
    })
    .join(', ');

  return {
    docsToRender: finalDocs,
    weeklyFilterSummary: `Semua Weekly Terpilih (${weeksSummary})`,
  };
}

interface ReviewLookupSlot {
  slotKey: string;
  secId: string;
  secCode: string;
  subId?: string;
  subLabel?: string;
}

function buildReviewLookupSlots(
  docs: ReviewMonitoringDoc[]
): ReviewLookupSlot[] {
  const slots: ReviewLookupSlot[] = [];
  const seen = new Set<string>();

  for (const doc of docs) {
    for (const sec of doc.sections || []) {
      if (sec.subRows && sec.subRows.length > 0) {
        for (const sub of sec.subRows) {
          const slotKey = `SUB::${sec.code.trim().toUpperCase()}::${sub.label.trim().toUpperCase()}`;
          if (!seen.has(slotKey)) {
            seen.add(slotKey);
            slots.push({
              slotKey,
              secId: sec.id,
              secCode: sec.code,
              subId: sub.id,
              subLabel: sub.label,
            });
          }
        }
      } else {
        const slotKey = `DIR::${sec.code.trim().toUpperCase()}`;
        if (!seen.has(slotKey)) {
          seen.add(slotKey);
          slots.push({
            slotKey,
            secId: sec.id,
            secCode: sec.code,
          });
        }
      }
    }
  }
  return slots;
}

function getReviewSlotFormattedText(
  doc: ReviewMonitoringDoc,
  slot: ReviewLookupSlot
): string {
  const sec = (doc.sections || []).find(
    (s) =>
      s.id === slot.secId ||
      s.code.trim().toUpperCase() === slot.secCode.trim().toUpperCase()
  );
  if (!sec) return '-';

  if (slot.subLabel !== undefined) {
    const sub = (sec.subRows || []).find(
      (sr) =>
        sr.id === slot.subId ||
        sr.label.trim().toUpperCase() === slot.subLabel!.trim().toUpperCase()
    );
    if (!sub || !Array.isArray(sub.items)) return '-';
    const joined = sub.items
      .map((it) => formatReviewBulletItem(it))
      .filter(Boolean)
      .join('\n');
    return joined || '-';
  }

  if (!Array.isArray(sec.directItems)) return '-';
  const joinedDirect = sec.directItems
    .map((it) => formatReviewBulletItem(it))
    .filter(Boolean)
    .join('\n');
  return joinedDirect || '-';
}

function buildReviewLookupRowValues(
  wk: number,
  allTables: WeeklySalesTable[],
  primaryReviewDoc: ReviewMonitoringDoc,
  poolDocs: ReviewMonitoringDoc[],
  slots: ReviewLookupSlot[],
  withFilterArrow: boolean
): string[] {
  const existingDoc = poolDocs.find((d) => getReviewDocWeekNumber(d) === wk);
  const docForWk =
    existingDoc || createReviewDocForWeek(wk, primaryReviewDoc, false);

  const arrowSuffix = withFilterArrow ? ' ▼' : '';
  const keyLabel = `WEEKLY ${wk}${arrowSuffix}`;
  const nextWk = wk < 53 ? wk + 1 : 53;
  const prevWk = wk > 1 ? wk - 1 : 1;
  const nextKeyLabel = `WEEKLY ${nextWk} ▼`;
  const prevKeyLabel = `WEEKLY ${prevWk} ▼`;

  const weeklyLabel = getReviewDocWeeklyLabel(docForWk);
  const matchedTable = allTables.find((t) => t.weekNumber === wk);
  const weekTotalSales = matchedTable
    ? calculateWeekMetrics(matchedTable).weeklySales
    : 0;
  const totalSalesInfo =
    matchedTable && weekTotalSales > 0
      ? ` — TOTAL SALES WEEKLY: Rp ${Math.round(weekTotalSales).toLocaleString('id-ID')}`
      : '';
  const bannerText = `${docForWk.title} — ${weeklyLabel}${totalSalesInfo}`;

  // Col 1 (DK): KeyLabel ("WEEKLY 1 ▼" .. "WEEKLY 53 ▼")
  // Col 2 (DL): NextSequentialKeyLabel ("WEEKLY 2 ▼" .. "WEEKLY 53 ▼")
  // Col 3 (DM): PrevSequentialKeyLabel ("WEEKLY 1 ▼" .. "WEEKLY 52 ▼")
  // Col 4 (DN): BannerText
  // Cols 5.. (DO..): Slot formatted items
  const rowVals: string[] = [
    keyLabel,
    nextKeyLabel,
    prevKeyLabel,
    bannerText,
  ];

  for (const slot of slots) {
    rowVals.push(getReviewSlotFormattedText(docForWk, slot));
  }

  return rowVals;
}

function populateWorksheetReviewLookupTable(
  ws: ExcelJS.Worksheet,
  allTables: WeeklySalesTable[],
  primaryReviewDoc: ReviewMonitoringDoc,
  poolDocs: ReviewMonitoringDoc[],
  slots: ReviewLookupSlot[],
  startCol: number = 115
): { startColLetter: string; endColLetter: string } {
  const totalCols = 4 + slots.length;
  const startColLetter = colToLetter(startCol);
  const endColLetter = colToLetter(startCol + totalCols - 1);

  for (let c = startCol; c < startCol + totalCols; c++) {
    const col = ws.getColumn(c);
    col.hidden = true;
    col.width = 18;
  }

  // Rows 2..54: "WEEKLY 1 ▼" .. "WEEKLY 53 ▼" (used by Data Validation)
  for (let wk = 1; wk <= 53; wk++) {
    const rowNum = 1 + wk; // 2..54
    const row = ws.getRow(rowNum);
    const vals = buildReviewLookupRowValues(
      wk,
      allTables,
      primaryReviewDoc,
      poolDocs,
      slots,
      true
    );
    for (let idx = 0; idx < vals.length; idx++) {
      row.getCell(startCol + idx).value = vals[idx];
    }
  }

  // Rows 55..107: "WEEKLY 1" .. "WEEKLY 53" (fallback without ▼ symbol)
  for (let wk = 1; wk <= 53; wk++) {
    const rowNum = 54 + wk; // 55..107
    const row = ws.getRow(rowNum);
    const vals = buildReviewLookupRowValues(
      wk,
      allTables,
      primaryReviewDoc,
      poolDocs,
      slots,
      false
    );
    for (let idx = 0; idx < vals.length; idx++) {
      row.getCell(startCol + idx).value = vals[idx];
    }
  }

  return { startColLetter, endColLetter };
}

function renderStaffAndReviewSection(
  ws: ExcelJS.Worksheet,
  tables: WeeklySalesTable[],
  reviewDoc: ReviewMonitoringDoc,
  allReviewDocs?: ReviewMonitoringDoc[],
  selectedReviewWeekFilter?: string,
  selectedMonthKey?: string,
  enableSheetAutoFilter: boolean = false,
  allTablesForLookup?: WeeklySalesTable[],
  resolvedActiveMonthKey?: string
): number {
  const { docsToRender } = resolveReviewDocsForExport(
    tables,
    reviewDoc,
    allReviewDocs,
    selectedReviewWeekFilter,
    selectedMonthKey
  );

  const poolDocs: ReviewMonitoringDoc[] =
    allReviewDocs && allReviewDocs.length > 0
      ? [...allReviewDocs]
      : [reviewDoc];
  if (reviewDoc && !poolDocs.some((d) => d.id === reviewDoc.id)) {
    poolDocs.unshift(reviewDoc);
  }
  const lookupTables =
    allTablesForLookup && allTablesForLookup.length > 0
      ? allTablesForLookup
      : tables;

  const slots = buildReviewLookupSlots([...docsToRender, ...poolDocs]);
  const slotIndexMap = new Map<string, number>();
  slots.forEach((s, idx) => {
    slotIndexMap.set(s.slotKey, 5 + idx);
  });

  // Compute lookup column letters upfront; IMPORTANT: call populateWorksheetReviewLookupTable
  // at the VERY END of renderStaffAndReviewSection AFTER all ws.addRow(...) calls so that
  // visible Review Monitoring Bazar rows start at Row 1 instead of Row 108!
  const lookupStartCol = enableSheetAutoFilter ? 8 : 115; // Col H on dedicated sheet, Col DK on combined sheet
  const totalLookupCols = 4 + slots.length;
  const startColLetter = colToLetter(lookupStartCol);
  const endColLetter = colToLetter(lookupStartCol + totalLookupCols - 1);
  const lookupRangeRef = `$${startColLetter}$2:$${endColLetter}$107`;
  const dropdownRangeRef = `$${startColLetter}$2:$${startColLetter}$54`;

  // On dedicated Review sheet, Rows 1, 2, 3 are the Red / Yellow / Dark Navy official report header + Row 4 blank (Filter Weekly row at top removed!)
  if (enableSheetAutoFilter) {
    addOfficialReportHeader(ws, 6);
  }

  // Main Banner for Review Monitoring Bazar
  addSectionBanner(
    ws,
    `BAGIAN 5: REVIEW MONITORING BAZAR — REPORT BAZAR MATAHARI (KLIK TANDA FILTER ▼ PADA WEEKLY UNTUK PILIH WEEKLY 1 S/D 53)`,
    6,
    'FF064E3B',
    'FFFDE047'
  );

  // Empty separator row before first Weekly Review table
  const sepRow = ws.addRow([]);
  sepRow.height = 15;

  // Render the single interactive Review Monitoring Bazar table prioritizing the LATEST filled/active weekly doc
  // while keeping the full WEEKLY 1..53 dropdown & VLOOKUP database accessible right in Excel.
  const singleDocsToRender = (() => {
    if (docsToRender.length === 0) return [reviewDoc];
    if (docsToRender.length === 1) return [docsToRender[0]];

    const sortedDocs = [...docsToRender].sort(
      (a, b) =>
        (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
    );
    const filledReviews = sortedDocs.filter((d) => hasReviewDocContent(d));
    if (filledReviews.length > 0) {
      return [filledReviews[filledReviews.length - 1]];
    }

    const activeSalesWeeks = new Set(
      tables.filter((t) => isWeeklyTableActive(t)).map((t) => t.weekNumber)
    );
    const matchingActiveSalesDocs = sortedDocs.filter((d) => {
      const w = getReviewDocWeekNumber(d);
      return w !== null && activeSalesWeeks.has(w);
    });
    if (matchingActiveSalesDocs.length > 0) {
      return [matchingActiveSalesDocs[matchingActiveSalesDocs.length - 1]];
    }

    const targetMonth =
      resolvedActiveMonthKey ||
      resolveLatestActiveMonthKey(selectedMonthKey, lookupTables);
    const monthMatched = sortedDocs.filter(
      (d) => getReviewDocMonthKey(d) === targetMonth
    );
    if (monthMatched.length > 0) {
      return [monthMatched[0]];
    }

    return [sortedDocs[sortedDocs.length - 1]];
  })();

  let firstReviewBannerRowNum: number | null = null;
  for (let dIdx = 0; dIdx < singleDocsToRender.length; dIdx++) {
    const doc = singleDocsToRender[dIdx];
    const wkNum = getReviewDocWeekNumber(doc) || 45;
    const weeklyLabel = getReviewDocWeeklyLabel(doc);
    const matchedTable = lookupTables.find((t) => t.weekNumber === wkNum);
    const weekTotalSales = matchedTable
      ? calculateWeekMetrics(matchedTable).weeklySales
      : 0;
    const totalSalesInfo =
      matchedTable && weekTotalSales > 0
        ? ` — TOTAL SALES WEEKLY: Rp ${Math.round(weekTotalSales).toLocaleString('id-ID')}`
        : '';

    const activeReviewKey = `WEEKLY ${wkNum} ▼`;
    const bannerText = `${doc.title} — ${weeklyLabel}${totalSalesInfo}`;

    // Sub-header banner per Weekly Review with interactive WEEKLY 1..53 dropdown & filter sign ▼
    if (dIdx > 0) {
      const gapRow = ws.addRow([]);
      gapRow.height = 14;
    }
    const periodBanner = ws.addRow([
      activeReviewKey,
      bannerText,
      '',
      '',
      '',
      '',
    ]);
    periodBanner.height = 25;
    const pRowNum = periodBanner.number;
    if (firstReviewBannerRowNum === null) {
      firstReviewBannerRowNum = pRowNum;
    }
    ws.mergeCells(pRowNum, 2, pRowNum, 6);

    // Keep pCell1 as a plain string ("WEEKLY 40 ▼", "WEEKLY 41 ▼", etc.) with Data Validation list
    // so there is never a circular formula and WPS Office immediately opens the dropdown on tap.
    const pCell1 = periodBanner.getCell(1);
    pCell1.value = activeReviewKey;

    pCell1.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' },
    };
    pCell1.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: 'FFFDE047' },
    };
    pCell1.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    pCell1.border = {
      top: { style: 'medium', color: { argb: 'FFFFC000' } },
      left: { style: 'medium', color: { argb: 'FFFFC000' } },
      bottom: { style: 'medium', color: { argb: 'FFFFC000' } },
      right: { style: 'medium', color: { argb: 'FFFFC000' } },
    };
    pCell1.dataValidation = {
      type: 'list',
      allowBlank: false,
      formulae: [dropdownRangeRef],
      showErrorMessage: false,
      promptTitle: 'Filter Pilih Weekly (1 s/d 53)',
      prompt: 'Pilih WEEKLY 1 s/d WEEKLY 53.',
    };

    periodBanner.getCell(2).value = {
      formula: `VLOOKUP($A$${pRowNum},${lookupRangeRef},4,FALSE)`,
      result: bannerText,
    };
    for (let c = 2; c <= 6; c++) {
      const cell = periodBanner.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F172A' },
      };
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: 'FFFDE047' },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'left',
        indent: 1,
      };
      cell.border = THIN_BORDER;
    }

    for (const sec of doc.sections) {
      if (sec.subRows && sec.subRows.length > 0) {
        // Section header row (e.g., "A. MERCHANDISE", "B.PROMO")
        const secHeaderRow = ws.addRow([sec.code, '', '', '', '', '']);
        secHeaderRow.height = 20;
        ws.mergeCells(secHeaderRow.number, 2, secHeaderRow.number, 6);

        const sh1 = secHeaderRow.getCell(1);
        sh1.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' },
        };
        sh1.font = {
          name: 'Calibri',
          size: 10.5,
          bold: false,
          color: { argb: 'FF000000' },
        };
        sh1.alignment = {
          vertical: 'middle',
          horizontal: 'left',
        };
        sh1.border = THIN_BORDER;

        for (let c = 2; c <= 6; c++) {
          const cell = secHeaderRow.getCell(c);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          cell.border = THIN_BORDER;
        }

        // Sub-rows (e.g., Divisi 1..4 right-aligned italic; FAST MOVING & SLOW MOVING left-aligned middle tall)
        for (const sub of sec.subRows) {
          const formattedItems = sub.items
            .map((it) => formatReviewBulletItem(it))
            .filter(Boolean);
          const itemsText = formattedItems.join('\n') || '-';
          const isDivisi =
            Boolean(sub.isItalic) ||
            sub.label.trim().toLowerCase().startsWith('divisi');
          const isFastOrSlow =
            sub.label.toUpperCase().includes('FAST MOVING') ||
            sub.label.toUpperCase().includes('SLOW MOVING');

          const subRow = ws.addRow([sub.label, itemsText, '', '', '', '']);
          subRow.height = isFastOrSlow
            ? Math.max(92, formattedItems.length * 15)
            : Math.max(20, formattedItems.length * 16);
          ws.mergeCells(subRow.number, 2, subRow.number, 6);

          const slotKey = `SUB::${sec.code.trim().toUpperCase()}::${sub.label.trim().toUpperCase()}`;
          const colOffset = slotIndexMap.get(slotKey);
          if (colOffset) {
            subRow.getCell(2).value = {
              formula: `VLOOKUP($A$${pRowNum},${lookupRangeRef},${colOffset},FALSE)`,
              result: itemsText,
            };
          }

          // Col 1: Label (Divisi 1..4 right-aligned italic; FAST/SLOW MOVING left-aligned middle)
          const c1 = subRow.getCell(1);
          c1.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          c1.font = {
            name: 'Calibri',
            size: 10.5,
            bold: false,
            italic: isDivisi,
            color: { argb: 'FF000000' },
          };
          c1.alignment = {
            vertical: 'middle',
            horizontal: isDivisi ? 'right' : 'left',
            wrapText: true,
          };
          c1.border = THIN_BORDER;

          // Col 2..6: Content cell
          for (let c = 2; c <= 6; c++) {
            const cell = subRow.getCell(c);
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFFFF' },
            };
            cell.font = {
              name: 'Calibri',
              size: 10,
              color: { argb: 'FF000000' },
            };
            cell.alignment = {
              vertical: isDivisi ? 'middle' : 'top',
              horizontal: 'left',
              wrapText: true,
            };
            cell.border = THIN_BORDER;
          }
        }
      } else {
        // Direct sections: C. PLATE/PEOPLE, D. ACTION PLAN, REASON/KENDALA, NEED SUPPORT
        const formattedDirect = (sec.directItems || [])
          .map((it) => formatReviewBulletItem(it))
          .filter(Boolean);
        const itemsText = formattedDirect.join('\n') || '-';
        const secRow = ws.addRow([sec.code, itemsText, '', '', '', '']);
        secRow.height = Math.max(76, formattedDirect.length * 15);
        ws.mergeCells(secRow.number, 2, secRow.number, 6);

        const slotKey = `DIR::${sec.code.trim().toUpperCase()}`;
        const colOffset = slotIndexMap.get(slotKey);
        if (colOffset) {
          secRow.getCell(2).value = {
            formula: `VLOOKUP($A$${pRowNum},${lookupRangeRef},${colOffset},FALSE)`,
            result: itemsText,
          };
        }

        // Col 1: Section Label (left-aligned, vertically centered)
        const c1 = secRow.getCell(1);
        c1.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' },
        };
        c1.font = {
          name: 'Calibri',
          size: 10.5,
          bold: false,
          color: { argb: 'FF000000' },
        };
        c1.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        };
        c1.border = THIN_BORDER;

        // Col 2..6: Content cell
        for (let c = 2; c <= 6; c++) {
          const cell = secRow.getCell(c);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' },
          };
          cell.font = {
            name: 'Calibri',
            size: 10,
            color: { argb: 'FF000000' },
          };
          cell.alignment = {
            vertical: 'top',
            horizontal: 'left',
            wrapText: true,
          };
          cell.border = THIN_BORDER;
        }
      }
    }
  }

  ws.addRow([]);

  // Populate the hidden lookup table AFTER all visible rows have been added so that
  // ws.addRow(...) places the Review Monitoring Bazar tables starting right at Row 1!
  populateWorksheetReviewLookupTable(
    ws,
    lookupTables,
    reviewDoc,
    poolDocs,
    slots,
    lookupStartCol
  );
  return firstReviewBannerRowNum ?? 7;
}

/* ============================================================================
   MAIN EXPORT FUNCTION: SUPPORTS CUSTOM SELECTION, WEEKEND FILTER & ALL-IN-1 EXCEL
   ============================================================================ */
export async function buildExportExcelWorkbook(
  allTables: WeeklySalesTable[],
  reviewDoc: ReviewMonitoringDoc,
  showDivZero: boolean = false,
  customOptions?: Partial<ExportExcelOptions>
): Promise<{
  workbook: ExcelJS.Workbook;
  fileName: string;
  resolvedMeta: {
    activeMonthKey: string;
    activeWeekNumber: number;
    weeklyMatrixWeekNumbers: number[];
  };
}> {
  const opts: ExportExcelOptions = {
    ...DEFAULT_EXPORT_EXCEL_OPTIONS,
    ...customOptions,
  };

  // Load live CV & Data CV Inventory early so period auto-detection reflects ALL modules
  const cvRecs =
    opts.cvRecords && opts.cvRecords.length > 0
      ? opts.cvRecords
      : loadCvDiv1Records();
  const cvOps =
    opts.cvInventoryOps && opts.cvInventoryOps.length > 0
      ? opts.cvInventoryOps
      : loadCvInventoryOps();

  // Filter tables according to period selection
  let filteredTables = [...allTables];

  if (opts.selectedWeekIds && opts.selectedWeekIds.length > 0) {
    const byId = new Map<string, WeeklySalesTable>();
    allTables.forEach((t) => {
      byId.set(t.id, t);
      byId.set(`week-${t.weekNumber}`, t);
    });
    const resolved: WeeklySalesTable[] = [];
    opts.selectedWeekIds.forEach((id) => {
      const existing = byId.get(id);
      if (existing) {
        resolved.push(existing);
      } else {
        const wkMatch = id.match(/(\d+)/);
        if (wkMatch) {
          const wkNum = Number(wkMatch[1]);
          const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum);
          if (sched) {
            resolved.push(createWeeklyTableFromAnnualSchedule(sched));
          }
        }
      }
    });
    if (resolved.length > 0) {
      filteredTables = resolved.sort((a, b) => a.weekNumber - b.weekNumber);
    }
  } else if (opts.selectedMonthKey && opts.selectedMonthKey !== 'ALL') {
    filteredTables = filteredTables.filter(
      (t) => getTableMonthKey(t) === opts.selectedMonthKey
    );
  }

  if (filteredTables.length === 0) {
    filteredTables = [...allTables];
  }

  // If no explicit single-week filter was chosen (or all tables were selected by default),
  // retain up to the LATEST active week (never clipping newer weeks on sparse/gap data)
  const isAllOrMultiDefault =
    !opts.selectedWeekIds ||
    opts.selectedWeekIds.length === 0 ||
    (opts.selectedWeekIds.length >= allTables.length && opts.selectedWeekIds.length > 1);
  if (isAllOrMultiDefault) {
    const filledWeeks = filteredTables.filter((t) => isWeeklyTableActive(t));
    if (filledWeeks.length > 0) {
      const minFilledWeek = Math.min(...filledWeeks.map((t) => t.weekNumber));
      const maxFilledWeek = Math.max(...filledWeeks.map((t) => t.weekNumber));
      const span = maxFilledWeek - minFilledWeek + 1;
      if (span <= 16) {
        const upperBoundWeek = Math.max(maxFilledWeek, minFilledWeek + 2);
        const cleaned = filteredTables.filter(
          (t) =>
            t.weekNumber >= minFilledWeek && t.weekNumber <= upperBoundWeek
        );
        if (cleaned.length > 0) {
          filteredTables = cleaned.sort((a, b) => a.weekNumber - b.weekNumber);
        }
      } else {
        filteredTables = [...filledWeeks].sort(
          (a, b) => a.weekNumber - b.weekNumber
        );
      }
    }
  }

  // Ensure at least 1 section is selected
  const includeDataCvResolved =
    opts.includeDataCv !== undefined
      ? opts.includeDataCv
      : Boolean(opts.includeCvReport);

  const hasAnySection =
    opts.includeWeeklyMatrix ||
    opts.includeDailyTable ||
    opts.includeMonthlyTable ||
    opts.includeReasonEvaluation ||
    opts.includeStaffAndReview ||
    opts.includeCvReport ||
    includeDataCvResolved ||
    opts.includeTotalConsolidated;
  if (!hasAnySection) {
    opts.includeWeeklyMatrix = true;
  }

  // Resolve latest active month across DP + CV + Data CV
  const activeMonthKey = resolveLatestActiveMonthKey(
    opts.selectedMonthKey,
    allTables,
    cvRecs,
    cvOps
  );

  // Build sequential weeklyMatrixTables: max 3 weekly periods, prioritizing the LATEST 3 active weeks
  // while keeping chronological ascending order so +1 week VLOOKUP links stay coherent.
  const weeklyMatrixTables: WeeklySalesTable[] = (() => {
    const sorted = [...filteredTables].sort(
      (a, b) => a.weekNumber - b.weekNumber
    );
    if (sorted.length === 1) {
      const startWk = sorted[0].weekNumber;
      const seq: WeeklySalesTable[] = [sorted[0]];
      for (let offset = 1; offset <= 2; offset++) {
        const nextWk = startWk + offset;
        if (nextWk <= 53) {
          const ex = allTables.find((t) => t.weekNumber === nextWk);
          const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === nextWk);
          if (ex) seq.push(ex);
          else if (sched) seq.push(createWeeklyTableFromAnnualSchedule(sched));
        }
      }
      return seq.slice(0, 3);
    }
    // Limit to at most 3 weekly periods, picking the LATEST 3
    if (sorted.length > 3) {
      // If user selected explicit custom weeks (>3 checkboxes), take the latest 3 selected
      if (
        opts.selectedWeekIds &&
        opts.selectedWeekIds.length > 0 &&
        opts.selectedWeekIds.length < allTables.length
      ) {
        return sorted.slice(-3);
      }
      // If "Semua Periode" or month filter with >3 weeks, prioritize the LATEST active weeks
      const filled = sorted.filter((t) => isWeeklyTableActive(t));
      if (filled.length >= 3) {
        return filled.slice(-3);
      } else if (filled.length > 0) {
        const maxFilledWk = filled[filled.length - 1].weekNumber;
        const minFilledWk = filled[0].weekNumber;
        const startWk = Math.max(minFilledWk, maxFilledWk - 2);
        const windowAroundLatest = sorted.filter(
          (t) =>
            t.weekNumber >= startWk &&
            t.weekNumber <= Math.max(maxFilledWk, startWk + 2)
        );
        if (windowAroundLatest.length >= 3) {
          return windowAroundLatest.slice(-3);
        }
        const fromFirstFilled = sorted.filter((t) => t.weekNumber >= minFilledWk);
        if (fromFirstFilled.length >= 3) {
          return fromFirstFilled.slice(0, 3);
        }
        return sorted.slice(-3);
      }
      // When no DP weeks are filled yet, show up to 3 weeks of the latest active month (e.g. Nov)
      const monthMatched = sorted.filter(
        (t) => getTableMonthKey(t) === activeMonthKey
      );
      if (monthMatched.length > 0) {
        return monthMatched.slice(-3);
      }
      return sorted.slice(-3);
    }
    return sorted.slice(-3);
  })();

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'REPORT BAZAR MATAHARI — Tanjung (PT MDS Retailing Tbk)';
  workbook.lastModifiedBy = 'PT MDS Retailing Tbk';
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const activeWkNum = resolveLatestActiveWeekNumber(
    opts.activeCvWeekNumber,
    weeklyMatrixTables,
    cvRecs,
    cvOps,
    activeMonthKey
  );

  if (opts.compactMode) {
    // =========================================================================
    // COMPACT 5-SHEET MODE:
    // 1. Report DP | 2. Report CV | 3. Total DP+CV | 4. Monitoring | 5. Data CV
    // =========================================================================
    const includeDpSheet =
      opts.includeWeeklyMatrix ||
      opts.includeDailyTable ||
      opts.includeMonthlyTable;

    if (includeDpSheet) {
      const wsDp = workbook.addWorksheet('Report DP', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      const dpWidths = [
        24, 15, 14.5, 14.5, 14.5, 16, 13, 13, 16, 15, 13, 13, 18, 13.5, 14, 17,
        14, 11.5, 11.5, 12.5, 12.5, 34, 34,
      ];
      dpWidths.forEach((w, i) => {
        wsDp.getColumn(i + 1).width = w;
      });
      addOfficialReportHeader(wsDp, 23);

      let dpFreeze: { x: number; y: number } | null = null;
      if (opts.includeWeeklyMatrix) {
        const wkTitleRow = renderWeeklyMatrixSection(
          workbook,
          wsDp,
          weeklyMatrixTables.slice(0, 3),
          showDivZero,
          true,
          opts.includeChartDiagrams !== false
        );
        dpFreeze = { x: 1, y: wkTitleRow };
      }
      if (opts.includeDailyTable) {
        if (opts.includeWeeklyMatrix) {
          const sp = wsDp.addRow([]);
          sp.height = 14;
        }
        const dailySubRow = renderDailyTableSection(
          workbook,
          wsDp,
          filteredTables,
          opts.includeChartDiagrams !== false,
          false, // AutoFilter off when multiple sections share Report DP
          opts.selectedMonthKey,
          allTables,
          activeMonthKey
        );
        if (!dpFreeze) {
          dpFreeze = { x: 4, y: dailySubRow };
        }
      }
      if (opts.includeMonthlyTable) {
        if (opts.includeWeeklyMatrix || opts.includeDailyTable) {
          const sp = wsDp.addRow([]);
          sp.height = 14;
        }
        const monthlySubRow = renderMonthlyTableSection(
          workbook,
          wsDp,
          filteredTables,
          opts.includeChartDiagrams !== false,
          false,
          opts.selectedMonthKey,
          allTables,
          activeMonthKey
        );
        if (!dpFreeze) {
          dpFreeze = { x: 2, y: monthlySubRow };
        }
      }
      if (dpFreeze) {
        applySheetFreeze(wsDp, dpFreeze.x, dpFreeze.y);
      }
      if (opts.includeWeeklyMatrix) {
        populateWorksheetWeeklyLookupTable(wsDp, allTables);
      }
    }

    if (opts.includeCvReport) {
      appendCvDiv1SheetsToWorkbook(
        workbook,
        cvRecs,
        activeWkNum,
        activeMonthKey,
        2026,
        cvOps,
        {
          compact: true,
          includeSales: true,
          includeDataCv: false,
          includeCharts: opts.includeChartDiagrams !== false,
          salesSheetName: 'Report CV',
        }
      );
    }

    if (opts.includeTotalConsolidated) {
      appendTotalDpCvSheetsToWorkbook(
        workbook,
        allTables,
        cvRecs,
        activeMonthKey,
        2026,
        {
          weeklyScope: 'month',
          compact: true,
          includePrefix: false,
          includeCharts: opts.includeChartDiagrams !== false,
          sheetName: 'Total DP+CV',
        }
      );
    }

    if (opts.includeStaffAndReview || opts.includeReasonEvaluation) {
      const wsMon = workbook.addWorksheet('Monitoring', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      const monWidths = [22, 34, 16, 20, 32, 16, 20, 32, 18, 24];
      monWidths.forEach((w, i) => {
        wsMon.getColumn(i + 1).width = w;
      });
      addOfficialReportHeader(wsMon, 10);

      let monFreeze: { x: number; y: number } | null = null;
      if (opts.includeStaffAndReview) {
        // enableSheetAutoFilter = false skips internal 6-col header and places Review lookup at Col 115 (DK)
        const reviewBannerRow = renderStaffAndReviewSection(
          wsMon,
          filteredTables,
          reviewDoc,
          opts.allReviewDocs,
          opts.selectedReviewWeekFilter,
          opts.selectedMonthKey,
          false,
          allTables,
          activeMonthKey
        );
        monFreeze = { x: 1, y: reviewBannerRow };
      }
      if (opts.includeReasonEvaluation) {
        if (opts.includeStaffAndReview) {
          const sp = wsMon.addRow([]);
          sp.height = 14;
        }
        const reasonHeaderRow = renderReasonEvaluationSection(
          wsMon,
          filteredTables,
          activeMonthKey
        );
        if (!monFreeze) {
          monFreeze = { x: 0, y: reasonHeaderRow };
        }
        // Weekly lookup uses Cols 26..114 (Z..DJ), non-overlapping with Cols 1..10 & Col 115+
        populateWorksheetWeeklyLookupTable(wsMon, allTables);
      }
      if (monFreeze) {
        applySheetFreeze(wsMon, monFreeze.x, monFreeze.y);
      }
    }

    if (includeDataCvResolved) {
      appendCvDiv1SheetsToWorkbook(
        workbook,
        cvRecs,
        activeWkNum,
        activeMonthKey,
        2026,
        cvOps,
        {
          compact: true,
          includeSales: false,
          includeDataCv: true,
          dataCvSheetName: 'Data CV',
        }
      );
    }
  } else {
    // =========================================================================
    // DETAILED MULTI-TAB MODE
    // =========================================================================
    if (
      opts.includeExecutiveCover ||
      opts.includeCvReport ||
      includeDataCvResolved ||
      opts.includeTotalConsolidated
    ) {
      renderExecutiveSummarySheet(
        workbook,
        allTables,
        cvRecs,
        activeMonthKey,
        opts
      );
    }

    if (opts.includeDailyTable) {
      const wsDaily = workbook.addWorksheet('Sales Daily', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      const dailyWidths = [
        7, 13.5, 12.5, 10.5,
        16, 17, 10.5, 10.5,
        16, 17, 10.5, 10.5,
        18.5, 12.5, 12.5, 14.5, 14.5,
        11.5, 11.5, 12.5, 13,
        36, 36,
      ];
      dailyWidths.forEach((w, i) => {
        wsDaily.getColumn(i + 1).width = w;
      });
      addOfficialReportHeader(wsDaily, 23);
      const dailySubRow = renderDailyTableSection(
        workbook,
        wsDaily,
        filteredTables,
        opts.includeChartDiagrams !== false,
        true,
        opts.selectedMonthKey,
        allTables,
        activeMonthKey
      );
      applySheetFreeze(wsDaily, 4, dailySubRow);
    }

    if (opts.includeWeeklyMatrix) {
      const wsWeekly = workbook.addWorksheet('Sales Weekly', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      wsWeekly.getColumn(1).width = 24;
      for (let c = 2; c <= 15; c++) wsWeekly.getColumn(c).width = 16;
      wsWeekly.getColumn(16).width = 22;
      addOfficialReportHeader(wsWeekly, 16);
      const wkTitleRow = renderWeeklyMatrixSection(
        workbook,
        wsWeekly,
        weeklyMatrixTables.slice(0, 3), // strictly max 3 weekly
        showDivZero,
        true,
        opts.includeChartDiagrams !== false
      );
      applySheetFreeze(wsWeekly, 1, wkTitleRow);
      populateWorksheetWeeklyLookupTable(wsWeekly, allTables);
    }

    if (opts.includeMonthlyTable) {
      const wsMonthly = workbook.addWorksheet('Sales Monthly', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      const mWidths = [
        22, 30, 14, 17.5, 11.5, 11.5, 17.5, 11.5, 11.5, 19.5, 13, 13, 16, 16, 15.5, 15.5,
      ];
      mWidths.forEach((w, i) => {
        wsMonthly.getColumn(i + 1).width = w;
      });
      addOfficialReportHeader(wsMonthly, 16);
      const monthlySubRow = renderMonthlyTableSection(
        workbook,
        wsMonthly,
        filteredTables,
        opts.includeChartDiagrams !== false,
        true,
        opts.selectedMonthKey,
        allTables,
        activeMonthKey
      );
      applySheetFreeze(wsMonthly, 2, monthlySubRow);
    }

    if (opts.includeReasonEvaluation) {
      const wsReason = workbook.addWorksheet('Evaluasi Reason & Event', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      const rWidths = [7, 21, 16, 22, 36, 16, 22, 36, 18.5, 25];
      rWidths.forEach((w, i) => {
        wsReason.getColumn(i + 1).width = w;
      });
      addOfficialReportHeader(wsReason, 10);
      const reasonHeaderRow = renderReasonEvaluationSection(
        wsReason,
        filteredTables,
        activeMonthKey
      );
      applySheetFreeze(wsReason, 0, reasonHeaderRow);
      populateWorksheetWeeklyLookupTable(wsReason, allTables);
    }

    if (opts.includeStaffAndReview) {
      const wsReview = workbook.addWorksheet('Review Monitoring Bazar', {
        views: [{ showGridLines: true }],
        pageSetup: {
          paperSize: 9,
          orientation: 'portrait',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
        },
      });
      [22, 46, 22, 20, 20, 20].forEach((w, i) => {
        wsReview.getColumn(i + 1).width = w;
      });
      const reviewBannerRow = renderStaffAndReviewSection(
        wsReview,
        filteredTables,
        reviewDoc,
        opts.allReviewDocs,
        opts.selectedReviewWeekFilter,
        opts.selectedMonthKey,
        true,
        allTables,
        activeMonthKey
      );
      applySheetFreeze(wsReview, 0, reviewBannerRow, { setPrintTitles: false });
    }

    // 2) OPTIONAL CV DIVISI 1 & DATA CV INVENTORY + ACTION PLAN SHEETS
    if (opts.includeCvReport || includeDataCvResolved) {
      appendCvDiv1SheetsToWorkbook(
        workbook,
        cvRecs,
        activeWkNum,
        activeMonthKey,
        2026,
        cvOps,
        {
          compact: false,
          includePrefix: true,
          includeSales: Boolean(opts.includeCvReport),
          includeDataCv: includeDataCvResolved,
          includeCharts: opts.includeChartDiagrams !== false,
        }
      );
    }

    // 3) OPTIONAL TOTAL KONSOLIDASI (DP + CV) & TRAFFIC / CONVERSION RATE SHEETS
    if (opts.includeTotalConsolidated) {
      appendTotalDpCvSheetsToWorkbook(
        workbook,
        allTables,
        cvRecs,
        activeMonthKey,
        2026,
        {
          compact: false,
          weeklyScope: 'month',
          includePrefix: true,
          includeCharts: opts.includeChartDiagrams !== false,
        }
      );
    }
  }

  // 4) BUILD HIDDEN DATABASE SHEET FOR "TOTAL WEEKLY 1" .. "TOTAL WEEKLY 53" DROPDOWN & VLOOKUP
  buildHiddenWeeklyDatabaseSheet(workbook, allTables);

  const dateStr = new Date().toISOString().slice(0, 10);
  const baseName = (opts.customFileName || '').trim()
    ? opts.customFileName!.trim().replace(/[^a-zA-Z0-9_-]/g, '_')
    : `Report_Bazar_Matahari_Tanjung_${dateStr}`;
  const fileName = baseName.endsWith('.xlsx') ? baseName : `${baseName}.xlsx`;

  return {
    workbook,
    fileName,
    resolvedMeta: {
      activeMonthKey,
      activeWeekNumber: activeWkNum,
      weeklyMatrixWeekNumbers: weeklyMatrixTables.map((t) => t.weekNumber),
    },
  };
}

export async function exportWeeklyTablesToWPSExcel(
  allTables: WeeklySalesTable[],
  reviewDoc: ReviewMonitoringDoc,
  showDivZero: boolean = false,
  customOptions?: Partial<ExportExcelOptions>
) {
  const { workbook, fileName } = await buildExportExcelWorkbook(
    allTables,
    reviewDoc,
    showDivZero,
    customOptions
  );
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function renderExecutiveSummarySheet(
  workbook: ExcelJS.Workbook,
  allTables: WeeklySalesTable[],
  cvRecords: CvDailyRecord[],
  activeMonthKey: string,
  opts: ExportExcelOptions
) {
  const wsCover = workbook.addWorksheet('0. Ringkasan Eksekutif', {
    views: [{ showGridLines: true }],
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
    },
  });

  [26, 22, 20, 18, 18, 34].forEach((w, idx) => {
    wsCover.getColumn(idx + 1).width = w;
  });

  const mName = (MONTH_NAMES_ID[activeMonthKey] || `Bulan ${activeMonthKey}`).toUpperCase();
  addUnifiedCorporateHeader(
    wsCover,
    6,
    'RINGKASAN EKSEKUTIF — MASTER LAPORAN BAZAR MATAHARI TANJUNG',
    `Bulan Aktif: ${mName} 2026 · Laporan Terpadu DP + CV + Inventory + Traffic & Conversion Rate`
  );

  const maps = buildConsolidatedMaps(allTables, cvRecords);
  let dpMonthSales = 0, dpMonthQty = 0, dpMonthTrx = 0;
  let cvMonthSales = 0, cvMonthQty = 0, cvMonthNota = 0;
  let trafPagi = 0, trafSiang = 0;
  const daysInMonth = new Date(2026, parseInt(activeMonthKey, 10) || 11, 0).getDate();

  for (let d = 1; d <= daysInMonth; d++) {
    const tgl = `${String(d).padStart(2, '0')}/${activeMonthKey}/2026`;
    const dp = maps.dpDateMap.get(tgl);
    if (dp) {
      dpMonthSales += dp.sales;
      dpMonthQty += dp.qty;
      dpMonthTrx += dp.trx;
      trafPagi += dp.pagiTraffic;
      trafSiang += dp.siangTraffic;
    }
    const cv = maps.cvDateMap.get(tgl);
    if (cv) {
      cvMonthSales += cv.sales;
      cvMonthQty += cv.qty;
      cvMonthNota += cv.trx;
    }
  }
  const grandSales = dpMonthSales + cvMonthSales;
  const grandQty = dpMonthQty + cvMonthQty;
  const grandTrxNota = dpMonthTrx + cvMonthNota;
  const totalTraf = trafPagi + trafSiang;
  const convRate = totalTraf > 0 ? grandTrxNota / totalTraf : 0;

  const sp = wsCover.addRow([]);
  sp.height = 8;

  addKpiStrip(
    wsCover,
    [
      {
        label: `1. REPORT DP (${mName})`,
        value: dpMonthSales,
        numFmt: NUM_FMT_IDR,
        bgArgb: EXCEL_COLORS.amberSubBg,
        fgArgb: EXCEL_COLORS.amberText,
      },
      {
        label: `2. REPORT CV DIV. 1 (${mName})`,
        value: cvMonthSales,
        numFmt: NUM_FMT_IDR,
        bgArgb: EXCEL_COLORS.purpleSubBg,
        fgArgb: EXCEL_COLORS.purpleText,
      },
      {
        label: 'GRAND TOTAL (DP + CV)',
        value: grandSales,
        numFmt: NUM_FMT_IDR,
        bgArgb: EXCEL_COLORS.goldLight,
        fgArgb: EXCEL_COLORS.black,
      },
      {
        label: 'TOTAL TRX+NOTA',
        value: grandTrxNota,
        numFmt: NUM_FMT_INTEGER,
        bgArgb: EXCEL_COLORS.emeraldSoft,
        fgArgb: EXCEL_COLORS.emeraldText,
      },
      {
        label: 'TOTAL TRAFFIC',
        value: totalTraf,
        numFmt: NUM_FMT_INTEGER,
        bgArgb: EXCEL_COLORS.skySubBg,
        fgArgb: EXCEL_COLORS.skyHeader,
      },
      {
        label: 'CONVERSION RATE %',
        value: convRate,
        numFmt: NUM_FMT_PERCENT,
        bgArgb: EXCEL_COLORS.indigoSubBg,
        fgArgb: EXCEL_COLORS.indigoText,
      },
    ],
    6
  );

  const sp2 = wsCover.addRow([]);
  sp2.height = 10;

  addSectionTitleBar(
    wsCover,
    `PERBANDINGAN KONTRIBUSI KATEGORI — BULAN ${mName} 2026`,
    6,
    EXCEL_COLORS.slate800,
    EXCEL_COLORS.yellowText
  );

  const kpiHdr = wsCover.addRow([
    'Kategori / Stream',
    'Total Sales (Rp)',
    'Total Qty (Pcs)',
    'Trx / Nota',
    'Share Sales %',
    'Keterangan / Catatan',
  ]);
  kpiHdr.height = 23;
  for (let c = 1; c <= 6; c++) {
    const cell = kpiHdr.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.crimsonDark);
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: EXCEL_COLORS.white } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = THIN_BORDER;
  }

  const dpRow = wsCover.addRow([
    '1. Report DP (Direct Purchase)',
    dpMonthSales,
    dpMonthQty,
    dpMonthTrx,
    null,
    'Satuan transaksi: Trx · Termasuk data Traffic Pagi & Siang',
  ]);
  const cvRow = wsCover.addRow([
    '2. Report CV (Konsinyasi Div. 1)',
    cvMonthSales,
    cvMonthQty,
    cvMonthNota,
    null,
    'Satuan transaksi: Nota · Dept 6201 Graphis & 6244 Simplicity',
  ]);
  const totRow = wsCover.addRow([
    'GRAND TOTAL BAZAR (DP + CV)',
    null,
    null,
    null,
    null,
    `Total Trx+Nota = ${grandTrxNota.toLocaleString('id-ID')} | Avg/Org = Rp ${totalTraf > 0 ? Math.round(grandSales / totalTraf).toLocaleString('id-ID') : '0'}`,
  ]);
  totRow.getCell(2).value = { formula: `B${dpRow.number}+B${cvRow.number}`, result: grandSales };
  totRow.getCell(3).value = { formula: `C${dpRow.number}+C${cvRow.number}`, result: grandQty };
  totRow.getCell(4).value = { formula: `D${dpRow.number}+D${cvRow.number}`, result: grandTrxNota };
  dpRow.getCell(5).value = { formula: `IF($B$${totRow.number}>0,B${dpRow.number}/$B$${totRow.number},0)` };
  cvRow.getCell(5).value = { formula: `IF($B$${totRow.number}>0,B${cvRow.number}/$B$${totRow.number},0)` };
  totRow.getCell(5).value = { formula: `E${dpRow.number}+E${cvRow.number}`, result: 1 };

  [dpRow, cvRow].forEach((r) => {
    r.height = 22;
    for (let c = 1; c <= 6; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = { name: 'Calibri', size: 10, bold: c === 1 || c === 2 };
      if (c === 2) cell.numFmt = NUM_FMT_IDR;
      else if (c === 3 || c === 4) cell.numFmt = NUM_FMT_INTEGER;
      else if (c === 5) cell.numFmt = NUM_FMT_PERCENT;
      cell.alignment = { vertical: 'middle', horizontal: c >= 2 && c <= 5 ? 'right' : 'left', indent: c === 1 || c === 6 ? 1 : 0 };
    }
  });

  totRow.height = 24;
  for (let c = 1; c <= 6; c++) {
    const cell = totRow.getCell(c);
    cell.fill = solidFill(EXCEL_COLORS.gold);
    cell.border = MEDIUM_OUTER_BORDER;
    cell.font = { name: 'Calibri', size: 10.5, bold: true };
    if (c === 2) cell.numFmt = NUM_FMT_IDR;
    else if (c === 3 || c === 4) cell.numFmt = NUM_FMT_INTEGER;
    else if (c === 5) cell.numFmt = NUM_FMT_PERCENT;
    cell.alignment = { vertical: 'middle', horizontal: c >= 2 && c <= 5 ? 'right' : 'center' };
  }

  const sp3 = wsCover.addRow([]);
  sp3.height = 10;

  addSectionTitleBar(
    wsCover,
    'DAFTAR ISI WORKSHEET (MODUL YANG DIEKSPOR)',
    6,
    EXCEL_COLORS.indigoHeader,
    EXCEL_COLORS.white
  );

  const modulesList: [string, string, string][] = [];
  if (opts.includeDailyTable) modulesList.push(['Sales Daily', 'Report DP Harian + Traffic & Conv. Rate', 'Tabel Kalender 1–30/31 + Diagram']);
  if (opts.includeWeeklyMatrix) modulesList.push(['Sales Weekly', 'Report DP Mingguan (Dropdown W1–W53)', 'Matriks Shift Pagi/Siang + Grafik']);
  if (opts.includeMonthlyTable) modulesList.push(['Sales Monthly', 'Rekap Bulanan & 12 Bulan DP', 'Kontribusi % + MoM Growth %']);
  if (opts.includeReasonEvaluation) modulesList.push(['Evaluasi Reason & Event', 'Evaluasi Shift Pagi vs Siang DP', 'Catatan Kendala & Dominasi Shift']);
  if (opts.includeStaffAndReview) modulesList.push(['Review Monitoring Bazar', 'Review Merchandise, Promo & Action Plan', 'VLOOKUP Interaktif W1–W53']);
  if (opts.includeCvReport) modulesList.push(['CV Harian / Weekly / Bulanan / Data CV', 'Report CV Divisi 1 & Inventory/Action Plan', 'Dept 6201 Graphis & 6244 Simplicity']);
  if (opts.includeTotalConsolidated) modulesList.push(['DP+CV Total Harian / Weekly / 12 Bln / Traffic', 'Konsolidasi Total DP + CV & Conversion Rate', 'Trx DP + Nota CV = Trx+Nota']);

  modulesList.forEach(([sheetTitle, desc, detail], idx) => {
    const r = wsCover.addRow([`${idx + 1}. ${sheetTitle}`, desc, detail, '', '', 'Aktif']);
    r.height = 20;
    wsCover.mergeCells(r.number, 3, r.number, 5);
    for (let c = 1; c <= 6; c++) {
      const cell = r.getCell(c);
      cell.border = THIN_BORDER;
      cell.font = { name: 'Calibri', size: 9.5, bold: c === 1 || c === 6 };
      if (idx % 2 === 1) cell.fill = solidFill(EXCEL_COLORS.slate50);
      cell.alignment = { vertical: 'middle', horizontal: c === 6 ? 'center' : 'left', indent: c < 6 ? 1 : 0 };
    }
  });
}
