/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * 
 * Data models and calculation engine for:
 * KATEGORI CV, DIVISI 1 (YOUNG FASHION, LADIES & INTIMATE)
 * Dept : 6201 - Brand : Graphis
 * Dept : 6244 - Brand : Simplicity
 */

import { ANNUAL_2026_WEEKS, getIndoDayName } from './initialData';

export interface CvBrandShift {
  sales: number | null;
  qty: number | null;
  trx: number | null;
}

export interface CvDailyBrandData {
  pagi: CvBrandShift;
  siang: CvBrandShift;
}

export interface CvDailyRecord {
  tanggal: string; // "DD/MM/YYYY"
  isoDate: string; // "YYYY-MM-DD"
  dayName: string; // "Senin", "Selasa", etc.
  weekNumber: number; // 1 - 53
  monthKey: string; // "01" - "12"
  petugasPagi?: string;
  petugasSiang?: string;
  targetDaily?: number;
  reasonEvent?: string;
  graphis: CvDailyBrandData; // Dept 6201 - Graphis
  simplicity: CvDailyBrandData; // Dept 6244 - Simplicity
}

export interface CvBrandTotals {
  sales: number;
  qty: number;
  trx: number;
  pricePoint: number;
  sharePct: number;
}

export interface CvDaySummary {
  tanggal: string;
  dayName: string;
  weekNumber: number;
  monthKey: string;
  petugasPagi: string;
  petugasSiang: string;
  targetDaily: number;
  reasonEvent: string;
  graphis: CvBrandTotals;
  simplicity: CvBrandTotals;
  total: {
    sales: number;
    qty: number;
    trx: number;
    pricePoint: number;
  };
  achievementPct: number;
}

export interface CvWeeklySummary {
  weekNumber: number;
  periodeText: string;
  startDate: string;
  endDate: string;
  monthKey: string;
  monthName: string;
  days: CvDaySummary[];
  graphis: CvBrandTotals;
  simplicity: CvBrandTotals;
  total: {
    sales: number;
    qty: number;
    trx: number;
    pricePoint: number;
  };
  targetWeekly: number;
  achievementPct: number;
  avgDailySales: number;
  weekdaySales: number;
  weekendSales: number;
}

export interface CvMonthlySummary {
  monthKey: string;
  monthName: string;
  year: number;
  daysCount: number;
  days: CvDaySummary[];
  graphis: CvBrandTotals;
  simplicity: CvBrandTotals;
  total: {
    sales: number;
    qty: number;
    trx: number;
    pricePoint: number;
  };
  targetMonthly: number;
  achievementPct: number;
  avgDailySales: number;
  bestDay: {
    tanggal: string;
    dayName: string;
    sales: number;
  } | null;
  weeklyBreakdown: {
    weekNumber: number;
    sales: number;
    qty: number;
    trx: number;
    graphisSales: number;
    simplicitySales: number;
  }[];
}

export interface CvQuarterSummary {
  quarter: 'Q1' | 'Q2' | 'Q3' | 'Q4';
  label: string;
  months: string[];
  graphisSales: number;
  simplicitySales: number;
  totalSales: number;
  totalQty: number;
  totalTrx: number;
  graphisSharePct: number;
  simplicitySharePct: number;
}

export interface CvAnnualSummary {
  year: number;
  monthlySummaries: CvMonthlySummary[];
  quarters: CvQuarterSummary[];
  graphis: CvBrandTotals;
  simplicity: CvBrandTotals;
  total: {
    sales: number;
    qty: number;
    trx: number;
    pricePoint: number;
  };
  targetAnnual: number;
  achievementPct: number;
  bestMonth: {
    monthKey: string;
    monthName: string;
    sales: number;
  } | null;
  dominantBrand: 'Graphis' | 'Simplicity' | 'Seimbang';
}

export const MONTH_NAMES_ID: Record<string, string> = {
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

export function calcCvBrandShiftTotals(brand: CvDailyBrandData): {
  sales: number;
  qty: number;
  trx: number;
  pricePoint: number;
} {
  const pSales = Number(brand.pagi.sales) || 0;
  const sSales = Number(brand.siang.sales) || 0;
  const pQty = Number(brand.pagi.qty) || 0;
  const sQty = Number(brand.siang.qty) || 0;
  const pTrx = Number(brand.pagi.trx) || 0;
  const sTrx = Number(brand.siang.trx) || 0;

  const sales = pSales + sSales;
  const qty = pQty + sQty;
  const trx = pTrx + sTrx;
  const pricePoint = qty > 0 ? Math.round(sales / qty) : 0;

  return { sales, qty, trx, pricePoint };
}

export function calcCvDaySummary(rec: CvDailyRecord): CvDaySummary {
  const gTotals = calcCvBrandShiftTotals(rec.graphis);
  const sTotals = calcCvBrandShiftTotals(rec.simplicity);

  const totalSales = gTotals.sales + sTotals.sales;
  const totalQty = gTotals.qty + sTotals.qty;
  const totalTrx = gTotals.trx + sTotals.trx;
  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;

  const gShare = totalSales > 0 ? (gTotals.sales / totalSales) * 100 : 0;
  const sShare = totalSales > 0 ? (sTotals.sales / totalSales) * 100 : 0;

  const target = Number(rec.targetDaily) || 1500000;
  const achievement = target > 0 ? (totalSales / target) * 100 : 0;

  return {
    tanggal: rec.tanggal,
    dayName: rec.dayName,
    weekNumber: rec.weekNumber,
    monthKey: rec.monthKey,
    petugasPagi: rec.petugasPagi || 'Petugas',
    petugasSiang: rec.petugasSiang || 'Petugas',
    targetDaily: target,
    reasonEvent: rec.reasonEvent || '',
    graphis: {
      sales: gTotals.sales,
      qty: gTotals.qty,
      trx: gTotals.trx,
      pricePoint: gTotals.pricePoint,
      sharePct: gShare,
    },
    simplicity: {
      sales: sTotals.sales,
      qty: sTotals.qty,
      trx: sTotals.trx,
      pricePoint: sTotals.pricePoint,
      sharePct: sShare,
    },
    total: {
      sales: totalSales,
      qty: totalQty,
      trx: totalTrx,
      pricePoint,
    },
    achievementPct: achievement,
  };
}

export function calcCvWeeklySummary(
  records: CvDailyRecord[],
  weekNumber: number
): CvWeeklySummary {
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNumber);
  const weekRecs = records.filter((r) => r.weekNumber === weekNumber);

  // If no records for week, generate dummy week slots
  const validRecs =
    weekRecs.length > 0
      ? weekRecs
      : (sched?.slashDates || []).map((dt, idx) => ({
          tanggal: dt,
          isoDate: '',
          dayName: getIndoDayName(dt),
          weekNumber,
          monthKey: sched?.monthKey || '10',
          graphis: { pagi: { sales: null, qty: null, trx: null }, siang: { sales: null, qty: null, trx: null } },
          simplicity: { pagi: { sales: null, qty: null, trx: null }, siang: { sales: null, qty: null, trx: null } },
        }));

  const daySummaries = validRecs.map(calcCvDaySummary);

  let gSales = 0,
    gQty = 0,
    gTrx = 0;
  let sSales = 0,
    sQty = 0,
    sTrx = 0;
  let weekdaySales = 0,
    weekendSales = 0;

  daySummaries.forEach((ds) => {
    gSales += ds.graphis.sales;
    gQty += ds.graphis.qty;
    gTrx += ds.graphis.trx;

    sSales += ds.simplicity.sales;
    sQty += ds.simplicity.qty;
    sTrx += ds.simplicity.trx;

    if (ds.dayName === 'Sabtu' || ds.dayName === 'Minggu') {
      weekendSales += ds.total.sales;
    } else {
      weekdaySales += ds.total.sales;
    }
  });

  const totalSales = gSales + sSales;
  const totalQty = gQty + sQty;
  const totalTrx = gTrx + sTrx;
  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;
  const gPricePoint = gQty > 0 ? Math.round(gSales / gQty) : 0;
  const sPricePoint = sQty > 0 ? Math.round(sSales / sQty) : 0;

  const targetWeekly = 10500000;
  const achievement = targetWeekly > 0 ? (totalSales / targetWeekly) * 100 : 0;
  const avgDaily = daySummaries.length > 0 ? Math.round(totalSales / daySummaries.length) : 0;

  return {
    weekNumber,
    periodeText: sched?.periodeText || `Weekly ${weekNumber}`,
    startDate: sched?.startDateIso || '',
    endDate: sched?.endDateIso || '',
    monthKey: sched?.monthKey || '10',
    monthName: MONTH_NAMES_ID[sched?.monthKey || '10'] || 'Oktober',
    days: daySummaries,
    graphis: {
      sales: gSales,
      qty: gQty,
      trx: gTrx,
      pricePoint: gPricePoint,
      sharePct: totalSales > 0 ? (gSales / totalSales) * 100 : 0,
    },
    simplicity: {
      sales: sSales,
      qty: sQty,
      trx: sTrx,
      pricePoint: sPricePoint,
      sharePct: totalSales > 0 ? (sSales / totalSales) * 100 : 0,
    },
    total: {
      sales: totalSales,
      qty: totalQty,
      trx: totalTrx,
      pricePoint,
    },
    targetWeekly,
    achievementPct: achievement,
    avgDailySales: avgDaily,
    weekdaySales,
    weekendSales,
  };
}

export function calcCvMonthlySummary(
  records: CvDailyRecord[],
  monthKey: string,
  year: number = 2026
): CvMonthlySummary {
  const monthRecs = records.filter(
    (r) => r.monthKey === monthKey || (r.tanggal && r.tanggal.split('/')[1] === monthKey)
  );
  const daySummaries = monthRecs.map(calcCvDaySummary);

  let gSales = 0,
    gQty = 0,
    gTrx = 0;
  let sSales = 0,
    sQty = 0,
    sTrx = 0;
  let bestDay: { tanggal: string; dayName: string; sales: number } | null = null;

  const weekMap = new Map<
    number,
    { sales: number; qty: number; trx: number; graphisSales: number; simplicitySales: number }
  >();

  daySummaries.forEach((ds) => {
    gSales += ds.graphis.sales;
    gQty += ds.graphis.qty;
    gTrx += ds.graphis.trx;

    sSales += ds.simplicity.sales;
    sQty += ds.simplicity.qty;
    sTrx += ds.simplicity.trx;

    if (!bestDay || ds.total.sales > bestDay.sales) {
      bestDay = {
        tanggal: ds.tanggal,
        dayName: ds.dayName,
        sales: ds.total.sales,
      };
    }

    const curr = weekMap.get(ds.weekNumber) || {
      sales: 0,
      qty: 0,
      trx: 0,
      graphisSales: 0,
      simplicitySales: 0,
    };
    curr.sales += ds.total.sales;
    curr.qty += ds.total.qty;
    curr.trx += ds.total.trx;
    curr.graphisSales += ds.graphis.sales;
    curr.simplicitySales += ds.simplicity.sales;
    weekMap.set(ds.weekNumber, curr);
  });

  const totalSales = gSales + sSales;
  const totalQty = gQty + sQty;
  const totalTrx = gTrx + sTrx;
  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;
  const gPricePoint = gQty > 0 ? Math.round(gSales / gQty) : 0;
  const sPricePoint = sQty > 0 ? Math.round(sSales / sQty) : 0;

  const targetMonthly = 45000000;
  const achievement = targetMonthly > 0 ? (totalSales / targetMonthly) * 100 : 0;
  const avgDaily = daySummaries.length > 0 ? Math.round(totalSales / daySummaries.length) : 0;

  const weeklyBreakdown = Array.from(weekMap.entries())
    .map(([wNum, data]) => ({
      weekNumber: wNum,
      ...data,
    }))
    .sort((a, b) => a.weekNumber - b.weekNumber);

  return {
    monthKey,
    monthName: MONTH_NAMES_ID[monthKey] || `Bulan ${monthKey}`,
    year,
    daysCount: daySummaries.length,
    days: daySummaries,
    graphis: {
      sales: gSales,
      qty: gQty,
      trx: gTrx,
      pricePoint: gPricePoint,
      sharePct: totalSales > 0 ? (gSales / totalSales) * 100 : 0,
    },
    simplicity: {
      sales: sSales,
      qty: sQty,
      trx: sTrx,
      pricePoint: sPricePoint,
      sharePct: totalSales > 0 ? (sSales / totalSales) * 100 : 0,
    },
    total: {
      sales: totalSales,
      qty: totalQty,
      trx: totalTrx,
      pricePoint,
    },
    targetMonthly,
    achievementPct: achievement,
    avgDailySales: avgDaily,
    bestDay,
    weeklyBreakdown,
  };
}

export function calcCvAnnualSummary(
  records: CvDailyRecord[],
  year: number = 2026
): CvAnnualSummary {
  const monthKeys = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
  const monthlySummaries = monthKeys.map((mk) => calcCvMonthlySummary(records, mk, year));

  let gSales = 0,
    gQty = 0,
    gTrx = 0;
  let sSales = 0,
    sQty = 0,
    sTrx = 0;
  let bestMonth: { monthKey: string; monthName: string; sales: number } | null = null;

  monthlySummaries.forEach((ms) => {
    gSales += ms.graphis.sales;
    gQty += ms.graphis.qty;
    gTrx += ms.graphis.trx;

    sSales += ms.simplicity.sales;
    sQty += ms.simplicity.qty;
    sTrx += ms.simplicity.trx;

    if (!bestMonth || ms.total.sales > bestMonth.sales) {
      bestMonth = {
        monthKey: ms.monthKey,
        monthName: ms.monthName,
        sales: ms.total.sales,
      };
    }
  });

  const totalSales = gSales + sSales;
  const totalQty = gQty + sQty;
  const totalTrx = gTrx + sTrx;
  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;
  const gPricePoint = gQty > 0 ? Math.round(gSales / gQty) : 0;
  const sPricePoint = sQty > 0 ? Math.round(sSales / sQty) : 0;

  const targetAnnual = 540000000;
  const achievement = targetAnnual > 0 ? (totalSales / targetAnnual) * 100 : 0;

  // Quarters
  const quarterDefs: { q: 'Q1' | 'Q2' | 'Q3' | 'Q4'; label: string; months: string[] }[] = [
    { q: 'Q1', label: 'Triwulan 1 (Jan - Mar)', months: ['01', '02', '03'] },
    { q: 'Q2', label: 'Triwulan 2 (Apr - Jun)', months: ['04', '05', '06'] },
    { q: 'Q3', label: 'Triwulan 3 (Jul - Sep)', months: ['07', '08', '09'] },
    { q: 'Q4', label: 'Triwulan 4 (Okt - Des)', months: ['10', '11', '12'] },
  ];

  const quarters: CvQuarterSummary[] = quarterDefs.map((qd) => {
    const qMonths = monthlySummaries.filter((ms) => qd.months.includes(ms.monthKey));
    const qGSales = qMonths.reduce((acc, m) => acc + m.graphis.sales, 0);
    const qSSales = qMonths.reduce((acc, m) => acc + m.simplicity.sales, 0);
    const qTotSales = qGSales + qSSales;
    const qTotQty = qMonths.reduce((acc, m) => acc + m.total.qty, 0);
    const qTotTrx = qMonths.reduce((acc, m) => acc + m.total.trx, 0);

    return {
      quarter: qd.q,
      label: qd.label,
      months: qd.months,
      graphisSales: qGSales,
      simplicitySales: qSSales,
      totalSales: qTotSales,
      totalQty: qTotQty,
      totalTrx: qTotTrx,
      graphisSharePct: qTotSales > 0 ? (qGSales / qTotSales) * 100 : 0,
      simplicitySharePct: qTotSales > 0 ? (qSSales / qTotSales) * 100 : 0,
    };
  });

  let dominantBrand: 'Graphis' | 'Simplicity' | 'Seimbang' = 'Seimbang';
  if (gSales > sSales * 1.05) dominantBrand = 'Graphis';
  else if (sSales > gSales * 1.05) dominantBrand = 'Simplicity';

  return {
    year,
    monthlySummaries,
    quarters,
    graphis: {
      sales: gSales,
      qty: gQty,
      trx: gTrx,
      pricePoint: gPricePoint,
      sharePct: totalSales > 0 ? (gSales / totalSales) * 100 : 0,
    },
    simplicity: {
      sales: sSales,
      qty: sQty,
      trx: sTrx,
      pricePoint: sPricePoint,
      sharePct: totalSales > 0 ? (sSales / totalSales) * 100 : 0,
    },
    total: {
      sales: totalSales,
      qty: totalQty,
      trx: totalTrx,
      pricePoint,
    },
    targetAnnual,
    achievementPct: achievement,
    bestMonth,
    dominantBrand,
  };
}

/**
 * Generate initial authentic 2026 daily records for CV Divisi 1
 * Pre-populated with realistic figures for Graphis (Dept 6201) and Simplicity (Dept 6244)
 */
export function generateInitialCvDiv1Records(): CvDailyRecord[] {
  const records: CvDailyRecord[] = [];

  // Seed sample base data across all 53 weeks
  for (const w of ANNUAL_2026_WEEKS) {
    const isSpecialFocusWeek = w.weekNumber >= 38 && w.weekNumber <= 43; // Sep - Oct Bazar period
    const isRamadanPeriod = w.weekNumber >= 11 && w.weekNumber <= 16; // Festive peak

    w.slashDates.forEach((slashDate, dIdx) => {
      const parts = slashDate.split('/');
      const dd = parts[0];
      const mm = parts[1];
      const yyyy = parts[2];
      const isoDate = `${yyyy}-${mm}-${dd}`;
      const dayName = getIndoDayName(slashDate);
      const isWeekend = dayName === 'Sabtu' || dayName === 'Minggu';

      records.push({
        tanggal: slashDate,
        isoDate,
        dayName,
        weekNumber: w.weekNumber,
        monthKey: w.monthKey,
        petugasPagi: '',
        petugasSiang: '',
        targetDaily: 0,
        reasonEvent: '',
        graphis: {
          pagi: {
            sales: null,
            qty: null,
            trx: null,
          },
          siang: {
            sales: null,
            qty: null,
            trx: null,
          },
        },
        simplicity: {
          pagi: {
            sales: null,
            qty: null,
            trx: null,
          },
          siang: {
            sales: null,
            qty: null,
            trx: null,
          },
        },
      });
    });
  }

  return records;
}

export const STORAGE_KEY_CV_DIV1 = 'bazar_tanjung_cv_div1_records_v2';

export function loadCvDiv1Records(): CvDailyRecord[] {
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CV_DIV1);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed;
        }
      }
    }
  } catch (err) {
    console.error('Error loading CV Divisi 1 records:', err);
  }
  const initial = generateInitialCvDiv1Records();
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY_CV_DIV1, JSON.stringify(initial));
    }
  } catch {}
  return initial;
}

export function saveCvDiv1Records(records: CvDailyRecord[]): void {
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY_CV_DIV1, JSON.stringify(records));
    }
  } catch (err) {
    console.error('Error saving CV Divisi 1 records:', err);
  }
}

// ============================================================================
// CV INVENTORY, PROMO, REASON/KENDALA & ACTION PLAN (PER WEEK & PER MONTHLY)
// ============================================================================

export type CvInventoryBrandKey = 'GRAPHIS' | 'SIMPLICITY';
export type CvInventoryPeriodType = 'weekly' | 'monthly';

export const CV_DIVISI_DEFAULT = 'Divisi 1';
export const CV_DIVISI_FULL = 'Divisi 1 · Young Fashion, Ladies & Intimate';

export const CV_BRAND_CATALOG: Record<
  CvInventoryBrandKey,
  {
    key: CvInventoryBrandKey;
    label: string;
    deptCode: '6201' | '6244';
    divisi: string;
  }
> = {
  GRAPHIS: {
    key: 'GRAPHIS',
    label: 'Graphis',
    deptCode: '6201',
    divisi: 'Divisi 1',
  },
  SIMPLICITY: {
    key: 'SIMPLICITY',
    label: 'Simplicity',
    deptCode: '6244',
    divisi: 'Divisi 1',
  },
};

export interface CvInventoryOpsEntry {
  id: string; // e.g. "weekly-45-GRAPHIS" | "monthly-11-SIMPLICITY"
  periodType: CvInventoryPeriodType;
  weekNumber?: number; // 1..53
  monthKey?: string; // "01".."12"
  divisi: string;
  brand: CvInventoryBrandKey;
  deptCode: '6201' | '6244';
  stockAwal: number;
  barangDatang: number;
  barangKeluar: number;
  /** Optional manual override for Sales Qty; when null/undefined, live-syncs from CvDailyRecord[] */
  salesQtyOverride: number | null;
  promo: string;
  reasonKendala: string;
  actionPlan: string;
  updatedAt?: string;
}

export const STORAGE_KEY_CV_INVENTORY_OPS = 'bazar_tanjung_cv_inventory_ops_v1';

export function getCvOpsEntryId(
  periodType: CvInventoryPeriodType,
  periodKey: number | string,
  brand: CvInventoryBrandKey
): string {
  return `${periodType}-${String(periodKey)}-${brand}`;
}

export function createDefaultCvOpsEntry(
  periodType: CvInventoryPeriodType,
  periodKey: number | string,
  brand: CvInventoryBrandKey
): CvInventoryOpsEntry {
  const brandMeta = CV_BRAND_CATALOG[brand];
  return {
    id: getCvOpsEntryId(periodType, periodKey, brand),
    periodType,
    weekNumber: periodType === 'weekly' ? Number(periodKey) : undefined,
    monthKey: periodType === 'monthly' ? String(periodKey).padStart(2, '0') : undefined,
    divisi: CV_DIVISI_DEFAULT,
    brand,
    deptCode: brandMeta.deptCode,
    stockAwal: 0,
    barangDatang: 0,
    barangKeluar: 0,
    salesQtyOverride: null,
    promo: '',
    reasonKendala: '',
    actionPlan: '',
    updatedAt: '',
  };
}

export function normalizeCvInventoryOpsList(rawList: unknown): CvInventoryOpsEntry[] {
  const byId = new Map<string, CvInventoryOpsEntry>();

  // Build baseline skeleton for all 53 weeks x 2 brands + 12 months x 2 brands
  const brands: CvInventoryBrandKey[] = ['GRAPHIS', 'SIMPLICITY'];
  ANNUAL_2026_WEEKS.forEach((w) => {
    brands.forEach((b) => {
      const def = createDefaultCvOpsEntry('weekly', w.weekNumber, b);
      byId.set(def.id, def);
    });
  });
  Object.keys(MONTH_NAMES_ID).forEach((mKey) => {
    brands.forEach((b) => {
      const def = createDefaultCvOpsEntry('monthly', mKey, b);
      byId.set(def.id, def);
    });
  });

  if (Array.isArray(rawList)) {
    rawList.forEach((item: any) => {
      if (!item || (item.brand !== 'GRAPHIS' && item.brand !== 'SIMPLICITY')) return;
      const pType: CvInventoryPeriodType = item.periodType === 'monthly' ? 'monthly' : 'weekly';
      const pKey =
        pType === 'weekly'
          ? Number(item.weekNumber || 0)
          : String(item.monthKey || '11').padStart(2, '0');
      if (pType === 'weekly' && (!pKey || Number(pKey) < 1 || Number(pKey) > 53)) return;
      if (pType === 'monthly' && !MONTH_NAMES_ID[String(pKey)]) return;

      const id = getCvOpsEntryId(pType, pKey, item.brand);
      const brandMeta = CV_BRAND_CATALOG[item.brand as CvInventoryBrandKey];
      byId.set(id, {
        id,
        periodType: pType,
        weekNumber: pType === 'weekly' ? Number(pKey) : undefined,
        monthKey: pType === 'monthly' ? String(pKey) : undefined,
        divisi: typeof item.divisi === 'string' && item.divisi.trim() ? item.divisi.trim() : CV_DIVISI_DEFAULT,
        brand: item.brand,
        deptCode: brandMeta.deptCode,
        stockAwal: Number.isFinite(Number(item.stockAwal)) ? Number(item.stockAwal) : 0,
        barangDatang: Number.isFinite(Number(item.barangDatang)) ? Number(item.barangDatang) : 0,
        barangKeluar: Number.isFinite(Number(item.barangKeluar)) ? Number(item.barangKeluar) : 0,
        salesQtyOverride:
          item.salesQtyOverride !== null &&
          item.salesQtyOverride !== undefined &&
          item.salesQtyOverride !== '' &&
          Number.isFinite(Number(item.salesQtyOverride))
            ? Number(item.salesQtyOverride)
            : null,
        promo: typeof item.promo === 'string' ? item.promo : '',
        reasonKendala: typeof item.reasonKendala === 'string' ? item.reasonKendala : '',
        actionPlan: typeof item.actionPlan === 'string' ? item.actionPlan : '',
        updatedAt: typeof item.updatedAt === 'string' ? item.updatedAt : '',
      });
    });
  }

  return Array.from(byId.values());
}

export function loadCvInventoryOps(): CvInventoryOpsEntry[] {
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(STORAGE_KEY_CV_INVENTORY_OPS);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          return normalizeCvInventoryOpsList(parsed);
        }
      }
    }
  } catch (err) {
    console.error('Error loading CV inventory ops:', err);
  }
  const initial = normalizeCvInventoryOpsList([]);
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY_CV_INVENTORY_OPS, JSON.stringify(initial));
    }
  } catch {}
  return initial;
}

export function saveCvInventoryOps(entries: CvInventoryOpsEntry[]): void {
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(STORAGE_KEY_CV_INVENTORY_OPS, JSON.stringify(entries));
    }
  } catch (err) {
    console.error('Error saving CV inventory ops:', err);
  }
}

export function calcCvLiveSalesQtyMap(records: CvDailyRecord[]): {
  weekly: Map<number, { GRAPHIS: number; SIMPLICITY: number; TOTAL: number }>;
  monthly: Map<string, { GRAPHIS: number; SIMPLICITY: number; TOTAL: number }>;
} {
  const byDate = new Map<string, CvDailyRecord>();
  records.forEach((r) => byDate.set(r.tanggal, r));

  const weekly = new Map<number, { GRAPHIS: number; SIMPLICITY: number; TOTAL: number }>();
  ANNUAL_2026_WEEKS.forEach((w) => {
    let gQty = 0;
    let sQty = 0;
    w.slashDates.forEach((dt) => {
      const rec = byDate.get(dt);
      if (rec) {
        gQty += (Number(rec.graphis?.pagi?.qty) || 0) + (Number(rec.graphis?.siang?.qty) || 0);
        sQty += (Number(rec.simplicity?.pagi?.qty) || 0) + (Number(rec.simplicity?.siang?.qty) || 0);
      }
    });
    weekly.set(w.weekNumber, { GRAPHIS: gQty, SIMPLICITY: sQty, TOTAL: gQty + sQty });
  });

  const monthly = new Map<string, { GRAPHIS: number; SIMPLICITY: number; TOTAL: number }>();
  Object.keys(MONTH_NAMES_ID).forEach((mKey) => {
    let gQty = 0;
    let sQty = 0;
    records.forEach((rec) => {
      const recMonth = rec.monthKey || (rec.tanggal ? rec.tanggal.split('/')[1] : '');
      if (recMonth === mKey) {
        gQty += (Number(rec.graphis?.pagi?.qty) || 0) + (Number(rec.graphis?.siang?.qty) || 0);
        sQty += (Number(rec.simplicity?.pagi?.qty) || 0) + (Number(rec.simplicity?.siang?.qty) || 0);
      }
    });
    monthly.set(mKey, { GRAPHIS: gQty, SIMPLICITY: sQty, TOTAL: gQty + sQty });
  });

  return { weekly, monthly };
}

