/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * 
 * REPORT KATEGORI CV, DIVISI 1 (YOUNG FASHION, LADIES & INTIMATE)
 * Dept : 6201 - Brand : Graphis
 * Dept : 6244 - Brand : Simplicity
 * Dedicated separate format with: Sales Harian, Weekly, Bulanan & Tahunan
 */

import React, { useState, useMemo } from 'react';
import { createPortal } from 'react-dom';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  calcCvDaySummary,
  calcCvWeeklySummary,
  calcCvMonthlySummary,
  calcCvAnnualSummary,
  calcCvBrandShiftTotals,
  loadCvInventoryOps,
  MONTH_NAMES_ID,
} from '../data/cvDiv1Data';
import { ANNUAL_2026_WEEKS, formatNumberId, getIndoDayName } from '../data/initialData';
import { exportCvDiv1ToExcel } from '../utils/exportCvDiv1Excel';
import {
  slashDateToIso,
  isoToSlashDate,
  formatIndoDateLong,
} from '../utils/dateUtils';
import { copyTextToClipboard } from '../utils/clipboard';
import { getGooglePublicCvReportUrl } from '../utils/shareUrl';
import {
  Calendar,
  CalendarDays,
  CalendarRange,
  FileSpreadsheet,
  Download,
  Share2,
  Save,
  PenSquare,
  CheckCircle2,
  TrendingUp,
  Tag,
  Sparkles,
  ChevronLeft,
  ChevronRight,
  Filter,
  BarChart3,
  Layers,
  ShoppingBag,
  Percent,
  X,
  Plus,
  Receipt,
  FileText,
  Printer,
  ArrowUpRight,
  ArrowDownRight,
  Eye,
  Link2,
} from 'lucide-react';
import { PhotoStyleSalesDiagram } from './PhotoStyleSalesDiagram';
import { GrowthAndContributionBarDiagram } from './GrowthAndContributionBarDiagram';

export type CvHorizon = 'harian' | 'weekly' | 'bulanan';

interface ReportCvDiv1ViewProps {
  records: CvDailyRecord[];
  onUpdateRecords: (updatedRecords: CvDailyRecord[]) => void;
  inventoryOps?: CvInventoryOpsEntry[];
  onUpdateInventoryOps?: (updatedOps: CvInventoryOpsEntry[]) => void;
  isAdmin: boolean;
  onTriggerSave?: () => void;
  activeHorizon?: CvHorizon;
  onChangeHorizon?: (horizon: CvHorizon) => void;
}

export const ReportCvDiv1View: React.FC<ReportCvDiv1ViewProps> = ({
  records,
  onUpdateRecords,
  inventoryOps,
  onUpdateInventoryOps,
  isAdmin,
  onTriggerSave,
  activeHorizon,
  onChangeHorizon,
}) => {
  const effectiveInventoryOps = useMemo(
    () => inventoryOps ?? loadCvInventoryOps(),
    [inventoryOps]
  );

  const [internalHorizon, setInternalHorizon] = useState<CvHorizon>('harian');
  const horizon: CvHorizon = activeHorizon !== undefined ? activeHorizon : internalHorizon;
  const setHorizon = (nextH: CvHorizon) => {
    setInternalHorizon(nextH);
    if (onChangeHorizon) onChangeHorizon(nextH);
  };
  const [selectedTanggal, setSelectedTanggal] = useState<string>('01/11/2026');
  const [selectedWeekNumber, setSelectedWeekNumber] = useState<number>(45);
  const [selectedMonthKey, setSelectedMonthKey] = useState<string>('11');
  const [brandFilter, setBrandFilter] = useState<'ALL' | 'GRAPHIS' | 'SIMPLICITY'>('ALL');
  const [dailySubView, setDailySubView] = useState<'matrix' | 'slip'>('matrix');

  // Weekly Tab Format Terbaru State
  const [weeklyFilterMonth, setWeeklyFilterMonth] = useState<string>('11');
  const [weeklyCalendarIso, setWeeklyCalendarIso] = useState<string>('');
  const [weeklyScope, setWeeklyScope] = useState<'month' | 'all'>('month');
  const [weeklyBrandFilter, setWeeklyBrandFilter] = useState<'TOTAL' | 'GRAPHIS' | 'SIMPLICITY' | 'SPLIT'>('TOTAL');
  const [selectedWeeklyRowNumber, setSelectedWeeklyRowNumber] = useState<number>(45);

  // Monthly Tab Format Terbaru State (Format Per Week & Format Perbulan 1 Tahun)
  const [monthlyViewFormat, setMonthlyViewFormat] = useState<'both' | 'per_week' | 'per_bulan' | 'harian'>('both');
  const [monthlyBrandFilter, setMonthlyBrandFilter] = useState<'TOTAL' | 'GRAPHIS' | 'SIMPLICITY' | 'SPLIT'>('TOTAL');

  // Copy toast status
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  // Edit Modal State (Input Sales Report CV - Tanpa Petugas & Shift & Tanpa Reason)
  interface CvEditFormState {
    tanggal: string;
    selectedBrand: 'GRAPHIS' | 'SIMPLICITY' | 'ALL';
    graphisSales: number | '';
    graphisQty: number | '';
    graphisNota: number | '';
    simplicitySales: number | '';
    simplicityQty: number | '';
    simplicityNota: number | '';
  }

  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [isTableEditMode, setIsTableEditMode] = useState(false);
  const [tableDateFilterIso, setTableDateFilterIso] = useState<string>('');
  const [editTanggal, setEditTanggal] = useState<string>('01/11/2026');
  const [editForm, setEditForm] = useState<CvEditFormState>({
    tanggal: '01/11/2026',
    selectedBrand: 'GRAPHIS',
    graphisSales: '',
    graphisQty: '',
    graphisNota: '',
    simplicitySales: '',
    simplicityQty: '',
    simplicityNota: '',
  });

  // Current records map
  const recordMap = useMemo(() => {
    const map = new Map<string, CvDailyRecord>();
    records.forEach((r) => map.set(r.tanggal, r));
    return map;
  }, [records]);

  // Selected Day Record
  const currentDayRecord = useMemo(() => {
    return (
      recordMap.get(selectedTanggal) ||
      records[0] || {
        tanggal: selectedTanggal,
        isoDate: '',
        dayName: getIndoDayName(selectedTanggal),
        weekNumber: selectedWeekNumber,
        monthKey: selectedMonthKey,
        graphis: {
          pagi: { sales: 0, qty: 0, trx: 0 },
          siang: { sales: 0, qty: 0, trx: 0 },
        },
        simplicity: {
          pagi: { sales: 0, qty: 0, trx: 0 },
          siang: { sales: 0, qty: 0, trx: 0 },
        },
      }
    );
  }, [recordMap, selectedTanggal, records, selectedWeekNumber, selectedMonthKey]);

  const currentDaySummary = useMemo(() => {
    return calcCvDaySummary(currentDayRecord);
  }, [currentDayRecord]);

  // Selected Weekly Summary
  const currentWeeklySummary = useMemo(() => {
    return calcCvWeeklySummary(records, selectedWeekNumber);
  }, [records, selectedWeekNumber]);

  // Computed Weekly Table Rows (Format Terbaru dengan Week, Tanggal, Senin s/d Minggu, Total Sales, Qty, Nota, Price Point, Big Size, Kontribusi, Growth)
  // Dilengkapi pemisahan sales week terpisah masing-masing brand (Graphis, Simplicity, & Total) serta tanggal di atas cell hari
  const weeklyTableData = useMemo(() => {
    // 1. Determine list of weeks based on weeklyScope & weeklyFilterMonth
    let weeksList = ANNUAL_2026_WEEKS;
    if (weeklyScope === 'month') {
      const filtered = ANNUAL_2026_WEEKS.filter(
        (w) => w.monthKey === weeklyFilterMonth || w.slashDates.some((dt) => dt.split('/')[1] === weeklyFilterMonth)
      );
      if (filtered.length > 0) weeksList = filtered;
      else weeksList = ANNUAL_2026_WEEKS.filter((w) => w.monthKey === weeklyFilterMonth);
    }

    // Maps across all 53 weeks for WoW growth calculation
    const graphisWeekSalesMap = new Map<number, number>();
    const simplicityWeekSalesMap = new Map<number, number>();
    const totalWeekSalesMap = new Map<number, number>();

    ANNUAL_2026_WEEKS.forEach((w) => {
      let gW = 0;
      let sW = 0;
      w.slashDates.forEach((dt) => {
        const rec = recordMap.get(dt);
        if (rec) {
          const gTot = calcCvBrandShiftTotals(rec.graphis);
          const sTot = calcCvBrandShiftTotals(rec.simplicity);
          gW += gTot.sales;
          sW += sTot.sales;
        }
      });
      graphisWeekSalesMap.set(w.weekNumber, gW);
      simplicityWeekSalesMap.set(w.weekNumber, sW);
      totalWeekSalesMap.set(w.weekNumber, gW + sW);
    });

    // Compute period grand total sales for each brand
    let periodGraphisGrandSales = 0;
    let periodSimplicityGrandSales = 0;
    let periodTotalGrandSales = 0;
    weeksList.forEach((w) => {
      periodGraphisGrandSales += graphisWeekSalesMap.get(w.weekNumber) || 0;
      periodSimplicityGrandSales += simplicityWeekSalesMap.get(w.weekNumber) || 0;
      periodTotalGrandSales += totalWeekSalesMap.get(w.weekNumber) || 0;
    });

    // Helper to build BrandRowMetrics
    const createBrandMetrics = (
      brandKey: 'GRAPHIS' | 'SIMPLICITY' | 'TOTAL',
      w: typeof ANNUAL_2026_WEEKS[0],
      periodBrandGrandSales: number,
      brandSalesMap: Map<number, number>
    ) => {
      const brandLabel =
        brandKey === 'GRAPHIS'
          ? 'Dept 6201 · Graphis'
          : brandKey === 'SIMPLICITY'
          ? 'Dept 6244 · Simplicity'
          : 'Total CV (Gabungan)';
      const deptCode = brandKey === 'GRAPHIS' ? '6201' : brandKey === 'SIMPLICITY' ? '6244' : 'DIV-1';

      const days = w.slashDates.map((dt) => {
        const dName = getIndoDayName(dt);
        const isWeekend = dName === 'Sabtu' || dName === 'Minggu';
        const dateShort = dt.slice(0, 5); // e.g. "02/11"
        const rec = recordMap.get(dt);
        let sales = 0;
        let qty = 0;
        let trx = 0;

        if (rec) {
          const gTot = calcCvBrandShiftTotals(rec.graphis);
          const sTot = calcCvBrandShiftTotals(rec.simplicity);
          if (brandKey === 'GRAPHIS') {
            sales = gTot.sales;
            qty = gTot.qty;
            trx = gTot.trx;
          } else if (brandKey === 'SIMPLICITY') {
            sales = sTot.sales;
            qty = sTot.qty;
            trx = sTot.trx;
          } else {
            sales = gTot.sales + sTot.sales;
            qty = gTot.qty + sTot.qty;
            trx = gTot.trx + sTot.trx;
          }
        }

        return {
          tanggal: dt,
          dateShort,
          dayName: dName,
          sales,
          qty,
          trx,
          isWeekend,
        };
      });

      const totalSalesWeek = days.reduce((sum, d) => sum + d.sales, 0);
      const totalQtyWeek = days.reduce((sum, d) => sum + d.qty, 0);
      const totalTrxWeek = days.reduce((sum, d) => sum + d.trx, 0);
      const pricePoint = totalQtyWeek > 0 ? Math.round(totalSalesWeek / totalQtyWeek) : 0;
      const bigSizeSales = Math.round(totalSalesWeek * 0.20);
      const bigSizeQty = Math.round(totalQtyWeek * 0.18);
      const bigSizePricePoint = bigSizeQty > 0 ? Math.round(bigSizeSales / bigSizeQty) : 0;
      const kontribusiPct = periodBrandGrandSales > 0 ? (totalSalesWeek / periodBrandGrandSales) * 100 : 0;

      const prevWeekSales = brandSalesMap.get(w.weekNumber - 1);
      let growthPct: number | null = null;
      if (prevWeekSales !== undefined && prevWeekSales > 0) {
        growthPct = ((totalSalesWeek - prevWeekSales) / prevWeekSales) * 100;
      }

      return {
        brandKey,
        brandLabel,
        deptCode,
        days,
        totalSalesWeek,
        totalQtyWeek,
        totalTrxWeek,
        pricePoint,
        bigSizeSales,
        bigSizeQty,
        bigSizePricePoint,
        kontribusiPct,
        growthPct,
      };
    };

    // 2. Map each week in weeksList to full row data with all brands
    const rows = weeksList.map((w, idx) => {
      const graphis = createBrandMetrics('GRAPHIS', w, periodGraphisGrandSales, graphisWeekSalesMap);
      const simplicity = createBrandMetrics('SIMPLICITY', w, periodSimplicityGrandSales, simplicityWeekSalesMap);
      const total = createBrandMetrics('TOTAL', w, periodTotalGrandSales, totalWeekSalesMap);

      const weekLabel = `WEEK ${idx + 1}`;

      return {
        weekNumber: w.weekNumber,
        weekIndexInMonth: idx + 1,
        weekLabel,
        periodeText: w.periodeText,
        monthKey: w.monthKey,
        monthName: MONTH_NAMES_ID[w.monthKey] || 'November',
        graphis,
        simplicity,
        total,
        // Legacy convenience accessors
        days: total.days,
        totalSalesWeek: total.totalSalesWeek,
        totalQtyWeek: total.totalQtyWeek,
        totalTrxWeek: total.totalTrxWeek,
        pricePoint: total.pricePoint,
        bigSizeSales: total.bigSizeSales,
        bigSizeQty: total.bigSizeQty,
        bigSizePricePoint: total.bigSizePricePoint,
        kontribusiPct: total.kontribusiPct,
        growthPct: total.growthPct,
      };
    });

    // Helper to summarize rows for a brand
    const summarizeBrand = (brandKey: 'graphis' | 'simplicity' | 'total') => {
      const brandRows = rows.map((r) => r[brandKey]);
      const senin = brandRows.reduce((s, b) => s + (b.days[0]?.sales || 0), 0);
      const selasa = brandRows.reduce((s, b) => s + (b.days[1]?.sales || 0), 0);
      const rabu = brandRows.reduce((s, b) => s + (b.days[2]?.sales || 0), 0);
      const kamis = brandRows.reduce((s, b) => s + (b.days[3]?.sales || 0), 0);
      const jumat = brandRows.reduce((s, b) => s + (b.days[4]?.sales || 0), 0);
      const sabtu = brandRows.reduce((s, b) => s + (b.days[5]?.sales || 0), 0);
      const minggu = brandRows.reduce((s, b) => s + (b.days[6]?.sales || 0), 0);
      const grandTotalSales = brandRows.reduce((s, b) => s + b.totalSalesWeek, 0);
      const grandTotalQty = brandRows.reduce((s, b) => s + b.totalQtyWeek, 0);
      const grandTotalTrx = brandRows.reduce((s, b) => s + b.totalTrxWeek, 0);
      const grandTotalBigSizeSales = brandRows.reduce((s, b) => s + b.bigSizeSales, 0);
      const grandTotalBigSizeQty = brandRows.reduce((s, b) => s + b.bigSizeQty, 0);
      const avgPricePoint = grandTotalQty > 0 ? Math.round(grandTotalSales / grandTotalQty) : 0;
      const gRows = brandRows.filter((b) => b.growthPct !== null);
      const avgGrowthPct = gRows.length > 0 ? gRows.reduce((s, b) => s + (b.growthPct || 0), 0) / gRows.length : 0;

      return {
        senin,
        selasa,
        rabu,
        kamis,
        jumat,
        sabtu,
        minggu,
        grandTotalSales,
        grandTotalQty,
        grandTotalTrx,
        grandTotalBigSizeSales,
        grandTotalBigSizeQty,
        avgPricePoint,
        avgGrowthPct,
      };
    };

    const graphisSummary = summarizeBrand('graphis');
    const simplicitySummary = summarizeBrand('simplicity');
    const totalSummary = summarizeBrand('total');

    return {
      rows,
      summary: {
        ...totalSummary,
        graphis: graphisSummary,
        simplicity: simplicitySummary,
        total: totalSummary,
      },
      periodGrandSales: periodTotalGrandSales,
      periodGraphisGrandSales,
      periodSimplicityGrandSales,
    };
  }, [ANNUAL_2026_WEEKS, recordMap, weeklyScope, weeklyFilterMonth]);

  // Selected Monthly Summary
  const currentMonthlySummary = useMemo(() => {
    return calcCvMonthlySummary(records, selectedMonthKey, 2026);
  }, [records, selectedMonthKey]);

  // Annual Summary
  const currentAnnualSummary = useMemo(() => {
    return calcCvAnnualSummary(records, 2026);
  }, [records]);

  // Monthly Weeks Detailed Breakdown for Format Per Week
  const monthlyWeeksData = useMemo(() => {
    const matchingWeeks = ANNUAL_2026_WEEKS.filter((w) => {
      return (
        w.monthKey === selectedMonthKey ||
        w.slashDates.some((d) => d.split('/')[1] === selectedMonthKey)
      );
    });

    let cumSalesMonth = 0;

    const weeks = matchingWeeks.map((w, index) => {
      let gSales = 0, gQty = 0, gTrx = 0;
      let sSales = 0, sQty = 0, sTrx = 0;
      let daysCountInMonth = 0;

      w.slashDates.forEach((dt) => {
        const inThisMonth = dt.split('/')[1] === selectedMonthKey;
        if (inThisMonth) daysCountInMonth++;
        const rec = recordMap.get(dt);
        if (rec) {
          const gs = (rec.graphis?.pagi?.sales || 0) + (rec.graphis?.siang?.sales || 0);
          const gq = (rec.graphis?.pagi?.qty || 0) + (rec.graphis?.siang?.qty || 0);
          const gt = (rec.graphis?.pagi?.trx || 0) + (rec.graphis?.siang?.trx || 0);

          const ss = (rec.simplicity?.pagi?.sales || 0) + (rec.simplicity?.siang?.sales || 0);
          const sq = (rec.simplicity?.pagi?.qty || 0) + (rec.simplicity?.siang?.qty || 0);
          const st = (rec.simplicity?.pagi?.trx || 0) + (rec.simplicity?.siang?.trx || 0);

          gSales += gs;
          gQty += gq;
          gTrx += gt;
          sSales += ss;
          sQty += sq;
          sTrx += st;
        }
      });

      const totalSales = gSales + sSales;
      const totalQty = gQty + sQty;
      const totalTrx = gTrx + sTrx;
      cumSalesMonth += totalSales;

      const gPricePoint = gQty > 0 ? Math.round(gSales / gQty) : 0;
      const sPricePoint = sQty > 0 ? Math.round(sSales / sQty) : 0;
      const totPricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : 0;
      const gShare = totalSales > 0 ? (gSales / totalSales) * 100 : 0;
      const sShare = totalSales > 0 ? (sSales / totalSales) * 100 : 0;
      const avgDaily = daysCountInMonth > 0 ? Math.round(totalSales / daysCountInMonth) : Math.round(totalSales / 7);
      const contribution = currentMonthlySummary.total.sales > 0 ? (totalSales / currentMonthlySummary.total.sales) * 100 : 0;

      return {
        weekNumber: w.weekNumber,
        weekIndexInMonth: index + 1,
        periodeText: w.periodeText,
        daysCountInMonth,
        graphis: {
          sales: gSales,
          qty: gQty,
          trx: gTrx,
          pricePoint: gPricePoint,
          sharePct: gShare,
        },
        simplicity: {
          sales: sSales,
          qty: sQty,
          trx: sTrx,
          pricePoint: sPricePoint,
          sharePct: sShare,
        },
        total: {
          sales: totalSales,
          qty: totalQty,
          trx: totalTrx,
          pricePoint: totPricePoint,
        },
        avgDaily,
        contribution,
        cumSalesMonth,
      };
    });

    const sumGraphisSales = weeks.reduce((acc, w) => acc + w.graphis.sales, 0);
    const sumGraphisQty = weeks.reduce((acc, w) => acc + w.graphis.qty, 0);
    const sumGraphisTrx = weeks.reduce((acc, w) => acc + w.graphis.trx, 0);
    const sumSimplicitySales = weeks.reduce((acc, w) => acc + w.simplicity.sales, 0);
    const sumSimplicityQty = weeks.reduce((acc, w) => acc + w.simplicity.qty, 0);
    const sumSimplicityTrx = weeks.reduce((acc, w) => acc + w.simplicity.trx, 0);
    const sumTotalSales = sumGraphisSales + sumSimplicitySales;
    const sumTotalQty = sumGraphisQty + sumSimplicityQty;
    const sumTotalTrx = sumGraphisTrx + sumSimplicityTrx;

    return {
      weeks,
      summary: {
        graphis: {
          sales: sumGraphisSales,
          qty: sumGraphisQty,
          trx: sumGraphisTrx,
          pricePoint: sumGraphisQty > 0 ? Math.round(sumGraphisSales / sumGraphisQty) : 0,
          sharePct: sumTotalSales > 0 ? (sumGraphisSales / sumTotalSales) * 100 : 0,
        },
        simplicity: {
          sales: sumSimplicitySales,
          qty: sumSimplicityQty,
          trx: sumSimplicityTrx,
          pricePoint: sumSimplicityQty > 0 ? Math.round(sumSimplicitySales / sumSimplicityQty) : 0,
          sharePct: sumTotalSales > 0 ? (sumSimplicitySales / sumTotalSales) * 100 : 0,
        },
        total: {
          sales: sumTotalSales,
          qty: sumTotalQty,
          trx: sumTotalTrx,
          pricePoint: sumTotalQty > 0 ? Math.round(sumTotalSales / sumTotalQty) : 0,
        },
      },
    };
  }, [ANNUAL_2026_WEEKS, recordMap, selectedMonthKey, currentMonthlySummary.total.sales]);

  // Format Perbulan dalam 1 Tahun (Januari s/d Desember 2026)
  const monthlyAnnualTableData = useMemo(() => {
    let cumSales = 0;
    const list = currentAnnualSummary.monthlySummaries;
    const annualTotalSales = currentAnnualSummary.total.sales;

    return list.map((m, idx) => {
      cumSales += m.total.sales;
      const prev = idx > 0 ? list[idx - 1] : null;
      let momGrowth: number | null = null;
      if (prev && prev.total.sales > 0) {
        momGrowth = ((m.total.sales - prev.total.sales) / prev.total.sales) * 100;
      }
      const contribution = annualTotalSales > 0 ? (m.total.sales / annualTotalSales) * 100 : 0;

      return {
        ...m,
        index: idx + 1,
        momGrowth,
        contribution,
        cumSales,
        isSelected: m.monthKey === selectedMonthKey,
      };
    });
  }, [currentAnnualSummary, selectedMonthKey]);

  // Show Toast
  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  // Helper to extract record data for form
  const getRecordDataForForm = (
    tgl: string,
    preferredBrand: 'GRAPHIS' | 'SIMPLICITY' | 'ALL' = 'GRAPHIS'
  ): CvEditFormState => {
    const rec = recordMap.get(tgl);
    if (!rec) {
      return {
        tanggal: tgl,
        selectedBrand: preferredBrand,
        graphisSales: '',
        graphisQty: '',
        graphisNota: '',
        simplicitySales: '',
        simplicityQty: '',
        simplicityNota: '',
      };
    }
    const gSales = (rec.graphis?.pagi?.sales || 0) + (rec.graphis?.siang?.sales || 0);
    const gQty = (rec.graphis?.pagi?.qty || 0) + (rec.graphis?.siang?.qty || 0);
    const gNota = (rec.graphis?.pagi?.trx || 0) + (rec.graphis?.siang?.trx || 0);

    const sSales = (rec.simplicity?.pagi?.sales || 0) + (rec.simplicity?.siang?.sales || 0);
    const sQty = (rec.simplicity?.pagi?.qty || 0) + (rec.simplicity?.siang?.qty || 0);
    const sNota = (rec.simplicity?.pagi?.trx || 0) + (rec.simplicity?.siang?.trx || 0);

    return {
      tanggal: tgl,
      selectedBrand: preferredBrand,
      graphisSales: gSales > 0 ? gSales : '',
      graphisQty: gQty > 0 ? gQty : '',
      graphisNota: gNota > 0 ? gNota : '',
      simplicitySales: sSales > 0 ? sSales : '',
      simplicityQty: sQty > 0 ? sQty : '',
      simplicityNota: sNota > 0 ? sNota : '',
    };
  };

  // Open Edit Modal for a specific day
  const handleOpenEditModal = (tgl: string, preferredBrand?: 'GRAPHIS' | 'SIMPLICITY' | 'ALL') => {
    const targetBrand = preferredBrand || (brandFilter !== 'ALL' ? brandFilter : 'GRAPHIS');
    const formData = getRecordDataForForm(tgl, targetBrand);
    setEditTanggal(tgl);
    setEditForm(formData);
    setIsEditModalOpen(true);
  };

  // Switch date inside the modal
  const handleDateChangeInModal = (newTgl: string) => {
    const formData = getRecordDataForForm(newTgl, editForm.selectedBrand);
    setEditTanggal(newTgl);
    setEditForm(formData);
  };

  const handlePrevDayInModal = () => {
    const idx = records.findIndex((r) => r.tanggal === editTanggal);
    if (idx > 0) {
      handleDateChangeInModal(records[idx - 1].tanggal);
    }
  };

  const handleNextDayInModal = () => {
    const idx = records.findIndex((r) => r.tanggal === editTanggal);
    if (idx < records.length - 1) {
      handleDateChangeInModal(records[idx + 1].tanggal);
    }
  };

  // Save Edit Form (Direct Sales, Qty, Nota per Brand, No Petugas, No Shift, No Reason)
  const handleSaveEdit = (e: React.FormEvent) => {
    e.preventDefault();
    let nextRecords = [...records];
    const existingIdx = nextRecords.findIndex((r) => r.tanggal === editTanggal);

    let baseRec: CvDailyRecord;
    if (existingIdx !== -1) {
      baseRec = { ...nextRecords[existingIdx] };
    } else {
      const dd = editTanggal.split('/')[0];
      const mm = editTanggal.split('/')[1] || selectedMonthKey;
      baseRec = {
        tanggal: editTanggal,
        isoDate: `2026-${mm}-${dd}`,
        dayName: getIndoDayName(editTanggal),
        weekNumber: selectedWeekNumber,
        monthKey: mm,
        graphis: { pagi: { sales: null, qty: null, trx: null }, siang: { sales: null, qty: null, trx: null } },
        simplicity: { pagi: { sales: null, qty: null, trx: null }, siang: { sales: null, qty: null, trx: null } },
      };
    }

    const updatedRec: CvDailyRecord = {
      ...baseRec,
      graphis: {
        pagi: {
          sales: editForm.graphisSales !== '' ? Number(editForm.graphisSales) : null,
          qty: editForm.graphisQty !== '' ? Number(editForm.graphisQty) : null,
          trx: editForm.graphisNota !== '' ? Number(editForm.graphisNota) : null,
        },
        siang: {
          sales: null,
          qty: null,
          trx: null,
        },
      },
      simplicity: {
        pagi: {
          sales: editForm.simplicitySales !== '' ? Number(editForm.simplicitySales) : null,
          qty: editForm.simplicityQty !== '' ? Number(editForm.simplicityQty) : null,
          trx: editForm.simplicityNota !== '' ? Number(editForm.simplicityNota) : null,
        },
        siang: {
          sales: null,
          qty: null,
          trx: null,
        },
      },
    };

    if (existingIdx !== -1) {
      nextRecords[existingIdx] = updatedRec;
    } else {
      nextRecords.push(updatedRec);
    }

    onUpdateRecords(nextRecords);
    setIsEditModalOpen(false);
    showToast(`Data Sales CV Divisi 1 tanggal ${editTanggal} berhasil disimpan!`);
    if (onTriggerSave) onTriggerSave();
  };

  // Direct manual editing inside the table format
  const handleDirectCellChange = (
    tanggal: string,
    field: 'gSales' | 'gQty' | 'gNota' | 'sSales' | 'sQty' | 'sNota' | 'reason',
    value: string | number
  ) => {
    let nextRecords = [...records];
    const existingIdx = nextRecords.findIndex((r) => r.tanggal === tanggal);

    let targetRec: CvDailyRecord;
    if (existingIdx !== -1) {
      targetRec = { ...nextRecords[existingIdx] };
    } else {
      const dd = tanggal.split('/')[0];
      const mm = tanggal.split('/')[1] || selectedMonthKey;
      targetRec = {
        tanggal,
        isoDate: `2026-${mm}-${dd}`,
        dayName: getIndoDayName(tanggal),
        weekNumber: selectedWeekNumber,
        monthKey: mm,
        reasonEvent: '',
        graphis: {
          pagi: { sales: null, qty: null, trx: null },
          siang: { sales: null, qty: null, trx: null },
        },
        simplicity: {
          pagi: { sales: null, qty: null, trx: null },
          siang: { sales: null, qty: null, trx: null },
        },
      };
    }

    if (field === 'reason') {
      targetRec.reasonEvent = String(value);
    } else {
      const num = value === '' ? null : Number(value);
      if (field === 'gSales') {
        targetRec.graphis = {
          ...targetRec.graphis,
          pagi: { ...targetRec.graphis.pagi, sales: num },
          siang: { ...targetRec.graphis.siang, sales: null },
        };
      } else if (field === 'gQty') {
        targetRec.graphis = {
          ...targetRec.graphis,
          pagi: { ...targetRec.graphis.pagi, qty: num },
          siang: { ...targetRec.graphis.siang, qty: null },
        };
      } else if (field === 'gNota') {
        targetRec.graphis = {
          ...targetRec.graphis,
          pagi: { ...targetRec.graphis.pagi, trx: num },
          siang: { ...targetRec.graphis.siang, trx: null },
        };
      } else if (field === 'sSales') {
        targetRec.simplicity = {
          ...targetRec.simplicity,
          pagi: { ...targetRec.simplicity.pagi, sales: num },
          siang: { ...targetRec.simplicity.siang, sales: null },
        };
      } else if (field === 'sQty') {
        targetRec.simplicity = {
          ...targetRec.simplicity,
          pagi: { ...targetRec.simplicity.pagi, qty: num },
          siang: { ...targetRec.simplicity.siang, qty: null },
        };
      } else if (field === 'sNota') {
        targetRec.simplicity = {
          ...targetRec.simplicity,
          pagi: { ...targetRec.simplicity.pagi, trx: num },
          siang: { ...targetRec.simplicity.siang, trx: null },
        };
      }
    }

    if (existingIdx !== -1) {
      nextRecords[existingIdx] = targetRec;
    } else {
      nextRecords.push(targetRec);
    }

    onUpdateRecords(nextRecords);
    if (onTriggerSave) onTriggerSave();
  };

  // Salin WhatsApp Text
  const handleCopyWhatsApp = () => {
    let text = '';
    if (horizon === 'harian') {
      const s = currentDaySummary;
      const monthNameUpper = (MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase();
      text = [
        `*LAPORAN SALES DAILY CV — DIVISI 1*`,
        `*BAZAR MATAHARI 311 TANJUNG — PT MDS RETAILING TBK*`,
        `----------------------------------------`,
        `• Periode      : ${monthNameUpper} 2026`,
        `• Tanggal      : ${s.tanggal}`,
        `• Hari         : ${s.dayName.toUpperCase()} (Minggu ke-${s.weekNumber})`,
        `• Divisi       : DIVISI 1 (Young Fashion, Ladies & Intimate)`,
        `----------------------------------------`,
        `*1. DEPT 6201 · BRAND GRAPHIS*`,
        `  • Shift Pagi : Rp ${formatNumberId(currentDayRecord.graphis.pagi.sales)} | ${currentDayRecord.graphis.pagi.qty || 0} pcs | ${currentDayRecord.graphis.pagi.trx || 0} nota`,
        `  • Shift Siang: Rp ${formatNumberId(currentDayRecord.graphis.siang.sales)} | ${currentDayRecord.graphis.siang.qty || 0} pcs | ${currentDayRecord.graphis.siang.trx || 0} nota`,
        `  *>> TOTAL BRAND GRAPHIS : Rp ${formatNumberId(s.graphis.sales)}*`,
        `    (Qty: ${s.graphis.qty} pcs | Nota: ${s.graphis.trx} nota | ASP: Rp ${formatNumberId(s.graphis.pricePoint)} | Share: ${s.graphis.sharePct.toFixed(1)}%)`,
        ``,
        `*2. DEPT 6244 · BRAND SIMPLICITY*`,
        `  • Shift Pagi : Rp ${formatNumberId(currentDayRecord.simplicity.pagi.sales)} | ${currentDayRecord.simplicity.pagi.qty || 0} pcs | ${currentDayRecord.simplicity.pagi.trx || 0} nota`,
        `  • Shift Siang: Rp ${formatNumberId(currentDayRecord.simplicity.siang.sales)} | ${currentDayRecord.simplicity.siang.qty || 0} pcs | ${currentDayRecord.simplicity.siang.trx || 0} nota`,
        `  *>> TOTAL BRAND SIMPLICITY : Rp ${formatNumberId(s.simplicity.sales)}*`,
        `    (Qty: ${s.simplicity.qty} pcs | Nota: ${s.simplicity.trx} nota | ASP: Rp ${formatNumberId(s.simplicity.pricePoint)} | Share: ${s.simplicity.sharePct.toFixed(1)}%)`,
        `----------------------------------------`,
        `*RINGKASAN TOTAL BRAND:*`,
        `• Total Graphis (6201)    : Rp ${formatNumberId(s.graphis.sales)} (${s.graphis.qty} pcs / ${s.graphis.trx} nota)`,
        `• Total Simplicity (6244) : Rp ${formatNumberId(s.simplicity.sales)} (${s.simplicity.qty} pcs / ${s.simplicity.trx} nota)`,
        `----------------------------------------`,
        `*TOTAL DAILY CV DIVISI 1 (GABUNGAN):*`,
        `★ Grand Total Sales : Rp ${formatNumberId(s.total.sales)}`,
        `★ Grand Total Qty   : ${s.total.qty} pcs`,
        `★ Grand Total Nota  : ${s.total.trx} nota`,
        `★ Price Point (ASP) : Rp ${formatNumberId(s.total.pricePoint)}`,
        `★ Target Daily      : Rp ${formatNumberId(s.targetDaily)} (${s.achievementPct.toFixed(1)}%)`,
        s.reasonEvent ? `• Catatan/Event     : ${s.reasonEvent}` : '',
        `Petugas: Pagi [${s.petugasPagi || '-'}] | Siang [${s.petugasSiang || '-'}]`,
      ]
        .filter(Boolean)
        .join('\n');
    } else if (horizon === 'weekly') {
      const monthTitle = (MONTH_NAMES_ID[weeklyFilterMonth] || 'NOVEMBER').toUpperCase();

      const lines = [
        `*LAPORAN SALES WEEKLY CV — DIVISI 1 (FORMAT TERBARU)*`,
        `*BAZAR MATAHARI TANJUNG 311 — PT MDS RETAILING TBK*`,
        `----------------------------------------`,
        `• Periode       : ${weeklyScope === 'month' ? `${monthTitle} 2026` : 'TAHUN 2026 (SEMUA WEEKLY)'}`,
        `• Format        : Sales Week Terpisah Masing-masing Brand (Graphis 6201 & Simplicity 6244)`,
        `• Grand Total CV: Rp ${formatNumberId(weeklyTableData.summary.grandTotalSales)} (${weeklyTableData.summary.grandTotalQty} pcs | ${weeklyTableData.summary.grandTotalTrx} nota)`,
        `  - Dept 6201 Graphis    : Rp ${formatNumberId(weeklyTableData.summary.graphis.grandTotalSales)} (${weeklyTableData.summary.graphis.grandTotalQty} pcs)`,
        `  - Dept 6244 Simplicity : Rp ${formatNumberId(weeklyTableData.summary.simplicity.grandTotalSales)} (${weeklyTableData.summary.simplicity.grandTotalQty} pcs)`,
        `• Average ASP   : Rp ${formatNumberId(weeklyTableData.summary.avgPricePoint)}`,
        `• Big Size Est. : Rp ${formatNumberId(weeklyTableData.summary.grandTotalBigSizeSales)} (${weeklyTableData.summary.grandTotalBigSizeQty} pcs)`,
        `----------------------------------------`,
        `*REKAPITULASI SALES PER WEEK (SENIN S/D MINGGU):*`,
      ];

      weeklyTableData.rows.forEach((r) => {
        const growthStr = r.total.growthPct !== null ? `${r.total.growthPct >= 0 ? '+' : ''}${r.total.growthPct.toFixed(1)}%` : 'Baseline';
        lines.push(
          `★ *${r.weekLabel} (${r.periodeText})* [W-${r.weekNumber}]`,
          `  🌸 *Graphis (6201)*   : Rp ${formatNumberId(r.graphis.totalSalesWeek)} (${r.graphis.totalQtyWeek} pcs | ${r.graphis.totalTrxWeek} nota | ASP Rp ${formatNumberId(r.graphis.pricePoint)})`,
          `  🔷 *Simplicity (6244)*: Rp ${formatNumberId(r.simplicity.totalSalesWeek)} (${r.simplicity.totalQtyWeek} pcs | ${r.simplicity.totalTrxWeek} nota | ASP Rp ${formatNumberId(r.simplicity.pricePoint)})`,
          `  ★ *TOTAL CV GABUNGAN* : Rp ${formatNumberId(r.total.totalSalesWeek)} (${r.total.totalQtyWeek} pcs | ${r.total.totalTrxWeek} nota | BigSize Rp ${formatNumberId(r.total.bigSizeSales)} | Growth: ${growthStr})`,
          `  • *Detail Harian (Cell Hari & Tanggal):*`,
          `    - Senin (${r.total.days[0].dateShort})   : Graphis Rp ${formatNumberId(r.graphis.days[0].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[0].sales)} | Total Rp ${formatNumberId(r.total.days[0].sales)}`,
          `    - Selasa (${r.total.days[1].dateShort})  : Graphis Rp ${formatNumberId(r.graphis.days[1].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[1].sales)} | Total Rp ${formatNumberId(r.total.days[1].sales)}`,
          `    - Rabu (${r.total.days[2].dateShort})    : Graphis Rp ${formatNumberId(r.graphis.days[2].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[2].sales)} | Total Rp ${formatNumberId(r.total.days[2].sales)}`,
          `    - Kamis (${r.total.days[3].dateShort})   : Graphis Rp ${formatNumberId(r.graphis.days[3].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[3].sales)} | Total Rp ${formatNumberId(r.total.days[3].sales)}`,
          `    - Jumat (${r.total.days[4].dateShort})   : Graphis Rp ${formatNumberId(r.graphis.days[4].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[4].sales)} | Total Rp ${formatNumberId(r.total.days[4].sales)}`,
          `    - Sabtu (${r.total.days[5].dateShort})   : Graphis Rp ${formatNumberId(r.graphis.days[5].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[5].sales)} | Total Rp ${formatNumberId(r.total.days[5].sales)} (WKND)`,
          `    - Minggu (${r.total.days[6].dateShort})  : Graphis Rp ${formatNumberId(r.graphis.days[6].sales)} | Simplicity Rp ${formatNumberId(r.simplicity.days[6].sales)} | Total Rp ${formatNumberId(r.total.days[6].sales)} (WKND)`,
          ``
        );
      });

      lines.push(
        `----------------------------------------`,
        `*GRAND TOTAL SENIN S/D MINGGU (KONSOLIDASI):*`,
        `• Senin  : Rp ${formatNumberId(weeklyTableData.summary.senin)}`,
        `• Selasa : Rp ${formatNumberId(weeklyTableData.summary.selasa)}`,
        `• Rabu   : Rp ${formatNumberId(weeklyTableData.summary.rabu)}`,
        `• Kamis  : Rp ${formatNumberId(weeklyTableData.summary.kamis)}`,
        `• Jumat  : Rp ${formatNumberId(weeklyTableData.summary.jumat)}`,
        `• Sabtu  : Rp ${formatNumberId(weeklyTableData.summary.sabtu)}`,
        `• Minggu : Rp ${formatNumberId(weeklyTableData.summary.minggu)}`,
        `★ GRAND TOTAL CV : Rp ${formatNumberId(weeklyTableData.summary.grandTotalSales)} (${weeklyTableData.summary.grandTotalQty} pcs | ${weeklyTableData.summary.grandTotalTrx} nota)`
      );

      text = lines.join('\n');
    } else if (horizon === 'bulanan') {
      const m = currentMonthlySummary;
      const weeklyLines = monthlyWeeksData.weeks.map(
        (w) => `• Week ${w.weekIndexInMonth} (W-${w.weekNumber}): Rp ${formatNumberId(w.total.sales)} (G: Rp ${formatNumberId(w.graphis.sales)} | S: Rp ${formatNumberId(w.simplicity.sales)})`
      );

      text = [
        `*LAPORAN SALES MONTHLY CV — DIVISI 1*`,
        `*BAZAR MATAHARI TANJUNG — PT MDS RETAILING TBK*`,
        `Periode: ${m.monthName.toUpperCase()} 2026 (${m.daysCount} Hari Aktif)`,
        `----------------------------------------`,
        `*1. Dept 6201 - GRAPHIS* : Rp ${formatNumberId(m.graphis.sales)} (${m.graphis.qty} pcs | ${m.graphis.sharePct.toFixed(1)}%)`,
        `*2. Dept 6244 - SIMPLICITY*: Rp ${formatNumberId(m.simplicity.sales)} (${m.simplicity.qty} pcs | ${m.simplicity.sharePct.toFixed(1)}%)`,
        `----------------------------------------`,
        `*TOTAL BULAN ${m.monthName.toUpperCase()}: Rp ${formatNumberId(m.total.sales)}*`,
        `• Total Qty     : ${m.total.qty} pcs`,
        `• Total Trx     : ${m.total.trx} nota`,
        `• Rata-rata/hari: Rp ${formatNumberId(m.avgDailySales)}`,
        m.bestDay ? `• Best Sales Day: ${m.bestDay.dayName}, ${m.bestDay.tanggal} (Rp ${formatNumberId(m.bestDay.sales)})` : '',
        `----------------------------------------`,
        `*REKAPITULASI PER WEEK:*`,
        ...weeklyLines,
        `----------------------------------------`,
        `*AKUMULASI YTD 1 TAHUN (2026):*`,
        `★ Total Sales 1 Tahun : Rp ${formatNumberId(currentAnnualSummary.total.sales)} (${currentAnnualSummary.total.qty} pcs)`,
        `★ Rata-rata Bulanan   : Rp ${formatNumberId(Math.round(currentAnnualSummary.total.sales / 12))}/bulan`,
      ]
        .filter(Boolean)
        .join('\n');
    } else {
      const a = currentAnnualSummary;
      text = [
        `*LAPORAN SALES TAHUNAN CV — DIVISI 1 (2026)*`,
        `*BAZAR MATAHARI TANJUNG — PT MDS RETAILING TBK*`,
        `Total Setahun: Rp ${formatNumberId(a.total.sales)} (${a.total.qty} pcs)`,
        `Dept 6201 Graphis    : Rp ${formatNumberId(a.graphis.sales)} (${a.graphis.sharePct.toFixed(1)}%)`,
        `Dept 6244 Simplicity : Rp ${formatNumberId(a.simplicity.sales)} (${a.simplicity.sharePct.toFixed(1)}%)`,
        `Brand Terkuat        : ${a.dominantBrand}`,
        a.bestMonth ? `Bulan Tertinggi      : ${a.bestMonth.monthName} (Rp ${formatNumberId(a.bestMonth.sales)})` : '',
      ]
        .filter(Boolean)
        .join('\n');
    }

    void copyTextToClipboard(text).then((ok) => {
      showToast(
        ok
          ? 'Teks laporan WhatsApp berhasil disalin ke clipboard!'
          : 'Gagal menyalin teks ke clipboard.'
      );
    });
  };

  // Ekspor Excel
  const handleExportExcel = async () => {
    try {
      await exportCvDiv1ToExcel(
        records,
        selectedWeekNumber,
        selectedMonthKey,
        2026,
        effectiveInventoryOps
      );
      showToast('File Excel Report CV Divisi 1 (.xlsx) berhasil diunduh!');
    } catch (err) {
      console.error('Excel Export Error:', err);
      showToast('Gagal mengunduh file Excel.');
    }
  };

  return (
    <div className="space-y-5 pb-12">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 flex items-center gap-2 px-4 py-3 bg-slate-900 text-white rounded-xl shadow-2xl border border-amber-400 text-xs font-bold animate-fade-in">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Grand Top Banner: Category CV, Divisi 1 */}
      <div className="rounded-2xl bg-gradient-to-r from-purple-950 via-slate-900 to-indigo-950 text-white p-4 sm:p-5 shadow-sm border border-purple-800/40">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          <div className="space-y-1.5 min-w-0">
            <div className="flex flex-wrap items-center gap-2">
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-purple-500/25 border border-purple-400/40 text-purple-200">
                <Tag className="w-3 h-3 text-purple-300" />
                Kategori CV (Konsinyasi)
              </span>
              <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-bold bg-amber-400/20 border border-amber-400/30 text-amber-300">
                Bazar Tanjung · PT MDS Retailing Tbk
              </span>
            </div>

            <h1 className="text-lg sm:text-2xl font-black text-white tracking-tight flex items-center gap-2 flex-wrap">
              <span>DIVISI 1</span>
              <span className="text-purple-300 font-extrabold text-sm sm:text-base">
                — YOUNG FASHION, LADIES &amp; INTIMATE
              </span>
            </h1>

            <div className="flex flex-wrap items-center gap-2 text-xs">
              <span className="px-2.5 py-1 rounded-lg bg-rose-950/80 border border-rose-700/60 text-rose-200 font-bold inline-flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse" />
                Dept 6201 : Brand Graphis
              </span>
              <span className="px-2.5 py-1 rounded-lg bg-indigo-950/80 border border-indigo-700/60 text-indigo-200 font-bold inline-flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-indigo-500 animate-pulse" />
                Dept 6244 : Brand Simplicity
              </span>
            </div>
          </div>

          {/* Action Buttons: Export & Share */}
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {isAdmin && (
              <button
                type="button"
                onClick={() => handleOpenEditModal(selectedTanggal)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 transition-colors shadow-2xs cursor-pointer"
                title="Input sales report CV (Tanggal, Filter Brand, Sales, Qty, Nota, Reason)"
              >
                <PenSquare className="w-3.5 h-3.5" />
                <span>Input Sales CV</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => {
                const url = getGooglePublicCvReportUrl({ date: slashDateToIso(selectedTanggal) });
                void copyTextToClipboard(url).then((ok) => {
                  showToast(
                    ok
                      ? `Link Pengisian Report CV (${selectedTanggal}) berhasil disalin!`
                      : 'Gagal menyalin link Report CV.'
                  );
                });
              }}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-purple-600 hover:bg-purple-500 text-white transition-colors shadow-2xs cursor-pointer"
              title="Salin link pengisian form Report CV untuk tanggal aktif"
            >
              <Link2 className="w-3.5 h-3.5" />
              <span>Salin Link Isi CV</span>
            </button>

            <button
              type="button"
              onClick={handleCopyWhatsApp}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white transition-colors shadow-2xs cursor-pointer"
              title="Salin ringkasan format WhatsApp untuk dikirim ke grup laporan"
            >
              <Share2 className="w-3.5 h-3.5" />
              <span>Salin WhatsApp</span>
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white transition-colors shadow-2xs cursor-pointer"
              title="Unduh laporan lengkap 4 sheet (Harian, Weekly, Bulanan, Tahunan) dalam format Excel (.xlsx)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Ekspor Excel (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* 3-Horizon Switcher Navigation Tabs */}
        <div className="mt-4 pt-3 border-t border-purple-800/40 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center p-1 bg-slate-950/80 rounded-xl border border-slate-800">
            <button
              type="button"
              onClick={() => setHorizon('harian')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                horizon === 'harian'
                  ? 'bg-gradient-to-r from-amber-500 to-orange-500 text-slate-950 shadow-xs'
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
              <span>Sales Daily</span>
            </button>

            <button
              type="button"
              onClick={() => setHorizon('weekly')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                horizon === 'weekly'
                  ? 'bg-gradient-to-r from-rose-600 to-pink-600 text-white shadow-xs'
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Sales Weekly</span>
            </button>

            <button
              type="button"
              onClick={() => setHorizon('bulanan')}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                horizon === 'bulanan'
                  ? 'bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow-xs'
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              <CalendarDays className="w-3.5 h-3.5" />
              <span>Sales Monthly</span>
            </button>
          </div>

          {/* Sub-Brand Filter */}
          <div className="flex items-center gap-1.5 text-xs">
            <span className="text-slate-400 font-semibold hidden sm:inline">Tampilan Brand:</span>
            <div className="inline-flex p-0.5 bg-slate-950/80 rounded-lg border border-slate-800">
              <button
                type="button"
                onClick={() => setBrandFilter('ALL')}
                className={`px-2 py-1 rounded text-[11px] font-bold transition-colors ${
                  brandFilter === 'ALL'
                    ? 'bg-purple-600 text-white'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Semua (Divisi 1)
              </button>
              <button
                type="button"
                onClick={() => setBrandFilter('GRAPHIS')}
                className={`px-2 py-1 rounded text-[11px] font-bold transition-colors ${
                  brandFilter === 'GRAPHIS'
                    ? 'bg-rose-600 text-white'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Graphis (6201)
              </button>
              <button
                type="button"
                onClick={() => setBrandFilter('SIMPLICITY')}
                className={`px-2 py-1 rounded text-[11px] font-bold transition-colors ${
                  brandFilter === 'SIMPLICITY'
                    ? 'bg-indigo-600 text-white'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Simplicity (6244)
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ============================================================ */}
      {/* HORIZON 1: SALES DAILY (FORMAT LENGKAP: PERIODE, TGL, HARI,  */}
      {/* DIVISI, DEPT, BRAND, SALES, QTY, NOTA, TOTAL BRAND, TOTAL DAILY) */}
      {/* ============================================================ */}
      {horizon === 'harian' && (() => {
        // Build full 31 calendar rows (1 to 31) matching the WPS Office template
        const currentMonthRecords = records.filter(
          (r) => r.monthKey === selectedMonthKey || (r.tanggal && r.tanggal.split('/')[1] === selectedMonthKey)
        );
        const mapByDay = new Map<number, CvDailyRecord>();
        currentMonthRecords.forEach((r) => {
          const dayNum = parseInt(r.tanggal.split('/')[0], 10);
          if (!isNaN(dayNum)) {
            mapByDay.set(dayNum, r);
          }
        });

        const displayMonthRecords: CvDailyRecord[] = [];
        for (let d = 1; d <= 31; d++) {
          if (mapByDay.has(d)) {
            displayMonthRecords.push(mapByDay.get(d)!);
          } else {
            const dd = String(d).padStart(2, '0');
            const mm = selectedMonthKey;
            const tgl = `${dd}/${mm}/2026`;
            displayMonthRecords.push({
              tanggal: tgl,
              isoDate: `2026-${mm}-${dd}`,
              dayName: getIndoDayName(tgl),
              weekNumber: 45,
              monthKey: selectedMonthKey,
              graphis: {
                pagi: { sales: 0, qty: 0, trx: 0 },
                siang: { sales: 0, qty: 0, trx: 0 },
              },
              simplicity: {
                pagi: { sales: 0, qty: 0, trx: 0 },
                siang: { sales: 0, qty: 0, trx: 0 },
              },
            });
          }
        }

        // Calculate totals across this month for Graphis, Simplicity, and Grand Daily
        let mTotalGraphisSales = 0;
        let mTotalGraphisQty = 0;
        let mTotalGraphisTrx = 0;
        let mTotalSimplicitySales = 0;
        let mTotalSimplicityQty = 0;
        let mTotalSimplicityTrx = 0;

        displayMonthRecords.forEach((r) => {
          const gSales = (Number(r.graphis.pagi.sales) || 0) + (Number(r.graphis.siang.sales) || 0);
          const gQty = (Number(r.graphis.pagi.qty) || 0) + (Number(r.graphis.siang.qty) || 0);
          const gTrx = (Number(r.graphis.pagi.trx) || 0) + (Number(r.graphis.siang.trx) || 0);

          const sSales = (Number(r.simplicity.pagi.sales) || 0) + (Number(r.simplicity.siang.sales) || 0);
          const sQty = (Number(r.simplicity.pagi.qty) || 0) + (Number(r.simplicity.siang.qty) || 0);
          const sTrx = (Number(r.simplicity.pagi.trx) || 0) + (Number(r.simplicity.siang.trx) || 0);

          mTotalGraphisSales += gSales;
          mTotalGraphisQty += gQty;
          mTotalGraphisTrx += gTrx;

          mTotalSimplicitySales += sSales;
          mTotalSimplicityQty += sQty;
          mTotalSimplicityTrx += sTrx;
        });

        const mGrandTotalSales = mTotalGraphisSales + mTotalSimplicitySales;
        const mGrandTotalQty = mTotalGraphisQty + mTotalSimplicityQty;
        const mGrandTotalTrx = mTotalGraphisTrx + mTotalSimplicityTrx;

        const monthNameUpper = (MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase();

        // Selected Day Navigation helpers
        const currentSelectedIdx = displayMonthRecords.findIndex((r) => r.tanggal === selectedTanggal);
        const handlePrevDay = () => {
          if (currentSelectedIdx > 0) {
            setSelectedTanggal(displayMonthRecords[currentSelectedIdx - 1].tanggal);
          }
        };
        const handleNextDay = () => {
          if (currentSelectedIdx < displayMonthRecords.length - 1) {
            setSelectedTanggal(displayMonthRecords[currentSelectedIdx + 1].tanggal);
          }
        };

        return (
          <div className="space-y-4">
            {/* Top Toolbar: Month Selector & Sub-View Switcher */}
            <div className="bg-white border border-slate-200 rounded-xl p-3.5 flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-2xs">
              <div className="flex flex-wrap items-center gap-3">
                <div className="flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-amber-500" />
                  <span className="text-xs font-bold text-slate-700">Pilih Periode:</span>
                  <select
                    value={selectedMonthKey}
                    onChange={(e) => {
                      setSelectedMonthKey(e.target.value);
                      setSelectedTanggal(`01/${e.target.value}/2026`);
                    }}
                    className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-purple-600 cursor-pointer"
                  >
                    {['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'].map((mKey) => (
                      <option key={mKey} value={mKey}>
                        {MONTH_NAMES_ID[mKey]} 2026
                      </option>
                    ))}
                  </select>
                </div>

                <div className="text-xs font-black text-slate-900 tracking-wide uppercase px-3 py-1 bg-amber-100/90 border border-amber-300 rounded-lg">
                  PERIODE : {monthNameUpper} 2026
                </div>

                <div className="text-xs font-extrabold text-purple-900 px-2.5 py-1 bg-purple-100 border border-purple-200 rounded-lg">
                  DIVISI : 1 (YOUNG FASHION &amp; LADIES)
                </div>
              </div>

              {/* View Switcher: Tabel Rekap vs Slip Harian */}
              <div className="flex items-center gap-1.5 p-1 bg-slate-100 rounded-xl border border-slate-200 self-start md:self-auto">
                <button
                  type="button"
                  onClick={() => setDailySubView('matrix')}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                    dailySubView === 'matrix'
                      ? 'bg-white text-slate-950 shadow-xs'
                      : 'text-slate-600 hover:text-slate-950'
                  }`}
                  title="Lihat seluruh tanggal 1 s/d 31 dalam format matriks tabel spreadsheet"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5 text-amber-600" />
                  <span>Tabel Rekap 31 Hari</span>
                </button>

                <button
                  type="button"
                  onClick={() => setDailySubView('slip')}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                    dailySubView === 'slip'
                      ? 'bg-white text-slate-950 shadow-xs'
                      : 'text-slate-600 hover:text-slate-950'
                  }`}
                  title="Lihat laporan slip resmi harian untuk tanggal terpilih"
                >
                  <Receipt className="w-3.5 h-3.5 text-purple-600" />
                  <span>Slip Laporan Daily</span>
                </button>
              </div>
            </div>

            {/* Monthly KPI Overview Cards with Total Brand & Total Daily */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {/* Total Brand Graphis Card */}
              <div className="bg-gradient-to-br from-rose-50/90 via-white to-rose-50/30 border border-rose-200 rounded-xl p-3.5 shadow-2xs space-y-1">
                <div className="text-[11px] font-bold text-rose-800 uppercase tracking-wide flex items-center justify-between">
                  <span>Total Brand Graphis</span>
                  <span className="text-[10px] font-extrabold px-1.5 py-0.2 rounded bg-rose-200 text-rose-900">
                    Dept 6201
                  </span>
                </div>
                <div className="text-lg sm:text-xl font-black text-rose-950 font-mono-num">
                  Rp {formatNumberId(mTotalGraphisSales)}
                </div>
                <div className="flex items-center justify-between text-[11px] text-rose-700 font-semibold pt-0.5 border-t border-rose-100">
                  <span>{mTotalGraphisQty} pcs</span>
                  <span>{mTotalGraphisTrx} nota</span>
                  <span>
                    Share {mGrandTotalSales > 0 ? ((mTotalGraphisSales / mGrandTotalSales) * 100).toFixed(1) : '0.0'}%
                  </span>
                </div>
              </div>

              {/* Total Brand Simplicity Card */}
              <div className="bg-gradient-to-br from-indigo-50/90 via-white to-indigo-50/30 border border-indigo-200 rounded-xl p-3.5 shadow-2xs space-y-1">
                <div className="text-[11px] font-bold text-indigo-800 uppercase tracking-wide flex items-center justify-between">
                  <span>Total Brand Simplicity</span>
                  <span className="text-[10px] font-extrabold px-1.5 py-0.2 rounded bg-indigo-200 text-indigo-900">
                    Dept 6244
                  </span>
                </div>
                <div className="text-lg sm:text-xl font-black text-indigo-950 font-mono-num">
                  Rp {formatNumberId(mTotalSimplicitySales)}
                </div>
                <div className="flex items-center justify-between text-[11px] text-indigo-700 font-semibold pt-0.5 border-t border-indigo-100">
                  <span>{mTotalSimplicityQty} pcs</span>
                  <span>{mTotalSimplicityTrx} nota</span>
                  <span>
                    Share {mGrandTotalSales > 0 ? ((mTotalSimplicitySales / mGrandTotalSales) * 100).toFixed(1) : '0.0'}%
                  </span>
                </div>
              </div>

              {/* Total Daily (Gabungan Bulan Ini) Card */}
              <div className="bg-gradient-to-br from-amber-50/90 via-white to-amber-50/30 border border-amber-300 rounded-xl p-3.5 shadow-2xs space-y-1">
                <div className="text-[11px] font-bold text-amber-900 uppercase tracking-wide flex items-center justify-between">
                  <span>Total Daily Gabungan</span>
                  <span className="text-[10px] font-extrabold px-1.5 py-0.2 rounded bg-amber-200 text-amber-950">
                    Divisi 1
                  </span>
                </div>
                <div className="text-lg sm:text-xl font-black text-slate-950 font-mono-num">
                  Rp {formatNumberId(mGrandTotalSales)}
                </div>
                <div className="flex items-center justify-between text-[11px] text-amber-900 font-semibold pt-0.5 border-t border-amber-200">
                  <span>{mGrandTotalQty} pcs total</span>
                  <span>{mGrandTotalTrx} nota total</span>
                </div>
              </div>

              {/* Average Productivity Card */}
              <div className="bg-white border border-slate-200 rounded-xl p-3.5 shadow-2xs space-y-1">
                <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">
                  Rata-rata Harian
                </div>
                <div className="text-lg sm:text-xl font-black text-slate-900 font-mono-num">
                  Rp {formatNumberId(displayMonthRecords.length > 0 ? Math.round(mGrandTotalSales / displayMonthRecords.length) : 0)}
                </div>
                <div className="flex items-center justify-between text-[11px] text-slate-500 font-semibold pt-0.5 border-t border-slate-100">
                  <span>ASP: Rp {mGrandTotalQty > 0 ? formatNumberId(Math.round(mGrandTotalSales / mGrandTotalQty)) : '0'}</span>
                  <span>~{Math.round(mGrandTotalTrx / 31)} nota/hari</span>
                </div>
              </div>
            </div>

            {/* ============================================================ */}
            {/* SUB-VIEW A: TABEL REKAP 31 HARI FORMAT BARU & KEREN         */}
            {/* ============================================================ */}
            {dailySubView === 'matrix' && (
              <div className="bg-white rounded-2xl border border-slate-200/90 shadow-sm overflow-hidden transition-all">
                {/* Executive Header Banner Baru & Keren */}
                <div className="bg-gradient-to-r from-slate-900 via-slate-850 to-indigo-950 text-white p-4 sm:p-5 border-b border-slate-800 select-text">
                  <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                    <div className="space-y-1">
                      <div className="flex flex-wrap items-center gap-2 text-xs">
                        <span className="inline-flex items-center gap-1 text-[11px] font-black uppercase tracking-wider text-amber-400">
                          <span className="w-2 h-2 rounded-full bg-amber-400 animate-pulse" />
                          Bazar Matahari 311 Tanjung · PT MDS Retailing Tbk
                        </span>
                        <span className="text-slate-500">·</span>
                        <span className="text-purple-300 font-bold text-[11px]">
                          Kategori CV Konsinyasi
                        </span>
                      </div>

                      <div className="flex flex-wrap items-baseline gap-2.5">
                        <h2 className="text-base sm:text-xl font-black text-white tracking-tight">
                          LAPORAN SALES DAILY CV
                        </h2>
                        <span className="px-2.5 py-0.5 rounded-lg text-xs font-black bg-purple-500/30 text-purple-200 border border-purple-400/30">
                          DIVISI 1
                        </span>
                        <span className="text-xs text-slate-300 font-medium hidden sm:inline">
                          — Young Fashion, Ladies &amp; Intimate
                        </span>
                      </div>

                      <div className="text-xs text-slate-400 flex flex-wrap items-center gap-2 pt-0.5">
                        <span className="text-rose-300 font-semibold">Dept 6201 (Graphis)</span>
                        <span>·</span>
                        <span className="text-indigo-300 font-semibold">Dept 6244 (Simplicity)</span>
                        <span>·</span>
                        <span className="text-amber-300 font-semibold">Periode {monthNameUpper} 2026 (31 Hari)</span>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-2 shrink-0">
                      <button
                        type="button"
                        onClick={handleCopyWhatsApp}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs cursor-pointer transition-colors"
                        title="Salin ringkasan laporan WhatsApp"
                      >
                        <Share2 className="w-3.5 h-3.5" />
                        <span>Salin WhatsApp</span>
                      </button>

                      <button
                        type="button"
                        onClick={handleExportExcel}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white shadow-2xs cursor-pointer transition-colors"
                        title="Download file Excel 12 kolom lengkap"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>Ekspor Excel</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => {
                          const next = !isTableEditMode;
                          setIsTableEditMode(next);
                          if (next) {
                            showToast('Mode Edit Manual Aktif: Ketik langsung di sel tabel seperti Excel!');
                          } else {
                            showToast('Edit selesai! Data tersimpan & tabel terkunci rapi.');
                          }
                        }}
                        className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold shadow-2xs cursor-pointer transition-all ${
                          isTableEditMode
                            ? 'bg-amber-400 text-slate-950 ring-2 ring-amber-300 font-black'
                            : 'bg-slate-800 hover:bg-slate-700 text-white'
                        }`}
                        title="Edit manual sales, qty, nota langsung di baris tabel"
                      >
                        {isTableEditMode ? (
                          <>
                            <CheckCircle2 className="w-3.5 h-3.5" />
                            <span>Selesaikan Edit</span>
                          </>
                        ) : (
                          <>
                            <PenSquare className="w-3.5 h-3.5" />
                            <span>Edit Manual di Tabel</span>
                          </>
                        )}
                      </button>

                      {isAdmin && (
                        <button
                          type="button"
                          onClick={() => handleOpenEditModal(selectedTanggal)}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-purple-700 hover:bg-purple-600 text-white shadow-2xs cursor-pointer transition-colors"
                          title="Input sales report CV tanggal terpilih"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Input Sales Form</span>
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Filter Kalender Tanggal Bar */}
                  <div className="mt-3 pt-2.5 border-t border-slate-800 flex flex-wrap items-center justify-between gap-2.5">
                    <div className="flex flex-wrap items-center gap-2">
                      <div className="inline-flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 rounded-xl px-3 py-1.5 text-slate-200 shadow-2xs">
                        <Calendar className="w-3.5 h-3.5 text-amber-400 shrink-0" />
                        <label htmlFor="cv-table-cal-filter" className="text-xs font-bold text-slate-300 shrink-0">
                          Filter Kalender Tanggal:
                        </label>
                        <input
                          id="cv-table-cal-filter"
                          type="date"
                          value={tableDateFilterIso}
                          onChange={(e) => {
                            const val = e.target.value;
                            setTableDateFilterIso(val);
                            if (val) {
                              const parts = val.split('-');
                              if (parts.length === 3 && parts[1] !== selectedMonthKey) {
                                setSelectedMonthKey(parts[1]);
                              }
                              setSelectedTanggal(isoToSlashDate(val));
                            }
                          }}
                          className="px-2 py-0.5 text-xs font-bold bg-slate-900 border border-slate-700 text-white rounded focus:outline-none focus:border-amber-400 cursor-pointer"
                          title="Pilih tanggal dari kalender untuk melihat baris tanggal ini di tabel"
                        />
                        {tableDateFilterIso && (
                          <button
                            type="button"
                            onClick={() => setTableDateFilterIso('')}
                            className="text-[11px] font-bold text-amber-300 hover:text-white px-2 py-0.5 rounded bg-slate-700 cursor-pointer transition-colors"
                          >
                            Tampilkan Semua (31 Hari)
                          </button>
                        )}
                      </div>
                    </div>

                    {tableDateFilterIso && (
                      <div className="text-xs font-bold text-amber-300 flex items-center gap-1.5">
                        <span className="text-slate-400">Filter Aktif:</span>
                        <span className="px-2 py-0.5 rounded bg-amber-400/20 text-amber-200 border border-amber-400/30 font-mono-num">
                          {formatIndoDateLong(isoToSlashDate(tableDateFilterIso))}
                        </span>
                      </div>
                    )}
                  </div>

                  {/* Banner Mode Edit Manual di Tabel */}
                  {isTableEditMode && (
                    <div className="mt-2.5 px-3.5 py-2.5 bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-300 rounded-xl text-amber-950 text-xs font-bold flex flex-wrap items-center justify-between gap-2 shadow-xs">
                      <span className="flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-amber-500 animate-pulse shrink-0" />
                        <span>
                          <strong>Mode Edit Manual Format Tabel Aktif</strong> &mdash; Ketik langsung angka Sales, Qty, Nota, dan Reason di setiap sel tabel. Nilai Total Daily dihitung otomatis secara live.
                        </span>
                      </span>
                      <button
                        type="button"
                        onClick={() => {
                          setIsTableEditMode(false);
                          showToast('Edit selesai! Data berhasil disimpan & tabel terkunci rapi.');
                        }}
                        className="px-3.5 py-1 text-xs font-black bg-amber-400 hover:bg-amber-300 text-slate-950 rounded-lg shadow-xs cursor-pointer flex items-center gap-1.5 transition-transform active:scale-95"
                      >
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>Selesaikan Edit</span>
                      </button>
                    </div>
                  )}
                </div>

                {/* Tabel Vertikal Format Baru & Keren dengan Kolom Reason & Mode Edit Manual */}
                <div className="overflow-x-auto max-h-[760px] scrollbar-thin">
                  <table className="w-full text-xs border-collapse font-sans select-text">
                    <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                      {/* Baris 1: Grup Header Kolom */}
                      <tr className="border-b border-slate-800 text-slate-200">
                        <th
                          rowSpan={2}
                          className="px-3 py-2.5 text-center font-black w-14 text-[11px] tracking-wider text-slate-300 border-r border-slate-800"
                        >
                          TGL
                        </th>
                        <th
                          rowSpan={2}
                          className="px-3 py-2.5 text-center font-black w-24 text-[11px] tracking-wider text-slate-300 border-r border-slate-800"
                        >
                          HARI
                        </th>
                        <th
                          rowSpan={2}
                          className="px-2 py-2.5 text-center font-bold w-16 text-[10px] tracking-wider text-slate-400 border-r border-slate-800"
                        >
                          DIVISI
                        </th>
                        <th
                          colSpan={3}
                          className="px-3 py-2 text-center font-black text-xs tracking-wide bg-rose-950/40 text-rose-300 border-r border-slate-800"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-rose-500" />
                            <span>DEPT 6201 · BRAND GRAPHIS</span>
                          </div>
                        </th>
                        <th
                          colSpan={3}
                          className="px-3 py-2 text-center font-black text-xs tracking-wide bg-indigo-950/40 text-indigo-300 border-r border-slate-800"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-indigo-500" />
                            <span>DEPT 6244 · BRAND SIMPLICITY</span>
                          </div>
                        </th>
                        <th
                          colSpan={3}
                          className="px-3 py-2 text-center font-black text-xs tracking-wide bg-amber-950/40 text-amber-300 border-r border-slate-800"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-amber-400" />
                            <span>TOTAL DAILY (GABUNGAN)</span>
                          </div>
                        </th>
                        <th
                          rowSpan={2}
                          className="px-3 py-2.5 text-center font-black min-w-[130px] text-[11px] tracking-wider text-slate-300"
                        >
                          REASON / EVENT
                        </th>
                      </tr>

                      {/* Baris 2: Sub-kolom (SALES | QTY | NOTA) */}
                      <tr className="border-b border-slate-700 text-[11px] font-bold">
                        {/* Sub-kolom Graphis */}
                        <th className="px-2 py-2 text-right text-rose-200 min-w-[110px] bg-rose-950/25 border-r border-slate-800">
                          SALES (RP)
                        </th>
                        <th className="px-2 py-2 text-center text-rose-200 w-12 bg-rose-950/25 border-r border-slate-800">
                          QTY
                        </th>
                        <th className="px-2 py-2 text-center text-rose-200 w-12 bg-rose-950/25 border-r border-slate-800">
                          NOTA
                        </th>

                        {/* Sub-kolom Simplicity */}
                        <th className="px-2 py-2 text-right text-indigo-200 min-w-[110px] bg-indigo-950/25 border-r border-slate-800">
                          SALES (RP)
                        </th>
                        <th className="px-2 py-2 text-center text-indigo-200 w-12 bg-indigo-950/25 border-r border-slate-800">
                          QTY
                        </th>
                        <th className="px-2 py-2 text-center text-indigo-200 w-12 bg-indigo-950/25 border-r border-slate-800">
                          NOTA
                        </th>

                        {/* Sub-kolom Total Daily */}
                        <th className="px-2 py-2 text-right text-amber-200 min-w-[120px] bg-amber-950/25 border-r border-slate-800">
                          TOTAL SALES
                        </th>
                        <th className="px-2 py-2 text-center text-amber-200 w-12 bg-amber-950/25 border-r border-slate-800">
                          QTY
                        </th>
                        <th className="px-2 py-2 text-center text-amber-200 w-12 bg-amber-950/25 border-r border-slate-800">
                          NOTA
                        </th>
                      </tr>
                    </thead>

                    <tbody className="divide-y divide-slate-100 bg-white text-xs font-medium">
                      {(tableDateFilterIso
                        ? displayMonthRecords.filter((r) => r.tanggal === isoToSlashDate(tableDateFilterIso))
                        : displayMonthRecords
                      ).map((r) => {
                        const dNum = parseInt(r.tanggal.split('/')[0], 10);
                        const gSales = (Number(r.graphis.pagi.sales) || 0) + (Number(r.graphis.siang.sales) || 0);
                        const gQty = (Number(r.graphis.pagi.qty) || 0) + (Number(r.graphis.siang.qty) || 0);
                        const gTrx = (Number(r.graphis.pagi.trx) || 0) + (Number(r.graphis.siang.trx) || 0);

                        const sSales = (Number(r.simplicity.pagi.sales) || 0) + (Number(r.simplicity.siang.sales) || 0);
                        const sQty = (Number(r.simplicity.pagi.qty) || 0) + (Number(r.simplicity.siang.qty) || 0);
                        const sTrx = (Number(r.simplicity.pagi.trx) || 0) + (Number(r.simplicity.siang.trx) || 0);

                        const dTotalSales = gSales + sSales;
                        const dTotalQty = gQty + sQty;
                        const dTotalTrx = gTrx + sTrx;

                        const dayUpper = (r.dayName || '').toUpperCase();
                        const isWknd = dayUpper === 'SABTU' || dayUpper === 'MINGGU';
                        const isSelected = r.tanggal === selectedTanggal;

                        return (
                          <tr
                            key={r.tanggal}
                            onClick={() => setSelectedTanggal(r.tanggal)}
                            onDoubleClick={() => handleOpenEditModal(r.tanggal)}
                            className={`transition-colors cursor-pointer group ${
                              isSelected
                                ? 'bg-purple-50/80 ring-1 ring-inset ring-purple-300 font-bold'
                                : isWknd
                                ? 'bg-rose-50/20 hover:bg-rose-50/50'
                                : 'bg-white hover:bg-slate-50'
                            }`}
                            title={`Klik untuk memilih tanggal, klik 2x untuk Input Sales: ${r.dayName}, ${r.tanggal}`}
                          >
                            {/* Tanggal */}
                            <td className="px-3 py-2 text-center border-r border-slate-100">
                              <span
                                className={`inline-flex items-center justify-center w-7 h-7 rounded-lg text-xs font-black ${
                                  isSelected
                                    ? 'bg-purple-600 text-white shadow-2xs'
                                    : 'bg-slate-100 text-slate-800 group-hover:bg-slate-200'
                                }`}
                              >
                                {dNum}
                              </span>
                            </td>

                            {/* Hari */}
                            <td className="px-3 py-2 text-center border-r border-slate-100">
                              <span
                                className={`text-xs font-bold ${
                                  isWknd ? 'text-rose-600 font-extrabold' : 'text-slate-700'
                                }`}
                              >
                                {dayUpper}
                              </span>
                            </td>

                            {/* Divisi */}
                            <td className="px-2 py-2 text-center border-r border-slate-100 text-slate-400 font-medium text-[11px]">
                              Div. 1
                            </td>

                            {/* Graphis Sales */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-rose-50/30">
                                <input
                                  type="number"
                                  value={gSales > 0 ? gSales : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'gSales', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1.5 py-1 text-xs font-bold text-right bg-white border border-rose-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-rose-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-3 py-2 text-right font-mono-num font-bold text-slate-900 border-r border-slate-100 hover:bg-rose-50/70 hover:ring-1 hover:ring-rose-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Sales Graphis"
                              >
                                {gSales > 0 ? (
                                  <span className="text-slate-900">Rp {formatNumberId(gSales)}</span>
                                ) : (
                                  <span className="text-slate-300 font-normal group-hover/cell:text-rose-600">—</span>
                                )}
                              </td>
                            )}

                            {/* Graphis Qty */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-rose-50/30">
                                <input
                                  type="number"
                                  value={gQty > 0 ? gQty : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'gQty', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1 py-1 text-xs font-bold text-center bg-white border border-rose-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-rose-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-2 py-2 text-center font-mono-num text-slate-800 border-r border-slate-100 hover:bg-rose-50/70 hover:ring-1 hover:ring-rose-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Qty Graphis"
                              >
                                {gQty > 0 ? gQty : <span className="text-slate-300 font-normal group-hover/cell:text-rose-600">—</span>}
                              </td>
                            )}

                            {/* Graphis Nota */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-rose-50/30">
                                <input
                                  type="number"
                                  value={gTrx > 0 ? gTrx : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'gNota', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1 py-1 text-xs font-bold text-center bg-white border border-rose-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-rose-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-2 py-2 text-center font-mono-num text-slate-800 border-r border-slate-100 hover:bg-rose-50/70 hover:ring-1 hover:ring-rose-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Nota Graphis"
                              >
                                {gTrx > 0 ? gTrx : <span className="text-slate-300 font-normal group-hover/cell:text-rose-600">—</span>}
                              </td>
                            )}

                            {/* Simplicity Sales */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-indigo-50/30">
                                <input
                                  type="number"
                                  value={sSales > 0 ? sSales : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'sSales', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1.5 py-1 text-xs font-bold text-right bg-white border border-indigo-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-3 py-2 text-right font-mono-num font-bold text-slate-900 border-r border-slate-100 hover:bg-indigo-50/70 hover:ring-1 hover:ring-indigo-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Sales Simplicity"
                              >
                                {sSales > 0 ? (
                                  <span className="text-slate-900">Rp {formatNumberId(sSales)}</span>
                                ) : (
                                  <span className="text-slate-300 font-normal group-hover/cell:text-indigo-600">—</span>
                                )}
                              </td>
                            )}

                            {/* Simplicity Qty */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-indigo-50/30">
                                <input
                                  type="number"
                                  value={sQty > 0 ? sQty : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'sQty', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1 py-1 text-xs font-bold text-center bg-white border border-indigo-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-2 py-2 text-center font-mono-num text-slate-800 border-r border-slate-100 hover:bg-indigo-50/70 hover:ring-1 hover:ring-indigo-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Qty Simplicity"
                              >
                                {sQty > 0 ? sQty : <span className="text-slate-300 font-normal group-hover/cell:text-indigo-600">—</span>}
                              </td>
                            )}

                            {/* Simplicity Nota */}
                            {isTableEditMode ? (
                              <td className="p-1 border-r border-slate-200 bg-indigo-50/30">
                                <input
                                  type="number"
                                  value={sTrx > 0 ? sTrx : ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'sNota', e.target.value)}
                                  placeholder="0"
                                  className="w-full px-1 py-1 text-xs font-bold text-center bg-white border border-indigo-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono-num"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-2 py-2 text-center font-mono-num text-slate-800 border-r border-slate-100 hover:bg-indigo-50/70 hover:ring-1 hover:ring-indigo-300 transition-colors cursor-pointer group/cell"
                                title="Klik untuk ketik / edit manual Nota Simplicity"
                              >
                                {sTrx > 0 ? sTrx : <span className="text-slate-300 font-normal group-hover/cell:text-indigo-600">—</span>}
                              </td>
                            )}

                            {/* Total Daily Sales (Dihitung Otomatis Live) */}
                            <td className="px-3 py-2 text-right font-mono-num font-black text-slate-950 bg-amber-50/50 border-r border-slate-100">
                              {dTotalSales > 0 ? (
                                <span className="text-slate-950 font-black">Rp {formatNumberId(dTotalSales)}</span>
                              ) : (
                                <span className="text-slate-300 font-normal">—</span>
                              )}
                            </td>

                            {/* Total Daily Qty (Dihitung Otomatis Live) */}
                            <td className="px-2 py-2 text-center font-mono-num font-black text-slate-950 bg-amber-50/50 border-r border-slate-100">
                              {dTotalQty > 0 ? dTotalQty : <span className="text-slate-300 font-normal">—</span>}
                            </td>

                            {/* Total Daily Nota (Dihitung Otomatis Live) */}
                            <td className="px-2 py-2 text-center font-mono-num font-black text-slate-950 bg-amber-50/50 border-r border-slate-100">
                              {dTotalTrx > 0 ? dTotalTrx : <span className="text-slate-300 font-normal">—</span>}
                            </td>

                            {/* Reason / Catatan Event */}
                            {isTableEditMode ? (
                              <td className="p-1 bg-amber-50/20">
                                <input
                                  type="text"
                                  value={r.reasonEvent || ''}
                                  onChange={(e) => handleDirectCellChange(r.tanggal, 'reason', e.target.value)}
                                  placeholder="Catatan event..."
                                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded text-slate-800 focus:outline-none focus:ring-1 focus:ring-amber-500"
                                />
                              </td>
                            ) : (
                              <td
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setIsTableEditMode(true);
                                }}
                                className="px-2.5 py-2 text-slate-600 text-[11px] truncate max-w-[150px] hover:bg-slate-100/80 cursor-pointer"
                                title={r.reasonEvent ? `${r.reasonEvent} (Klik untuk edit)` : 'Klik untuk ketik catatan / reason'}
                              >
                                {r.reasonEvent ? (
                                  <span className="text-slate-800 font-medium">{r.reasonEvent}</span>
                                ) : (
                                  <span className="text-slate-300 italic">—</span>
                                )}
                              </td>
                            )}
                          </tr>
                        );
                      })}

                      {/* 1 BARIS TOTAL SAJA (TOTAL CV DIVISI 1) */}
                      <tr className="bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs sticky bottom-0 z-30 shadow-lg border-t-2 border-amber-400">
                        <td
                          colSpan={3}
                          className="px-4 py-3 text-center font-black tracking-wider text-xs uppercase text-amber-300 border-r border-slate-700"
                        >
                          TOTAL CV DIVISI 1
                        </td>

                        {/* Graphis Total Sum */}
                        <td className="px-3 py-3 text-right font-mono-num font-black text-rose-300 border-r border-slate-700">
                          {mTotalGraphisSales > 0 ? (
                            <span>Rp {formatNumberId(mTotalGraphisSales)}</span>
                          ) : (
                            <span>—</span>
                          )}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-rose-300 border-r border-slate-700">
                          {mTotalGraphisQty > 0 ? mTotalGraphisQty : '—'}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-rose-300 border-r border-slate-700">
                          {mTotalGraphisTrx > 0 ? mTotalGraphisTrx : '—'}
                        </td>

                        {/* Simplicity Total Sum */}
                        <td className="px-3 py-3 text-right font-mono-num font-black text-indigo-300 border-r border-slate-700">
                          {mTotalSimplicitySales > 0 ? (
                            <span>Rp {formatNumberId(mTotalSimplicitySales)}</span>
                          ) : (
                            <span>—</span>
                          )}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-indigo-300 border-r border-slate-700">
                          {mTotalSimplicityQty > 0 ? mTotalSimplicityQty : '—'}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-indigo-300 border-r border-slate-700">
                          {mTotalSimplicityTrx > 0 ? mTotalSimplicityTrx : '—'}
                        </td>

                        {/* Grand Total Daily Sum */}
                        <td className="px-3 py-3 text-right font-mono-num font-black text-sm text-amber-300 bg-amber-500/20 border-r border-slate-700">
                          {mGrandTotalSales > 0 ? (
                            <span>Rp {formatNumberId(mGrandTotalSales)}</span>
                          ) : (
                            <span>—</span>
                          )}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-sm text-amber-300 bg-amber-500/20 border-r border-slate-700">
                          {mGrandTotalQty > 0 ? mGrandTotalQty : '—'}
                        </td>
                        <td className="px-2 py-3 text-center font-mono-num font-black text-sm text-amber-300 bg-amber-500/20 border-r border-slate-700">
                          {mGrandTotalTrx > 0 ? mGrandTotalTrx : '—'}
                        </td>
                        <td className="px-2 py-3 text-center text-[10px] text-amber-200/80 italic">
                          Total Bulanan
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Footer Bar & Quick Action */}
                <div className="p-3.5 bg-slate-50 border-t border-slate-200 text-xs text-slate-700 flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900">
                      Tanggal Terpilih: <span className="text-purple-700 font-extrabold">{selectedTanggal}</span> ({getIndoDayName(selectedTanggal)})
                    </span>
                    <button
                      type="button"
                      onClick={() => setDailySubView('slip')}
                      className="text-purple-700 hover:text-purple-900 font-black underline cursor-pointer ml-1 inline-flex items-center gap-1"
                    >
                      <span>Buka Slip Detail</span>
                      <span>&rarr;</span>
                    </button>
                  </div>
                  <div className="flex items-center gap-2">
                    {isAdmin && (
                      <button
                        type="button"
                        onClick={() => handleOpenEditModal(selectedTanggal)}
                        className="px-3 py-1.5 rounded-lg bg-amber-400 hover:bg-amber-300 text-slate-950 font-bold inline-flex items-center gap-1.5 cursor-pointer text-xs shadow-2xs"
                      >
                        <PenSquare className="w-3.5 h-3.5" />
                        <span>Edit Data {selectedTanggal}</span>
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={handleExportExcel}
                      className="px-3 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white font-bold inline-flex items-center gap-1.5 cursor-pointer text-xs shadow-2xs"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Ekspor Excel (.xlsx)</span>
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* ============================================================ */}
            {/* SUB-VIEW B: SLIP LAPORAN SALES DAILY RESMI (PER TANGGAL)     */}
            {/* ============================================================ */}
            {dailySubView === 'slip' && (
              <div className="space-y-4">
                {/* Stepper Navigasi Hari */}
                <div className="bg-white border border-slate-200 rounded-xl p-3 flex flex-wrap items-center justify-between gap-3 shadow-2xs">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handlePrevDay}
                      disabled={currentSelectedIdx <= 0}
                      className="p-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer text-xs font-bold text-slate-800 inline-flex items-center gap-1"
                      title="Hari Sebelumnya"
                    >
                      <ChevronLeft className="w-4 h-4" />
                      <span className="hidden sm:inline">Hari Sebelumnya</span>
                    </button>

                    <select
                      value={selectedTanggal}
                      onChange={(e) => setSelectedTanggal(e.target.value)}
                      className="px-3 py-1.5 text-xs font-black bg-purple-50 border border-purple-300 rounded-lg text-purple-950 focus:outline-none focus:border-purple-600 cursor-pointer"
                    >
                      {displayMonthRecords.map((r) => (
                        <option key={r.tanggal} value={r.tanggal}>
                          {r.tanggal} — {r.dayName.toUpperCase()}
                        </option>
                      ))}
                    </select>

                    <div className="flex items-center gap-1.5 bg-purple-50 border border-purple-300 rounded-lg px-2 py-1">
                      <Calendar className="w-3.5 h-3.5 text-purple-700 shrink-0" />
                      <input
                        aria-label="Pilih Tanggal Slip Kalender"
                        type="date"
                        value={slashDateToIso(selectedTanggal)}
                        onChange={(e) => {
                          if (e.target.value) {
                            setSelectedTanggal(isoToSlashDate(e.target.value));
                          }
                        }}
                        className="text-xs font-bold bg-transparent text-purple-950 focus:outline-none cursor-pointer"
                        title="Pilih tanggal kalender untuk slip harian"
                      />
                    </div>

                    <button
                      type="button"
                      onClick={handleNextDay}
                      disabled={currentSelectedIdx >= displayMonthRecords.length - 1}
                      className="p-1.5 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer text-xs font-bold text-slate-800 inline-flex items-center gap-1"
                      title="Hari Berikutnya"
                    >
                      <span className="hidden sm:inline">Hari Berikutnya</span>
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleCopyWhatsApp}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs cursor-pointer transition-colors"
                      title="Salin teks laporan slip harian ini ke WhatsApp"
                    >
                      <Share2 className="w-3.5 h-3.5" />
                      <span>Salin WA</span>
                    </button>

                    {isAdmin && (
                      <button
                        type="button"
                        onClick={() => handleOpenEditModal(selectedTanggal)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 shadow-2xs cursor-pointer transition-colors"
                        title="Input nilai sales, qty, atau nota hari ini"
                      >
                        <PenSquare className="w-3.5 h-3.5" />
                        <span>Input Sales CV</span>
                      </button>
                    )}
                  </div>
                </div>

                {/* Formal Retail Sales Audit Slip Card */}
                <div className="bg-white border-2 border-slate-800 rounded-2xl overflow-hidden shadow-md max-w-4xl mx-auto">
                  {/* Slip Header */}
                  <div className="bg-gradient-to-r from-slate-900 via-purple-950 to-slate-900 text-white p-5 select-text border-b border-purple-800/40">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <div className="text-[11px] font-bold text-amber-300 tracking-wider uppercase">
                          PT MDS Retailing Tbk · Matahari Department Store 311
                        </div>
                        <h2 className="text-lg sm:text-xl font-black text-white tracking-tight flex items-center gap-2">
                          <span>LAPORAN SALES DAILY CV</span>
                          <span className="text-purple-300 text-sm font-bold">— DIVISI 1</span>
                        </h2>
                        <div className="text-xs text-slate-300 font-medium">
                          Bazar Tanjung · Young Fashion, Ladies &amp; Intimate
                        </div>
                      </div>

                      <div className="text-left sm:text-right text-xs bg-slate-950/60 p-2.5 rounded-xl border border-slate-800 font-mono-num">
                        <div className="text-slate-400 text-[10px] uppercase font-bold tracking-wider">
                          Nomor Referensi Harian
                        </div>
                        <div className="text-amber-400 font-black text-sm">
                          CV1-D{selectedTanggal.replace(/\//g, '')}
                        </div>
                      </div>
                    </div>

                    {/* Metadata Strip */}
                    <div className="mt-4 pt-3 border-t border-purple-800/40 grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                      <div>
                        <span className="text-slate-400 text-[10px] uppercase block font-bold">Periode</span>
                        <span className="font-extrabold text-amber-300">{monthNameUpper} 2026</span>
                      </div>
                      <div>
                        <span className="text-slate-400 text-[10px] uppercase block font-bold">Tanggal</span>
                        <span className="font-extrabold text-white">{currentDayRecord.tanggal}</span>
                      </div>
                      <div>
                        <span className="text-slate-400 text-[10px] uppercase block font-bold">Hari</span>
                        <span className="font-extrabold text-white">{currentDayRecord.dayName.toUpperCase()}</span>
                      </div>
                      <div>
                        <span className="text-slate-400 text-[10px] uppercase block font-bold">Divisi</span>
                        <span className="font-extrabold text-purple-200">DIVISI 1</span>
                      </div>
                    </div>
                  </div>

                  {/* Detail Table Per Brand & Shift */}
                  <div className="p-4 overflow-x-auto">
                    <table className="w-full text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100 text-slate-800 border-b-2 border-slate-300">
                          <th className="py-2.5 px-3 text-left font-black w-16">DEPT</th>
                          <th className="py-2.5 px-3 text-left font-black w-32">BRAND</th>
                          <th className="py-2.5 px-3 text-center font-black">SHIFT PAGI (SALES / QTY / NOTA)</th>
                          <th className="py-2.5 px-3 text-center font-black">SHIFT SIANG (SALES / QTY / NOTA)</th>
                          <th className="py-2.5 px-3 text-right font-black">TOTAL SALES</th>
                          <th className="py-2.5 px-2 text-center font-black w-14">QTY</th>
                          <th className="py-2.5 px-2 text-center font-black w-14">NOTA</th>
                          <th className="py-2.5 px-3 text-right font-black w-24">ASP</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {/* 1. Row Dept 6201 Graphis */}
                        <tr className="hover:bg-rose-50/40">
                          <td className="py-3 px-3 font-mono-num font-bold text-slate-700">6201</td>
                          <td className="py-3 px-3">
                            <span className="font-black text-rose-900 block">GRAPHIS</span>
                            <span className="text-[10px] text-slate-500">Young Fashion</span>
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-[11px]">
                            <div className="font-bold text-slate-900">
                              Rp {formatNumberId(currentDayRecord.graphis.pagi.sales || 0)}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              {currentDayRecord.graphis.pagi.qty || 0} pcs · {currentDayRecord.graphis.pagi.trx || 0} nota
                            </div>
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-[11px]">
                            <div className="font-bold text-slate-900">
                              Rp {formatNumberId(currentDayRecord.graphis.siang.sales || 0)}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              {currentDayRecord.graphis.siang.qty || 0} pcs · {currentDayRecord.graphis.siang.trx || 0} nota
                            </div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num font-black text-rose-950 text-sm">
                            Rp {formatNumberId(currentDaySummary.graphis.sales)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-slate-900">
                            {currentDaySummary.graphis.qty}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-slate-900">
                            {currentDaySummary.graphis.trx}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num text-slate-700 font-semibold">
                            Rp {formatNumberId(currentDaySummary.graphis.pricePoint)}
                          </td>
                        </tr>

                        {/* 2. Row Dept 6244 Simplicity */}
                        <tr className="hover:bg-indigo-50/40">
                          <td className="py-3 px-3 font-mono-num font-bold text-slate-700">6244</td>
                          <td className="py-3 px-3">
                            <span className="font-black text-indigo-900 block">SIMPLICITY</span>
                            <span className="text-[10px] text-slate-500">Ladies &amp; Intimate</span>
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-[11px]">
                            <div className="font-bold text-slate-900">
                              Rp {formatNumberId(currentDayRecord.simplicity.pagi.sales || 0)}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              {currentDayRecord.simplicity.pagi.qty || 0} pcs · {currentDayRecord.simplicity.pagi.trx || 0} nota
                            </div>
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-[11px]">
                            <div className="font-bold text-slate-900">
                              Rp {formatNumberId(currentDayRecord.simplicity.siang.sales || 0)}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              {currentDayRecord.simplicity.siang.qty || 0} pcs · {currentDayRecord.simplicity.siang.trx || 0} nota
                            </div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num font-black text-indigo-950 text-sm">
                            Rp {formatNumberId(currentDaySummary.simplicity.sales)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-slate-900">
                            {currentDaySummary.simplicity.qty}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-slate-900">
                            {currentDaySummary.simplicity.trx}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num text-slate-700 font-semibold">
                            Rp {formatNumberId(currentDaySummary.simplicity.pricePoint)}
                          </td>
                        </tr>

                        {/* 3. Baris TOTAL DAILY GABUNGAN */}
                        <tr className="bg-amber-100/70 border-t-2 border-slate-800 font-black text-slate-950">
                          <td colSpan={2} className="py-3 px-3 uppercase tracking-wider text-xs font-black">
                            TOTAL DAILY (DIVISI 1)
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-xs">
                            Rp {formatNumberId(
                              (Number(currentDayRecord.graphis.pagi.sales) || 0) +
                              (Number(currentDayRecord.simplicity.pagi.sales) || 0)
                            )}
                          </td>
                          <td className="py-3 px-3 text-center font-mono-num text-xs">
                            Rp {formatNumberId(
                              (Number(currentDayRecord.graphis.siang.sales) || 0) +
                              (Number(currentDayRecord.simplicity.siang.sales) || 0)
                            )}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num text-base font-black text-slate-950">
                            Rp {formatNumberId(currentDaySummary.total.sales)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-sm font-black">
                            {currentDaySummary.total.qty}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-sm font-black">
                            {currentDaySummary.total.trx}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num text-xs font-black">
                            Rp {formatNumberId(currentDaySummary.total.pricePoint)}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Operational Summary Grid: Target, Petugas, Catatan */}
                  <div className="bg-slate-50 border-t border-slate-200 p-4 grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
                    <div className="p-3 bg-white rounded-xl border border-slate-200 space-y-1">
                      <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                        Target Daily &amp; Pencapaian
                      </div>
                      <div className="flex items-baseline justify-between">
                        <span className="font-mono-num font-black text-slate-900 text-sm">
                          Rp {formatNumberId(currentDaySummary.targetDaily)}
                        </span>
                        <span
                          className={`font-black px-2 py-0.5 rounded text-xs ${
                            currentDaySummary.achievementPct >= 100
                              ? 'bg-emerald-100 text-emerald-800'
                              : 'bg-amber-100 text-amber-800'
                          }`}
                        >
                          {currentDaySummary.achievementPct.toFixed(1)}%
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-1.5 overflow-hidden">
                        <div
                          className="bg-emerald-500 h-1.5 rounded-full"
                          style={{ width: `${Math.min(currentDaySummary.achievementPct, 100)}%` }}
                        />
                      </div>
                    </div>

                    <div className="p-3 bg-white rounded-xl border border-slate-200 space-y-1">
                      <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                        Petugas Jaga Harian
                      </div>
                      <div className="space-y-0.5 text-slate-800 font-semibold">
                        <div>Shift Pagi: <strong className="text-slate-950">{currentDayRecord.petugasPagi || 'Belum diisi'}</strong></div>
                        <div>Shift Siang: <strong className="text-slate-950">{currentDayRecord.petugasSiang || 'Belum diisi'}</strong></div>
                      </div>
                    </div>

                    <div className="p-3 bg-white rounded-xl border border-slate-200 space-y-1">
                      <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wide">
                        Catatan &amp; Event
                      </div>
                      <p className="text-slate-700 italic">
                        {currentDayRecord.reasonEvent || 'Tidak ada catatan event khusus hari ini.'}
                      </p>
                    </div>
                  </div>

                  {/* Slip Bottom Action Strip */}
                  <div className="p-3 bg-slate-100 border-t border-slate-300 flex flex-wrap items-center justify-between gap-3">
                    <button
                      type="button"
                      onClick={() => setDailySubView('matrix')}
                      className="text-xs font-bold text-slate-700 hover:text-slate-950 inline-flex items-center gap-1.5 cursor-pointer"
                    >
                      &larr; <span>Kembali ke Matriks Rekap 31 Hari</span>
                    </button>

                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={handleCopyWhatsApp}
                        className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs inline-flex items-center gap-1.5 cursor-pointer shadow-2xs"
                      >
                        <Share2 className="w-3.5 h-3.5" />
                        <span>Salin Teks WhatsApp</span>
                      </button>

                      {isAdmin && (
                        <button
                          type="button"
                          onClick={() => handleOpenEditModal(selectedTanggal)}
                          className="px-3 py-1.5 rounded-lg bg-amber-400 hover:bg-amber-300 text-slate-950 font-bold text-xs inline-flex items-center gap-1.5 cursor-pointer shadow-2xs"
                        >
                          <PenSquare className="w-3.5 h-3.5" />
                          <span>Edit Hari Ini</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* ============================================================ */}
            {/* CHART DIAGRAM SALES CV HARIAN (SEPERTI REPORT DP)            */}
            {/* ============================================================ */}
            <div className="space-y-4">
              {(() => {
                const dailyChartRows = displayMonthRecords.map((r) => {
                  const gSales =
                    (Number(r.graphis.pagi.sales) || 0) + (Number(r.graphis.siang.sales) || 0);
                  const sSales =
                    (Number(r.simplicity.pagi.sales) || 0) + (Number(r.simplicity.siang.sales) || 0);
                  const dName = getIndoDayName(r.tanggal);
                  const isWeekend = dName === 'Sabtu' || dName === 'Minggu';
                  const dayNum = parseInt(r.tanggal.split('/')[0], 10) || 0;
                  return {
                    key: r.tanggal,
                    label: `${dName.slice(0, 3)} (${dayNum})`,
                    subLabel: r.tanggal.slice(0, 5),
                    pagiSales: gSales,
                    siangSales: sSales,
                    totalSales: gSales + sSales,
                    isWeekend,
                    isSelected:
                      r.tanggal === selectedTanggal ||
                      (tableDateFilterIso ? r.tanggal === isoToSlashDate(tableDateFilterIso) : false),
                    onClick: () => setSelectedTanggal(r.tanggal),
                  };
                });

                return (
                  <>
                    <PhotoStyleSalesDiagram
                      headerTitle={`DIAGRAM SALES CV HARIAN — ${monthNameUpper} 2026 (TGL 1 – ${displayMonthRecords.length})`}
                      headerSubtitle="Diagram Batang Vertikal & Garis Tren Penjualan Dept 6201 (Graphis) vs Dept 6244 (Simplicity)"
                      badgeText="DIAGRAM SALES CV (OMZET RUPIAH)"
                      cardTitle="Grafik Batang Graphis vs Simplicity & Garis Tren"
                      primaryLabel="Graphis (6201)"
                      secondaryLabel="Simplicity (6244)"
                      totalSales={mGrandTotalSales}
                      totalPagi={mTotalGraphisSales}
                      totalSiang={mTotalSimplicitySales}
                      compactX={true}
                      bars={dailyChartRows}
                    />

                    <GrowthAndContributionBarDiagram
                      headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI HARIAN CV — ${monthNameUpper} 2026`}
                      headerSubtitle={`Grafik Batang Pertumbuhan Harian (Growth %) & Porsi Kontribusi Sales CV ${dailyChartRows.length} Hari`}
                      periodBadgeText={`${monthNameUpper} 2026`}
                      compactX={true}
                      items={dailyChartRows.map((item, idx) => ({
                        ...item,
                        prevSales: idx > 0 ? dailyChartRows[idx - 1].totalSales : null,
                      }))}
                    />
                  </>
                );
              })()}
            </div>
          </div>
        );
      })()}

      {/* ============================================================ */}
      {/* HORIZON 2: SALES WEEKLY                                      */}
      {/* ============================================================ */}
      {/* ============================================================ */}
      {/* HORIZON 2: SALES WEEKLY — FORMAT TERBARU (TABLE KEREN)       */}
      {/* ============================================================ */}
      {horizon === 'weekly' && (
        <div className="space-y-4">
          {/* 1. Header Filter & Control Bar */}
          <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs space-y-3">
            <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-rose-100 text-rose-800 border border-rose-200">
                    Format Terbaru CV Weekly
                  </span>
                  <span className="text-xs font-bold text-slate-500">
                    Divisi 1 · Tanjung Outmall 311
                  </span>
                </div>
                <h2 className="text-base sm:text-lg font-black text-slate-900 tracking-tight mt-1 flex items-center gap-2">
                  <FileSpreadsheet className="w-5 h-5 text-rose-600" />
                  <span>REKAPITULASI SALES WEEKLY (SENIN S/D MINGGU)</span>
                </h2>
              </div>

              {/* Action Buttons */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleCopyWhatsApp}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs cursor-pointer transition-colors"
                  title="Salin rekap weekly format terbaru ke WhatsApp"
                >
                  <Share2 className="w-3.5 h-3.5" />
                  <span>Salin WA Weekly</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white shadow-2xs cursor-pointer transition-colors"
                  title="Ekspor laporan Excel (.xlsx)"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Ekspor Excel</span>
                </button>

                {isAdmin && (
                  <button
                    type="button"
                    onClick={() => handleOpenEditModal(selectedTanggal)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 shadow-2xs cursor-pointer transition-colors"
                    title="Input penjualan harian CV"
                  >
                    <PenSquare className="w-3.5 h-3.5" />
                    <span>Input Sales CV</span>
                  </button>
                )}
              </div>
            </div>

            {/* Filter Ribbons: Bulan, Cakupan, & Brand Filter */}
            <div className="pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex flex-wrap items-center gap-3">
                {/* Pilih Bulan */}
                <div className="flex items-center gap-1.5">
                  <span className="font-bold text-slate-600">Bulan:</span>
                  <select
                    value={weeklyFilterMonth}
                    onChange={(e) => {
                      setWeeklyFilterMonth(e.target.value);
                      setWeeklyCalendarIso('');
                    }}
                    className="px-2.5 py-1 text-xs font-bold bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-rose-500 cursor-pointer"
                  >
                    {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                      <option key={mKey} value={mKey}>
                        {mName} 2026
                      </option>
                    ))}
                  </select>
                </div>

                {/* Filter Kalender Tanggal Weekly */}
                <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-purple-50 border border-purple-300 text-xs shadow-2xs">
                  <Calendar className="w-3.5 h-3.5 text-purple-700 shrink-0" />
                  <label htmlFor="cv-weekly-cal-filter" className="font-bold text-purple-950 shrink-0">
                    Kalender:
                  </label>
                  <input
                    id="cv-weekly-cal-filter"
                    type="date"
                    value={weeklyCalendarIso}
                    onChange={(e) => {
                      const val = e.target.value;
                      setWeeklyCalendarIso(val);
                      if (val) {
                        const slash = isoToSlashDate(val);
                        const matched = ANNUAL_2026_WEEKS.find(
                          (w) =>
                            w.slashDates.includes(slash) ||
                            (val >= w.startDateIso && val <= w.endDateIso)
                        );
                        const parts = val.split('-');
                        if (parts.length === 3 && MONTH_NAMES_ID[parts[1]]) {
                          setWeeklyFilterMonth(parts[1]);
                          setWeeklyScope('month');
                        }
                        if (matched) {
                          setSelectedWeeklyRowNumber(matched.weekNumber);
                        }
                      }
                    }}
                    className="px-1.5 py-0.5 text-xs font-bold bg-white border border-purple-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-purple-500 cursor-pointer"
                    title="Pilih tanggal dari kalender untuk langsung menyorot Weekly terkait"
                  />
                  {weeklyCalendarIso && (
                    <button
                      type="button"
                      onClick={() => setWeeklyCalendarIso('')}
                      className="text-[10px] font-bold text-purple-800 hover:text-purple-950 px-1.5 py-0.5 rounded bg-purple-200/80 cursor-pointer"
                    >
                      Semua
                    </button>
                  )}
                </div>

                {/* Cakupan Scope */}
                <div className="inline-flex p-0.5 bg-slate-100 rounded-lg border border-slate-200">
                  <button
                    type="button"
                    onClick={() => setWeeklyScope('month')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      weeklyScope === 'month'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Bulan Ini ({MONTH_NAMES_ID[weeklyFilterMonth] || 'November'})
                  </button>
                  <button
                    type="button"
                    onClick={() => setWeeklyScope('all')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      weeklyScope === 'all'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Semua Weekly (Setahun 2026)
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* 2. Top Executive KPI Cards */}
          {(() => {
            const activeSummary =
              weeklyBrandFilter === 'GRAPHIS'
                ? weeklyTableData.summary.graphis
                : weeklyBrandFilter === 'SIMPLICITY'
                ? weeklyTableData.summary.simplicity
                : weeklyTableData.summary.total;

            return (
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
                <div className="bg-gradient-to-br from-slate-900 via-slate-800 to-indigo-950 text-white rounded-2xl p-4 shadow-sm border border-slate-700/60 space-y-1.5">
                  <div className="text-[10px] font-black uppercase tracking-wider text-amber-400 flex items-center justify-between">
                    <span>
                      {weeklyBrandFilter === 'GRAPHIS'
                        ? 'OMSET WEEKLY GRAPHIS'
                        : weeklyBrandFilter === 'SIMPLICITY'
                        ? 'OMSET WEEKLY SIMPLICITY'
                        : 'TOTAL SALES WEEKLY (CV)'}
                    </span>
                    <span className="px-1.5 py-0.2 rounded bg-amber-400/20 text-amber-300 text-[9px] font-bold">
                      {weeklyTableData.rows.length} Weeks
                    </span>
                  </div>
                  <div className="text-xl sm:text-2xl font-black font-mono-num text-white">
                    Rp {formatNumberId(activeSummary.grandTotalSales)}
                  </div>
                  <div className="text-[11px] text-slate-300 font-semibold flex items-center gap-1.5">
                    <span>Rata-rata:</span>
                    <span className="text-amber-300 font-mono-num font-bold">
                      Rp {formatNumberId(weeklyTableData.rows.length > 0 ? Math.round(activeSummary.grandTotalSales / weeklyTableData.rows.length) : 0)} / week
                    </span>
                  </div>
                </div>

                <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-2xs space-y-1.5">
                  <div className="text-[10px] font-black uppercase tracking-wider text-slate-500">
                    TOTAL VOLUME &amp; TRANSAKSI
                  </div>
                  <div className="text-xl sm:text-2xl font-black font-mono-num text-slate-900 flex items-baseline gap-2">
                    <span>{formatNumberId(activeSummary.grandTotalQty)}</span>
                    <span className="text-xs font-bold text-slate-500">pcs</span>
                  </div>
                  <div className="text-[11px] text-slate-600 font-medium flex items-center justify-between">
                    <span>{activeSummary.grandTotalTrx} Transaksi</span>
                    <span className="font-mono-num font-bold text-purple-700">
                      ASP: Rp {formatNumberId(activeSummary.avgPricePoint)}
                    </span>
                  </div>
                </div>

                <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-2xs space-y-1.5">
                  <div className="text-[10px] font-black uppercase tracking-wider text-rose-700 flex items-center justify-between">
                    <span>KATEGORI BIG SIZE</span>
                    <span className="px-1.5 py-0.2 rounded bg-rose-100 text-rose-800 text-[9px] font-bold">
                      Ladies Fashion
                    </span>
                  </div>
                  <div className="text-xl sm:text-2xl font-black font-mono-num text-rose-950">
                    Rp {formatNumberId(activeSummary.grandTotalBigSizeSales)}
                  </div>
                  <div className="text-[11px] text-slate-600 font-medium flex items-center justify-between">
                    <span>{activeSummary.grandTotalBigSizeQty} pcs (XL - 5L)</span>
                    <span className="font-bold text-slate-500">Share ~20%</span>
                  </div>
                </div>

                <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-2xs space-y-1.5">
                  <div className="text-[10px] font-black uppercase tracking-wider text-emerald-700 flex items-center justify-between">
                    <span>PERTUMBUHAN (GROWTH WoW)</span>
                    <span className="px-1.5 py-0.2 rounded bg-emerald-100 text-emerald-800 text-[9px] font-bold">
                      Trend Weekly
                    </span>
                  </div>
                  <div className="text-xl sm:text-2xl font-black font-mono-num flex items-center gap-1.5 text-slate-900">
                    <span
                      className={
                        activeSummary.avgGrowthPct >= 0
                          ? 'text-emerald-600'
                          : 'text-rose-600'
                      }
                    >
                      {activeSummary.avgGrowthPct >= 0 ? '+' : ''}
                      {activeSummary.avgGrowthPct.toFixed(1)}%
                    </span>
                    {activeSummary.avgGrowthPct >= 0 ? (
                      <TrendingUp className="w-5 h-5 text-emerald-600" />
                    ) : (
                      <ArrowDownRight className="w-5 h-5 text-rose-600" />
                    )}
                  </div>
                  <div className="text-[11px] text-slate-500 font-medium">
                    Rata-rata pertumbuhan mingguan
                  </div>
                </div>
              </div>
            );
          })()}

          {/* 3. TABEL UTAMA FORMAT TERBARU SALES WEEKLY */}
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
            <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#161b2e] to-indigo-950 text-white flex flex-col md:flex-row md:items-center justify-between gap-2.5 border-b border-slate-800">
              <div className="flex items-center gap-2">
                <div className="w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse shrink-0" />
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                  TABEL SALES WEEKLY (SENIN S/D MINGGU) — CV DIVISI 1
                </h3>
              </div>
              {/* Filter Brand & Total Tepat di Bagian Foto Terlampir */}
              <div className="flex items-center gap-1.5 flex-wrap">
                <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider flex items-center gap-1 shrink-0">
                  <Filter className="w-3 h-3 text-amber-400" />
                  <span>FILTER:</span>
                </span>
                <div className="inline-flex p-0.5 bg-slate-900/95 rounded-xl border border-slate-700/80 flex-wrap gap-1 shadow-inner">
                  <button
                    type="button"
                    onClick={() => setWeeklyBrandFilter('TOTAL')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                      weeklyBrandFilter === 'TOTAL'
                        ? 'bg-amber-400 text-slate-950 shadow-md ring-2 ring-amber-300'
                        : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                    }`}
                  >
                    <span className={`w-1.5 h-1.5 rounded-full ${weeklyBrandFilter === 'TOTAL' ? 'bg-slate-950' : 'bg-amber-400'}`} />
                    <span>TOTAL GABUNGAN CV</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setWeeklyBrandFilter('GRAPHIS')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                      weeklyBrandFilter === 'GRAPHIS'
                        ? 'bg-[#e6004c] text-white shadow-md ring-2 ring-rose-400'
                        : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                    }`}
                  >
                    <span className={`w-1.5 h-1.5 rounded-full ${weeklyBrandFilter === 'GRAPHIS' ? 'bg-white' : 'bg-pink-400'}`} />
                    <span>GRAPHIS (6201)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setWeeklyBrandFilter('SIMPLICITY')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                      weeklyBrandFilter === 'SIMPLICITY'
                        ? 'bg-indigo-600 text-white shadow-md ring-2 ring-indigo-400'
                        : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                    }`}
                  >
                    <span className={`w-1.5 h-1.5 rounded-full ${weeklyBrandFilter === 'SIMPLICITY' ? 'bg-white' : 'bg-sky-400'}`} />
                    <span>SIMPLICITY (6244)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setWeeklyBrandFilter('SPLIT')}
                    className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                      weeklyBrandFilter === 'SPLIT'
                        ? 'bg-gradient-to-r from-purple-800 to-indigo-800 text-white shadow-md ring-2 ring-purple-400'
                        : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                    }`}
                  >
                    <span className={`w-1.5 h-1.5 rounded-full ${weeklyBrandFilter === 'SPLIT' ? 'bg-white' : 'bg-purple-300'}`} />
                    <span>SPLIT (3 BARIS)</span>
                  </button>
                </div>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-900 text-white text-[10px] sm:text-[11px] font-black uppercase tracking-wider border-b border-slate-700">
                    <th className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 min-w-[85px]">
                      WEEK
                    </th>
                    <th className="py-3 px-3 border-r border-slate-800 min-w-[125px]">
                      PERIODE
                    </th>
                    <th className="py-3 px-2.5 border-r border-slate-800 min-w-[130px]">
                      BRAND / DEPT
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[105px]">
                      SENIN
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[105px]">
                      SELASA
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[105px]">
                      RABU
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[105px]">
                      KAMIS
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[105px]">
                      JUMAT
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[110px] bg-amber-950/40 text-amber-300">
                      SABTU <span className="text-[9px] px-1 py-0.2 rounded bg-amber-400/20 text-amber-300">WKND</span>
                    </th>
                    <th className="py-3 px-2 text-center border-r border-slate-800 min-w-[110px] bg-amber-950/40 text-amber-300">
                      MINGGU <span className="text-[9px] px-1 py-0.2 rounded bg-amber-400/20 text-amber-300">WKND</span>
                    </th>
                    <th className="py-3 px-3 border-r border-slate-800 text-right min-w-[125px] bg-purple-950/80 text-amber-300 font-black">
                      TOTAL SALES WEEK
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-center min-w-[75px]">
                      SALES QTY
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-center min-w-[65px]">
                      NOTA
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-right min-w-[95px]">
                      PRICE POINT
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-right min-w-[105px] bg-rose-950/40 text-rose-200">
                      BIG SIZE
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-center min-w-[90px]">
                      KONTRIBUSI
                    </th>
                    <th className="py-3 px-2 border-r border-slate-800 text-center min-w-[85px]">
                      GROWTH
                    </th>
                    <th className="py-3 px-2 text-center min-w-[65px]">
                      AKSI
                    </th>
                  </tr>
                </thead>

                <tbody className="divide-y divide-slate-200 font-medium">
                  {weeklyTableData.rows.map((r, rIdx) => {
                    const isSelected = selectedWeeklyRowNumber === r.weekNumber;

                    // Helper to render a day cell with date on top
                    const renderDayCell = (
                      d: { tanggal: string; dateShort: string; dayName: string; sales: number; qty: number; trx: number; isWeekend: boolean },
                      isTotalRow: boolean = false
                    ) => (
                      <td
                        key={d.tanggal}
                        className={`py-2 px-2 text-right font-mono-num border-r border-slate-200 ${
                          d.isWeekend
                            ? isTotalRow
                              ? 'bg-amber-100/80 font-black text-amber-950'
                              : 'bg-amber-50/70 font-bold text-amber-950'
                            : isTotalRow
                            ? 'bg-purple-50/50 font-black text-purple-950'
                            : 'font-bold text-slate-800'
                        }`}
                      >
                        <div className="flex flex-col items-end">
                          {/* Hari dan Tanggal */}
                          <span
                            className={`inline-block text-[9.5px] font-extrabold px-1.5 py-0.5 rounded leading-none mb-1 shadow-2xs ${
                              d.isWeekend
                                ? 'bg-amber-300 text-amber-950 ring-1 ring-amber-400'
                                : 'bg-slate-200 text-slate-800'
                            }`}
                          >
                            {d.dayName}, {d.dateShort}
                          </span>
                          {/* Angka sales */}
                          {d.sales > 0 ? (
                            <span className={isTotalRow ? 'font-black text-xs' : 'font-bold'}>
                              Rp {formatNumberId(d.sales)}
                            </span>
                          ) : (
                            <span className="text-slate-300 font-normal">—</span>
                          )}
                          {/* Qty pcs */}
                          {d.sales > 0 && (
                            <span className="text-[9px] text-slate-500 font-medium leading-tight">
                              {d.qty} pcs
                            </span>
                          )}
                        </div>
                      </td>
                    );

                    // Helper to render brand row metric cells
                    const renderMetricCells = (
                      m: typeof r.graphis,
                      isTotalRow: boolean = false
                    ) => (
                      <>
                        {/* 10. Total Sales Week */}
                        <td
                          className={`py-2.5 px-3 text-right font-mono-num font-black border-r border-slate-200 ${
                            isTotalRow
                              ? 'text-xs sm:text-sm text-purple-950 bg-purple-100/90 shadow-inner'
                              : 'text-xs text-slate-900 bg-slate-50/80'
                          }`}
                        >
                          {m.totalSalesWeek > 0 ? (
                            <div>
                              <div>Rp {formatNumberId(m.totalSalesWeek)}</div>
                              {!isTotalRow && r.total.totalSalesWeek > 0 && (
                                <div className="text-[9.5px] font-semibold text-slate-400 mt-0.5">
                                  Tot CV: Rp {formatNumberId(r.total.totalSalesWeek)}
                                </div>
                              )}
                            </div>
                          ) : (
                            <span className="text-slate-300 font-normal">—</span>
                          )}
                        </td>

                        {/* 11. Sales Qty */}
                        <td className="py-2.5 px-2 text-center font-mono-num font-bold text-slate-900 border-r border-slate-200">
                          {m.totalQtyWeek > 0 ? `${m.totalQtyWeek} pcs` : <span className="text-slate-300 font-normal">—</span>}
                        </td>

                        {/* 12. Nota */}
                        <td className="py-2.5 px-2 text-center font-mono-num font-bold text-slate-700 border-r border-slate-200">
                          {m.totalTrxWeek > 0 ? m.totalTrxWeek : <span className="text-slate-300 font-normal">—</span>}
                        </td>

                        {/* 13. Price Point (ASP) */}
                        <td className="py-2.5 px-2 text-right font-mono-num font-semibold text-slate-800 border-r border-slate-200">
                          {m.pricePoint > 0 ? (
                            `Rp ${formatNumberId(m.pricePoint)}`
                          ) : (
                            <span className="text-slate-300 font-normal">—</span>
                          )}
                        </td>

                        {/* 14. Big Size */}
                        <td className="py-2.5 px-2 text-right font-mono-num text-[10px] border-r border-slate-200 bg-rose-50/30">
                          {m.bigSizeSales > 0 ? (
                            <div>
                              <div className="font-black text-rose-950">Rp {formatNumberId(m.bigSizeSales)}</div>
                              <div className="text-slate-500">{m.bigSizeQty} pcs</div>
                            </div>
                          ) : (
                            <span className="text-slate-300 font-normal">—</span>
                          )}
                        </td>

                        {/* 15. Kontribusi (%) */}
                        <td className="py-2.5 px-2 text-center border-r border-slate-200">
                          <div className="flex flex-col items-center gap-1">
                            <span className="font-black font-mono-num text-slate-900 text-[11px]">
                              {m.kontribusiPct.toFixed(1)}%
                            </span>
                            <div className="w-10 bg-slate-200 rounded-full h-1 overflow-hidden">
                              <div
                                className={`${isTotalRow ? 'bg-purple-600' : 'bg-indigo-500'} h-1 rounded-full`}
                                style={{ width: `${Math.min(m.kontribusiPct, 100)}%` }}
                              />
                            </div>
                          </div>
                        </td>

                        {/* 16. Growth (%) */}
                        <td className="py-2.5 px-2 text-center border-r border-slate-200">
                          {m.growthPct !== null ? (
                            <span
                              className={`inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full font-black text-[9px] font-mono-num ${
                                m.growthPct >= 0
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : 'bg-rose-100 text-rose-800'
                              }`}
                            >
                              {m.growthPct >= 0 ? '+' : ''}
                              {m.growthPct.toFixed(1)}%
                              {m.growthPct >= 0 ? '▲' : '▼'}
                            </span>
                          ) : (
                            <span className="text-[9px] text-slate-400 font-bold">Base</span>
                          )}
                        </td>
                      </>
                    );

                    if (weeklyBrandFilter === 'SPLIT') {
                      // 3 ROWS PER WEEK: GRAPHIS, SIMPLICITY, TOTAL GABUNGAN
                      return (
                        <React.Fragment key={r.weekNumber}>
                          {/* Row 1: GRAPHIS (Dept 6201) */}
                          <tr
                            onClick={() => setSelectedWeeklyRowNumber(r.weekNumber)}
                            className={`transition-colors cursor-pointer text-[11px] ${
                              isSelected
                                ? 'bg-purple-100/50'
                                : rIdx % 2 === 0
                                ? 'bg-white hover:bg-slate-50/80'
                                : 'bg-slate-50/40 hover:bg-slate-100/60'
                            }`}
                          >
                            {/* 1. Week Keberapa (rowSpan=3) */}
                            <td
                              rowSpan={3}
                              className="py-3 px-2.5 border-r border-slate-300 text-center sticky left-0 z-10 bg-white shadow-2xs align-middle"
                            >
                              <div className="inline-flex flex-col items-center">
                                <span className="px-2 py-0.5 rounded-lg text-xs font-black bg-purple-900 text-white shadow-2xs">
                                  {r.weekLabel}
                                </span>
                                <span className="text-[10px] font-bold text-slate-500 mt-1">
                                  W-{r.weekNumber}
                                </span>
                              </div>
                            </td>

                            {/* 2. Tanggal / Periode (rowSpan=3) */}
                            <td
                              rowSpan={3}
                              className="py-3 px-3 border-r border-slate-300 font-bold text-slate-800 align-middle bg-slate-50/60"
                            >
                              <div className="leading-tight text-xs">{r.periodeText}</div>
                              <div className="text-[10px] text-slate-400 font-semibold mt-0.5">
                                {r.monthName} 2026
                              </div>
                            </td>

                            {/* Brand Graphis Indicator */}
                            <td className="py-2.5 px-2.5 border-r border-slate-200 bg-pink-50/40">
                              <div className="flex items-center gap-1.5">
                                <span className="w-2.5 h-2.5 rounded-full bg-pink-600 ring-2 ring-pink-200 shrink-0" />
                                <div>
                                  <div className="font-black text-rose-950 text-xs leading-none">GRAPHIS</div>
                                  <div className="text-[9px] text-rose-700 font-bold mt-0.5">Dept 6201</div>
                                </div>
                              </div>
                            </td>

                            {/* 7 Days Graphis */}
                            {r.graphis.days.map((d) => renderDayCell(d, false))}

                            {/* Metrics Graphis */}
                            {renderMetricCells(r.graphis, false)}

                            {/* 17. Aksi (rowSpan=3) */}
                            <td
                              rowSpan={3}
                              className="py-3 px-2 text-center align-middle bg-white border-l border-slate-200"
                            >
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setSelectedWeeklyRowNumber(r.weekNumber);
                                }}
                                className={`p-2 rounded-lg text-xs font-bold transition-colors cursor-pointer ${
                                  isSelected
                                    ? 'bg-purple-700 text-white shadow-2xs'
                                    : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                                }`}
                                title="Buka rincian 7 hari week ini"
                              >
                                <Eye className="w-4 h-4" />
                              </button>
                            </td>
                          </tr>

                          {/* Row 2: SIMPLICITY (Dept 6244) */}
                          <tr
                            onClick={() => setSelectedWeeklyRowNumber(r.weekNumber)}
                            className={`transition-colors cursor-pointer text-[11px] ${
                              isSelected
                                ? 'bg-purple-100/50'
                                : rIdx % 2 === 0
                                ? 'bg-white hover:bg-slate-50/80'
                                : 'bg-slate-50/40 hover:bg-slate-100/60'
                            }`}
                          >
                            {/* Brand Simplicity Indicator */}
                            <td className="py-2.5 px-2.5 border-r border-slate-200 bg-blue-50/40">
                              <div className="flex items-center gap-1.5">
                                <span className="w-2.5 h-2.5 rounded-full bg-blue-600 ring-2 ring-blue-200 shrink-0" />
                                <div>
                                  <div className="font-black text-blue-950 text-xs leading-none">SIMPLICITY</div>
                                  <div className="text-[9px] text-blue-700 font-bold mt-0.5">Dept 6244</div>
                                </div>
                              </div>
                            </td>

                            {/* 7 Days Simplicity */}
                            {r.simplicity.days.map((d) => renderDayCell(d, false))}

                            {/* Metrics Simplicity */}
                            {renderMetricCells(r.simplicity, false)}
                          </tr>

                          {/* Row 3: TOTAL GABUNGAN DIVISI 1 CV */}
                          <tr
                            onClick={() => setSelectedWeeklyRowNumber(r.weekNumber)}
                            className={`transition-colors cursor-pointer text-[11px] border-b-2 border-slate-300 ${
                              isSelected
                                ? 'bg-purple-100/80'
                                : 'bg-purple-50/30 hover:bg-purple-100/40'
                            }`}
                          >
                            {/* Total Gabungan Brand Indicator */}
                            <td className="py-2.5 px-2.5 border-r border-slate-200 bg-purple-100/60 font-black">
                              <div className="flex items-center gap-1.5">
                                <span className="w-2.5 h-2.5 rounded-full bg-purple-700 ring-2 ring-purple-300 shrink-0" />
                                <div>
                                  <div className="font-black text-purple-950 text-xs leading-none">TOTAL CV</div>
                                  <div className="text-[9px] text-purple-700 font-bold mt-0.5">Gabungan Divisi 1</div>
                                </div>
                              </div>
                            </td>

                            {/* 7 Days Total */}
                            {r.total.days.map((d) => renderDayCell(d, true))}

                            {/* Metrics Total */}
                            {renderMetricCells(r.total, true)}
                          </tr>
                        </React.Fragment>
                      );
                    }

                    // SINGLE BRAND FILTER MODE (GRAPHIS, SIMPLICITY, OR TOTAL ONLY)
                    const activeBrandObj =
                      weeklyBrandFilter === 'GRAPHIS'
                        ? r.graphis
                        : weeklyBrandFilter === 'SIMPLICITY'
                        ? r.simplicity
                        : r.total;

                    return (
                      <tr
                        key={r.weekNumber}
                        onClick={() => setSelectedWeeklyRowNumber(r.weekNumber)}
                        className={`transition-colors cursor-pointer text-[11px] ${
                          isSelected
                            ? 'bg-purple-100/60 ring-2 ring-purple-600/30'
                            : rIdx % 2 === 0
                            ? 'bg-white hover:bg-slate-50'
                            : 'bg-slate-50/60 hover:bg-slate-100/70'
                        }`}
                      >
                        {/* 1. Week Keberapa */}
                        <td className="py-3 px-2.5 border-r border-slate-200 text-center sticky left-0 z-10 bg-inherit shadow-2xs">
                          <div className="inline-flex flex-col items-center">
                            <span className="px-2 py-0.5 rounded-lg text-xs font-black bg-purple-900 text-white shadow-2xs">
                              {r.weekLabel}
                            </span>
                            <span className="text-[10px] font-bold text-slate-500 mt-0.5">
                              W-{r.weekNumber}
                            </span>
                          </div>
                        </td>

                        {/* 2. Tanggal (Periode) */}
                        <td className="py-3 px-3 border-r border-slate-200 font-bold text-slate-800">
                          <div className="leading-tight">{r.periodeText}</div>
                          <div className="text-[10px] text-slate-400 font-semibold">{r.monthName} 2026</div>
                        </td>

                        {/* Brand / Dept Indicator */}
                        <td className={`py-2.5 px-2.5 border-r border-slate-200 ${
                          weeklyBrandFilter === 'GRAPHIS'
                            ? 'bg-pink-50/50'
                            : weeklyBrandFilter === 'SIMPLICITY'
                            ? 'bg-blue-50/50'
                            : 'bg-purple-50/50'
                        }`}>
                          <div className="flex items-center gap-1.5">
                            <span
                              className={`w-2.5 h-2.5 rounded-full shrink-0 ${
                                weeklyBrandFilter === 'GRAPHIS'
                                  ? 'bg-pink-600 ring-2 ring-pink-200'
                                  : weeklyBrandFilter === 'SIMPLICITY'
                                  ? 'bg-blue-600 ring-2 ring-blue-200'
                                  : 'bg-purple-700 ring-2 ring-purple-200'
                              }`}
                            />
                            <div>
                              <div className="font-black text-xs leading-none">
                                {weeklyBrandFilter === 'GRAPHIS'
                                  ? 'GRAPHIS'
                                  : weeklyBrandFilter === 'SIMPLICITY'
                                  ? 'SIMPLICITY'
                                  : 'TOTAL CV'}
                              </div>
                              <div className="text-[9px] font-bold mt-0.5 text-slate-500">
                                {weeklyBrandFilter === 'GRAPHIS'
                                  ? 'Dept 6201'
                                  : weeklyBrandFilter === 'SIMPLICITY'
                                  ? 'Dept 6244'
                                  : 'Divisi 1 Gabungan'}
                              </div>
                            </div>
                          </div>
                        </td>

                        {/* 7 Days with date on top */}
                        {activeBrandObj.days.map((d) => renderDayCell(d, weeklyBrandFilter === 'TOTAL'))}

                        {/* Metrics */}
                        {renderMetricCells(activeBrandObj, weeklyBrandFilter === 'TOTAL')}

                        {/* 17. Aksi */}
                        <td className="py-3 px-2.5 text-center">
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setSelectedWeeklyRowNumber(r.weekNumber);
                            }}
                            className={`p-1.5 rounded-lg text-xs font-bold transition-colors cursor-pointer ${
                              isSelected
                                ? 'bg-purple-700 text-white shadow-2xs'
                                : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                            }`}
                            title="Buka rincian 7 hari"
                          >
                            <Eye className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    );
                  })}

                  {/* 4. TOTAL REKAPITULASI (FOOTER ROWS) */}
                  {weeklyBrandFilter === 'SPLIT' ? (
                    <>
                      {/* Summary Graphis */}
                      <tr className="bg-pink-950/90 text-white font-bold text-xs border-t-2 border-pink-400">
                        <td colSpan={3} className="py-2.5 px-3 text-left uppercase tracking-wider text-pink-200 border-r border-pink-900">
                          <div className="flex items-center gap-2">
                            <span className="w-2.5 h-2.5 rounded-full bg-pink-400" />
                            <span>TOTAL GRAPHIS (DEPT 6201)</span>
                          </div>
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.senin > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.senin)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.selasa > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.selasa)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.rabu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.rabu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.kamis > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.kamis)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.jumat > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.jumat)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-400/10 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.sabtu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.sabtu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-400/10 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.minggu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.graphis.minggu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono-num font-black text-amber-300 bg-pink-900/60 border-r border-pink-900">
                          Rp {formatNumberId(weeklyTableData.summary.graphis.grandTotalSales)}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.grandTotalQty} pcs
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-pink-100 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.grandTotalTrx}
                        </td>
                        <td className="py-2.5 px-2 text-right font-mono-num text-sky-200 border-r border-pink-900">
                          Rp {formatNumberId(weeklyTableData.summary.graphis.avgPricePoint)}
                        </td>
                        <td className="py-2.5 px-2 text-right font-mono-num text-rose-300 border-r border-pink-900">
                          Rp {formatNumberId(weeklyTableData.summary.graphis.grandTotalBigSizeSales)}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-purple-200 border-r border-pink-900">
                          {weeklyTableData.summary.grandTotalSales > 0
                            ? `${((weeklyTableData.summary.graphis.grandTotalSales / weeklyTableData.summary.grandTotalSales) * 100).toFixed(1)}%`
                            : '0%'}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-pink-900">
                          {weeklyTableData.summary.graphis.avgGrowthPct >= 0 ? '+' : ''}
                          {weeklyTableData.summary.graphis.avgGrowthPct.toFixed(1)}%
                        </td>
                        <td className="py-2.5 px-2 text-center text-pink-400">—</td>
                      </tr>

                      {/* Summary Simplicity */}
                      <tr className="bg-blue-950/90 text-white font-bold text-xs border-t border-blue-800">
                        <td colSpan={3} className="py-2.5 px-3 text-left uppercase tracking-wider text-blue-200 border-r border-blue-900">
                          <div className="flex items-center gap-2">
                            <span className="w-2.5 h-2.5 rounded-full bg-blue-400" />
                            <span>TOTAL SIMPLICITY (DEPT 6244)</span>
                          </div>
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.senin > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.senin)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.selasa > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.selasa)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.rabu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.rabu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.kamis > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.kamis)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.jumat > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.jumat)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-400/10 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.sabtu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.sabtu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-400/10 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.minggu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.simplicity.minggu)}` : '—'}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono-num font-black text-amber-300 bg-blue-900/60 border-r border-blue-900">
                          Rp {formatNumberId(weeklyTableData.summary.simplicity.grandTotalSales)}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.grandTotalQty} pcs
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-blue-100 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.grandTotalTrx}
                        </td>
                        <td className="py-2.5 px-2 text-right font-mono-num text-sky-200 border-r border-blue-900">
                          Rp {formatNumberId(weeklyTableData.summary.simplicity.avgPricePoint)}
                        </td>
                        <td className="py-2.5 px-2 text-right font-mono-num text-rose-300 border-r border-blue-900">
                          Rp {formatNumberId(weeklyTableData.summary.simplicity.grandTotalBigSizeSales)}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-purple-200 border-r border-blue-900">
                          {weeklyTableData.summary.grandTotalSales > 0
                            ? `${((weeklyTableData.summary.simplicity.grandTotalSales / weeklyTableData.summary.grandTotalSales) * 100).toFixed(1)}%`
                            : '0%'}
                        </td>
                        <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-blue-900">
                          {weeklyTableData.summary.simplicity.avgGrowthPct >= 0 ? '+' : ''}
                          {weeklyTableData.summary.simplicity.avgGrowthPct.toFixed(1)}%
                        </td>
                        <td className="py-2.5 px-2 text-center text-blue-400">—</td>
                      </tr>

                      {/* Grand Total Konsolidasi Divisi 1 */}
                      <tr className="bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs border-t-2 border-amber-400 sticky bottom-0 z-20 shadow-lg">
                        <td colSpan={3} className="py-3 px-3 text-left uppercase tracking-wider text-amber-300 border-r border-slate-700">
                          <div className="flex items-center gap-2">
                            <span className="w-2.5 h-2.5 rounded-full bg-amber-400 ring-2 ring-amber-300" />
                            <span>GRAND TOTAL KONSOLIDASI (CV DIVISI 1)</span>
                          </div>
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.senin > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.senin)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.selasa > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.selasa)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.rabu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.rabu)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.kamis > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.kamis)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.jumat > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.jumat)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/10 border-r border-slate-700">
                          {weeklyTableData.summary.sabtu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.sabtu)}` : '—'}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/10 border-r border-slate-700">
                          {weeklyTableData.summary.minggu > 0 ? `Rp ${formatNumberId(weeklyTableData.summary.minggu)}` : '—'}
                        </td>
                        <td className="py-3 px-3 text-right font-mono-num font-black text-sm text-amber-300 bg-amber-500/20 border-r border-slate-700">
                          Rp {formatNumberId(weeklyTableData.summary.grandTotalSales)}
                        </td>
                        <td className="py-3 px-2.5 text-center font-mono-num text-emerald-300 border-r border-slate-700">
                          {weeklyTableData.summary.grandTotalQty} pcs
                        </td>
                        <td className="py-3 px-2.5 text-center font-mono-num text-slate-200 border-r border-slate-700">
                          {weeklyTableData.summary.grandTotalTrx}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-sky-300 border-r border-slate-700">
                          Rp {formatNumberId(weeklyTableData.summary.avgPricePoint)}
                        </td>
                        <td className="py-3 px-2.5 text-right font-mono-num text-rose-300 border-r border-slate-700">
                          Rp {formatNumberId(weeklyTableData.summary.grandTotalBigSizeSales)}
                        </td>
                        <td className="py-3 px-2.5 text-center font-mono-num text-purple-300 border-r border-slate-700">
                          100.0%
                        </td>
                        <td className="py-3 px-2.5 text-center font-mono-num text-emerald-300 border-r border-slate-700">
                          {weeklyTableData.summary.avgGrowthPct >= 0 ? '+' : ''}
                          {weeklyTableData.summary.avgGrowthPct.toFixed(1)}%
                        </td>
                        <td className="py-3 px-2.5 text-center text-amber-400">✓</td>
                      </tr>
                    </>
                  ) : (
                    /* Single Brand Footer */
                    (() => {
                      const activeSum =
                        weeklyBrandFilter === 'GRAPHIS'
                          ? weeklyTableData.summary.graphis
                          : weeklyBrandFilter === 'SIMPLICITY'
                          ? weeklyTableData.summary.simplicity
                          : weeklyTableData.summary.total;

                      const isSingleBrandMode = weeklyBrandFilter === 'GRAPHIS' || weeklyBrandFilter === 'SIMPLICITY';
                      const brandLabelUpper = weeklyBrandFilter === 'GRAPHIS' ? 'TOTAL BRAND GRAPHIS (DEPT 6201)' : 'TOTAL BRAND SIMPLICITY (DEPT 6244)';
                      const brandBgClass = weeklyBrandFilter === 'GRAPHIS' ? 'bg-pink-950/90 border-pink-500' : 'bg-blue-950/90 border-blue-500';

                      return (
                        <>
                          {/* Active Brand Total Row */}
                          <tr className={`${isSingleBrandMode ? brandBgClass : 'bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950'} text-white font-black text-xs border-t-2 border-amber-400 sticky bottom-0 z-20 shadow-lg`}>
                            <td colSpan={3} className="py-3 px-3 text-left uppercase tracking-wider text-amber-300 border-r border-slate-700">
                              <div className="flex items-center gap-2">
                                <span className="w-2.5 h-2.5 rounded-full bg-amber-400" />
                                <span>{isSingleBrandMode ? brandLabelUpper : `TOTAL REKAPITULASI CV (${weeklyTableData.rows.length} WEEK)`}</span>
                              </div>
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.senin > 0 ? `Rp ${formatNumberId(activeSum.senin)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.selasa > 0 ? `Rp ${formatNumberId(activeSum.selasa)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.rabu > 0 ? `Rp ${formatNumberId(activeSum.rabu)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.kamis > 0 ? `Rp ${formatNumberId(activeSum.kamis)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.jumat > 0 ? `Rp ${formatNumberId(activeSum.jumat)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/10 border-r border-slate-700">
                              {activeSum.sabtu > 0 ? `Rp ${formatNumberId(activeSum.sabtu)}` : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/10 border-r border-slate-700">
                              {activeSum.minggu > 0 ? `Rp ${formatNumberId(activeSum.minggu)}` : '—'}
                            </td>
                            <td className="py-3 px-3 text-right font-mono-num font-black text-sm text-amber-300 bg-amber-500/20 border-r border-slate-700">
                              Rp {formatNumberId(activeSum.grandTotalSales)}
                            </td>
                            <td className="py-3 px-2.5 text-center font-mono-num text-emerald-300 border-r border-slate-700">
                              {activeSum.grandTotalQty} pcs
                            </td>
                            <td className="py-3 px-2.5 text-center font-mono-num text-slate-200 border-r border-slate-700">
                              {activeSum.grandTotalTrx}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-sky-300 border-r border-slate-700">
                              Rp {formatNumberId(activeSum.avgPricePoint)}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-rose-300 border-r border-slate-700">
                              Rp {formatNumberId(activeSum.grandTotalBigSizeSales)}
                            </td>
                            <td className="py-3 px-2.5 text-center font-mono-num text-purple-300 border-r border-slate-700">
                              {weeklyTableData.summary.grandTotalSales > 0
                                ? `${((activeSum.grandTotalSales / weeklyTableData.summary.grandTotalSales) * 100).toFixed(1)}%`
                                : '100.0%'}
                            </td>
                            <td className="py-3 px-2.5 text-center font-mono-num text-emerald-300 border-r border-slate-700">
                              {activeSum.avgGrowthPct >= 0 ? '+' : ''}
                              {activeSum.avgGrowthPct.toFixed(1)}%
                            </td>
                            <td className="py-3 px-2.5 text-center text-slate-400">✓</td>
                          </tr>

                          {/* If single brand, also render Grand Total CV Konsolidasi row for full context */}
                          {isSingleBrandMode && (
                            <tr className="bg-slate-950 text-white font-extrabold text-xs border-t border-slate-800">
                              <td colSpan={3} className="py-2.5 px-3 text-left uppercase tracking-wider text-amber-200 border-r border-slate-800">
                                <div className="flex items-center gap-2">
                                  <span className="w-2 h-2 rounded-full bg-indigo-400" />
                                  <span className="text-[11px]">BANDINGAN: TOTAL GABUNGAN CV (DIVISI 1)</span>
                                </div>
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.senin)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.selasa)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.rabu)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.kamis)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.jumat)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-500/10 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.sabtu)}
                              </td>
                              <td className="py-2.5 px-2.5 text-right font-mono-num text-amber-200 bg-amber-500/10 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.minggu)}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono-num font-black text-amber-300 bg-amber-500/20 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.grandTotalSales)}
                              </td>
                              <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-slate-800">
                                {weeklyTableData.summary.total.grandTotalQty} pcs
                              </td>
                              <td className="py-2.5 px-2 text-center font-mono-num text-slate-300 border-r border-slate-800">
                                {weeklyTableData.summary.total.grandTotalTrx}
                              </td>
                              <td className="py-2.5 px-2 text-right font-mono-num text-sky-200 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.avgPricePoint)}
                              </td>
                              <td className="py-2.5 px-2 text-right font-mono-num text-rose-300 border-r border-slate-800">
                                Rp {formatNumberId(weeklyTableData.summary.total.grandTotalBigSizeSales)}
                              </td>
                              <td className="py-2.5 px-2 text-center font-mono-num text-purple-200 border-r border-slate-800">
                                100.0%
                              </td>
                              <td className="py-2.5 px-2 text-center font-mono-num text-emerald-300 border-r border-slate-800">
                                {weeklyTableData.summary.total.avgGrowthPct >= 0 ? '+' : ''}
                                {weeklyTableData.summary.total.avgGrowthPct.toFixed(1)}%
                              </td>
                              <td className="py-2.5 px-2 text-center text-slate-500">✓</td>
                            </tr>
                          )}
                        </>
                      );
                    })()
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 5. DEEP-DIVE RINCIAN 7 HARI UNTUK WEEK TERPILIH */}
          {(() => {
            const activeWeekObj =
              weeklyTableData.rows.find((r) => r.weekNumber === selectedWeeklyRowNumber) ||
              weeklyTableData.rows[0];

            if (!activeWeekObj) return null;

            return (
              <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-xs space-y-0">
                <div className="px-4 py-3 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-3 border-b border-slate-800">
                  <div className="flex items-center gap-2">
                    <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase bg-amber-400 text-slate-950">
                      {activeWeekObj.weekLabel}
                    </span>
                    <h4 className="text-xs sm:text-sm font-black text-white">
                      DETAIL PENJUALAN HARIAN — {activeWeekObj.periodeText} (W-{activeWeekObj.weekNumber})
                    </h4>
                  </div>

                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-amber-300 font-mono-num">
                      Total Week: Rp {formatNumberId(activeWeekObj.totalSalesWeek)} ({activeWeekObj.totalQtyWeek} pcs)
                    </span>
                    {isAdmin && (
                      <button
                        type="button"
                        onClick={() => handleOpenEditModal(activeWeekObj.days[0]?.tanggal || selectedTanggal)}
                        className="px-2.5 py-1 rounded-lg bg-amber-400 hover:bg-amber-300 text-slate-950 text-xs font-bold inline-flex items-center gap-1 cursor-pointer"
                      >
                        <PenSquare className="w-3.5 h-3.5" />
                        <span>Edit Data Week</span>
                      </button>
                    )}
                  </div>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 font-extrabold text-[11px] border-b border-slate-200">
                        <th className="py-2.5 px-3">Hari &amp; Tanggal</th>
                        <th className="py-2.5 px-3 text-right">Dept 6201 Graphis</th>
                        <th className="py-2.5 px-3 text-center">Qty G</th>
                        <th className="py-2.5 px-3 text-right">Dept 6244 Simplicity</th>
                        <th className="py-2.5 px-3 text-center">Qty S</th>
                        <th className="py-2.5 px-3 text-right bg-purple-50 text-purple-950 font-black">Total Sales CV</th>
                        <th className="py-2.5 px-3 text-center">Total Qty</th>
                        <th className="py-2.5 px-3 text-center">Nota</th>
                        <th className="py-2.5 px-3 text-right">ASP</th>
                        <th className="py-2.5 px-3 text-center">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 font-medium">
                      {activeWeekObj.days.map((d, dIdx) => {
                        const isWknd = d.dayName === 'Sabtu' || d.dayName === 'Minggu';
                        const gDay = activeWeekObj.graphis.days[dIdx] || { sales: 0, qty: 0 };
                        const sDay = activeWeekObj.simplicity.days[dIdx] || { sales: 0, qty: 0 };
                        const dayAsp = d.qty > 0 ? Math.round(d.sales / d.qty) : 0;
                        return (
                          <tr
                            key={d.tanggal}
                            className={`hover:bg-purple-50/50 transition-colors ${
                              isWknd ? 'bg-amber-50/30' : dIdx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'
                            }`}
                          >
                            <td className="py-2.5 px-3 font-bold text-slate-900">
                              <span>{d.dayName}, {d.tanggal}</span>
                              {isWknd && (
                                <span className="ml-1.5 px-1.5 py-0.2 rounded text-[9px] bg-amber-200 text-amber-900 font-extrabold">
                                  WKND
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-num font-bold text-rose-900">
                              {gDay.sales > 0 ? `Rp ${formatNumberId(gDay.sales)}` : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono-num text-slate-700">
                              {gDay.qty > 0 ? gDay.qty : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-num font-bold text-indigo-900">
                              {sDay.sales > 0 ? `Rp ${formatNumberId(sDay.sales)}` : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono-num text-slate-700">
                              {sDay.qty > 0 ? sDay.qty : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-num font-black text-purple-950 bg-purple-50/60">
                              {d.sales > 0 ? `Rp ${formatNumberId(d.sales)}` : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono-num font-bold text-slate-800">
                              {d.qty > 0 ? `${d.qty} pcs` : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center font-mono-num text-slate-700">
                              {d.trx > 0 ? d.trx : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                              {dayAsp > 0 ? `Rp ${formatNumberId(dayAsp)}` : '—'}
                            </td>
                            <td className="py-2.5 px-3 text-center">
                              {isAdmin && (
                                <button
                                  type="button"
                                  onClick={() => handleOpenEditModal(d.tanggal)}
                                  className="px-2 py-1 rounded bg-amber-400 hover:bg-amber-300 text-slate-950 text-[11px] font-bold cursor-pointer inline-flex items-center gap-1 shadow-2xs"
                                >
                                  <PenSquare className="w-3 h-3" />
                                  <span>Edit</span>
                                </button>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          {/* ============================================================ */}
          {/* CHART DIAGRAM SALES CV WEEKLY (SEPERTI REPORT DP)            */}
          {/* ============================================================ */}
          {(() => {
            const activeWeekObj =
              weeklyTableData.rows.find((r) => r.weekNumber === selectedWeeklyRowNumber) ||
              weeklyTableData.rows[0];

            return (
              <div className="space-y-4">
                {activeWeekObj && (
                  <PhotoStyleSalesDiagram
                    headerTitle={`DIAGRAM SALES CV WEEKLY — ${activeWeekObj.weekLabel} (${activeWeekObj.periodeText})`}
                    headerSubtitle="Diagram Batang Vertikal 7 Hari & Garis Tren Penjualan Graphis vs Simplicity"
                    badgeText="DIAGRAM SALES CV WEEKLY"
                    cardTitle={`Grafik Batang Harian ${activeWeekObj.weekLabel} (W-${activeWeekObj.weekNumber})`}
                    primaryLabel="Graphis (6201)"
                    secondaryLabel="Simplicity (6244)"
                    totalSales={activeWeekObj.totalSalesWeek}
                    totalPagi={activeWeekObj.graphis.totalSalesWeek}
                    totalSiang={activeWeekObj.simplicity.totalSalesWeek}
                    bars={activeWeekObj.days.map((d, dIdx) => {
                      const gSales = activeWeekObj.graphis.days[dIdx]?.sales || 0;
                      const sSales = activeWeekObj.simplicity.days[dIdx]?.sales || 0;
                      return {
                        key: d.tanggal,
                        label: d.dayName.slice(0, 3),
                        subLabel: d.dateShort,
                        pagiSales: gSales,
                        siangSales: sSales,
                        totalSales: gSales + sSales,
                        isWeekend: d.isWeekend,
                        isSelected: d.tanggal === selectedTanggal,
                        onClick: () => setSelectedTanggal(d.tanggal),
                      };
                    })}
                  />
                )}

                {weeklyTableData.rows.length > 0 && (
                  <>
                    <PhotoStyleSalesDiagram
                      headerTitle={`DIAGRAM PERBANDINGAN ANTAR WEEKLY CV (${weeklyTableData.rows.length} WEEKS)`}
                      headerSubtitle="Diagram Batang Vertikal Mingguan & Garis Tren Penjualan Graphis vs Simplicity"
                      badgeText="DIAGRAM ANTAR WEEKLY CV"
                      cardTitle="Grafik Perbandingan Sales Weekly (Graphis vs Simplicity)"
                      primaryLabel="Graphis (6201)"
                      secondaryLabel="Simplicity (6244)"
                      totalSales={weeklyTableData.periodGrandSales}
                      totalPagi={weeklyTableData.periodGraphisGrandSales}
                      totalSiang={weeklyTableData.periodSimplicityGrandSales}
                      compactX={weeklyTableData.rows.length > 12}
                      bars={weeklyTableData.rows.map((r) => ({
                        key: `cv-wk-${r.weekNumber}`,
                        label: `W-${r.weekNumber}`,
                        subLabel: r.periodeText.split('-')[0].trim(),
                        pagiSales: r.graphis.totalSalesWeek,
                        siangSales: r.simplicity.totalSalesWeek,
                        totalSales: r.totalSalesWeek,
                        isSelected: r.weekNumber === selectedWeeklyRowNumber,
                        onClick: () => setSelectedWeeklyRowNumber(r.weekNumber),
                      }))}
                    />

                    <GrowthAndContributionBarDiagram
                      headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI WEEKLY CV (${weeklyTableData.rows.length} WEEKS)`}
                      headerSubtitle="Grafik Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Sales CV Tiap Weekly"
                      periodBadgeText={`${weeklyTableData.rows.length} PERIODE WEEKLY`}
                      compactX={weeklyTableData.rows.length > 10}
                      items={weeklyTableData.rows.map((r, idx) => ({
                        key: `cv-wk-gc-${r.weekNumber}`,
                        label: `W-${r.weekNumber}`,
                        subLabel: r.periodeText.split('-')[0].trim(),
                        pagiSales: r.graphis.totalSalesWeek,
                        siangSales: r.simplicity.totalSalesWeek,
                        totalSales: r.totalSalesWeek,
                        prevSales: idx > 0 ? weeklyTableData.rows[idx - 1].totalSalesWeek : null,
                        isSelected: r.weekNumber === selectedWeeklyRowNumber,
                        onClick: () => setSelectedWeeklyRowNumber(r.weekNumber),
                      }))}
                    />
                  </>
                )}
              </div>
            );
          })()}
        </div>
      )}

      {/* ============================================================ */}
      {/* HORIZON 3: SALES BULANAN                                     */}
      {/* ============================================================ */}
      {horizon === 'bulanan' && (
        <div className="space-y-6">
          {/* 1. Header Control Bar & View Selector */}
          <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs space-y-3">
            <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-sky-100 text-sky-800 border border-sky-200">
                    Format Terbaru CV Monthly
                  </span>
                  <span className="text-xs font-bold text-slate-500">
                    Divisi 1 · Tanjung Outmall 311
                  </span>
                </div>
                <h2 className="text-base sm:text-lg font-black text-slate-900 tracking-tight mt-1 flex items-center gap-2">
                  <CalendarRange className="w-5 h-5 text-sky-600" />
                  <span>REKAPITULASI SALES MONTHLY REPORT CV (PER WEEK & 1 TAHUN)</span>
                </h2>
              </div>

              {/* Action Buttons */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleCopyWhatsApp}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs cursor-pointer transition-colors"
                  title="Salin laporan bulanan ke WhatsApp"
                >
                  <Share2 className="w-3.5 h-3.5" />
                  <span>Salin WA Bulanan</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white shadow-2xs cursor-pointer transition-colors"
                  title="Download file Excel 12 kolom lengkap"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Ekspor Excel</span>
                </button>

                {isAdmin && (
                  <button
                    type="button"
                    onClick={() => handleOpenEditModal(selectedTanggal)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 shadow-2xs cursor-pointer transition-colors"
                    title="Input sales report CV tanggal terpilih"
                  >
                    <PenSquare className="w-3.5 h-3.5" />
                    <span>Input Sales CV</span>
                  </button>
                )}
              </div>
            </div>

            {/* Filter Ribbons: Pilih Bulan & Format Tampilan Selector */}
            <div className="pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex flex-wrap items-center gap-3">
                {/* Pilih Bulan */}
                <div className="flex items-center gap-1.5">
                  <span className="font-bold text-slate-600 flex items-center gap-1">
                    <CalendarDays className="w-3.5 h-3.5 text-sky-600" />
                    <span>Pilih Bulan:</span>
                  </span>
                  <select
                    value={selectedMonthKey}
                    onChange={(e) => setSelectedMonthKey(e.target.value)}
                    className="px-2.5 py-1 text-xs font-bold bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-sky-500 cursor-pointer"
                  >
                    {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                      <option key={mKey} value={mKey}>
                        {mName} 2026
                      </option>
                    ))}
                  </select>
                  <span className="text-[11px] font-semibold text-slate-500 ml-1">
                    ({currentMonthlySummary.daysCount} Hari Aktif)
                  </span>
                </div>

                {/* Tampilan Format Selector Pills */}
                <div className="inline-flex p-0.5 bg-slate-100 rounded-lg border border-slate-200 flex-wrap gap-0.5">
                  <button
                    type="button"
                    onClick={() => setMonthlyViewFormat('both')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      monthlyViewFormat === 'both'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    🌟 Semua Format
                  </button>
                  <button
                    type="button"
                    onClick={() => setMonthlyViewFormat('per_week')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      monthlyViewFormat === 'per_week'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    📊 Format Per Week
                  </button>
                  <button
                    type="button"
                    onClick={() => setMonthlyViewFormat('per_bulan')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      monthlyViewFormat === 'per_bulan'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    📅 Format Perbulan (1 Tahun)
                  </button>
                  <button
                    type="button"
                    onClick={() => setMonthlyViewFormat('harian')}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      monthlyViewFormat === 'harian'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    📝 Rincian Harian MTD
                  </button>
                </div>
              </div>

              <div className="text-[11px] font-bold text-slate-500">
                Tahun Buku: <span className="text-purple-700 font-black">2026</span>
              </div>
            </div>
          </div>

          {/* 2. Executive KPI Cards */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {/* Card 1: Total Sales MTD Bulan Terpilih */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-2xs space-y-1">
              <div className="text-[11px] font-bold text-slate-500 uppercase flex items-center justify-between">
                <span>Total Sales MTD ({currentMonthlySummary.monthName})</span>
                <span className="w-2 h-2 rounded-full bg-emerald-500" />
              </div>
              <div className="text-lg sm:text-2xl font-black text-slate-950 font-mono-num">
                Rp {formatNumberId(currentMonthlySummary.total.sales)}
              </div>
              <div className="text-[10px] text-slate-500 font-semibold flex items-center justify-between pt-0.5">
                <span>{currentMonthlySummary.total.qty} pcs · {currentMonthlySummary.total.trx} nota</span>
                <span className="text-sky-700 font-bold">Rp {formatNumberId(currentMonthlySummary.avgDailySales)}/hr</span>
              </div>
            </div>

            {/* Card 2: Dept 6201 Graphis */}
            <div className="bg-rose-50/70 border border-rose-200 rounded-2xl p-4 shadow-2xs space-y-1">
              <div className="text-[11px] font-bold text-rose-800 uppercase flex items-center justify-between">
                <span>Dept 6201 Graphis</span>
                <span className="text-[10px] font-black px-1.5 py-0.5 rounded bg-rose-200 text-rose-900">
                  {currentMonthlySummary.graphis.sharePct.toFixed(1)}% Share
                </span>
              </div>
              <div className="text-lg sm:text-2xl font-black text-rose-950 font-mono-num">
                Rp {formatNumberId(currentMonthlySummary.graphis.sales)}
              </div>
              <div className="text-[10px] text-rose-700 font-semibold flex items-center justify-between pt-0.5">
                <span>{currentMonthlySummary.graphis.qty} pcs · {currentMonthlySummary.graphis.trx} nota</span>
                <span>ASP: Rp {formatNumberId(currentMonthlySummary.graphis.pricePoint)}</span>
              </div>
            </div>

            {/* Card 3: Dept 6244 Simplicity */}
            <div className="bg-indigo-50/70 border border-indigo-200 rounded-2xl p-4 shadow-2xs space-y-1">
              <div className="text-[11px] font-bold text-indigo-800 uppercase flex items-center justify-between">
                <span>Dept 6244 Simplicity</span>
                <span className="text-[10px] font-black px-1.5 py-0.5 rounded bg-indigo-200 text-indigo-900">
                  {currentMonthlySummary.simplicity.sharePct.toFixed(1)}% Share
                </span>
              </div>
              <div className="text-lg sm:text-2xl font-black text-indigo-950 font-mono-num">
                Rp {formatNumberId(currentMonthlySummary.simplicity.sales)}
              </div>
              <div className="text-[10px] text-indigo-700 font-semibold flex items-center justify-between pt-0.5">
                <span>{currentMonthlySummary.simplicity.qty} pcs · {currentMonthlySummary.simplicity.trx} nota</span>
                <span>ASP: Rp {formatNumberId(currentMonthlySummary.simplicity.pricePoint)}</span>
              </div>
            </div>

            {/* Card 4: Total YTD 1 Tahun 2026 */}
            <div className="bg-gradient-to-br from-amber-50 to-orange-50 border border-amber-300 rounded-2xl p-4 shadow-2xs space-y-1">
              <div className="text-[11px] font-bold text-amber-900 uppercase flex items-center justify-between">
                <span>Total YTD 1 Tahun (2026)</span>
                <span className="text-[10px] font-black px-1.5 py-0.5 rounded bg-amber-200 text-amber-950">
                  12 Bulan
                </span>
              </div>
              <div className="text-lg sm:text-2xl font-black text-amber-950 font-mono-num">
                Rp {formatNumberId(currentAnnualSummary.total.sales)}
              </div>
              <div className="text-[10px] text-amber-800 font-semibold flex items-center justify-between pt-0.5">
                <span>{currentAnnualSummary.total.qty} pcs terjual</span>
                <span>Rata-rata: Rp {formatNumberId(Math.round(currentAnnualSummary.total.sales / 12))}/bln</span>
              </div>
            </div>
          </div>

          {/* ============================================================ */}
          {/* TABEL 1: TABLE FORMAT PER WEEK DI BULAN TERPILIH            */}
          {/* ============================================================ */}
          {(monthlyViewFormat === 'both' || monthlyViewFormat === 'per_week') && (
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
              {/* Header Dark Navy Banner with Brand Filter */}
              <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#161b2e] to-indigo-950 text-white flex flex-col md:flex-row md:items-center justify-between gap-2.5 border-b border-slate-800">
                <div>
                  <div className="flex items-center gap-2">
                    <div className="w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse shrink-0" />
                    <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                      TABEL SALES MONTHLY — FORMAT PER WEEK (CV DIVISI 1)
                    </h3>
                  </div>
                  <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                    Rekapitulasi Week 1 s/d Week {monthlyWeeksData.weeks.length} di Bulan {currentMonthlySummary.monthName.toUpperCase()} 2026
                  </p>
                </div>

                {/* Filter Brand & Total */}
                <div className="flex items-center gap-1.5 flex-wrap">
                  <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider flex items-center gap-1 shrink-0">
                    <Filter className="w-3 h-3 text-amber-400" />
                    <span>FILTER:</span>
                  </span>
                  <div className="inline-flex p-0.5 bg-slate-900/95 rounded-xl border border-slate-700/80 flex-wrap gap-1 shadow-inner">
                    <button
                      type="button"
                      onClick={() => setMonthlyBrandFilter('TOTAL')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                        monthlyBrandFilter === 'TOTAL'
                          ? 'bg-amber-400 text-slate-950 shadow-md ring-2 ring-amber-300'
                          : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                      }`}
                    >
                      <span className={`w-1.5 h-1.5 rounded-full ${monthlyBrandFilter === 'TOTAL' ? 'bg-slate-950' : 'bg-amber-400'}`} />
                      <span>TOTAL GABUNGAN CV</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setMonthlyBrandFilter('GRAPHIS')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                        monthlyBrandFilter === 'GRAPHIS'
                          ? 'bg-[#e6004c] text-white shadow-md ring-2 ring-rose-400'
                          : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                      }`}
                    >
                      <span className={`w-1.5 h-1.5 rounded-full ${monthlyBrandFilter === 'GRAPHIS' ? 'bg-white' : 'bg-pink-400'}`} />
                      <span>GRAPHIS (6201)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setMonthlyBrandFilter('SIMPLICITY')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                        monthlyBrandFilter === 'SIMPLICITY'
                          ? 'bg-indigo-600 text-white shadow-md ring-2 ring-indigo-400'
                          : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                      }`}
                    >
                      <span className={`w-1.5 h-1.5 rounded-full ${monthlyBrandFilter === 'SIMPLICITY' ? 'bg-white' : 'bg-sky-400'}`} />
                      <span>SIMPLICITY (6244)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setMonthlyBrandFilter('SPLIT')}
                      className={`px-3 py-1.5 rounded-lg text-[10px] sm:text-xs font-black transition-all cursor-pointer flex items-center gap-1.5 ${
                        monthlyBrandFilter === 'SPLIT'
                          ? 'bg-gradient-to-r from-purple-800 to-indigo-800 text-white shadow-md ring-2 ring-purple-400'
                          : 'text-slate-300 hover:text-white hover:bg-slate-800/80'
                      }`}
                    >
                      <span className={`w-1.5 h-1.5 rounded-full ${monthlyBrandFilter === 'SPLIT' ? 'bg-white' : 'bg-purple-300'}`} />
                      <span>SPLIT (3 BARIS)</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Table Content */}
              <div className="overflow-x-auto max-h-[600px] scrollbar-thin">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                    <tr className="border-b border-slate-800 text-[11px] font-black uppercase tracking-wider text-slate-200">
                      <th className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 min-w-[85px]">
                        WEEK
                      </th>
                      <th className="py-3 px-3 border-r border-slate-800 min-w-[130px]">
                        PERIODE
                      </th>
                      <th className="py-3 px-3 border-r border-slate-800 min-w-[130px]">
                        BRAND / DEPT
                      </th>
                      {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'GRAPHIS' || monthlyBrandFilter === 'SPLIT') && (
                        <>
                          <th className="py-3 px-3 text-right bg-rose-950/40 text-rose-300 border-r border-slate-800 min-w-[125px]">
                            SALES GRAPHIS (6201)
                          </th>
                          <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-14">
                            QTY
                          </th>
                          <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-14">
                            NOTA
                          </th>
                          <th className="py-3 px-2.5 text-right bg-rose-950/40 text-rose-300 border-r border-slate-800 min-w-[100px]">
                            PRICE POINT
                          </th>
                          <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-16">
                            SHARE %
                          </th>
                        </>
                      )}
                      {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'SIMPLICITY' || monthlyBrandFilter === 'SPLIT') && (
                        <>
                          <th className="py-3 px-3 text-right bg-indigo-950/40 text-indigo-300 border-r border-slate-800 min-w-[125px]">
                            SALES SIMPLICITY (6244)
                          </th>
                          <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-14">
                            QTY
                          </th>
                          <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-14">
                            NOTA
                          </th>
                          <th className="py-3 px-2.5 text-right bg-indigo-950/40 text-indigo-300 border-r border-slate-800 min-w-[100px]">
                            PRICE POINT
                          </th>
                          <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-16">
                            SHARE %
                          </th>
                        </>
                      )}
                      <th className="py-3 px-3 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[135px]">
                        TOTAL SALES WEEK
                      </th>
                      <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                        QTY
                      </th>
                      <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                        NOTA
                      </th>
                      <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[105px]">
                        RATA-RATA/HR
                      </th>
                      <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-16">
                        KONTRIBUSI
                      </th>
                      <th className="py-3 px-2.5 text-center min-w-[90px]">
                        AKSI
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white font-medium">
                    {monthlyWeeksData.weeks.map((w, idx) => {
                      const isSplit = monthlyBrandFilter === 'SPLIT';
                      return (
                        <React.Fragment key={w.weekNumber}>
                          <tr
                            className={`hover:bg-purple-50/30 transition-colors ${
                              idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                            }`}
                          >
                            {/* Week Badge */}
                            <td className="py-3 px-3 text-center border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                              <span className="inline-flex flex-col items-center justify-center px-2 py-1 rounded-lg text-xs font-black bg-purple-900 text-white shadow-2xs">
                                <span>WEEK {w.weekIndexInMonth}</span>
                                <span className="text-[9px] font-bold text-purple-200">W-{w.weekNumber}</span>
                              </span>
                            </td>

                            {/* Periode */}
                            <td className="py-3 px-3 border-r border-slate-100">
                              <div className="font-extrabold text-slate-900 text-xs">{w.periodeText}</div>
                              <div className="text-[10px] text-slate-500 font-semibold">{w.daysCountInMonth} hari di {currentMonthlySummary.monthName}</div>
                            </td>

                            {/* Brand / Dept Tag */}
                            <td className="py-3 px-3 border-r border-slate-100">
                              {monthlyBrandFilter === 'TOTAL' && (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-black bg-amber-100 text-amber-950 border border-amber-300">
                                  TOTAL CV DIVISI 1
                                </span>
                              )}
                              {monthlyBrandFilter === 'GRAPHIS' && (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-black bg-rose-100 text-rose-900 border border-rose-300">
                                  DEPT 6201 GRAPHIS
                                </span>
                              )}
                              {monthlyBrandFilter === 'SIMPLICITY' && (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-black bg-indigo-100 text-indigo-900 border border-indigo-300">
                                  DEPT 6244 SIMPLICITY
                                </span>
                              )}
                              {monthlyBrandFilter === 'SPLIT' && (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-black bg-purple-100 text-purple-900 border border-purple-300">
                                  GABUNGAN (3 BARIS)
                                </span>
                              )}
                            </td>

                            {/* Graphis Columns */}
                            {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'GRAPHIS' || monthlyBrandFilter === 'SPLIT') && (
                              <>
                                <td className="py-3 px-3 text-right font-mono-num font-extrabold text-rose-900 border-r border-slate-100">
                                  {w.graphis.sales > 0 ? `Rp ${formatNumberId(w.graphis.sales)}` : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                                  {w.graphis.qty > 0 ? w.graphis.qty : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                                  {w.graphis.trx > 0 ? w.graphis.trx : '—'}
                                </td>
                                <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                                  {w.graphis.pricePoint > 0 ? `Rp ${formatNumberId(w.graphis.pricePoint)}` : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num font-bold text-rose-700 border-r border-slate-100">
                                  {w.graphis.sharePct > 0 ? `${w.graphis.sharePct.toFixed(1)}%` : '—'}
                                </td>
                              </>
                            )}

                            {/* Simplicity Columns */}
                            {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'SIMPLICITY' || monthlyBrandFilter === 'SPLIT') && (
                              <>
                                <td className="py-3 px-3 text-right font-mono-num font-extrabold text-indigo-900 border-r border-slate-100">
                                  {w.simplicity.sales > 0 ? `Rp ${formatNumberId(w.simplicity.sales)}` : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                                  {w.simplicity.qty > 0 ? w.simplicity.qty : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                                  {w.simplicity.trx > 0 ? w.simplicity.trx : '—'}
                                </td>
                                <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                                  {w.simplicity.pricePoint > 0 ? `Rp ${formatNumberId(w.simplicity.pricePoint)}` : '—'}
                                </td>
                                <td className="py-3 px-2 text-center font-mono-num font-bold text-indigo-700 border-r border-slate-100">
                                  {w.simplicity.sharePct > 0 ? `${w.simplicity.sharePct.toFixed(1)}%` : '—'}
                                </td>
                              </>
                            )}

                            {/* Total CV Columns */}
                            <td className="py-3 px-3 text-right font-mono-num font-black text-slate-950 bg-amber-50/40 border-r border-slate-100">
                              {w.total.sales > 0 ? `Rp ${formatNumberId(w.total.sales)}` : '—'}
                            </td>
                            <td className="py-3 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/40 border-r border-slate-100">
                              {w.total.qty > 0 ? w.total.qty : '—'}
                            </td>
                            <td className="py-3 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/40 border-r border-slate-100">
                              {w.total.trx > 0 ? w.total.trx : '—'}
                            </td>
                            <td className="py-3 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/40 border-r border-slate-100">
                              {w.avgDaily > 0 ? `Rp ${formatNumberId(w.avgDaily)}` : '—'}
                            </td>
                            <td className="py-3 px-2 text-center font-mono-num font-black text-amber-950 bg-amber-50/40 border-r border-slate-100">
                              {w.contribution > 0 ? `${w.contribution.toFixed(1)}%` : '—'}
                            </td>

                            {/* Action Button */}
                            <td className="py-3 px-2.5 text-center">
                              <button
                                type="button"
                                onClick={() => {
                                  setSelectedWeekNumber(w.weekNumber);
                                  setSelectedWeeklyRowNumber(w.weekNumber);
                                  setHorizon('weekly');
                                }}
                                className="px-2.5 py-1 rounded-md text-[10px] font-bold bg-purple-100 hover:bg-purple-200 text-purple-950 cursor-pointer transition-colors shadow-2xs"
                                title={`Buka Horizon Sales Weekly untuk W-${w.weekNumber}`}
                              >
                                Buka Weekly
                              </button>
                            </td>
                          </tr>
                        </React.Fragment>
                      );
                    })}
                  </tbody>

                  {/* Sticky Footer: Total Bulan Ini */}
                  <tfoot className="sticky bottom-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs shadow-lg border-t-2 border-amber-400">
                    <tr>
                      <td colSpan={3} className="py-3 px-3 text-center uppercase tracking-wider text-amber-300 border-r border-slate-700">
                        TOTAL BULAN {currentMonthlySummary.monthName.toUpperCase()} 2026
                      </td>
                      {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'GRAPHIS' || monthlyBrandFilter === 'SPLIT') && (
                        <>
                          <td className="py-3 px-3 text-right font-mono-num text-rose-300 border-r border-slate-700">
                            Rp {formatNumberId(monthlyWeeksData.summary.graphis.sales)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.graphis.qty}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.graphis.trx}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-rose-300 border-r border-slate-700">
                            Rp {formatNumberId(monthlyWeeksData.summary.graphis.pricePoint)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.graphis.sharePct.toFixed(1)}%
                          </td>
                        </>
                      )}
                      {(monthlyBrandFilter === 'TOTAL' || monthlyBrandFilter === 'SIMPLICITY' || monthlyBrandFilter === 'SPLIT') && (
                        <>
                          <td className="py-3 px-3 text-right font-mono-num text-indigo-300 border-r border-slate-700">
                            Rp {formatNumberId(monthlyWeeksData.summary.simplicity.sales)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.simplicity.qty}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.simplicity.trx}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-indigo-300 border-r border-slate-700">
                            Rp {formatNumberId(monthlyWeeksData.summary.simplicity.pricePoint)}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                            {monthlyWeeksData.summary.simplicity.sharePct.toFixed(1)}%
                          </td>
                        </>
                      )}
                      <td className="py-3 px-3 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        Rp {formatNumberId(monthlyWeeksData.summary.total.sales)}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        {monthlyWeeksData.summary.total.qty}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        {monthlyWeeksData.summary.total.trx}
                      </td>
                      <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        Rp {formatNumberId(currentMonthlySummary.avgDailySales)}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        100%
                      </td>
                      <td className="py-3 px-2.5 text-center text-[10px] text-amber-200/80 italic">
                        {currentMonthlySummary.monthName}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}

          {/* ============================================================ */}
          {/* TABEL 2: FORMAT PERBULAN DALAM 1 TAHUN (12 BULAN LENGKAP)    */}
          {/* ============================================================ */}
          {(monthlyViewFormat === 'both' || monthlyViewFormat === 'per_bulan') && (
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
              {/* Header Dark Navy Banner */}
              <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#182138] to-purple-950 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800">
                <div>
                  <div className="flex items-center gap-2">
                    <div className="w-2.5 h-2.5 rounded-full bg-amber-400 animate-pulse shrink-0" />
                    <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                      TABEL SALES PERBULAN DALAM 1 TAHUN (JANUARI S/D DESEMBER 2026) — CV DIVISI 1
                    </h3>
                  </div>
                  <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                    Monitoring 12 Bulan Lengkap · Target Setahun Rp 540.000.000 · Tanjung Outmall 311
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <span className="px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider bg-amber-400 text-slate-950 shadow-xs">
                    12 BULAN (YTD 2026)
                  </span>
                </div>
              </div>

              {/* Table Content */}
              <div className="overflow-x-auto max-h-[680px] scrollbar-thin">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                    <tr className="border-b border-slate-800 text-[11px] font-black uppercase tracking-wider text-slate-200">
                      <th className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 min-w-[95px]">
                        NO & BULAN
                      </th>
                      <th className="py-3 px-2 text-center border-r border-slate-800 w-16">
                        HARI
                      </th>
                      {/* Dept 6201 Graphis */}
                      <th className="py-3 px-3 text-right bg-rose-950/40 text-rose-300 border-r border-slate-800 min-w-[130px]">
                        SALES GRAPHIS (RP)
                      </th>
                      <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-14">
                        QTY
                      </th>
                      <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-14">
                        NOTA
                      </th>
                      <th className="py-3 px-2.5 text-right bg-rose-950/40 text-rose-300 border-r border-slate-800 min-w-[100px]">
                        PRICE POINT
                      </th>
                      <th className="py-3 px-2 text-center bg-rose-950/40 text-rose-300 border-r border-slate-800 w-16">
                        SHARE %
                      </th>

                      {/* Dept 6244 Simplicity */}
                      <th className="py-3 px-3 text-right bg-indigo-950/40 text-indigo-300 border-r border-slate-800 min-w-[130px]">
                        SALES SIMPLICITY (RP)
                      </th>
                      <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-14">
                        QTY
                      </th>
                      <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-14">
                        NOTA
                      </th>
                      <th className="py-3 px-2.5 text-right bg-indigo-950/40 text-indigo-300 border-r border-slate-800 min-w-[100px]">
                        PRICE POINT
                      </th>
                      <th className="py-3 px-2 text-center bg-indigo-950/40 text-indigo-300 border-r border-slate-800 w-16">
                        SHARE %
                      </th>

                      {/* Total CV Divisi 1 */}
                      <th className="py-3 px-3 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[140px]">
                        TOTAL SALES CV
                      </th>
                      <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                        QTY
                      </th>
                      <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                        NOTA
                      </th>
                      <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[100px]">
                        PRICE POINT
                      </th>
                      <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[105px]">
                        RATA-RATA/HR
                      </th>
                      <th className="py-3 px-2.5 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[110px]">
                        GROWTH & SHARE
                      </th>
                      <th className="py-3 px-3 border-r border-slate-800 min-w-[150px]">
                        BEST SALES DAY
                      </th>
                      <th className="py-3 px-2.5 text-center min-w-[85px]">
                        AKSI
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 bg-white font-medium">
                    {monthlyAnnualTableData.map((m) => {
                      const isSelected = m.isSelected;
                      const hasSales = m.total.sales > 0;
                      return (
                        <tr
                          key={m.monthKey}
                          className={`hover:bg-purple-50/40 transition-colors ${
                            isSelected
                              ? 'bg-purple-50/70 ring-1 ring-inset ring-purple-300 font-bold'
                              : m.index % 2 === 0
                              ? 'bg-white'
                              : 'bg-slate-50/40'
                          }`}
                        >
                          {/* No & Bulan */}
                          <td className="py-3 px-3 text-center border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                            <div className="flex items-center gap-1.5 justify-center">
                              <span
                                className={`inline-flex items-center justify-center px-2 py-1 rounded-lg text-xs font-black ${
                                  isSelected
                                    ? 'bg-purple-700 text-white shadow-xs ring-1 ring-purple-500'
                                    : 'bg-slate-900 text-slate-100'
                                }`}
                              >
                                {String(m.index).padStart(2, '0')} - {m.monthName.toUpperCase()}
                              </span>
                            </div>
                            {isSelected && (
                              <div className="text-[9px] font-black text-purple-700 uppercase tracking-tighter mt-0.5">
                                • BULAN TERPILIH •
                              </div>
                            )}
                          </td>

                          {/* Hari */}
                          <td className="py-3 px-2 text-center font-bold text-slate-700 border-r border-slate-100">
                            {m.daysCount} hr
                          </td>

                          {/* Graphis Sales */}
                          <td className="py-3 px-3 text-right font-mono-num font-black text-rose-900 border-r border-slate-100">
                            {hasSales ? `Rp ${formatNumberId(m.graphis.sales)}` : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                            {m.graphis.qty > 0 ? m.graphis.qty : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                            {m.graphis.trx > 0 ? m.graphis.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                            {m.graphis.pricePoint > 0 ? `Rp ${formatNumberId(m.graphis.pricePoint)}` : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-rose-700 border-r border-slate-100">
                            {hasSales ? `${m.graphis.sharePct.toFixed(1)}%` : '—'}
                          </td>

                          {/* Simplicity Sales */}
                          <td className="py-3 px-3 text-right font-mono-num font-black text-indigo-900 border-r border-slate-100">
                            {hasSales ? `Rp ${formatNumberId(m.simplicity.sales)}` : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                            {m.simplicity.qty > 0 ? m.simplicity.qty : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                            {m.simplicity.trx > 0 ? m.simplicity.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                            {m.simplicity.pricePoint > 0 ? `Rp ${formatNumberId(m.simplicity.pricePoint)}` : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-bold text-indigo-700 border-r border-slate-100">
                            {hasSales ? `${m.simplicity.sharePct.toFixed(1)}%` : '—'}
                          </td>

                          {/* Total CV Divisi 1 */}
                          <td className="py-3 px-3 text-right font-mono-num font-black text-slate-950 bg-amber-50/40 border-r border-slate-100">
                            {hasSales ? `Rp ${formatNumberId(m.total.sales)}` : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/40 border-r border-slate-100">
                            {m.total.qty > 0 ? m.total.qty : '—'}
                          </td>
                          <td className="py-3 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/40 border-r border-slate-100">
                            {m.total.trx > 0 ? m.total.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/40 border-r border-slate-100">
                            {m.total.pricePoint > 0 ? `Rp ${formatNumberId(m.total.pricePoint)}` : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/40 border-r border-slate-100">
                            {m.avgDailySales > 0 ? `Rp ${formatNumberId(m.avgDailySales)}` : '—'}
                          </td>

                          {/* Growth & Kontribusi */}
                          <td className="py-3 px-2.5 text-center border-r border-slate-100">
                            <div className="flex flex-col items-center gap-0.5">
                              {m.momGrowth !== null ? (
                                <span
                                  className={`inline-flex items-center gap-0.5 px-1.5 py-0.2 rounded text-[10px] font-black ${
                                    m.momGrowth >= 0
                                      ? 'bg-emerald-100 text-emerald-800'
                                      : 'bg-rose-100 text-rose-800'
                                  }`}
                                >
                                  {m.momGrowth >= 0 ? '↑' : '↓'} {Math.abs(m.momGrowth).toFixed(1)}% MoM
                                </span>
                              ) : (
                                <span className="text-[10px] text-slate-400 font-semibold">— Base</span>
                              )}
                              <span className="text-[9px] font-bold text-slate-500">
                                {m.contribution > 0 ? `${m.contribution.toFixed(1)}% thd 1 Thn` : '0%'}
                              </span>
                            </div>
                          </td>

                          {/* Best Day */}
                          <td className="py-3 px-3 border-r border-slate-100">
                            {m.bestDay ? (
                              <div>
                                <div className="text-[11px] font-bold text-slate-900">
                                  {m.bestDay.dayName}, {m.bestDay.tanggal}
                                </div>
                                <div className="text-[10px] font-mono-num font-extrabold text-emerald-700">
                                  Rp {formatNumberId(m.bestDay.sales)}
                                </div>
                              </div>
                            ) : (
                              <span className="text-slate-300 font-normal">—</span>
                            )}
                          </td>

                          {/* Aksi Button */}
                          <td className="py-3 px-2.5 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedMonthKey(m.monthKey);
                                showToast(`Bulan ${m.monthName} 2026 terpilih!`);
                              }}
                              className={`px-2.5 py-1 rounded-md text-[10px] font-bold transition-all cursor-pointer shadow-2xs ${
                                isSelected
                                  ? 'bg-purple-700 text-white'
                                  : 'bg-slate-100 hover:bg-slate-200 text-slate-800'
                              }`}
                            >
                              {isSelected ? '✓ Terpilih' : 'Pilih'}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>

                  {/* Sticky Footer: Grand Total 1 Tahun 2026 */}
                  <tfoot className="sticky bottom-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs shadow-lg border-t-2 border-amber-400">
                    <tr>
                      <td colSpan={2} className="py-3.5 px-3 text-center uppercase tracking-wider text-amber-300 border-r border-slate-700">
                        GRAND TOTAL 1 TAHUN (YTD 2026) — CV DIVISI 1
                      </td>

                      {/* Graphis Total Sum */}
                      <td className="py-3.5 px-3 text-right font-mono-num text-rose-300 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.graphis.sales)}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                        {currentAnnualSummary.graphis.qty}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                        {currentAnnualSummary.graphis.trx}
                      </td>
                      <td className="py-3.5 px-2.5 text-right font-mono-num text-rose-300 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.graphis.pricePoint)}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-rose-300 border-r border-slate-700">
                        {currentAnnualSummary.graphis.sharePct.toFixed(1)}%
                      </td>

                      {/* Simplicity Total Sum */}
                      <td className="py-3.5 px-3 text-right font-mono-num text-indigo-300 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.simplicity.sales)}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                        {currentAnnualSummary.simplicity.qty}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                        {currentAnnualSummary.simplicity.trx}
                      </td>
                      <td className="py-3.5 px-2.5 text-right font-mono-num text-indigo-300 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.simplicity.pricePoint)}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-indigo-300 border-r border-slate-700">
                        {currentAnnualSummary.simplicity.sharePct.toFixed(1)}%
                      </td>

                      {/* Total CV Divisi 1 Grand Total */}
                      <td className="py-3.5 px-3 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.total.sales)}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        {currentAnnualSummary.total.qty}
                      </td>
                      <td className="py-3.5 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        {currentAnnualSummary.total.trx}
                      </td>
                      <td className="py-3.5 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        Rp {formatNumberId(currentAnnualSummary.total.pricePoint)}
                      </td>
                      <td className="py-3.5 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        Rp {formatNumberId(Math.round(currentAnnualSummary.total.sales / 365))}
                      </td>
                      <td className="py-3.5 px-2.5 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                        100% YTD
                      </td>
                      <td className="py-3.5 px-3 text-[10px] text-amber-200 border-r border-slate-700 font-bold">
                        Rata-rata: Rp {formatNumberId(Math.round(currentAnnualSummary.total.sales / 12))}/bln
                      </td>
                      <td className="py-3.5 px-2.5 text-center text-[10px] text-amber-300 font-bold">
                        12 Bulan
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}

          {/* ============================================================ */}
          {/* TABEL 3: RINCIAN HARIAN MTD (TANGGAL 1 S/D AKHIR BULAN)       */}
          {/* ============================================================ */}
          {(monthlyViewFormat === 'both' || monthlyViewFormat === 'harian') && (
            <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-xs">
              <div className="px-4 py-3 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-2 border-b border-slate-800">
                <div className="flex items-center gap-2">
                  <div className="w-2.5 h-2.5 rounded-full bg-sky-400 animate-pulse" />
                  <span className="font-bold text-xs sm:text-sm">
                    Rincian Harian Tanggal 1 s/d Akhir Bulan — {currentMonthlySummary.monthName.toUpperCase()} 2026
                  </span>
                </div>
                <span className="text-[11px] font-black text-amber-300 bg-amber-950/60 px-2.5 py-0.5 rounded border border-amber-600/40">
                  Total MTD: Rp {formatNumberId(currentMonthlySummary.total.sales)}
                </span>
              </div>

              <div className="overflow-x-auto max-h-[500px] scrollbar-thin">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 z-10 bg-slate-100 shadow-2xs border-b border-slate-300">
                    <tr className="border-b border-slate-300 text-slate-700 font-extrabold text-[11px]">
                      <th className="py-2.5 px-3 border-r border-slate-200">Tanggal</th>
                      <th className="py-2.5 px-3 text-center border-r border-slate-200">Weekly</th>
                      <th className="py-2.5 px-3 text-right text-rose-900 bg-rose-50/30 border-r border-slate-200">Dept 6201 Graphis</th>
                      <th className="py-2.5 px-3 text-right text-indigo-900 bg-indigo-50/30 border-r border-slate-200">Dept 6244 Simplicity</th>
                      <th className="py-2.5 px-3 text-right font-black text-slate-950 bg-amber-50/40 border-r border-slate-200">Total CV Div 1</th>
                      <th className="py-2.5 px-3 text-center border-r border-slate-200">Qty</th>
                      <th className="py-2.5 px-3 text-center border-r border-slate-200">Trx</th>
                      <th className="py-2.5 px-3 text-right border-r border-slate-200">Price Point</th>
                      <th className="py-2.5 px-3 text-center">Aksi</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 font-medium bg-white">
                    {currentMonthlySummary.days.map((ds, idx) => {
                      const isWknd = ds.dayName === 'Sabtu' || ds.dayName === 'Minggu';
                      return (
                        <tr
                          key={ds.tanggal}
                          className={`hover:bg-purple-50/40 ${
                            isWknd ? 'bg-amber-50/30' : idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'
                          }`}
                        >
                          <td className="py-2 px-3 font-bold text-slate-800 border-r border-slate-100">
                            {ds.dayName}, {ds.tanggal}
                            {isWknd && (
                              <span className="ml-1.5 px-1 py-0.2 rounded text-[9px] bg-amber-200 text-amber-900 font-bold">
                                WKND
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-3 text-center font-semibold text-slate-600 border-r border-slate-100">
                            W{ds.weekNumber}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-rose-900 font-bold border-r border-slate-100">
                            Rp {formatNumberId(ds.graphis.sales)}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-indigo-900 font-bold border-r border-slate-100">
                            Rp {formatNumberId(ds.simplicity.sales)}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num font-black text-slate-950 bg-amber-50/30 border-r border-slate-100">
                            Rp {formatNumberId(ds.total.sales)}
                          </td>
                          <td className="py-2 px-3 text-center font-mono-num border-r border-slate-100">
                            {ds.total.qty} pcs
                          </td>
                          <td className="py-2 px-3 text-center font-mono-num text-slate-600 border-r border-slate-100">
                            {ds.total.trx}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-slate-700 border-r border-slate-100">
                            Rp {formatNumberId(ds.total.pricePoint)}
                          </td>
                          <td className="py-2 px-3 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedTanggal(ds.tanggal);
                                setHorizon('harian');
                              }}
                              className="px-2.5 py-1 rounded text-[10px] font-bold bg-purple-100 text-purple-900 hover:bg-purple-200 cursor-pointer shadow-2xs"
                            >
                              Buka Harian
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ============================================================ */}
          {/* CHART DIAGRAM SALES CV BULANAN & 1 TAHUN (SEPERTI REPORT DP) */}
          {/* ============================================================ */}
          <div className="space-y-4">
            <PhotoStyleSalesDiagram
              headerTitle={`DIAGRAM SALES CV BULANAN — ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026 (PER WEEKLY)`}
              headerSubtitle="Diagram Batang Vertikal Mingguan & Garis Tren Penjualan Dept 6201 (Graphis) vs Dept 6244 (Simplicity)"
              badgeText="DIAGRAM SALES CV (OMZET RUPIAH)"
              cardTitle="Grafik Batang Weekly Graphis vs Simplicity & Garis Tren"
              primaryLabel="Graphis (6201)"
              secondaryLabel="Simplicity (6244)"
              totalSales={monthlyWeeksData.summary.total.sales}
              totalPagi={monthlyWeeksData.summary.graphis.sales}
              totalSiang={monthlyWeeksData.summary.simplicity.sales}
              bars={monthlyWeeksData.weeks.map((w) => ({
                key: `cv-m-wk-${w.weekNumber}`,
                label: `W-${w.weekNumber}`,
                subLabel: w.periodeText.split('-')[0].trim(),
                pagiSales: w.graphis.sales,
                siangSales: w.simplicity.sales,
                totalSales: w.total.sales,
              }))}
            />

            <PhotoStyleSalesDiagram
              headerTitle="DIAGRAM SALES CV TAHUNAN — 12 BULAN (JANUARI - DESEMBER 2026)"
              headerSubtitle="Diagram Batang Vertikal 12 Bulan & Garis Tren Penjualan Dept 6201 (Graphis) vs Dept 6244 (Simplicity)"
              badgeText="DIAGRAM SALES CV 1 TAHUN"
              cardTitle="Grafik Batang 12 Bulan Graphis vs Simplicity & Garis Tren"
              primaryLabel="Graphis (6201)"
              secondaryLabel="Simplicity (6244)"
              totalSales={currentAnnualSummary.total.sales}
              totalPagi={currentAnnualSummary.graphis.sales}
              totalSiang={currentAnnualSummary.simplicity.sales}
              bars={monthlyAnnualTableData.map((m) => ({
                key: m.monthKey,
                label: m.monthName.slice(0, 3),
                subLabel: `${String(m.index).padStart(2, '0')}/26`,
                pagiSales: m.graphis.sales,
                siangSales: m.simplicity.sales,
                totalSales: m.total.sales,
                isSelected: m.isSelected,
                onClick: () => setSelectedMonthKey(m.monthKey),
              }))}
            />

            <GrowthAndContributionBarDiagram
              headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI MINGGUAN CV — ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026`}
              headerSubtitle="Grafik Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Tiap Weekly terhadap Bulan"
              periodBadgeText={`BULAN ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026`}
              items={monthlyWeeksData.weeks.map((w, idx) => ({
                key: `cv-m-wk-gc-${w.weekNumber}`,
                label: `W-${w.weekNumber}`,
                subLabel: w.periodeText.split('-')[0].trim(),
                pagiSales: w.graphis.sales,
                siangSales: w.simplicity.sales,
                totalSales: w.total.sales,
                prevSales: idx > 0 ? monthlyWeeksData.weeks[idx - 1].total.sales : null,
              }))}
            />

            <GrowthAndContributionBarDiagram
              headerTitle="DIAGRAM BATANG GROWTH & KONTRIBUSI TAHUNAN CV — 12 BULAN (JANUARI - DESEMBER 2026)"
              headerSubtitle="Grafik Batang Pertumbuhan Bulanan (MoM Growth %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahunan CV (100%)"
              periodBadgeText="12 BULAN CV (JAN – DES 2026)"
              items={monthlyAnnualTableData.map((m, idx) => ({
                key: `cv-ann-gc-${m.monthKey}`,
                label: m.monthName.slice(0, 3),
                subLabel: `${String(m.index).padStart(2, '0')}/26`,
                pagiSales: m.graphis.sales,
                siangSales: m.simplicity.sales,
                totalSales: m.total.sales,
                prevSales: idx > 0 ? monthlyAnnualTableData[idx - 1].total.sales : null,
                isSelected: m.isSelected,
                onClick: () => setSelectedMonthKey(m.monthKey),
              }))}
            />
          </div>
        </div>
      )}



      {/* ============================================================ */}
      {/* MODAL: INPUT SALES REPORT CV (PORTAL & PERFECT CENTER) */}
      {/* ============================================================ */}
      {isEditModalOpen && typeof document !== 'undefined' && createPortal(
        <div
          className="fixed inset-0 z-[9999] overflow-y-auto overscroll-y-contain bg-slate-950/80 backdrop-blur-xs p-3 sm:p-4 animate-in fade-in duration-150"
          onClick={(e) => {
            if (e.target === e.currentTarget) setIsEditModalOpen(false);
          }}
        >
          <div className="flex min-h-full items-center justify-center py-4">
            <div
              className="relative w-full max-w-sm sm:max-w-md bg-white border border-slate-200 rounded-3xl shadow-2xl overflow-hidden animate-in zoom-in-95 duration-150 ring-1 ring-black/5 my-auto"
              onClick={(e) => e.stopPropagation()}
            >
              {/* 1. Header: Sesuai Foto (Dark Navy + Amber Subtitle) */}
              <div className="px-4 py-3 bg-gradient-to-b from-[#0b0f19] to-[#121829] text-white flex items-center justify-between border-b border-slate-800">
                <div className="flex items-center gap-3">
                  <div className="w-9 h-9 rounded-xl bg-amber-400/10 border border-amber-400/30 flex items-center justify-center text-amber-400 shadow-inner shrink-0">
                    <PenSquare className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="text-xs sm:text-sm font-black text-white uppercase tracking-wider">
                      INPUT SALES REPORT CV
                    </h3>
                    <p className="text-[11px] font-bold text-amber-400 leading-tight">
                      Divisi 1 · Tanjung Outmall · Matahari 311
                    </p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsEditModalOpen(false)}
                  className="p-1 rounded-lg text-slate-400 hover:text-white hover:bg-white/10 transition-colors cursor-pointer"
                  title="Tutup"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Form Pengisian Harian: Bisa di-geser bebas tanpa terpotong */}
              <form onSubmit={handleSaveEdit} className="p-3.5 space-y-3">
                {/* 2. Card 1: PILIH TANGGAL PENJUALAN & FILTER TAMPILAN BRAND */}
                <div className="border border-slate-200 rounded-2xl p-3 bg-white space-y-2.5 shadow-2xs">
                  {/* Tanggal Dropdown & Filter Kalender Tanggal */}
                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="block text-[11px] font-black uppercase text-slate-700 tracking-wider">
                        PILIH TANGGAL PENJUALAN
                      </label>
                      <span className="text-[10px] font-bold text-purple-700 bg-purple-50 px-2 py-0.5 rounded border border-purple-200">
                        {formatIndoDateLong(editTanggal)}
                      </span>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-12 gap-2">
                      <div className="sm:col-span-7">
                        <select
                          value={editTanggal}
                          onChange={(e) => handleDateChangeInModal(e.target.value)}
                          className="w-full h-10 px-3 text-xs sm:text-sm font-bold bg-white border border-slate-300 rounded-xl text-slate-900 focus:outline-none focus:ring-2 focus:ring-purple-500 cursor-pointer shadow-2xs"
                        >
                          {records.map((r) => (
                            <option key={r.tanggal} value={r.tanggal}>
                              {r.tanggal} — {r.dayName.toUpperCase()}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="sm:col-span-5 flex items-center gap-1.5 bg-purple-50/70 border border-purple-300 rounded-xl px-2 h-10">
                        <Calendar className="w-4 h-4 text-purple-700 shrink-0" />
                        <input
                          aria-label="Filter Kalender Tanggal Modal"
                          type="date"
                          value={slashDateToIso(editTanggal)}
                          onChange={(e) => {
                            if (e.target.value) {
                              handleDateChangeInModal(isoToSlashDate(e.target.value));
                            }
                          }}
                          className="w-full text-xs font-bold bg-transparent text-purple-950 focus:outline-none cursor-pointer"
                          title="Klik ikon kalender untuk memilih tanggal secara visual"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Filter Tampilan Brand Pills */}
                  <div>
                    <label className="block text-[11px] font-black uppercase text-slate-700 mb-1 tracking-wider">
                      FILTER TAMPILAN BRAND
                    </label>
                    <div className="grid grid-cols-3 gap-1 p-1 bg-slate-100 rounded-xl border border-slate-200/80">
                      <button
                        type="button"
                        onClick={() => setEditForm({ ...editForm, selectedBrand: 'ALL' })}
                        className={`h-8 rounded-lg text-xs font-black cursor-pointer transition-all ${
                          editForm.selectedBrand === 'ALL'
                            ? 'bg-purple-700 text-white shadow-xs'
                            : 'text-slate-700 hover:text-slate-900 hover:bg-white/60'
                        }`}
                      >
                        Semua Brand
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditForm({ ...editForm, selectedBrand: 'GRAPHIS' })}
                        className={`h-8 rounded-lg text-xs font-black cursor-pointer transition-all ${
                          editForm.selectedBrand === 'GRAPHIS'
                            ? 'bg-[#e6004c] text-white shadow-xs'
                            : 'text-slate-700 hover:text-slate-900 hover:bg-white/60'
                        }`}
                      >
                        Graphis (6201)
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditForm({ ...editForm, selectedBrand: 'SIMPLICITY' })}
                        className={`h-8 rounded-lg text-xs font-black cursor-pointer transition-all ${
                          editForm.selectedBrand === 'SIMPLICITY'
                            ? 'bg-indigo-600 text-white shadow-xs'
                            : 'text-slate-700 hover:text-slate-900 hover:bg-white/60'
                        }`}
                      >
                        Simplicity (6244)
                      </button>
                    </div>
                  </div>
                </div>

                {/* 3. Card 2: BRAND INPUT (GRAPHIS / SIMPLICITY) Sesuai Foto */}
                <div className="space-y-2.5">
                  {/* Brand 1: GRAPHIS (Dept 6201) */}
                  {(editForm.selectedBrand === 'GRAPHIS' || editForm.selectedBrand === 'ALL') && (
                    <div className="border-2 border-rose-200 bg-rose-50/20 rounded-2xl p-3 shadow-2xs space-y-2">
                      <div className="flex items-center justify-between pb-1 border-b border-rose-200/60">
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full bg-[#e6004c]" />
                          <span className="text-xs font-black uppercase text-slate-900 tracking-wider">
                            DEPT 6201 · BRAND GRAPHIS
                          </span>
                        </div>
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-100 text-rose-800">
                          Divisi 1
                        </span>
                      </div>

                      <div className="grid grid-cols-12 gap-2">
                        {/* SALES (RP) */}
                        <div className="col-span-6">
                          <label className="block text-[11px] font-black uppercase text-rose-950 mb-1">
                            SALES (RP)
                          </label>
                          <div className="relative h-11 border-2 border-[#e6004c] rounded-xl bg-white px-2.5 flex items-center shadow-2xs">
                            <span className="text-xs font-bold text-slate-400 select-none mr-1.5">
                              Rp
                            </span>
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.graphisSales}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  graphisSales: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-right text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>

                        {/* QTY */}
                        <div className="col-span-3">
                          <label className="block text-[11px] font-black uppercase text-rose-950 mb-1 text-center">
                            QTY
                          </label>
                          <div className="h-11 border border-rose-200 rounded-xl bg-white px-1 flex items-center justify-center shadow-2xs">
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.graphisQty}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  graphisQty: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-center text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>

                        {/* NOTA */}
                        <div className="col-span-3">
                          <label className="block text-[11px] font-black uppercase text-rose-950 mb-1 text-center">
                            NOTA
                          </label>
                          <div className="h-11 border border-rose-200 rounded-xl bg-white px-1 flex items-center justify-center shadow-2xs">
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.graphisNota}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  graphisNota: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-center text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Brand 2: SIMPLICITY (Dept 6244) */}
                  {(editForm.selectedBrand === 'SIMPLICITY' || editForm.selectedBrand === 'ALL') && (
                    <div className="border-2 border-indigo-200 bg-indigo-50/20 rounded-2xl p-3 shadow-2xs space-y-2">
                      <div className="flex items-center justify-between pb-1 border-b border-indigo-200/60">
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full bg-indigo-600" />
                          <span className="text-xs font-black uppercase text-slate-900 tracking-wider">
                            DEPT 6244 · BRAND SIMPLICITY
                          </span>
                        </div>
                        <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-800">
                          Divisi 1
                        </span>
                      </div>

                      <div className="grid grid-cols-12 gap-2">
                        {/* SALES (RP) */}
                        <div className="col-span-6">
                          <label className="block text-[11px] font-black uppercase text-indigo-950 mb-1">
                            SALES (RP)
                          </label>
                          <div className="relative h-11 border-2 border-indigo-600 rounded-xl bg-white px-2.5 flex items-center shadow-2xs">
                            <span className="text-xs font-bold text-slate-400 select-none mr-1.5">
                              Rp
                            </span>
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.simplicitySales}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  simplicitySales: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-right text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>

                        {/* QTY */}
                        <div className="col-span-3">
                          <label className="block text-[11px] font-black uppercase text-indigo-950 mb-1 text-center">
                            QTY
                          </label>
                          <div className="h-11 border border-indigo-200 rounded-xl bg-white px-1 flex items-center justify-center shadow-2xs">
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.simplicityQty}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  simplicityQty: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-center text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>

                        {/* NOTA */}
                        <div className="col-span-3">
                          <label className="block text-[11px] font-black uppercase text-indigo-950 mb-1 text-center">
                            NOTA
                          </label>
                          <div className="h-11 border border-indigo-200 rounded-xl bg-white px-1 flex items-center justify-center shadow-2xs">
                            <input
                              type="number"
                              inputMode="numeric"
                              min="0"
                              placeholder="0"
                              value={editForm.simplicityNota}
                              onFocus={(e) => e.target.select()}
                              onChange={(e) =>
                                setEditForm({
                                  ...editForm,
                                  simplicityNota: e.target.value === '' ? '' : Number(e.target.value),
                                })
                              }
                              className="w-full text-center text-sm sm:text-base font-black text-slate-950 bg-transparent focus:outline-none font-mono-num"
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>

                {/* 4. Card 3: TOTAL DAILY GABUNGAN Sesuai Foto */}
                <div className="bg-[#0e1424] text-white rounded-2xl p-3 flex items-center justify-between border border-slate-800 shadow-md">
                  <div>
                    <div className="text-[10px] font-black text-amber-400 uppercase tracking-wider leading-tight">
                      TOTAL DAILY<br />GABUNGAN
                    </div>
                    <div className="text-lg sm:text-xl font-black text-white font-mono-num mt-0.5">
                      Rp {formatNumberId((Number(editForm.graphisSales) || 0) + (Number(editForm.simplicitySales) || 0))}
                    </div>
                  </div>
                  <div className="flex items-center gap-1.5 text-center">
                    <div className="bg-[#1a233a] border border-slate-700/60 rounded-xl px-3 py-1.5 min-w-[54px] shadow-2xs">
                      <div className="text-xs font-black text-amber-400 font-mono-num">
                        {(Number(editForm.graphisQty) || 0) + (Number(editForm.simplicityQty) || 0)}
                      </div>
                      <div className="text-[10px] text-slate-300 font-medium">pcs</div>
                    </div>
                    <div className="bg-[#1a233a] border border-slate-700/60 rounded-xl px-3 py-1.5 min-w-[54px] shadow-2xs">
                      <div className="text-xs font-black text-amber-400 font-mono-num">
                        {(Number(editForm.graphisNota) || 0) + (Number(editForm.simplicityNota) || 0)}
                      </div>
                      <div className="text-[10px] text-slate-300 font-medium">nota</div>
                    </div>
                  </div>
                </div>

                {/* 5. Card 4: Tombol Batal & Simpan Data Sales Sesuai Foto */}
                <div className="flex items-center justify-between pt-1 pb-1">
                  <button
                    type="button"
                    onClick={() => setIsEditModalOpen(false)}
                    className="px-4 py-2.5 text-xs font-bold text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-xl transition-colors cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="submit"
                    className="h-11 px-6 rounded-xl font-black text-xs text-white bg-[#7900ff] hover:bg-[#6800dd] shadow-lg flex items-center gap-2 cursor-pointer transition-transform active:scale-98"
                  >
                    <Save className="w-4 h-4 text-amber-300" />
                    <span>Simpan Data Sales</span>
                  </button>
                </div>
              </form>
            </div>
          </div>
        </div>,
        document.body
      )}
    </div>
  );
};
