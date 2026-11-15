/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * 
 * TOTAL REPORT (REPORT DP + REPORT CV) — KONSOLIDASI BAZAR MATAHARI TANJUNG
 * Menggabungkan seluruh data penjualan:
 * - REPORT DP (Direct Purchase)
 * - REPORT CV (Konsinyasi: Dept 6201 Graphis + Dept 6244 Simplicity)
 * Menyajikan 3 Horizon: Total Daily, Total Weekly, dan Total Monthly & 1 Tahun (2026)
 */

import React, { useState, useMemo } from 'react';
import { exportTotalDpCvToExcel } from '../utils/exportTotalDpCvExcel';
import {
  Calendar,
  CalendarDays,
  CalendarRange,
  FileSpreadsheet,
  Download,
  Share2,
  TrendingUp,
  BarChart3,
  Layers,
  ArrowRight,
  ArrowUpRight,
  Filter,
  CheckCircle2,
  PieChart,
  Users,
  Percent,
} from 'lucide-react';
import { WeeklySalesTable, ANNUAL_2026_WEEKS, formatNumberId, getIndoDayName } from '../data/initialData';
import { CvDailyRecord, MONTH_NAMES_ID } from '../data/cvDiv1Data';
import { slashDateToIso, isoToSlashDate, formatIndoDateLong } from '../utils/dateUtils';
import { PhotoStyleSalesDiagram } from './PhotoStyleSalesDiagram';
import { copyTextToClipboard } from '../utils/clipboard';
import { GrowthAndContributionBarDiagram } from './GrowthAndContributionBarDiagram';

export type TotalHorizon = 'harian' | 'weekly' | 'bulanan';

export interface TotalDpCvReportViewProps {
  weeklyTables: WeeklySalesTable[];
  cvRecords: CvDailyRecord[];
  activeHorizon?: TotalHorizon;
  onChangeHorizon?: (horizon: TotalHorizon) => void;
  onNavigateToDp?: (horizon: 'harian' | 'weekly' | 'bulanan', weekId?: string) => void;
  onNavigateToCv?: (horizon: 'harian' | 'weekly' | 'bulanan') => void;
  onUpdateTraffic?: (tanggalSlash: string, shift: 'pagi' | 'siang', trafficVal: number | null) => void;
  isAdmin?: boolean;
}

export const TotalDpCvReportView: React.FC<TotalDpCvReportViewProps> = ({
  weeklyTables,
  cvRecords,
  activeHorizon = 'harian',
  onChangeHorizon,
  onNavigateToDp,
  onNavigateToCv,
  onUpdateTraffic,
  isAdmin = false,
}) => {
  const [internalHorizon, setInternalHorizon] = useState<TotalHorizon>('harian');
  const horizon: TotalHorizon = activeHorizon !== undefined ? activeHorizon : internalHorizon;
  const setHorizon = (nextH: TotalHorizon) => {
    setInternalHorizon(nextH);
    if (onChangeHorizon) onChangeHorizon(nextH);
  };

  const [selectedMonthKey, setSelectedMonthKey] = useState<string>('11');
  const [dailyCalendarFilterIso, setDailyCalendarFilterIso] = useState<string>('');
  const [weeklyCalendarFilterIso, setWeeklyCalendarFilterIso] = useState<string>('');
  const [weeklyScope, setWeeklyScope] = useState<'month' | 'all'>('month');
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  // 1. Build fast Lookup Maps
  // Map date "DD/MM/YYYY" -> DP metrics + Kasir Pagi & Siang Traffic
  const dpDateMap = useMemo(() => {
    const map = new Map<
      string,
      {
        sales: number;
        qty: number;
        trx: number;
        pagiSales: number;
        siangSales: number;
        pagiQty: number;
        siangQty: number;
        pagiTrx: number;
        siangTrx: number;
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
      if (Array.isArray(table.days)) {
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

          const nextEntry = {
            sales: pSales + sSales,
            qty: pQty + sQty,
            trx: pTrx + sTrx,
            pagiSales: pSales,
            siangSales: sSales,
            pagiQty: pQty,
            siangQty: sQty,
            pagiTrx: pTrx,
            siangTrx: sTrx,
            pagiTraffic: pTraffic,
            siangTraffic: sTraffic,
            trafficTotal: pTraffic + sTraffic,
            pagiStaff: (day.pagi?.nama || '').trim(),
            siangStaff: (day.siang?.nama || '').trim(),
            reasons,
            weekNumber: table.weekNumber,
          };

          const prev = map.get(tgl);
          if (
            !prev ||
            nextEntry.sales > 0 ||
            nextEntry.trafficTotal > 0 ||
            (prev.sales === 0 && prev.trafficTotal === 0)
          ) {
            map.set(tgl, nextEntry);
          }
        });
      }
    });

    return map;
  }, [weeklyTables]);

  // Map date "DD/MM/YYYY" -> CV metrics
  const cvDateMap = useMemo(() => {
    const map = new Map<
      string,
      {
        sales: number;
        qty: number;
        trx: number;
        graphisSales: number;
        graphisQty: number;
        graphisTrx: number;
        simplicitySales: number;
        simplicityQty: number;
        simplicityTrx: number;
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

      map.set(tgl, {
        sales: gSales + sSales,
        qty: gQty + sQty,
        trx: gTrx + sTrx,
        graphisSales: gSales,
        graphisQty: gQty,
        graphisTrx: gTrx,
        simplicitySales: sSales,
        simplicityQty: sQty,
        simplicityTrx: sTrx,
        reasonEvent: r.reasonEvent || '',
        weekNumber: r.weekNumber,
      });
    });

    return map;
  }, [cvRecords]);

  // 2. Data Harian Konsolidasi (Daily Consolidated) untuk bulan terpilih (01 s/d 31)
  const dailyConsolidatedData = useMemo(() => {
    const mm = selectedMonthKey;
    const year = 2026;
    const daysInMonth = new Date(year, parseInt(mm, 10), 0).getDate();

    const days = [];
    let totDpSales = 0, totDpQty = 0, totDpTrx = 0;
    let totCvSales = 0, totCvQty = 0, totCvTrx = 0;
    let totGrandSales = 0, totGrandQty = 0, totGrandTrx = 0;
    let totTrafficPagi = 0, totTrafficSiang = 0, totTrafficTotal = 0;
    let bestDay: { tanggal: string; dayName: string; sales: number } | null = null;

    for (let day = 1; day <= daysInMonth; day++) {
      const dd = String(day).padStart(2, '0');
      const tanggal = `${dd}/${mm}/${year}`;
      const dayName = getIndoDayName(tanggal);

      const dp = dpDateMap.get(tanggal) || {
        sales: 0,
        qty: 0,
        trx: 0,
        pagiTraffic: 0,
        siangTraffic: 0,
        trafficTotal: 0,
        pagiStaff: '',
        siangStaff: '',
        reasons: [],
        weekNumber: 0,
      };
      const cv = cvDateMap.get(tanggal) || { sales: 0, qty: 0, trx: 0, reasonEvent: '', weekNumber: 0 };

      const grandSales = dp.sales + cv.sales;
      const grandQty = dp.qty + cv.qty;
      const grandTrx = dp.trx + cv.trx;
      const grandPricePoint = grandQty > 0 ? Math.round(grandSales / grandQty) : 0;
      const dpShare = grandSales > 0 ? (dp.sales / grandSales) * 100 : 0;
      const cvShare = grandSales > 0 ? (cv.sales / grandSales) * 100 : 0;

      const dayTrafficPagi = dp.pagiTraffic || 0;
      const dayTrafficSiang = dp.siangTraffic || 0;
      const dayTrafficTotal = dayTrafficPagi + dayTrafficSiang;
      const dayConvRatePct = dayTrafficTotal > 0 ? (grandTrx / dayTrafficTotal) * 100 : 0;
      const daySalesPerVisitor = dayTrafficTotal > 0 ? Math.round(grandSales / dayTrafficTotal) : 0;

      totDpSales += dp.sales;
      totDpQty += dp.qty;
      totDpTrx += dp.trx;
      totCvSales += cv.sales;
      totCvQty += cv.qty;
      totCvTrx += cv.trx;
      totGrandSales += grandSales;
      totGrandQty += grandQty;
      totGrandTrx += grandTrx;
      totTrafficPagi += dayTrafficPagi;
      totTrafficSiang += dayTrafficSiang;
      totTrafficTotal += dayTrafficTotal;

      if (!bestDay || grandSales > bestDay.sales) {
        bestDay = { tanggal, dayName, sales: grandSales };
      }

      // Combine reasons
      const allReasons = [...dp.reasons];
      if (cv.reasonEvent) allReasons.push(`CV: ${cv.reasonEvent}`);

      days.push({
        dayNumber: day,
        tanggal,
        dayName,
        weekNumber: dp.weekNumber || cv.weekNumber || 0,
        dp: {
          sales: dp.sales,
          qty: dp.qty,
          trx: dp.trx,
          sharePct: dpShare,
        },
        cv: {
          sales: cv.sales,
          qty: cv.qty,
          trx: cv.trx,
          sharePct: cvShare,
        },
        total: {
          sales: grandSales,
          qty: grandQty,
          trx: grandTrx,
          pricePoint: grandPricePoint,
        },
        traffic: {
          pagi: dayTrafficPagi,
          siang: dayTrafficSiang,
          total: dayTrafficTotal,
          pagiStaff: dp.pagiStaff,
          siangStaff: dp.siangStaff,
          conversionRatePct: dayConvRatePct,
          salesPerVisitor: daySalesPerVisitor,
        },
        reasonsText: allReasons.join(' | '),
      });
    }

    const overallDpShare = totGrandSales > 0 ? (totDpSales / totGrandSales) * 100 : 0;
    const overallCvShare = totGrandSales > 0 ? (totCvSales / totGrandSales) * 100 : 0;
    const avgDailySales = daysInMonth > 0 ? Math.round(totGrandSales / daysInMonth) : 0;
    const overallPricePoint = totGrandQty > 0 ? Math.round(totGrandSales / totGrandQty) : 0;
    const overallConvRatePct = totTrafficTotal > 0 ? (totGrandTrx / totTrafficTotal) * 100 : 0;
    const overallSalesPerVisitor = totTrafficTotal > 0 ? Math.round(totGrandSales / totTrafficTotal) : 0;

    return {
      days,
      daysCount: daysInMonth,
      summary: {
        dp: {
          sales: totDpSales,
          qty: totDpQty,
          trx: totDpTrx,
          sharePct: overallDpShare,
          pricePoint: totDpQty > 0 ? Math.round(totDpSales / totDpQty) : 0,
        },
        cv: {
          sales: totCvSales,
          qty: totCvQty,
          trx: totCvTrx,
          sharePct: overallCvShare,
          pricePoint: totCvQty > 0 ? Math.round(totCvSales / totCvQty) : 0,
        },
        total: {
          sales: totGrandSales,
          qty: totGrandQty,
          trx: totGrandTrx,
          pricePoint: overallPricePoint,
          avgDailySales,
        },
        traffic: {
          pagi: totTrafficPagi,
          siang: totTrafficSiang,
          total: totTrafficTotal,
          conversionRatePct: overallConvRatePct,
          salesPerVisitor: overallSalesPerVisitor,
        },
        bestDay,
      },
    };
  }, [selectedMonthKey, dpDateMap, cvDateMap]);

  // 3. Data Mingguan Konsolidasi (Weekly Consolidated)
  const weeklyConsolidatedData = useMemo(() => {
    let weekList = ANNUAL_2026_WEEKS;
    if (weeklyCalendarFilterIso) {
      const slash = isoToSlashDate(weeklyCalendarFilterIso);
      const matched = ANNUAL_2026_WEEKS.filter(
        (w) =>
          w.slashDates.includes(slash) ||
          (weeklyCalendarFilterIso >= w.startDateIso &&
            weeklyCalendarFilterIso <= w.endDateIso)
      );
      if (matched.length > 0) {
        weekList = matched;
      }
    } else if (weeklyScope === 'month') {
      weekList = ANNUAL_2026_WEEKS.filter((w) => {
        return (
          w.monthKey === selectedMonthKey ||
          w.slashDates.some((d) => d.split('/')[1] === selectedMonthKey)
        );
      });
    }

    let totDpSales = 0, totDpQty = 0, totDpTrx = 0;
    let totCvSales = 0, totCvQty = 0, totCvTrx = 0;
    let totGrandSales = 0, totGrandQty = 0, totGrandTrx = 0;
    let totTrafficPagi = 0, totTrafficSiang = 0, totTrafficTotal = 0;

    const weeks = weekList.map((w, index) => {
      let wDpSales = 0, wDpQty = 0, wDpTrx = 0;
      let wCvSales = 0, wCvQty = 0, wCvTrx = 0;
      let wTrafficPagi = 0, wTrafficSiang = 0;

      w.slashDates.forEach((dt) => {
        const dp = dpDateMap.get(dt);
        if (dp) {
          wDpSales += dp.sales;
          wDpQty += dp.qty;
          wDpTrx += dp.trx;
          wTrafficPagi += dp.pagiTraffic || 0;
          wTrafficSiang += dp.siangTraffic || 0;
        }
        const cv = cvDateMap.get(dt);
        if (cv) {
          wCvSales += cv.sales;
          wCvQty += cv.qty;
          wCvTrx += cv.trx;
        }
      });

      const wGrandSales = wDpSales + wCvSales;
      const wGrandQty = wDpQty + wCvQty;
      const wGrandTrx = wDpTrx + wCvTrx;
      const wTrafficTotal = wTrafficPagi + wTrafficSiang;
      const wConvRatePct = wTrafficTotal > 0 ? (wGrandTrx / wTrafficTotal) * 100 : 0;
      const wSalesPerVisitor = wTrafficTotal > 0 ? Math.round(wGrandSales / wTrafficTotal) : 0;
      const dpShare = wGrandSales > 0 ? (wDpSales / wGrandSales) * 100 : 0;
      const cvShare = wGrandSales > 0 ? (wCvSales / wGrandSales) * 100 : 0;
      const wPricePoint = wGrandQty > 0 ? Math.round(wGrandSales / wGrandQty) : 0;
      const avgDaily = Math.round(wGrandSales / 7);

      totDpSales += wDpSales;
      totDpQty += wDpQty;
      totDpTrx += wDpTrx;
      totCvSales += wCvSales;
      totCvQty += wCvQty;
      totCvTrx += wCvTrx;
      totGrandSales += wGrandSales;
      totGrandQty += wGrandQty;
      totGrandTrx += wGrandTrx;
      totTrafficPagi += wTrafficPagi;
      totTrafficSiang += wTrafficSiang;
      totTrafficTotal += wTrafficTotal;

      return {
        weekNumber: w.weekNumber,
        weekIndex: index + 1,
        periodeText: w.periodeText,
        monthName: w.monthName,
        dp: {
          sales: wDpSales,
          qty: wDpQty,
          trx: wDpTrx,
          pricePoint: wDpQty > 0 ? Math.round(wDpSales / wDpQty) : 0,
          sharePct: dpShare,
        },
        cv: {
          sales: wCvSales,
          qty: wCvQty,
          trx: wCvTrx,
          pricePoint: wCvQty > 0 ? Math.round(wCvSales / wCvQty) : 0,
          sharePct: cvShare,
        },
        total: {
          sales: wGrandSales,
          qty: wGrandQty,
          trx: wGrandTrx,
          pricePoint: wPricePoint,
          avgDaily,
        },
        traffic: {
          pagi: wTrafficPagi,
          siang: wTrafficSiang,
          total: wTrafficTotal,
          conversionRatePct: wConvRatePct,
          salesPerVisitor: wSalesPerVisitor,
        },
      };
    });

    const overallDpShare = totGrandSales > 0 ? (totDpSales / totGrandSales) * 100 : 0;
    const overallCvShare = totGrandSales > 0 ? (totCvSales / totGrandSales) * 100 : 0;
    const overallPricePoint = totGrandQty > 0 ? Math.round(totGrandSales / totGrandQty) : 0;
    const overallConvRatePct = totTrafficTotal > 0 ? (totGrandTrx / totTrafficTotal) * 100 : 0;
    const overallSalesPerVisitor = totTrafficTotal > 0 ? Math.round(totGrandSales / totTrafficTotal) : 0;

    return {
      weeks,
      summary: {
        dp: {
          sales: totDpSales,
          qty: totDpQty,
          trx: totDpTrx,
          pricePoint: totDpQty > 0 ? Math.round(totDpSales / totDpQty) : 0,
          sharePct: overallDpShare,
        },
        cv: {
          sales: totCvSales,
          qty: totCvQty,
          trx: totCvTrx,
          pricePoint: totCvQty > 0 ? Math.round(totCvSales / totCvQty) : 0,
          sharePct: overallCvShare,
        },
        total: {
          sales: totGrandSales,
          qty: totGrandQty,
          trx: totGrandTrx,
          pricePoint: overallPricePoint,
        },
        traffic: {
          pagi: totTrafficPagi,
          siang: totTrafficSiang,
          total: totTrafficTotal,
          conversionRatePct: overallConvRatePct,
          salesPerVisitor: overallSalesPerVisitor,
        },
      },
    };
  }, [weeklyScope, weeklyCalendarFilterIso, selectedMonthKey, dpDateMap, cvDateMap]);

  // 4. Data Bulanan & 1 Tahun Konsolidasi (Monthly 12 Months Consolidated)
  const monthlyAnnualConsolidatedData = useMemo(() => {
    const months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
    const year = 2026;

    let totDpSales = 0, totDpQty = 0, totDpTrx = 0;
    let totCvSales = 0, totCvQty = 0, totCvTrx = 0;
    let totGrandSales = 0, totGrandQty = 0, totGrandTrx = 0;
    let totTrafficPagi = 0, totTrafficSiang = 0, totTrafficTotal = 0;

    const list = months.map((mKey, idx) => {
      const monthName = MONTH_NAMES_ID[mKey] || `Bulan ${mKey}`;
      const daysCount = new Date(year, parseInt(mKey, 10), 0).getDate();

      let mDpSales = 0, mDpQty = 0, mDpTrx = 0;
      let mCvSales = 0, mCvQty = 0, mCvTrx = 0;
      let mTrafficPagi = 0, mTrafficSiang = 0;

      for (let day = 1; day <= daysCount; day++) {
        const dd = String(day).padStart(2, '0');
        const tanggal = `${dd}/${mKey}/${year}`;

        const dp = dpDateMap.get(tanggal);
        if (dp) {
          mDpSales += dp.sales;
          mDpQty += dp.qty;
          mDpTrx += dp.trx;
          mTrafficPagi += dp.pagiTraffic || 0;
          mTrafficSiang += dp.siangTraffic || 0;
        }

        const cv = cvDateMap.get(tanggal);
        if (cv) {
          mCvSales += cv.sales;
          mCvQty += cv.qty;
          mCvTrx += cv.trx;
        }
      }

      const mGrandSales = mDpSales + mCvSales;
      const mGrandQty = mDpQty + mCvQty;
      const mGrandTrx = mDpTrx + mCvTrx;
      const mTrafficTotal = mTrafficPagi + mTrafficSiang;
      const mConvRatePct = mTrafficTotal > 0 ? (mGrandTrx / mTrafficTotal) * 100 : 0;
      const mSalesPerVisitor = mTrafficTotal > 0 ? Math.round(mGrandSales / mTrafficTotal) : 0;
      const dpShare = mGrandSales > 0 ? (mDpSales / mGrandSales) * 100 : 0;
      const cvShare = mGrandSales > 0 ? (mCvSales / mGrandSales) * 100 : 0;
      const mPricePoint = mGrandQty > 0 ? Math.round(mGrandSales / mGrandQty) : 0;
      const avgDaily = daysCount > 0 ? Math.round(mGrandSales / daysCount) : 0;

      totDpSales += mDpSales;
      totDpQty += mDpQty;
      totDpTrx += mDpTrx;
      totCvSales += mCvSales;
      totCvQty += mCvQty;
      totCvTrx += mCvTrx;
      totGrandSales += mGrandSales;
      totGrandQty += mGrandQty;
      totGrandTrx += mGrandTrx;
      totTrafficPagi += mTrafficPagi;
      totTrafficSiang += mTrafficSiang;
      totTrafficTotal += mTrafficTotal;

      return {
        monthKey: mKey,
        monthName,
        monthIndex: idx + 1,
        daysCount,
        dp: {
          sales: mDpSales,
          qty: mDpQty,
          trx: mDpTrx,
          pricePoint: mDpQty > 0 ? Math.round(mDpSales / mDpQty) : 0,
          sharePct: dpShare,
        },
        cv: {
          sales: mCvSales,
          qty: mCvQty,
          trx: mCvTrx,
          pricePoint: mCvQty > 0 ? Math.round(mCvSales / mCvQty) : 0,
          sharePct: cvShare,
        },
        total: {
          sales: mGrandSales,
          qty: mGrandQty,
          trx: mGrandTrx,
          pricePoint: mPricePoint,
          avgDaily,
        },
        traffic: {
          pagi: mTrafficPagi,
          siang: mTrafficSiang,
          total: mTrafficTotal,
          conversionRatePct: mConvRatePct,
          salesPerVisitor: mSalesPerVisitor,
        },
        isSelected: mKey === selectedMonthKey,
      };
    });

    const monthsWithGrowth = list.map((m, idx) => {
      const prev = idx > 0 ? list[idx - 1] : null;
      let momGrowth: number | null = null;
      if (prev && prev.total.sales > 0) {
        momGrowth = ((m.total.sales - prev.total.sales) / prev.total.sales) * 100;
      }
      const contribution = totGrandSales > 0 ? (m.total.sales / totGrandSales) * 100 : 0;

      return {
        ...m,
        momGrowth,
        contribution,
      };
    });

    const bestMonth = monthsWithGrowth.reduce<{ monthKey: string; monthName: string; sales: number } | null>(
      (best, m) => {
        if (!best || m.total.sales > best.sales) {
          return { monthKey: m.monthKey, monthName: m.monthName, sales: m.total.sales };
        }
        return best;
      },
      null
    );

    const overallDpShare = totGrandSales > 0 ? (totDpSales / totGrandSales) * 100 : 0;
    const overallCvShare = totGrandSales > 0 ? (totCvSales / totGrandSales) * 100 : 0;
    const overallPricePoint = totGrandQty > 0 ? Math.round(totGrandSales / totGrandQty) : 0;
    const overallConvRatePct = totTrafficTotal > 0 ? (totGrandTrx / totTrafficTotal) * 100 : 0;
    const overallSalesPerVisitor = totTrafficTotal > 0 ? Math.round(totGrandSales / totTrafficTotal) : 0;

    return {
      months: monthsWithGrowth,
      summary: {
        dp: {
          sales: totDpSales,
          qty: totDpQty,
          trx: totDpTrx,
          pricePoint: totDpQty > 0 ? Math.round(totDpSales / totDpQty) : 0,
          sharePct: overallDpShare,
        },
        cv: {
          sales: totCvSales,
          qty: totCvQty,
          trx: totCvTrx,
          pricePoint: totCvQty > 0 ? Math.round(totCvSales / totCvQty) : 0,
          sharePct: overallCvShare,
        },
        total: {
          sales: totGrandSales,
          qty: totGrandQty,
          trx: totGrandTrx,
          pricePoint: overallPricePoint,
          avgMonthlySales: Math.round(totGrandSales / 12),
          avgDailySales: Math.round(totGrandSales / 365),
        },
        traffic: {
          pagi: totTrafficPagi,
          siang: totTrafficSiang,
          total: totTrafficTotal,
          conversionRatePct: overallConvRatePct,
          salesPerVisitor: overallSalesPerVisitor,
        },
        bestMonth,
      },
    };
  }, [dpDateMap, cvDateMap, selectedMonthKey]);

  // Salin WhatsApp Text Konsolidasi
  const handleCopyWhatsApp = () => {
    let text = '';
    const currentMonthName = (MONTH_NAMES_ID[selectedMonthKey] || 'November').toUpperCase();

    if (horizon === 'harian') {
      const s = dailyConsolidatedData.summary;
      text = [
        `*LAPORAN TOTAL KONSOLIDASI SALES DAILY (DP + CV)*`,
        `*BAZAR MATAHARI TANJUNG — PT MDS RETAILING TBK*`,
        `Periode: ${currentMonthName} 2026 (${dailyConsolidatedData.daysCount} Hari)`,
        `----------------------------------------`,
        `*1. REPORT DP (DIRECT PURCHASE)* : Rp ${formatNumberId(s.dp.sales)} (${s.dp.qty} pcs | ${s.dp.sharePct.toFixed(1)}%)`,
        `*2. REPORT CV (KONSINYASI)*      : Rp ${formatNumberId(s.cv.sales)} (${s.cv.qty} pcs | ${s.cv.sharePct.toFixed(1)}%)`,
        `----------------------------------------`,
        `★ *GRAND TOTAL BAZAR (DP + CV) : Rp ${formatNumberId(s.total.sales)}*`,
        `• Total Qty        : ${s.total.qty} pcs`,
        `• Total Trx+Nota   : ${s.total.trx} trx+nota`,
        `• Average ASP      : Rp ${formatNumberId(s.total.pricePoint)}`,
        `• Rata-rata/hr     : Rp ${formatNumberId(s.total.avgDailySales)}`,
        `• Traffic Pengunjung: ${formatNumberId(s.traffic.total) || '0'} org (Pagi: ${formatNumberId(s.traffic.pagi) || '0'} | Siang: ${formatNumberId(s.traffic.siang) || '0'})`,
        `★ *Total Conversion Rate : ${s.traffic.total > 0 ? s.traffic.conversionRatePct.toFixed(1) + '%' : '—'}*`,
        s.bestDay ? `• Puncak Sales : ${s.bestDay.dayName}, ${s.bestDay.tanggal} (Rp ${formatNumberId(s.bestDay.sales)})` : '',
      ]
        .filter(Boolean)
        .join('\n');
    } else if (horizon === 'weekly') {
      const s = weeklyConsolidatedData.summary;
      const weekRows = weeklyConsolidatedData.weeks.map(
        (w) =>
          `• W-${w.weekNumber} (${w.periodeText}): Rp ${formatNumberId(w.total.sales)} [DP: Rp ${formatNumberId(w.dp.sales)} | CV: Rp ${formatNumberId(w.cv.sales)}] · Traffic: ${formatNumberId(w.traffic.total) || '0'} org (${w.traffic.total > 0 ? w.traffic.conversionRatePct.toFixed(1) + '%' : '—'})`
      );

      text = [
        `*LAPORAN TOTAL KONSOLIDASI SALES WEEKLY (DP + CV)*`,
        `*BAZAR MATAHARI TANJUNG — PT MDS RETAILING TBK*`,
        `Cakupan: ${weeklyScope === 'month' ? `Bulan ${currentMonthName}` : 'Semua 52 Weekly (Setahun 2026)'}`,
        `----------------------------------------`,
        `*RINGKASAN KONTRIBUSI:*`,
        `• Total DP : Rp ${formatNumberId(s.dp.sales)} (${s.dp.sharePct.toFixed(1)}%)`,
        `• Total CV : Rp ${formatNumberId(s.cv.sales)} (${s.cv.sharePct.toFixed(1)}%)`,
        `★ *GRAND TOTAL WEEKLY : Rp ${formatNumberId(s.total.sales)}* (${s.total.qty} pcs | ${s.total.trx} trx+nota)`,
        `• Traffic (Pagi/Siang): ${formatNumberId(s.traffic.pagi) || '0'} / ${formatNumberId(s.traffic.siang) || '0'} (Total: ${formatNumberId(s.traffic.total) || '0'} org)`,
        `★ *Total Conversion Rate : ${s.traffic.total > 0 ? s.traffic.conversionRatePct.toFixed(1) + '%' : '—'}*`,
        `----------------------------------------`,
        `*RINCIAN PER WEEK:*`,
        ...weekRows,
      ]
        .filter(Boolean)
        .join('\n');
    } else {
      const s = monthlyAnnualConsolidatedData.summary;
      const mRows = monthlyAnnualConsolidatedData.months.map(
        (m) =>
          `• ${m.monthName}: Rp ${formatNumberId(m.total.sales)} (DP: Rp ${formatNumberId(m.dp.sales)} | CV: Rp ${formatNumberId(m.cv.sales)}) · Traffic: ${formatNumberId(m.traffic.total) || '0'} (${m.traffic.total > 0 ? m.traffic.conversionRatePct.toFixed(1) + '%' : '—'})`
      );

      text = [
        `*LAPORAN TOTAL KONSOLIDASI SALES 1 TAHUN (DP + CV 2026)*`,
        `*BAZAR MATAHARI TANJUNG — PT MDS RETAILING TBK*`,
        `Target Setahun: Rp 540.000.000`,
        `----------------------------------------`,
        `*1. TOTAL REPORT DP* : Rp ${formatNumberId(s.dp.sales)} (${s.dp.qty} pcs | ${s.dp.sharePct.toFixed(1)}%)`,
        `*2. TOTAL REPORT CV* : Rp ${formatNumberId(s.cv.sales)} (${s.cv.qty} pcs | ${s.cv.sharePct.toFixed(1)}%)`,
        `----------------------------------------`,
        `★ *GRAND TOTAL BAZAR 1 TAHUN : Rp ${formatNumberId(s.total.sales)}* (${s.total.qty} pcs | ${s.total.trx} trx+nota)`,
        `• Total Traffic 1 Thn   : ${formatNumberId(s.traffic.total) || '0'} org (Pagi: ${formatNumberId(s.traffic.pagi) || '0'} | Siang: ${formatNumberId(s.traffic.siang) || '0'})`,
        `★ *Total Conversion Rate : ${s.traffic.total > 0 ? s.traffic.conversionRatePct.toFixed(1) + '%' : '—'}*`,
        `• Rata-rata/Bulan       : Rp ${formatNumberId(s.total.avgMonthlySales)}`,
        `• Rata-rata/Hari        : Rp ${formatNumberId(s.total.avgDailySales)}`,
        s.bestMonth ? `• Bulan Tertinggi       : ${s.bestMonth.monthName} (Rp ${formatNumberId(s.bestMonth.sales)})` : '',
        `----------------------------------------`,
        `*RINCIAN 12 BULAN:*`,
        ...mRows,
      ]
        .filter(Boolean)
        .join('\n');
    }

    void copyTextToClipboard(text).then((ok) => {
      showToast(
        ok
          ? 'Teks laporan WhatsApp Total (DP + CV + Traffic) berhasil disalin!'
          : 'Gagal menyalin teks ke clipboard.'
      );
    });
  };

  // Ekspor Excel (.xlsx) Konsolidasi + Weekly + 12 Bulan + Traffic & Conversion Rate
  const handleExportExcel = async () => {
    try {
      await exportTotalDpCvToExcel(
        weeklyTables,
        cvRecords,
        selectedMonthKey,
        2026,
        weeklyScope
      );
      showToast('File Excel Total Konsolidasi (DP + CV + Traffic) berhasil diunduh!');
    } catch (err) {
      console.error('Error export excel konsolidasi:', err);
      showToast('Gagal mengunduh Excel konsolidasi.');
    }
  };

  return (
    <div className="space-y-6 pb-12">
      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed top-5 right-5 z-[100] px-4 py-2.5 bg-slate-900 text-white text-xs font-bold rounded-xl shadow-2xl flex items-center gap-2 border border-slate-700 animate-in fade-in slide-in-from-top-4">
          <CheckCircle2 className="w-4 h-4 text-emerald-400" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* 1. Top Header & Control Card */}
      <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs space-y-3">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div>
            <div className="flex items-center gap-2">
              <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-gradient-to-r from-amber-400 to-orange-400 text-slate-950 shadow-2xs">
                TOTAL KONSOLIDASI BAZAR MATAHARI
              </span>
              <span className="text-xs font-bold text-slate-500">
                PT MDS Retailing Tbk · Tanjung Outmall 311
              </span>
            </div>
            <h2 className="text-base sm:text-xl font-black text-slate-900 tracking-tight mt-1 flex items-center gap-2">
              <Layers className="w-5 h-5 text-amber-500" />
              <span>TOTAL REPORT (DP + CV) — GABUNGAN PENJUALAN</span>
            </h2>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleCopyWhatsApp}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xs cursor-pointer transition-colors"
              title="Salin laporan konsolidasi ke WhatsApp"
            >
              <Share2 className="w-3.5 h-3.5" />
              <span>Salin WA Total</span>
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white shadow-2xs cursor-pointer transition-colors"
              title="Ekspor laporan Excel gabungan DP + CV"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Ekspor Excel Total</span>
            </button>
          </div>
        </div>

        {/* Filter Ribbons: Horizon Selector, Bulan, & Focus */}
        <div className="pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex flex-wrap items-center gap-3">
            {/* Horizon Switcher */}
            <div className="inline-flex p-0.5 bg-slate-900/95 rounded-xl border border-slate-800 shadow-inner">
              <button
                type="button"
                onClick={() => setHorizon('harian')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer ${
                  horizon === 'harian'
                    ? 'bg-amber-400 text-slate-950 shadow-md ring-1 ring-amber-300'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <Calendar className="w-3.5 h-3.5" />
                <span>Total Daily</span>
              </button>

              <button
                type="button"
                onClick={() => setHorizon('weekly')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer ${
                  horizon === 'weekly'
                    ? 'bg-amber-400 text-slate-950 shadow-md ring-1 ring-amber-300'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>Total Weekly</span>
              </button>

              <button
                type="button"
                onClick={() => setHorizon('bulanan')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer ${
                  horizon === 'bulanan'
                    ? 'bg-amber-400 text-slate-950 shadow-md ring-1 ring-amber-300'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <CalendarDays className="w-3.5 h-3.5" />
                <span>Total Monthly (1 Tahun)</span>
              </button>
            </div>

            {/* Pilih Bulan (Bila pada Horizon Harian & Weekly) */}
            <div className="flex items-center gap-1.5">
              <span className="font-bold text-slate-600 flex items-center gap-1">
                <CalendarDays className="w-3.5 h-3.5 text-sky-600" />
                <span>Bulan:</span>
              </span>
              <select
                value={selectedMonthKey}
                onChange={(e) => setSelectedMonthKey(e.target.value)}
                className="px-2.5 py-1 text-xs font-bold bg-slate-50 border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:ring-2 focus:ring-amber-500 cursor-pointer"
              >
                {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                  <option key={mKey} value={mKey}>
                    {mName} 2026
                  </option>
                ))}
              </select>
            </div>

            {/* Filter Kalender Tanggal untuk Total Daily */}
            {horizon === 'harian' && (
              <div className="flex items-center gap-1.5 bg-amber-50 border border-amber-300 rounded-lg px-2.5 py-1 text-xs shadow-2xs">
                <Calendar className="w-3.5 h-3.5 text-amber-700 shrink-0" />
                <label htmlFor="total-cal-filter" className="font-bold text-amber-950 shrink-0">
                  Filter Kalender:
                </label>
                <input
                  id="total-cal-filter"
                  type="date"
                  value={dailyCalendarFilterIso}
                  onChange={(e) => {
                    const val = e.target.value;
                    setDailyCalendarFilterIso(val);
                    if (val) {
                      const parts = val.split('-');
                      if (parts.length === 3 && parts[1] !== selectedMonthKey) {
                        setSelectedMonthKey(parts[1]);
                      }
                    }
                  }}
                  className="px-1.5 py-0.5 text-xs font-bold bg-white border border-amber-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                  title="Filter tabel ke tanggal tertentu menggunakan kalender"
                />
                {dailyCalendarFilterIso && (
                  <button
                    type="button"
                    onClick={() => setDailyCalendarFilterIso('')}
                    className="text-[10px] font-bold text-amber-800 hover:text-amber-950 px-1.5 py-0.5 rounded bg-amber-200/80 cursor-pointer shrink-0"
                  >
                    Semua
                  </button>
                )}
              </div>
            )}

            {/* Scope & Kalender Tanggal untuk Weekly */}
            {horizon === 'weekly' && (
              <>
                <div className="flex items-center gap-1.5 bg-amber-50 border border-amber-300 rounded-lg px-2.5 py-1 text-xs shadow-2xs">
                  <Calendar className="w-3.5 h-3.5 text-amber-700 shrink-0" />
                  <label htmlFor="total-weekly-cal-filter" className="font-bold text-amber-950 shrink-0">
                    Kalender Weekly:
                  </label>
                  <input
                    id="total-weekly-cal-filter"
                    type="date"
                    value={weeklyCalendarFilterIso}
                    onChange={(e) => {
                      const val = e.target.value;
                      setWeeklyCalendarFilterIso(val);
                      if (val) {
                        const parts = val.split('-');
                        if (parts.length === 3 && parts[1] !== selectedMonthKey) {
                          setSelectedMonthKey(parts[1]);
                        }
                      }
                    }}
                    className="px-1.5 py-0.5 text-xs font-bold bg-white border border-amber-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                    title="Filter Weekly ke minggu yang memuat tanggal kalender pilihan"
                  />
                  {weeklyCalendarFilterIso && (
                    <button
                      type="button"
                      onClick={() => setWeeklyCalendarFilterIso('')}
                      className="text-[10px] font-bold text-amber-800 hover:text-amber-950 px-1.5 py-0.5 rounded bg-amber-200/80 cursor-pointer shrink-0"
                    >
                      Semua
                    </button>
                  )}
                </div>
                <div className="inline-flex p-0.5 bg-slate-100 rounded-lg border border-slate-200">
                  <button
                    type="button"
                    onClick={() => {
                      setWeeklyCalendarFilterIso('');
                      setWeeklyScope('month');
                    }}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      !weeklyCalendarFilterIso && weeklyScope === 'month'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Bulan Ini ({MONTH_NAMES_ID[selectedMonthKey]})
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setWeeklyCalendarFilterIso('');
                      setWeeklyScope('all');
                    }}
                    className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                      !weeklyCalendarFilterIso && weeklyScope === 'all'
                        ? 'bg-white text-slate-900 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Semua 52 Weekly (Setahun)
                  </button>
                </div>
              </>
            )}
          </div>

          {/* Quick Links to Source Reports */}
          <div className="flex items-center gap-2">
            {onNavigateToDp && (
              <button
                type="button"
                onClick={() => onNavigateToDp(horizon)}
                className="px-2.5 py-1 rounded-lg text-[10px] font-bold bg-blue-50 text-blue-800 border border-blue-200 hover:bg-blue-100 cursor-pointer inline-flex items-center gap-1 transition-colors"
                title="Buka Report DP (Direct Purchase)"
              >
                <span>Buka Report DP</span>
                <ArrowRight className="w-3 h-3" />
              </button>
            )}
            {onNavigateToCv && (
              <button
                type="button"
                onClick={() => onNavigateToCv(horizon)}
                className="px-2.5 py-1 rounded-lg text-[10px] font-bold bg-purple-50 text-purple-800 border border-purple-200 hover:bg-purple-100 cursor-pointer inline-flex items-center gap-1 transition-colors"
                title="Buka Report CV (Konsinyasi)"
              >
                <span>Buka Report CV</span>
                <ArrowRight className="w-3 h-3" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* 2. Executive KPI Cards Konsolidasi */}
      {(() => {
        const activeSummary =
          horizon === 'harian'
            ? dailyConsolidatedData.summary
            : horizon === 'weekly'
            ? weeklyConsolidatedData.summary
            : monthlyAnnualConsolidatedData.summary;

        return (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
            {/* Card 1: Grand Total Sales Gabungan */}
            <div className="bg-slate-900 text-white rounded-2xl p-4 border border-slate-800 space-y-1">
              <div className="text-[11px] font-bold text-amber-300 uppercase flex items-center justify-between">
                <span>GRAND TOTAL BAZAR (DP + CV)</span>
                <TrendingUp className="w-3.5 h-3.5 text-amber-300 shrink-0" />
              </div>
              <div className="text-lg sm:text-xl font-black text-white font-mono-num tabular-nums">
                Rp {formatNumberId(activeSummary.total.sales)}
              </div>
              <div className="text-[10px] text-slate-300 font-semibold flex items-center justify-between pt-0.5 tabular-nums">
                <span>{activeSummary.total.qty} pcs · {activeSummary.total.trx} trx+nota</span>
                <span className="text-amber-300 font-bold">ASP: Rp {formatNumberId(activeSummary.total.pricePoint)}</span>
              </div>
            </div>

            {/* Card 2: Report DP (Direct Purchase) */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 space-y-1">
              <div className="text-[11px] font-bold text-blue-900 uppercase flex items-center justify-between">
                <span>1. REPORT DP</span>
                <span className="text-[10px] font-black text-blue-700 tabular-nums">
                  {activeSummary.dp.sharePct.toFixed(1)}% Share
                </span>
              </div>
              <div className="text-lg sm:text-xl font-black text-slate-900 font-mono-num tabular-nums">
                Rp {formatNumberId(activeSummary.dp.sales)}
              </div>
              <div className="text-[10px] text-slate-600 font-semibold flex items-center justify-between pt-0.5 tabular-nums">
                <span>{activeSummary.dp.qty} pcs · {activeSummary.dp.trx} trx</span>
                <span>ASP: Rp {formatNumberId(activeSummary.dp.pricePoint)}</span>
              </div>
            </div>

            {/* Card 3: Report CV (Konsinyasi) */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 space-y-1">
              <div className="text-[11px] font-bold text-purple-900 uppercase flex items-center justify-between">
                <span>2. REPORT CV</span>
                <span className="text-[10px] font-black text-purple-700 tabular-nums">
                  {activeSummary.cv.sharePct.toFixed(1)}% Share
                </span>
              </div>
              <div className="text-lg sm:text-xl font-black text-slate-900 font-mono-num tabular-nums">
                Rp {formatNumberId(activeSummary.cv.sales)}
              </div>
              <div className="text-[10px] text-slate-600 font-semibold flex items-center justify-between pt-0.5 tabular-nums">
                <span>{activeSummary.cv.qty} pcs · {activeSummary.cv.trx} nota</span>
                <span>ASP: Rp {formatNumberId(activeSummary.cv.pricePoint)}</span>
              </div>
            </div>

            {/* Card 4: Komposisi Kontribusi */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 space-y-1.5">
              <div className="text-[11px] font-bold text-slate-600 uppercase flex items-center justify-between">
                <span>KONTRIBUSI DP vs CV</span>
                <PieChart className="w-3.5 h-3.5 text-amber-600" />
              </div>
              <div className="flex items-center gap-2 pt-1.5">
                <div className="flex-1 bg-slate-100 rounded-full h-2.5 overflow-hidden flex">
                  <div
                    className="bg-blue-600 h-full transition-all"
                    style={{ width: `${Math.max(activeSummary.dp.sharePct > 0 ? 5 : 0, activeSummary.dp.sharePct)}%` }}
                    title={`DP: ${activeSummary.dp.sharePct.toFixed(1)}%`}
                  />
                  <div
                    className="bg-purple-600 h-full transition-all"
                    style={{ width: `${Math.max(activeSummary.cv.sharePct > 0 ? 5 : 0, activeSummary.cv.sharePct)}%` }}
                    title={`CV: ${activeSummary.cv.sharePct.toFixed(1)}%`}
                  />
                </div>
              </div>
              <div className="text-[10px] text-slate-600 font-bold flex items-center justify-between pt-1 tabular-nums">
                <span className="text-blue-700">DP: {activeSummary.dp.sharePct.toFixed(1)}%</span>
                <span className="text-purple-700">CV: {activeSummary.cv.sharePct.toFixed(1)}%</span>
              </div>
            </div>

            {/* Card 5: Total Traffic & Conversion Rate (DP + CV) */}
            <div className="bg-white border border-indigo-200 rounded-2xl p-4 space-y-1">
              <div className="text-[11px] font-bold text-indigo-900 uppercase flex items-center justify-between">
                <span>TOTAL CONVERSION RATE</span>
                <Percent className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
              </div>
              <div className="text-lg sm:text-xl font-black text-indigo-900 font-mono-num tabular-nums flex items-baseline gap-2">
                <span>
                  {activeSummary.traffic.total > 0
                    ? `${activeSummary.traffic.conversionRatePct.toFixed(1)}%`
                    : '—'}
                </span>
                <span className="text-[11px] font-semibold text-slate-500">
                  ({formatNumberId(activeSummary.traffic.total) || '0'} org)
                </span>
              </div>
              <div className="text-[10px] text-slate-600 font-semibold flex items-center justify-between pt-0.5 tabular-nums">
                <span>Pagi: {formatNumberId(activeSummary.traffic.pagi) || '0'} · Siang: {formatNumberId(activeSummary.traffic.siang) || '0'}</span>
                <span className="text-indigo-700 font-bold">
                  {activeSummary.traffic.salesPerVisitor > 0
                    ? `Rp ${formatNumberId(activeSummary.traffic.salesPerVisitor)}/org`
                    : '0 org'}
                </span>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ============================================================ */}
      {/* HORIZON 1: TOTAL SALES DAILY (KONSOLIDASI HARIAN DP + CV)    */}
      {/* ============================================================ */}
      {horizon === 'harian' && (
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          {/* Header Dark Navy Banner */}
          <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#161b2e] to-indigo-950 text-white flex flex-col md:flex-row md:items-center justify-between gap-2 border-b border-slate-800">
            <div>
              <div className="flex items-center gap-2">
                <Calendar className="w-4 h-4 text-amber-400 shrink-0" />
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                  TABEL KONSOLIDASI SALES DAILY — TOTAL REPORT (DP + CV)
                </h3>
              </div>
              <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                Bulan {(MONTH_NAMES_ID[selectedMonthKey] || 'November').toUpperCase()} 2026 ({dailyConsolidatedData.daysCount} Hari) · Direct Purchase + Konsinyasi
              </p>
            </div>
            <span className="px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider bg-amber-400 text-slate-950 shadow-xs">
              TOTAL GABUNGAN BAZAR
            </span>
          </div>

          {/* Table Container */}
          <div className="overflow-x-auto max-h-[700px] scrollbar-thin">
            <table className="w-full text-left text-xs border-collapse font-sans">
              <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                {/* Baris 1: Header Grup */}
                <tr className="border-b border-slate-800 text-slate-200 text-[11px] font-black uppercase tracking-wider">
                  <th rowSpan={2} className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 w-12">
                    TGL
                  </th>
                  <th rowSpan={2} className="py-3 px-3 border-r border-slate-800 text-center min-w-[85px]">
                    HARI
                  </th>
                  <th rowSpan={2} className="py-3 px-2 text-center border-r border-slate-800 w-16 text-[10px]">
                    WEEK
                  </th>
                  {/* Grup 1: Report DP */}
                  <th colSpan={4} className="py-2.5 px-3 text-center bg-blue-950/60 text-blue-300 border-r border-slate-800">
                    1. REPORT DP (DIRECT PURCHASE)
                  </th>
                  {/* Grup 2: Report CV */}
                  <th colSpan={4} className="py-2.5 px-3 text-center bg-purple-950/60 text-purple-300 border-r border-slate-800">
                    2. REPORT CV (KONSINYASI DIVISI 1)
                  </th>
                  {/* Grup 3: Total Konsolidasi */}
                  <th colSpan={6} className="py-2.5 px-3 text-center bg-amber-950/60 text-amber-300 border-r border-slate-800">
                    ★ TOTAL GABUNGAN BAZAR (DP + CV) &amp; CONV. RATE
                  </th>
                  <th rowSpan={2} className="py-3 px-3 min-w-[150px]">
                    CATATAN / EVENT
                  </th>
                </tr>

                {/* Baris 2: Sub-kolom */}
                <tr className="border-b border-slate-700 text-[10px] font-black uppercase tracking-wider text-slate-300">
                  {/* Sub-kolom DP */}
                  <th className="py-2 px-3 text-right bg-blue-950/40 text-blue-200 border-r border-slate-800 min-w-[110px]">SALES (RP)</th>
                  <th className="py-2 px-2 text-center bg-blue-950/40 text-blue-200 border-r border-slate-800 w-12">QTY</th>
                  <th className="py-2 px-2 text-center bg-blue-950/40 text-blue-200 border-r border-slate-800 w-12">TRX</th>
                  <th className="py-2 px-2 text-center bg-blue-950/40 text-blue-200 border-r border-slate-800 w-14">SHARE %</th>

                  {/* Sub-kolom CV */}
                  <th className="py-2 px-3 text-right bg-purple-950/40 text-purple-200 border-r border-slate-800 min-w-[110px]">SALES (RP)</th>
                  <th className="py-2 px-2 text-center bg-purple-950/40 text-purple-200 border-r border-slate-800 w-12">QTY</th>
                  <th className="py-2 px-2 text-center bg-purple-950/40 text-purple-200 border-r border-slate-800 w-12">NOTA</th>
                  <th className="py-2 px-2 text-center bg-purple-950/40 text-purple-200 border-r border-slate-800 w-14">SHARE %</th>

                  {/* Sub-kolom Total */}
                  <th className="py-2 px-3 text-right bg-amber-950/40 text-amber-200 border-r border-slate-800 min-w-[125px]">TOTAL SALES</th>
                  <th className="py-2 px-2 text-center bg-amber-950/40 text-amber-200 border-r border-slate-800 w-12">QTY</th>
                  <th className="py-2 px-2 text-center bg-amber-950/40 text-amber-200 border-r border-slate-800 min-w-[74px]">TRX+NOTA</th>
                  <th className="py-2 px-2.5 text-right bg-amber-950/40 text-amber-200 border-r border-slate-800 min-w-[95px]">PRICE POINT</th>
                  <th className="py-2 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-16">TRAFFIC</th>
                  <th className="py-2 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-20">CONV. RATE</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white font-medium">
                {(dailyCalendarFilterIso
                  ? dailyConsolidatedData.days.filter((d) => d.tanggal === isoToSlashDate(dailyCalendarFilterIso))
                  : dailyConsolidatedData.days
                ).map((d, idx) => {
                  const isWknd = d.dayName === 'Sabtu' || d.dayName === 'Minggu';
                  const hasSales = d.total.sales > 0;
                  return (
                    <tr
                      key={d.tanggal}
                      className={`hover:bg-amber-50/30 transition-colors ${
                        isWknd ? 'bg-amber-50/20' : idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                      }`}
                    >
                      {/* Tanggal */}
                      <td className="py-2.5 px-3 text-center border-r border-slate-100 sticky left-0 z-10 bg-inherit font-black">
                        <span className="inline-flex items-center justify-center w-7 h-7 rounded-lg text-xs font-black bg-slate-900 text-white">
                          {d.dayNumber}
                        </span>
                      </td>

                      {/* Hari */}
                      <td className="py-2.5 px-3 text-center border-r border-slate-100">
                        <span className={`text-xs font-bold ${isWknd ? 'text-rose-600 font-black' : 'text-slate-800'}`}>
                          {d.dayName}
                        </span>
                      </td>

                      {/* Week */}
                      <td className="py-2.5 px-2 text-center border-r border-slate-100 text-[10px] font-bold text-slate-500">
                        W-{d.weekNumber}
                      </td>

                      {/* DP Columns */}
                      <td className="py-2.5 px-3 text-right font-mono-num font-bold text-blue-900 border-r border-slate-100">
                        {d.dp.sales > 0 ? `Rp ${formatNumberId(d.dp.sales)}` : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {d.dp.qty > 0 ? d.dp.qty : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {d.dp.trx > 0 ? d.dp.trx : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-bold text-blue-700 border-r border-slate-100">
                        {d.dp.sharePct > 0 ? `${d.dp.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* CV Columns */}
                      <td className="py-2.5 px-3 text-right font-mono-num font-bold text-purple-900 border-r border-slate-100">
                        {d.cv.sales > 0 ? `Rp ${formatNumberId(d.cv.sales)}` : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {d.cv.qty > 0 ? d.cv.qty : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {d.cv.trx > 0 ? d.cv.trx : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-bold text-purple-700 border-r border-slate-100">
                        {d.cv.sharePct > 0 ? `${d.cv.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* Grand Total Columns */}
                      <td className="py-2.5 px-3 text-right font-mono-num font-black text-slate-950 bg-amber-50/50 border-r border-slate-100">
                        {hasSales ? `Rp ${formatNumberId(d.total.sales)}` : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/50 border-r border-slate-100">
                        {d.total.qty > 0 ? d.total.qty : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-black text-slate-900 bg-amber-50/50 border-r border-slate-100">
                        {d.total.trx > 0 ? d.total.trx : '—'}
                      </td>
                      <td className="py-2.5 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/50 border-r border-slate-100">
                        {d.total.pricePoint > 0 ? `Rp ${formatNumberId(d.total.pricePoint)}` : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-bold text-indigo-800 bg-indigo-50/30 border-r border-slate-100">
                        {d.traffic.total > 0 ? formatNumberId(d.traffic.total) : '—'}
                      </td>
                      <td className="py-2.5 px-2 text-center font-mono-num font-black text-indigo-900 bg-indigo-50/30 border-r border-slate-100">
                        {d.traffic.total > 0 ? `${d.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* Reasons */}
                      <td className="py-2.5 px-3 text-slate-600 text-[11px] truncate max-w-[200px]" title={d.reasonsText || 'Tidak ada catatan'}>
                        {d.reasonsText || <span className="text-slate-300 italic">—</span>}
                      </td>
                    </tr>
                  );
                })}
              </tbody>

              {/* Sticky Footer Grand Total */}
              <tfoot className="sticky bottom-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs shadow-lg border-t-2 border-amber-400">
                <tr>
                  <td colSpan={3} className="py-3 px-3 text-center uppercase tracking-wider text-amber-300 border-r border-slate-700">
                    TOTAL KONSOLIDASI (DP + CV) — {MONTH_NAMES_ID[selectedMonthKey]?.toUpperCase()} 2026
                  </td>

                  {/* DP Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-blue-300 border-r border-slate-700">
                    Rp {formatNumberId(dailyConsolidatedData.summary.dp.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.dp.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.dp.trx}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.dp.sharePct.toFixed(1)}%
                  </td>

                  {/* CV Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-purple-300 border-r border-slate-700">
                    Rp {formatNumberId(dailyConsolidatedData.summary.cv.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.cv.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.cv.trx}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {dailyConsolidatedData.summary.cv.sharePct.toFixed(1)}%
                  </td>

                  {/* Grand Total Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(dailyConsolidatedData.summary.total.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {dailyConsolidatedData.summary.total.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {dailyConsolidatedData.summary.total.trx}
                  </td>
                  <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(dailyConsolidatedData.summary.total.pricePoint)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {formatNumberId(dailyConsolidatedData.summary.traffic.total) || '0'}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {dailyConsolidatedData.summary.traffic.total > 0
                      ? `${dailyConsolidatedData.summary.traffic.conversionRatePct.toFixed(1)}%`
                      : '—'}
                  </td>
                  <td className="py-3 px-3 text-[10px] text-amber-200/80 italic font-bold">
                    Rata-rata: Rp {formatNumberId(dailyConsolidatedData.summary.total.avgDailySales)}/hari
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* CHART DIAGRAM TOTAL DAILY (DP + CV)                          */}
      {/* ============================================================ */}
      {horizon === 'harian' && (
        <div className="space-y-4">
          <PhotoStyleSalesDiagram
            headerTitle={`DIAGRAM TOTAL SALES HARIAN (DP + CV) — ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026 (TGL 1 – ${dailyConsolidatedData.daysCount})`}
            headerSubtitle={`Diagram Batang Vertikal ${dailyConsolidatedData.days.length} Hari & Garis Tren Penjualan 1. Report DP vs 2. Report CV`}
            badgeText="DIAGRAM TOTAL SALES (DP + CV)"
            cardTitle="Grafik Batang Report DP vs Report CV & Garis Tren"
            primaryLabel="1. Report DP"
            secondaryLabel="2. Report CV"
            totalSales={dailyConsolidatedData.summary.total.sales}
            totalPagi={dailyConsolidatedData.summary.dp.sales}
            totalSiang={dailyConsolidatedData.summary.cv.sales}
            compactX={true}
            bars={dailyConsolidatedData.days.map((d) => ({
              key: `tot-d-${d.tanggal}`,
              label: `${d.dayName.slice(0, 3)} (${d.dayNumber})`,
              subLabel: d.tanggal.slice(0, 5),
              pagiSales: d.dp.sales,
              siangSales: d.cv.sales,
              totalSales: d.total.sales,
              isWeekend: d.dayName === 'Sabtu' || d.dayName === 'Minggu',
              isSelected: dailyCalendarFilterIso ? d.tanggal === isoToSlashDate(dailyCalendarFilterIso) : false,
              onClick: () => setDailyCalendarFilterIso(slashDateToIso(d.tanggal)),
            }))}
          />

          <GrowthAndContributionBarDiagram
            headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI HARIAN (DP + CV) — ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026`}
            headerSubtitle={`Grafik Batang Pertumbuhan Harian (Growth %) & Porsi Kontribusi Total Sales (DP + CV) ${dailyConsolidatedData.days.length} Hari`}
            periodBadgeText={`${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026`}
            compactX={true}
            items={dailyConsolidatedData.days.map((d, idx) => ({
              key: `tot-d-gc-${d.tanggal}`,
              label: `${d.dayName.slice(0, 3)} (${d.dayNumber})`,
              subLabel: d.tanggal.slice(0, 5),
              pagiSales: d.dp.sales,
              siangSales: d.cv.sales,
              totalSales: d.total.sales,
              prevSales: idx > 0 ? dailyConsolidatedData.days[idx - 1].total.sales : null,
              isWeekend: d.dayName === 'Sabtu' || d.dayName === 'Minggu',
              isSelected: dailyCalendarFilterIso ? d.tanggal === isoToSlashDate(dailyCalendarFilterIso) : false,
              onClick: () => setDailyCalendarFilterIso(slashDateToIso(d.tanggal)),
            }))}
          />
        </div>
      )}

      {/* ============================================================ */}
      {/* HORIZON 2: TOTAL SALES WEEKLY (KONSOLIDASI MINGGUAN DP + CV) */}
      {/* ============================================================ */}
      {horizon === 'weekly' && (
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          {/* Header Dark Navy Banner */}
          <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#161b2e] to-indigo-950 text-white flex flex-col md:flex-row md:items-center justify-between gap-2 border-b border-slate-800">
            <div>
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                  TABEL KONSOLIDASI SALES WEEKLY — TOTAL REPORT (DP + CV)
                </h3>
              </div>
              <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                Rekapitulasi Mingguan Senin s/d Minggu · Perbandingan &amp; Kontribusi DP vs CV
              </p>
            </div>
            <span className="px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider bg-emerald-400 text-slate-950 shadow-xs">
              {weeklyScope === 'month' ? `BULAN ${MONTH_NAMES_ID[selectedMonthKey]?.toUpperCase()}` : '52 WEEKLY SETAHUN'}
            </span>
          </div>

          {/* Table Container */}
          <div className="overflow-x-auto max-h-[700px] scrollbar-thin">
            <table className="w-full text-left text-xs border-collapse font-sans">
              <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                <tr className="border-b border-slate-800 text-[11px] font-black uppercase tracking-wider text-slate-200">
                  <th className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 min-w-[85px]">
                    WEEK
                  </th>
                  <th className="py-3 px-3 border-r border-slate-800 min-w-[130px]">
                    PERIODE
                  </th>

                  {/* DP Columns */}
                  <th className="py-3 px-3 text-right bg-blue-950/40 text-blue-300 border-r border-slate-800 min-w-[125px]">
                    SALES DP (RP)
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-14">
                    TRX
                  </th>
                  <th className="py-3 px-2.5 text-right bg-blue-950/40 text-blue-300 border-r border-slate-800 min-w-[95px]">
                    ASP
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-16">
                    SHARE DP
                  </th>

                  {/* CV Columns */}
                  <th className="py-3 px-3 text-right bg-purple-950/40 text-purple-300 border-r border-slate-800 min-w-[125px]">
                    SALES CV (RP)
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-14">
                    NOTA
                  </th>
                  <th className="py-3 px-2.5 text-right bg-purple-950/40 text-purple-300 border-r border-slate-800 min-w-[95px]">
                    ASP
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-16">
                    SHARE CV
                  </th>

                  {/* Total Columns */}
                  <th className="py-3 px-3 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[135px]">
                    TOTAL SALES BAZAR
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[78px]">
                    TRX+NOTA
                  </th>
                  <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[95px]">
                    ASP TOTAL
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-16">
                    TRAFFIC
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-20">
                    CONV. RATE
                  </th>
                  <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 min-w-[100px]">
                    RATA-RATA/HR
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white font-medium">
                {weeklyConsolidatedData.weeks.map((w, idx) => {
                  return (
                    <tr
                      key={w.weekNumber}
                      className={`hover:bg-purple-50/30 transition-colors ${
                        idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                      }`}
                    >
                      {/* Week Badge */}
                      <td className="py-3 px-3 text-center border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                        <span className="inline-flex flex-col items-center justify-center px-2 py-1 rounded-lg text-xs font-black bg-purple-900 text-white shadow-2xs">
                          <span>WEEK {w.weekIndex}</span>
                          <span className="text-[9px] font-bold text-purple-200">W-{w.weekNumber}</span>
                        </span>
                      </td>

                      {/* Periode */}
                      <td className="py-3 px-3 border-r border-slate-100">
                        <div className="font-extrabold text-slate-900 text-xs">{w.periodeText}</div>
                        <div className="text-[10px] text-slate-500 font-semibold">{w.monthName} 2026</div>
                      </td>

                      {/* DP Columns */}
                      <td className="py-3 px-3 text-right font-mono-num font-bold text-blue-900 border-r border-slate-100">
                        {w.dp.sales > 0 ? `Rp ${formatNumberId(w.dp.sales)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {w.dp.qty > 0 ? w.dp.qty : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {w.dp.trx > 0 ? w.dp.trx : '—'}
                      </td>
                      <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                        {w.dp.pricePoint > 0 ? `Rp ${formatNumberId(w.dp.pricePoint)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-blue-700 border-r border-slate-100">
                        {w.dp.sharePct > 0 ? `${w.dp.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* CV Columns */}
                      <td className="py-3 px-3 text-right font-mono-num font-bold text-purple-900 border-r border-slate-100">
                        {w.cv.sales > 0 ? `Rp ${formatNumberId(w.cv.sales)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {w.cv.qty > 0 ? w.cv.qty : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {w.cv.trx > 0 ? w.cv.trx : '—'}
                      </td>
                      <td className="py-3 px-2.5 text-right font-mono-num text-slate-700 border-r border-slate-100">
                        {w.cv.pricePoint > 0 ? `Rp ${formatNumberId(w.cv.pricePoint)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-purple-700 border-r border-slate-100">
                        {w.cv.sharePct > 0 ? `${w.cv.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* Grand Total Columns */}
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
                        {w.total.pricePoint > 0 ? `Rp ${formatNumberId(w.total.pricePoint)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-indigo-800 bg-indigo-50/30 border-r border-slate-100">
                        {w.traffic.total > 0 ? formatNumberId(w.traffic.total) : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-black text-indigo-900 bg-indigo-50/30 border-r border-slate-100">
                        {w.traffic.total > 0 ? `${w.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                      </td>
                      <td className="py-3 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/40">
                        {w.total.avgDaily > 0 ? `Rp ${formatNumberId(w.total.avgDaily)}` : '—'}
                      </td>
                    </tr>
                  );
                })}
              </tbody>

              {/* Sticky Footer Total */}
              <tfoot className="sticky bottom-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs shadow-lg border-t-2 border-amber-400">
                <tr>
                  <td colSpan={2} className="py-3 px-3 text-center uppercase tracking-wider text-amber-300 border-r border-slate-700">
                    TOTAL KONSOLIDASI WEEKLY ({weeklyConsolidatedData.weeks.length} WEEKS)
                  </td>

                  {/* DP Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-blue-300 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.dp.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.dp.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.dp.trx}
                  </td>
                  <td className="py-3 px-2.5 text-right font-mono-num text-blue-300 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.dp.pricePoint)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.dp.sharePct.toFixed(1)}%
                  </td>

                  {/* CV Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-purple-300 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.cv.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.cv.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.cv.trx}
                  </td>
                  <td className="py-3 px-2.5 text-right font-mono-num text-purple-300 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.cv.pricePoint)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.cv.sharePct.toFixed(1)}%
                  </td>

                  {/* Grand Total Sum */}
                  <td className="py-3 px-3 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.total.sales)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.total.qty}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.total.trx}
                  </td>
                  <td className="py-3 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(weeklyConsolidatedData.summary.total.pricePoint)}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {formatNumberId(weeklyConsolidatedData.summary.traffic.total) || '0'}
                  </td>
                  <td className="py-3 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {weeklyConsolidatedData.summary.traffic.total > 0
                      ? `${weeklyConsolidatedData.summary.traffic.conversionRatePct.toFixed(1)}%`
                      : '—'}
                  </td>
                  <td className="py-3 px-2.5 text-center text-[10px] text-amber-200/80 italic font-bold">
                    100% Total
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* ============================================================ */}
      {/* CHART DIAGRAM TOTAL WEEKLY (DP + CV)                         */}
      {/* ============================================================ */}
      {horizon === 'weekly' && weeklyConsolidatedData.weeks.length > 0 && (
        <div className="space-y-4">
          <PhotoStyleSalesDiagram
            headerTitle={`DIAGRAM TOTAL SALES WEEKLY (DP + CV) — ${weeklyScope === 'month' ? `${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026` : 'SEMUA WEEKLY 2026'} (${weeklyConsolidatedData.weeks.length} WEEKS)`}
            headerSubtitle="Diagram Batang Vertikal Mingguan & Garis Tren Penjualan 1. Report DP vs 2. Report CV"
            badgeText="DIAGRAM TOTAL WEEKLY (DP + CV)"
            cardTitle="Grafik Batang Report DP vs Report CV & Garis Tren"
            primaryLabel="1. Report DP"
            secondaryLabel="2. Report CV"
            totalSales={weeklyConsolidatedData.summary.total.sales}
            totalPagi={weeklyConsolidatedData.summary.dp.sales}
            totalSiang={weeklyConsolidatedData.summary.cv.sales}
            compactX={weeklyConsolidatedData.weeks.length > 12}
            bars={weeklyConsolidatedData.weeks.map((w) => ({
              key: `tot-w-${w.weekNumber}`,
              label: `W-${w.weekNumber}`,
              subLabel: w.periodeText.split('-')[0].trim(),
              pagiSales: w.dp.sales,
              siangSales: w.cv.sales,
              totalSales: w.total.sales,
              onClick: () => onNavigateToDp?.('weekly', `weekly-${w.weekNumber}`),
            }))}
          />

          <GrowthAndContributionBarDiagram
            headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI WEEKLY (DP + CV) — ${weeklyConsolidatedData.weeks.length} WEEKS`}
            headerSubtitle="Grafik Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Total Sales (DP + CV) Tiap Weekly"
            periodBadgeText={`${weeklyConsolidatedData.weeks.length} PERIODE WEEKLY`}
            compactX={weeklyConsolidatedData.weeks.length > 10}
            items={weeklyConsolidatedData.weeks.map((w, idx) => ({
              key: `tot-w-gc-${w.weekNumber}`,
              label: `W-${w.weekNumber}`,
              subLabel: w.periodeText.split('-')[0].trim(),
              pagiSales: w.dp.sales,
              siangSales: w.cv.sales,
              totalSales: w.total.sales,
              prevSales: idx > 0 ? weeklyConsolidatedData.weeks[idx - 1].total.sales : null,
              onClick: () => onNavigateToDp?.('weekly', `weekly-${w.weekNumber}`),
            }))}
          />
        </div>
      )}

      {/* ============================================================ */}
      {/* HORIZON 3: TOTAL SALES BULANAN (12 BULAN / 1 TAHUN 2026)    */}
      {/* ============================================================ */}
      {horizon === 'bulanan' && (
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          {/* Header Dark Navy Banner */}
          <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-[#182138] to-purple-950 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800">
            <div>
              <div className="flex items-center gap-2">
                <CalendarDays className="w-4 h-4 text-amber-400 shrink-0" />
                <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                  TABEL KONSOLIDASI SALES BULANAN (1 TAHUN 2026) — TOTAL REPORT (DP + CV)
                </h3>
              </div>
              <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                Monitoring 12 Bulan (Januari s/d Desember 2026) · Direct Purchase + Konsinyasi
              </p>
            </div>
            <span className="px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider bg-amber-400 text-slate-950 shadow-xs">
              12 BULAN (YTD 2026)
            </span>
          </div>

          {/* Table Container */}
          <div className="overflow-x-auto max-h-[700px] scrollbar-thin">
            <table className="w-full text-left text-xs border-collapse font-sans">
              <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-sm border-b border-slate-700">
                <tr className="border-b border-slate-800 text-[11px] font-black uppercase tracking-wider text-slate-200">
                  <th className="py-3 px-3 border-r border-slate-800 text-center sticky left-0 z-10 bg-slate-900 min-w-[95px]">
                    NO &amp; BULAN
                  </th>
                  <th className="py-3 px-2 text-center border-r border-slate-800 w-16">
                    HARI
                  </th>

                  {/* DP Columns */}
                  <th className="py-3 px-3 text-right bg-blue-950/40 text-blue-300 border-r border-slate-800 min-w-[130px]">
                    SALES DP (RP)
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-14">
                    TRX
                  </th>
                  <th className="py-3 px-2 text-center bg-blue-950/40 text-blue-300 border-r border-slate-800 w-16">
                    SHARE DP
                  </th>

                  {/* CV Columns */}
                  <th className="py-3 px-3 text-right bg-purple-950/40 text-purple-300 border-r border-slate-800 min-w-[130px]">
                    SALES CV (RP)
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-14">
                    NOTA
                  </th>
                  <th className="py-3 px-2 text-center bg-purple-950/40 text-purple-300 border-r border-slate-800 w-16">
                    SHARE CV
                  </th>

                  {/* Total Columns */}
                  <th className="py-3 px-3 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[140px]">
                    TOTAL SALES BAZAR
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 w-14">
                    QTY
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[78px]">
                    TRX+NOTA
                  </th>
                  <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[95px]">
                    ASP
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-16">
                    TRAFFIC
                  </th>
                  <th className="py-3 px-2 text-center bg-amber-950/40 text-indigo-200 border-r border-slate-800 w-20">
                    CONV. RATE
                  </th>
                  <th className="py-3 px-2.5 text-right bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[105px]">
                    RATA-RATA/HR
                  </th>
                  <th className="py-3 px-2.5 text-center bg-amber-950/40 text-amber-300 border-r border-slate-800 min-w-[110px]">
                    GROWTH &amp; SHARE
                  </th>
                  <th className="py-3 px-2.5 text-center min-w-[85px]">
                    AKSI
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white font-medium">
                {monthlyAnnualConsolidatedData.months.map((m) => {
                  const isSelected = m.isSelected;
                  const hasSales = m.total.sales > 0;
                  return (
                    <tr
                      key={m.monthKey}
                      className={`hover:bg-purple-50/40 transition-colors ${
                        isSelected
                          ? 'bg-purple-50/70 ring-1 ring-inset ring-purple-300 font-bold'
                          : m.monthIndex % 2 === 0
                          ? 'bg-white'
                          : 'bg-slate-50/40'
                      }`}
                    >
                      {/* No & Bulan */}
                      <td className="py-3 px-3 text-center border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                        <span
                          className={`inline-flex items-center justify-center px-2 py-1 rounded-lg text-xs font-black ${
                            isSelected
                              ? 'bg-purple-700 text-white shadow-xs ring-1 ring-purple-500'
                              : 'bg-slate-900 text-slate-100'
                          }`}
                        >
                          {String(m.monthIndex).padStart(2, '0')} - {m.monthName.toUpperCase()}
                        </span>
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

                      {/* DP Columns */}
                      <td className="py-3 px-3 text-right font-mono-num font-black text-blue-900 border-r border-slate-100">
                        {m.dp.sales > 0 ? `Rp ${formatNumberId(m.dp.sales)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {m.dp.qty > 0 ? m.dp.qty : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {m.dp.trx > 0 ? m.dp.trx : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-blue-700 border-r border-slate-100">
                        {m.dp.sharePct > 0 ? `${m.dp.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* CV Columns */}
                      <td className="py-3 px-3 text-right font-mono-num font-black text-purple-900 border-r border-slate-100">
                        {m.cv.sales > 0 ? `Rp ${formatNumberId(m.cv.sales)}` : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {m.cv.qty > 0 ? m.cv.qty : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num text-slate-800 border-r border-slate-100">
                        {m.cv.trx > 0 ? m.cv.trx : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-purple-700 border-r border-slate-100">
                        {m.cv.sharePct > 0 ? `${m.cv.sharePct.toFixed(1)}%` : '—'}
                      </td>

                      {/* Grand Total Columns */}
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
                      <td className="py-3 px-2 text-center font-mono-num font-bold text-indigo-800 bg-indigo-50/30 border-r border-slate-100">
                        {m.traffic.total > 0 ? formatNumberId(m.traffic.total) : '—'}
                      </td>
                      <td className="py-3 px-2 text-center font-mono-num font-black text-indigo-900 bg-indigo-50/30 border-r border-slate-100">
                        {m.traffic.total > 0 ? `${m.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                      </td>
                      <td className="py-3 px-2.5 text-right font-mono-num text-slate-800 bg-amber-50/40 border-r border-slate-100">
                        {m.total.avgDaily > 0 ? `Rp ${formatNumberId(m.total.avgDaily)}` : '—'}
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

                      {/* Aksi */}
                      <td className="py-3 px-2.5 text-center">
                        <button
                          type="button"
                          onClick={() => {
                            setSelectedMonthKey(m.monthKey);
                            setHorizon('harian');
                            showToast(`Bulan ${m.monthName} 2026 dipilih!`);
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

              {/* Sticky Footer Grand Total 1 Tahun */}
              <tfoot className="sticky bottom-0 z-20 bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white font-black text-xs shadow-lg border-t-2 border-amber-400">
                <tr>
                  <td colSpan={2} className="py-3.5 px-3 text-center uppercase tracking-wider text-amber-300 border-r border-slate-700">
                    GRAND TOTAL 1 TAHUN (DP + CV 2026) — BAZAR MATAHARI
                  </td>

                  {/* DP Sum */}
                  <td className="py-3.5 px-3 text-right font-mono-num text-blue-300 border-r border-slate-700">
                    Rp {formatNumberId(monthlyAnnualConsolidatedData.summary.dp.sales)}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.dp.qty}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.dp.trx}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-blue-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.dp.sharePct.toFixed(1)}%
                  </td>

                  {/* CV Sum */}
                  <td className="py-3.5 px-3 text-right font-mono-num text-purple-300 border-r border-slate-700">
                    Rp {formatNumberId(monthlyAnnualConsolidatedData.summary.cv.sales)}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.cv.qty}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.cv.trx}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-purple-300 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.cv.sharePct.toFixed(1)}%
                  </td>

                  {/* Grand Total Sum */}
                  <td className="py-3.5 px-3 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(monthlyAnnualConsolidatedData.summary.total.sales)}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.total.qty}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.total.trx}
                  </td>
                  <td className="py-3.5 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(monthlyAnnualConsolidatedData.summary.total.pricePoint)}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {formatNumberId(monthlyAnnualConsolidatedData.summary.traffic.total) || '0'}
                  </td>
                  <td className="py-3.5 px-2 text-center font-mono-num text-indigo-200 bg-indigo-500/20 border-r border-slate-700">
                    {monthlyAnnualConsolidatedData.summary.traffic.total > 0
                      ? `${monthlyAnnualConsolidatedData.summary.traffic.conversionRatePct.toFixed(1)}%`
                      : '—'}
                  </td>
                  <td className="py-3.5 px-2.5 text-right font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    Rp {formatNumberId(monthlyAnnualConsolidatedData.summary.total.avgDailySales)}
                  </td>
                  <td className="py-3.5 px-2.5 text-center font-mono-num text-amber-300 bg-amber-500/20 border-r border-slate-700">
                    100% YTD
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
      {/* CHART DIAGRAM TOTAL BULANAN & 1 TAHUN (DP + CV)              */}
      {/* ============================================================ */}
      {horizon === 'bulanan' && (() => {
        const monthWeeks = ANNUAL_2026_WEEKS.filter(
          (w) =>
            w.monthKey === selectedMonthKey ||
            w.slashDates.some((dt) => dt.split('/')[1] === selectedMonthKey)
        ).map((w) => {
          let dpS = 0;
          let cvS = 0;
          w.slashDates.forEach((dt) => {
            dpS += dpDateMap.get(dt)?.sales || 0;
            cvS += cvDateMap.get(dt)?.sales || 0;
          });
          return {
            weekNumber: w.weekNumber,
            periodeText: w.periodeText,
            dpSales: dpS,
            cvSales: cvS,
            totalSales: dpS + cvS,
          };
        });
        const monthWeeksTotal = monthWeeks.reduce((s, w) => s + w.totalSales, 0);
        const monthWeeksDpTotal = monthWeeks.reduce((s, w) => s + w.dpSales, 0);
        const monthWeeksCvTotal = monthWeeks.reduce((s, w) => s + w.cvSales, 0);
        const monthNameUpper = (MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase();

        return (
          <div className="space-y-4">
            <PhotoStyleSalesDiagram
              headerTitle={`DIAGRAM TOTAL SALES BULANAN (DP + CV) — ${monthNameUpper} 2026 (PER WEEKLY)`}
              headerSubtitle="Diagram Batang Vertikal Mingguan & Garis Tren Penjualan 1. Report DP vs 2. Report CV"
              badgeText="DIAGRAM TOTAL BULANAN (DP + CV)"
              cardTitle={`Grafik Batang Weekly ${monthNameUpper} 2026 (DP vs CV)`}
              primaryLabel="1. Report DP"
              secondaryLabel="2. Report CV"
              totalSales={monthWeeksTotal}
              totalPagi={monthWeeksDpTotal}
              totalSiang={monthWeeksCvTotal}
              bars={monthWeeks.map((w) => ({
                key: `tot-bw-${w.weekNumber}`,
                label: `W-${w.weekNumber}`,
                subLabel: w.periodeText.split('-')[0].trim(),
                pagiSales: w.dpSales,
                siangSales: w.cvSales,
                totalSales: w.totalSales,
              }))}
            />

            <PhotoStyleSalesDiagram
              headerTitle="DIAGRAM TOTAL SALES TAHUNAN (DP + CV) — 12 BULAN (JANUARI - DESEMBER 2026)"
              headerSubtitle="Diagram Batang Vertikal 12 Bulan & Garis Tren Penjualan 1. Report DP vs 2. Report CV"
              badgeText="DIAGRAM TOTAL 1 TAHUN (DP + CV)"
              cardTitle="Grafik Batang 12 Bulan Report DP vs Report CV & Garis Tren"
              primaryLabel="1. Report DP"
              secondaryLabel="2. Report CV"
              totalSales={monthlyAnnualConsolidatedData.summary.total.sales}
              totalPagi={monthlyAnnualConsolidatedData.summary.dp.sales}
              totalSiang={monthlyAnnualConsolidatedData.summary.cv.sales}
              bars={monthlyAnnualConsolidatedData.months.map((m) => ({
                key: `tot-m-${m.monthKey}`,
                label: m.monthName.slice(0, 3),
                subLabel: `${String(m.monthIndex).padStart(2, '0')}/26`,
                pagiSales: m.dp.sales,
                siangSales: m.cv.sales,
                totalSales: m.total.sales,
                isSelected: m.isSelected,
                onClick: () => setSelectedMonthKey(m.monthKey),
              }))}
            />

            <GrowthAndContributionBarDiagram
              headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI MINGGUAN (DP + CV) — ${monthNameUpper} 2026`}
              headerSubtitle="Grafik Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Tiap Weekly terhadap Bulan"
              periodBadgeText={`BULAN ${monthNameUpper} 2026`}
              items={monthWeeks.map((w, idx) => ({
                key: `tot-bw-gc-${w.weekNumber}`,
                label: `W-${w.weekNumber}`,
                subLabel: w.periodeText.split('-')[0].trim(),
                pagiSales: w.dpSales,
                siangSales: w.cvSales,
                totalSales: w.totalSales,
                prevSales: idx > 0 ? monthWeeks[idx - 1].totalSales : null,
              }))}
            />

            <GrowthAndContributionBarDiagram
              headerTitle="DIAGRAM BATANG GROWTH & KONTRIBUSI TAHUNAN (DP + CV) — 12 BULAN (JANUARI - DESEMBER 2026)"
              headerSubtitle="Grafik Batang Pertumbuhan Bulanan (MoM Growth %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahunan (100%)"
              periodBadgeText="12 BULAN (JAN – DES 2026)"
              items={monthlyAnnualConsolidatedData.months.map((m, idx) => ({
                key: `tot-m-gc-${m.monthKey}`,
                label: m.monthName.slice(0, 3),
                subLabel: `${String(m.monthIndex).padStart(2, '0')}/26`,
                pagiSales: m.dp.sales,
                siangSales: m.cv.sales,
                totalSales: m.total.sales,
                prevSales: idx > 0 ? monthlyAnnualConsolidatedData.months[idx - 1].total.sales : null,
                isSelected: m.isSelected,
                onClick: () => setSelectedMonthKey(m.monthKey),
              }))}
            />
          </div>
        );
      })()}

      {/* ========================================================================== */}
      {/* TABEL TRAFFIC KASIR (PAGI & SIANG) & TOTAL CONVERSION RATE (DP + CV)      */}
      {/* ========================================================================== */}
      {(() => {
        const renderConversionLabel = (ratePct: number, hasTraffic: boolean) => {
          if (!hasTraffic) {
            return <span className="text-[10px] font-semibold text-slate-400">Belum Ada Traffic</span>;
          }
          if (ratePct >= 30) {
            return <span className="text-[10px] font-bold text-emerald-700">Tinggi</span>;
          }
          if (ratePct >= 15) {
            return <span className="text-[10px] font-bold text-indigo-700">Normal</span>;
          }
          return <span className="text-[10px] font-bold text-amber-700">Perlu Optimasi</span>;
        };

        const activeTrafficSummary =
          horizon === 'harian'
            ? {
                pagi: dailyConsolidatedData.summary.traffic.pagi,
                siang: dailyConsolidatedData.summary.traffic.siang,
                total: dailyConsolidatedData.summary.traffic.total,
                dpTrx: dailyConsolidatedData.summary.dp.trx,
                cvTrx: dailyConsolidatedData.summary.cv.trx,
                totalTrx: dailyConsolidatedData.summary.total.trx,
                totalSales: dailyConsolidatedData.summary.total.sales,
                conversionRatePct: dailyConsolidatedData.summary.traffic.conversionRatePct,
                salesPerVisitor: dailyConsolidatedData.summary.traffic.salesPerVisitor,
                label: `BULAN ${(MONTH_NAMES_ID[selectedMonthKey] || 'NOVEMBER').toUpperCase()} 2026`,
              }
            : horizon === 'weekly'
            ? {
                pagi: weeklyConsolidatedData.summary.traffic.pagi,
                siang: weeklyConsolidatedData.summary.traffic.siang,
                total: weeklyConsolidatedData.summary.traffic.total,
                dpTrx: weeklyConsolidatedData.summary.dp.trx,
                cvTrx: weeklyConsolidatedData.summary.cv.trx,
                totalTrx: weeklyConsolidatedData.summary.total.trx,
                totalSales: weeklyConsolidatedData.summary.total.sales,
                conversionRatePct: weeklyConsolidatedData.summary.traffic.conversionRatePct,
                salesPerVisitor: weeklyConsolidatedData.summary.traffic.salesPerVisitor,
                label: `${weeklyConsolidatedData.weeks.length} WEEKS`,
              }
            : {
                pagi: monthlyAnnualConsolidatedData.summary.traffic.pagi,
                siang: monthlyAnnualConsolidatedData.summary.traffic.siang,
                total: monthlyAnnualConsolidatedData.summary.traffic.total,
                dpTrx: monthlyAnnualConsolidatedData.summary.dp.trx,
                cvTrx: monthlyAnnualConsolidatedData.summary.cv.trx,
                totalTrx: monthlyAnnualConsolidatedData.summary.total.trx,
                totalSales: monthlyAnnualConsolidatedData.summary.total.sales,
                conversionRatePct: monthlyAnnualConsolidatedData.summary.traffic.conversionRatePct,
                salesPerVisitor: monthlyAnnualConsolidatedData.summary.traffic.salesPerVisitor,
                label: '12 BULAN (1 TAHUN 2026)',
              };

        const visibleDailyRows = dailyCalendarFilterIso
          ? dailyConsolidatedData.days.filter((d) => d.tanggal === isoToSlashDate(dailyCalendarFilterIso))
          : dailyConsolidatedData.days;

        return (
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
            {/* Section Header */}
            <div className="px-4 py-3.5 bg-slate-900 text-white flex flex-col lg:flex-row lg:items-center justify-between gap-2 border-b border-slate-800">
              <div>
                <div className="flex items-center gap-2">
                  <Users className="w-4 h-4 text-indigo-300 shrink-0" />
                  <h3 className="text-xs sm:text-sm font-black uppercase tracking-wider text-white">
                    TABEL TRAFFIC PENGUNJUNG (KASIR PAGI &amp; SIANG) &amp; TOTAL CONVERSION RATE (DP + CV)
                  </h3>
                </div>
                <p className="text-[11px] text-slate-300 font-medium mt-0.5">
                  Diisi via Link Pengisian Report Shift oleh Petugas Kasir Pagi &amp; Siang · Rumus Conversion Rate = Total Trx+Nota (DP + CV) ÷ Total Traffic × 100%
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-4 text-xs font-mono-num tabular-nums">
                <div>
                  <span className="text-slate-400">Total Traffic: </span>
                  <strong className="text-white">{formatNumberId(activeTrafficSummary.total) || '0'} org</strong>
                </div>
                <div>
                  <span className="text-slate-400">Total Trx+Nota: </span>
                  <strong className="text-amber-300">{formatNumberId(activeTrafficSummary.totalTrx) || '0'} trx+nota</strong>
                </div>
                <div>
                  <span className="text-slate-400">Total Conv. Rate: </span>
                  <strong className="text-emerald-300">
                    {activeTrafficSummary.total > 0
                      ? `${activeTrafficSummary.conversionRatePct.toFixed(1)}%`
                      : '—'}
                  </strong>
                </div>
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto max-h-[600px] scrollbar-thin">
              <table className="w-full text-left text-xs border-collapse font-sans">
                <thead className="sticky top-0 z-20 bg-slate-900 text-white border-b border-slate-700">
                  <tr className="text-[11px] font-black uppercase tracking-wider text-slate-200">
                    <th className="py-3 px-3 border-r border-slate-800 sticky left-0 z-10 bg-slate-900 min-w-[130px]">
                      {horizon === 'harian' ? 'TANGGAL & HARI' : horizon === 'weekly' ? 'WEEK & PERIODE' : 'BULAN'}
                    </th>
                    <th className="py-3 px-3 text-right bg-amber-950/30 text-amber-200 border-r border-slate-800 min-w-[135px]">
                      TRAFFIC PAGI (ORG)
                    </th>
                    <th className="py-3 px-3 text-right bg-indigo-950/40 text-indigo-200 border-r border-slate-800 min-w-[135px]">
                      TRAFFIC SIANG (ORG)
                    </th>
                    <th className="py-3 px-3 text-right bg-slate-800 text-white border-r border-slate-700 min-w-[125px]">
                      TOTAL TRAFFIC
                    </th>
                    <th className="py-3 px-3 text-center border-r border-slate-800 min-w-[140px]">
                      RASIO PAGI / SIANG
                    </th>
                    <th className="py-3 px-2.5 text-center text-blue-200 border-r border-slate-800 w-20">
                      TRX DP
                    </th>
                    <th className="py-3 px-2.5 text-center text-purple-200 border-r border-slate-800 w-20">
                      NOTA CV
                    </th>
                    <th className="py-3 px-2.5 text-center text-amber-200 border-r border-slate-800 min-w-[105px]">
                      TOTAL TRX+NOTA
                    </th>
                    <th className="py-3 px-3 text-center bg-emerald-950/40 text-emerald-200 border-r border-slate-800 min-w-[145px]">
                      TOTAL CONVERSION RATE
                    </th>
                    <th className="py-3 px-3 text-right text-slate-200 min-w-[130px]">
                      SALES / PENGUNJUNG
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white font-medium">
                  {horizon === 'harian' &&
                    visibleDailyRows.map((d, idx) => {
                      const pShare =
                        d.traffic.total > 0 ? (d.traffic.pagi / d.traffic.total) * 100 : 0;
                      const sShare =
                        d.traffic.total > 0 ? (d.traffic.siang / d.traffic.total) * 100 : 0;
                      const isWknd = d.dayName === 'Sabtu' || d.dayName === 'Minggu';
                      return (
                        <tr
                          key={`traf-${d.tanggal}`}
                          className={`hover:bg-indigo-50/20 transition-colors ${
                            isWknd ? 'bg-amber-50/15' : idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                          }`}
                        >
                          <td className="py-2.5 px-3 border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                            <div className="font-bold text-slate-900 font-mono-num">
                              {d.tanggal}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              <span className={isWknd ? 'text-rose-600 font-bold' : ''}>{d.dayName}</span>
                              {d.weekNumber > 0 && ` · W-${d.weekNumber}`}
                            </div>
                          </td>
                          <td className="py-2.5 px-3 text-right border-r border-slate-100 font-mono-num tabular-nums">
                            {isAdmin && onUpdateTraffic ? (
                              <input
                                type="number"
                                min={0}
                                value={d.traffic.pagi > 0 ? d.traffic.pagi : ''}
                                onChange={(e) => {
                                  const raw = e.target.value;
                                  onUpdateTraffic(
                                    d.tanggal,
                                    'pagi',
                                    raw === '' ? null : Math.max(0, Number(raw))
                                  );
                                }}
                                placeholder="0"
                                className="w-20 px-2 py-1 text-right text-xs font-bold font-mono-num bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                              />
                            ) : (
                              <span className="font-bold text-slate-900">
                                {d.traffic.pagi > 0 ? formatNumberId(d.traffic.pagi) : '—'}
                              </span>
                            )}
                            {d.traffic.pagiStaff && (
                              <div className="text-[10px] text-slate-500 truncate max-w-[125px] ml-auto">
                                Kasir: {d.traffic.pagiStaff}
                              </div>
                            )}
                          </td>
                          <td className="py-2.5 px-3 text-right border-r border-slate-100 font-mono-num tabular-nums">
                            {isAdmin && onUpdateTraffic ? (
                              <input
                                type="number"
                                min={0}
                                value={d.traffic.siang > 0 ? d.traffic.siang : ''}
                                onChange={(e) => {
                                  const raw = e.target.value;
                                  onUpdateTraffic(
                                    d.tanggal,
                                    'siang',
                                    raw === '' ? null : Math.max(0, Number(raw))
                                  );
                                }}
                                placeholder="0"
                                className="w-20 px-2 py-1 text-right text-xs font-bold font-mono-num bg-slate-50 border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                              />
                            ) : (
                              <span className="font-bold text-slate-900">
                                {d.traffic.siang > 0 ? formatNumberId(d.traffic.siang) : '—'}
                              </span>
                            )}
                            {d.traffic.siangStaff && (
                              <div className="text-[10px] text-slate-500 truncate max-w-[125px] ml-auto">
                                Kasir: {d.traffic.siangStaff}
                              </div>
                            )}
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono-num tabular-nums font-black text-slate-950 bg-slate-50/80 border-r border-slate-100">
                            {d.traffic.total > 0 ? `${formatNumberId(d.traffic.total)} org` : '—'}
                          </td>
                          <td className="py-2.5 px-3 border-r border-slate-100">
                            {d.traffic.total > 0 ? (
                              <div className="space-y-1">
                                <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden flex">
                                  <div className="bg-amber-500 h-full" style={{ width: `${pShare}%` }} />
                                  <div className="bg-indigo-600 h-full" style={{ width: `${sShare}%` }} />
                                </div>
                                <div className="flex justify-between text-[10px] text-slate-600 font-mono-num tabular-nums">
                                  <span>P {pShare.toFixed(0)}%</span>
                                  <span>S {sShare.toFixed(0)}%</span>
                                </div>
                              </div>
                            ) : (
                              <div className="text-center text-slate-300">—</div>
                            )}
                          </td>
                          <td className="py-2.5 px-2.5 text-center font-mono-num tabular-nums text-blue-800 border-r border-slate-100">
                            {d.dp.trx > 0 ? d.dp.trx : '—'}
                          </td>
                          <td className="py-2.5 px-2.5 text-center font-mono-num tabular-nums text-purple-800 border-r border-slate-100">
                            {d.cv.trx > 0 ? d.cv.trx : '—'}
                          </td>
                          <td className="py-2.5 px-2.5 text-center font-mono-num tabular-nums font-black text-slate-900 border-r border-slate-100">
                            {d.total.trx > 0 ? d.total.trx : '—'}
                          </td>
                          <td className="py-2.5 px-3 text-center bg-emerald-50/30 border-r border-slate-100 font-mono-num tabular-nums">
                            <div className="font-black text-slate-900 text-xs">
                              {d.traffic.total > 0 ? `${d.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                            </div>
                            <div>{renderConversionLabel(d.traffic.conversionRatePct, d.traffic.total > 0)}</div>
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono-num tabular-nums text-slate-800">
                            {d.traffic.salesPerVisitor > 0
                              ? `Rp ${formatNumberId(d.traffic.salesPerVisitor)}`
                              : '—'}
                          </td>
                        </tr>
                      );
                    })}

                  {horizon === 'weekly' &&
                    weeklyConsolidatedData.weeks.map((w, idx) => {
                      const pShare =
                        w.traffic.total > 0 ? (w.traffic.pagi / w.traffic.total) * 100 : 0;
                      const sShare =
                        w.traffic.total > 0 ? (w.traffic.siang / w.traffic.total) * 100 : 0;
                      return (
                        <tr
                          key={`traf-w-${w.weekNumber}`}
                          className={`hover:bg-indigo-50/20 transition-colors ${
                            idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                          }`}
                        >
                          <td className="py-3 px-3 border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                            <div className="font-bold text-slate-900">WEEK {w.weekIndex} (W-{w.weekNumber})</div>
                            <div className="text-[11px] text-slate-500">{w.periodeText}</div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-bold text-slate-900 border-r border-slate-100">
                            {w.traffic.pagi > 0 ? formatNumberId(w.traffic.pagi) : '—'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-bold text-slate-900 border-r border-slate-100">
                            {w.traffic.siang > 0 ? formatNumberId(w.traffic.siang) : '—'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-black text-slate-950 bg-slate-50/80 border-r border-slate-100">
                            {w.traffic.total > 0 ? `${formatNumberId(w.traffic.total)} org` : '—'}
                          </td>
                          <td className="py-3 px-3 border-r border-slate-100">
                            {w.traffic.total > 0 ? (
                              <div className="space-y-1">
                                <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden flex">
                                  <div className="bg-amber-500 h-full" style={{ width: `${pShare}%` }} />
                                  <div className="bg-indigo-600 h-full" style={{ width: `${sShare}%` }} />
                                </div>
                                <div className="flex justify-between text-[10px] text-slate-600 font-mono-num tabular-nums">
                                  <span>P {pShare.toFixed(0)}%</span>
                                  <span>S {sShare.toFixed(0)}%</span>
                                </div>
                              </div>
                            ) : (
                              <div className="text-center text-slate-300">—</div>
                            )}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-blue-800 border-r border-slate-100">
                            {w.dp.trx > 0 ? w.dp.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-purple-800 border-r border-slate-100">
                            {w.cv.trx > 0 ? w.cv.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums font-black text-slate-900 border-r border-slate-100">
                            {w.total.trx > 0 ? w.total.trx : '—'}
                          </td>
                          <td className="py-3 px-3 text-center bg-emerald-50/30 border-r border-slate-100 font-mono-num tabular-nums">
                            <div className="font-black text-slate-900 text-xs">
                              {w.traffic.total > 0 ? `${w.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                            </div>
                            <div>{renderConversionLabel(w.traffic.conversionRatePct, w.traffic.total > 0)}</div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums text-slate-800">
                            {w.traffic.salesPerVisitor > 0
                              ? `Rp ${formatNumberId(w.traffic.salesPerVisitor)}`
                              : '—'}
                          </td>
                        </tr>
                      );
                    })}

                  {horizon === 'bulanan' &&
                    monthlyAnnualConsolidatedData.months.map((m, idx) => {
                      const pShare =
                        m.traffic.total > 0 ? (m.traffic.pagi / m.traffic.total) * 100 : 0;
                      const sShare =
                        m.traffic.total > 0 ? (m.traffic.siang / m.traffic.total) * 100 : 0;
                      return (
                        <tr
                          key={`traf-m-${m.monthKey}`}
                          className={`hover:bg-indigo-50/20 transition-colors ${
                            m.isSelected
                              ? 'bg-indigo-50/40 font-bold'
                              : idx % 2 === 0
                              ? 'bg-white'
                              : 'bg-slate-50/40'
                          }`}
                        >
                          <td className="py-3 px-3 border-r border-slate-100 sticky left-0 z-10 bg-inherit">
                            <div className="font-bold text-slate-900">
                              {String(m.monthIndex).padStart(2, '0')} · {m.monthName} 2026
                            </div>
                            <div className="text-[11px] text-slate-500">{m.daysCount} hari</div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-bold text-slate-900 border-r border-slate-100">
                            {m.traffic.pagi > 0 ? formatNumberId(m.traffic.pagi) : '—'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-bold text-slate-900 border-r border-slate-100">
                            {m.traffic.siang > 0 ? formatNumberId(m.traffic.siang) : '—'}
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums font-black text-slate-950 bg-slate-50/80 border-r border-slate-100">
                            {m.traffic.total > 0 ? `${formatNumberId(m.traffic.total)} org` : '—'}
                          </td>
                          <td className="py-3 px-3 border-r border-slate-100">
                            {m.traffic.total > 0 ? (
                              <div className="space-y-1">
                                <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden flex">
                                  <div className="bg-amber-500 h-full" style={{ width: `${pShare}%` }} />
                                  <div className="bg-indigo-600 h-full" style={{ width: `${sShare}%` }} />
                                </div>
                                <div className="flex justify-between text-[10px] text-slate-600 font-mono-num tabular-nums">
                                  <span>P {pShare.toFixed(0)}%</span>
                                  <span>S {sShare.toFixed(0)}%</span>
                                </div>
                              </div>
                            ) : (
                              <div className="text-center text-slate-300">—</div>
                            )}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-blue-800 border-r border-slate-100">
                            {m.dp.trx > 0 ? m.dp.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-purple-800 border-r border-slate-100">
                            {m.cv.trx > 0 ? m.cv.trx : '—'}
                          </td>
                          <td className="py-3 px-2.5 text-center font-mono-num tabular-nums font-black text-slate-900 border-r border-slate-100">
                            {m.total.trx > 0 ? m.total.trx : '—'}
                          </td>
                          <td className="py-3 px-3 text-center bg-emerald-50/30 border-r border-slate-100 font-mono-num tabular-nums">
                            <div className="font-black text-slate-900 text-xs">
                              {m.traffic.total > 0 ? `${m.traffic.conversionRatePct.toFixed(1)}%` : '—'}
                            </div>
                            <div>{renderConversionLabel(m.traffic.conversionRatePct, m.traffic.total > 0)}</div>
                          </td>
                          <td className="py-3 px-3 text-right font-mono-num tabular-nums text-slate-800">
                            {m.traffic.salesPerVisitor > 0
                              ? `Rp ${formatNumberId(m.traffic.salesPerVisitor)}`
                              : '—'}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>

                {/* Sticky Footer Grand Total Traffic & Conversion Rate */}
                <tfoot className="sticky bottom-0 z-20 bg-slate-950 text-white font-black text-xs border-t-2 border-indigo-400">
                  <tr>
                    <td className="py-3 px-3 uppercase tracking-wider text-amber-300 border-r border-slate-700">
                      TOTAL ({activeTrafficSummary.label})
                    </td>
                    <td className="py-3 px-3 text-right font-mono-num tabular-nums text-amber-200 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.pagi) || '0'} org
                    </td>
                    <td className="py-3 px-3 text-right font-mono-num tabular-nums text-indigo-200 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.siang) || '0'} org
                    </td>
                    <td className="py-3 px-3 text-right font-mono-num tabular-nums text-white bg-slate-800 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.total) || '0'} org
                    </td>
                    <td className="py-3 px-3 text-center font-mono-num tabular-nums text-slate-300 border-r border-slate-700">
                      {activeTrafficSummary.total > 0
                        ? `P ${((activeTrafficSummary.pagi / activeTrafficSummary.total) * 100).toFixed(0)}% · S ${((activeTrafficSummary.siang / activeTrafficSummary.total) * 100).toFixed(0)}%`
                        : '—'}
                    </td>
                    <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-blue-300 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.dpTrx) || '0'}
                    </td>
                    <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-purple-300 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.cvTrx) || '0'}
                    </td>
                    <td className="py-3 px-2.5 text-center font-mono-num tabular-nums text-amber-300 border-r border-slate-700">
                      {formatNumberId(activeTrafficSummary.totalTrx) || '0'} trx+nota
                    </td>
                    <td className="py-3 px-3 text-center font-mono-num tabular-nums text-emerald-300 bg-emerald-950/60 border-r border-slate-700 text-sm">
                      {activeTrafficSummary.total > 0
                        ? `${activeTrafficSummary.conversionRatePct.toFixed(1)}%`
                        : '—'}
                    </td>
                    <td className="py-3 px-3 text-right font-mono-num tabular-nums text-amber-200">
                      {activeTrafficSummary.salesPerVisitor > 0
                        ? `Rp ${formatNumberId(activeTrafficSummary.salesPerVisitor)}`
                        : '—'}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        );
      })()}
    </div>
  );
};
