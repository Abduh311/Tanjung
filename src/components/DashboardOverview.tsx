import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  ANNUAL_2026_WEEKS,
  createWeeklyTableFromAnnualSchedule,
  getIndoDayName,
  isWeekendDate,
  formatNumberId,
  formatReviewBulletItem,
  getReviewDocWeekNumber,
  getReviewDocWeeklyLabel,
} from '../data/initialData';
import {
  CvDailyRecord,
  CvInventoryBrandKey,
  CvInventoryOpsEntry,
  CV_BRAND_CATALOG,
  MONTH_NAMES_ID,
  loadCvDiv1Records,
  loadCvInventoryOps,
  calcCvLiveSalesQtyMap,
} from '../data/cvDiv1Data';
import {
  resolveLatestActiveMonthKey,
  resolveLatestActiveWeekNumber,
  hasReviewDocContent,
} from '../utils/exportExcel';
import { isoToSlashDate, slashDateToIso } from '../utils/dateUtils';
import {
  Calendar,
  CalendarRange,
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  Download,
  Camera,
  Maximize2,
  ArrowUpRight,
  ArrowDownRight,
  FileSpreadsheet,
  PlusCircle,
  Search,
  SlidersHorizontal,
  Eye,
  EyeOff,
  RotateCcw,
  Users,
  Layers,
  ClipboardCheck,
  BarChart3,
  Tag,
} from 'lucide-react';
import {
  DashboardScreenshotModal,
  MonthlyScreenshotFocus,
  WeeklyScreenshotFocus,
} from './DashboardScreenshotModal';
import { PhotoStyleSalesDiagram, SalesDiagramBarItem } from './PhotoStyleSalesDiagram';
import {
  GrowthAndContributionBarDiagram,
  GrowthContributionBarItem,
} from './GrowthAndContributionBarDiagram';
import { CvHorizon } from './ReportCvDiv1View';
import { TotalHorizon } from './TotalDpCvReportView';

export type DashboardHorizon = 'harian' | 'weekly' | 'bulanan';
export type DashboardReportStream = 'ALL' | 'DP' | 'CV';

export type SummaryMetricKey =
  | 'sales'
  | 'qty'
  | 'trx'
  | 'pricePoint'
  | 'bigSize';

export interface VisibleSummaryMetrics {
  sales: boolean;
  qty: boolean;
  trx: boolean;
  pricePoint: boolean;
  bigSize: boolean;
}

export type MonthlyKpiKgLayout = 'grid' | 'list';

const DEFAULT_VISIBLE_METRICS: VisibleSummaryMetrics = {
  sales: true,
  qty: true,
  trx: true,
  pricePoint: true,
  bigSize: true,
};

const DASHBOARD_METRICS_STORAGE_KEY = 'bazar_dashboard_visible_metrics_v1';

const SUMMARY_METRIC_OPTIONS: {
  key: SummaryMetricKey;
  label: string;
  shortDesc: string;
}[] = [
  { key: 'sales', label: 'Omzet Sales', shortDesc: 'Sales Rp & Growth' },
  { key: 'qty', label: 'Quantity', shortDesc: 'Qty Pcs' },
  { key: 'trx', label: 'Trx / Nota', shortDesc: 'Jumlah Struk' },
  { key: 'pricePoint', label: 'Price Point', shortDesc: 'Sales ÷ Qty' },
  { key: 'bigSize', label: 'Big Size', shortDesc: 'Sales ÷ Trx/Nota' },
];

export interface DashboardOverviewProps {
  tables: WeeklySalesTable[];
  selectedWeeklyFilter: string;
  onSelectWeeklyFilter: (filterId: string) => void;
  reviewDoc: ReviewMonitoringDoc;
  reviewDocs?: ReviewMonitoringDoc[];
  horizon: DashboardHorizon;
  onChangeHorizon: (h: DashboardHorizon) => void;
  onNavigate: (
    menu:
      | 'dashboard'
      | 'sales_matrix'
      | 'report_cv_div1'
      | 'total_dp_cv'
      | 'data_cv_ops'
      | 'annual_schedule'
      | 'reason_evaluation'
      | 'daily_input'
      | 'report_link'
      | 'review_monitoring'
      | 'staff_analytics'
  ) => void;
  onChangeTables?: (updated: WeeklySalesTable[]) => void;
  isPublicReadOnly?: boolean;
  cvRecords?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  onNavigateToCv?: (h: CvHorizon) => void;
  onNavigateToTotalDpCv?: (h: TotalHorizon) => void;
  onOpenExportModal?: () => void;
}

interface DayConsolidatedMetrics {
  tanggal: string;
  dayName: string;
  isWeekend: boolean;
  weekNumber: number;
  monthKey: string;
  // DP
  dpSales: number;
  dpPagiSales: number;
  dpSiangSales: number;
  dpQty: number;
  dpPagiQty: number;
  dpSiangQty: number;
  dpTrx: number;
  dpPagiTrx: number;
  dpSiangTrx: number;
  pagiTraffic: number;
  siangTraffic: number;
  trafficTotal: number;
  pagiStaff: string;
  siangStaff: string;
  reasonPagi: string;
  reasonSiang: string;
  // CV
  cvSales: number;
  cvQty: number;
  cvNota: number;
  graphisSales: number;
  graphisPagiSales: number;
  graphisSiangSales: number;
  graphisQty: number;
  graphisNota: number;
  simplicitySales: number;
  simplicityPagiSales: number;
  simplicitySiangSales: number;
  simplicityQty: number;
  simplicityNota: number;
  cvReason: string;
  // Combined DP + CV
  grandSales: number;
  totalQty: number;
  totalTrxNota: number;
  convRatePct: number;
}

function getMonthLastDayNum(monthKey: string, year = 2026): number {
  const m = parseInt(monthKey, 10) || 11;
  return new Date(year, m, 0).getDate();
}

function formatShortRp(val: number): string {
  if (val >= 1000000000) return `${(val / 1000000000).toFixed(2)}M`;
  if (val >= 1000000) return `${(val / 1000000).toFixed(1)}jt`;
  if (val >= 1000) return `${Math.round(val / 1000)}rb`;
  return String(val || 0);
}

function calcGrowthPct(current: number, prev: number | null): number | null {
  if (prev === null || prev === undefined) return null;
  if (prev <= 0 && current <= 0) return 0;
  if (prev <= 0 && current > 0) return 100;
  return ((current - prev) / prev) * 100;
}

export const DashboardOverview: React.FC<DashboardOverviewProps> = ({
  tables,
  selectedWeeklyFilter,
  onSelectWeeklyFilter,
  reviewDoc,
  reviewDocs,
  horizon,
  onChangeHorizon,
  onNavigate,
  onChangeTables,
  isPublicReadOnly = false,
  cvRecords,
  cvInventoryOps,
  onNavigateToCv,
  onNavigateToTotalDpCv,
  onOpenExportModal,
}) => {
  const effectiveCvRecords = useMemo(
    () => (cvRecords && cvRecords.length > 0 ? cvRecords : loadCvDiv1Records()),
    [cvRecords]
  );
  const effectiveCvOps = useMemo(
    () => (cvInventoryOps && cvInventoryOps.length > 0 ? cvInventoryOps : loadCvInventoryOps()),
    [cvInventoryOps]
  );

  // Resolve latest active month & week across DP + CV + Data CV
  const latestMeta = useMemo(() => {
    const activeMonthKey = resolveLatestActiveMonthKey(
      'ALL',
      tables,
      effectiveCvRecords,
      effectiveCvOps
    );
    const activeWeekNumber = resolveLatestActiveWeekNumber(
      undefined,
      tables,
      effectiveCvRecords,
      effectiveCvOps,
      activeMonthKey
    );
    return { activeMonthKey, activeWeekNumber };
  }, [tables, effectiveCvRecords, effectiveCvOps]);

  // Core UI State
  const [reportStream, setReportStream] = useState<DashboardReportStream>('ALL');
  const [weeklyMonthFilter, setWeeklyMonthFilter] = useState<string>('ALL');
  const [selectedMonthKey, setSelectedMonthKey] = useState<string>(
    () => latestMeta.activeMonthKey
  );
  const [selectedWeekNumber, setSelectedWeekNumber] = useState<number>(
    () => latestMeta.activeWeekNumber
  );
  const [selectedDayIdx, setSelectedDayIdx] = useState<number>(0);
  const hasUserPickedPeriodRef = useRef<boolean>(false);

  // Monthly sub-view & table filter
  const [monthlyTableMode, setMonthlyTableMode] = useState<
    'harian' | 'per_week' | '12_bulan'
  >('harian');
  const [dayTypeFilter, setDayTypeFilter] = useState<'ALL' | 'WEEKEND' | 'WEEKDAY'>('ALL');
  const [tableSearch, setTableSearch] = useState<string>('');
  const [showOnlyActiveRows, setShowOnlyActiveRows] = useState<boolean>(false);
  const [brandSectionTab, setBrandSectionTab] = useState<'summary' | 'trend'>('summary');

  // Screenshot & Portrait frame controls
  const [isScreenshotModalOpen, setIsScreenshotModalOpen] = useState<boolean>(false);
  const [monthlyScreenshotFocus, setMonthlyScreenshotFocus] =
    useState<MonthlyScreenshotFocus>('ringkasan');
  const [weeklyScreenshotFocus, setWeeklyScreenshotFocus] =
    useState<WeeklyScreenshotFocus>('ringkasan');
  const [isPortrait3x4Frame, setIsPortrait3x4Frame] = useState<boolean>(false);

  // Metric visibility customizer (collapsed by default)
  const [isMetricCustomizerOpen, setIsMetricCustomizerOpen] = useState<boolean>(false);
  const [visibleMetrics, setVisibleMetrics] = useState<VisibleSummaryMetrics>(() => {
    try {
      const saved = localStorage.getItem(DASHBOARD_METRICS_STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        return {
          sales: parsed.sales !== false,
          qty: parsed.qty !== false,
          trx: parsed.trx !== false,
          pricePoint: parsed.pricePoint !== false,
          bigSize: parsed.bigSize !== false,
        };
      }
    } catch {}
    return DEFAULT_VISIBLE_METRICS;
  });

  useEffect(() => {
    try {
      localStorage.setItem(DASHBOARD_METRICS_STORAGE_KEY, JSON.stringify(visibleMetrics));
    } catch {}
  }, [visibleMetrics]);

  // Auto-sync latest active period after async data hydration until user manually picks another period
  useEffect(() => {
    if (!hasUserPickedPeriodRef.current) {
      setSelectedMonthKey(latestMeta.activeMonthKey);
      if (selectedWeeklyFilter === 'ALL') {
        setSelectedWeekNumber(latestMeta.activeWeekNumber);
      }
    }
  }, [latestMeta.activeMonthKey, latestMeta.activeWeekNumber, selectedWeeklyFilter]);

  // Sync week selection with global selectedWeeklyFilter if explicit table selected
  useEffect(() => {
    if (selectedWeeklyFilter && selectedWeeklyFilter !== 'ALL') {
      const found = tables.find((t) => t.id === selectedWeeklyFilter);
      if (found && found.weekNumber !== selectedWeekNumber) {
        setSelectedWeekNumber(found.weekNumber);
      }
    }
  }, [selectedWeeklyFilter, tables]);

  // Fast Date Lookup Maps (DP + CV)
  const { dpDateMap, cvDateMap } = useMemo(() => {
    const dpMap = new Map<
      string,
      {
        sales: number;
        pagiSales: number;
        siangSales: number;
        qty: number;
        pagiQty: number;
        siangQty: number;
        trx: number;
        pagiTrx: number;
        siangTrx: number;
        pagiTraffic: number;
        siangTraffic: number;
        trafficTotal: number;
        pagiStaff: string;
        siangStaff: string;
        reasonPagi: string;
        reasonSiang: string;
        weekNumber: number;
        score: number;
      }
    >();

    tables.forEach((t) => {
      (t.days || []).forEach((d) => {
        const tgl = (d.tanggal || '').trim();
        if (!tgl) return;
        const pS = Number(d.pagi?.sales) || 0;
        const sS = Number(d.siang?.sales) || 0;
        const pQ = Number(d.pagi?.qty) || 0;
        const sQ = Number(d.siang?.qty) || 0;
        const pT = Number(d.pagi?.trx) || 0;
        const sT = Number(d.siang?.trx) || 0;
        const pTrf = Number(d.pagi?.traffic) || 0;
        const sTrf = Number(d.siang?.traffic) || 0;
        const reasonP = (d.reasonPagi || '').trim();
        const reasonS = (d.reasonSiang || '').trim();
        const pStaff = (d.pagi?.nama || '').trim();
        const sStaff = (d.siang?.nama || '').trim();

        const score =
          (pS + sS > 0 ? 100 : 0) +
          (pTrf + sTrf > 0 ? 50 : 0) +
          (pQ + sQ + pT + sT > 0 ? 20 : 0) +
          (reasonP || reasonS || pStaff || sStaff ? 5 : 0);

        const prev = dpMap.get(tgl);
        if (!prev || score >= prev.score) {
          dpMap.set(tgl, {
            sales: pS + sS,
            pagiSales: pS,
            siangSales: sS,
            qty: pQ + sQ,
            pagiQty: pQ,
            siangQty: sQ,
            trx: pT + sT,
            pagiTrx: pT,
            siangTrx: sT,
            pagiTraffic: pTrf,
            siangTraffic: sTrf,
            trafficTotal: pTrf + sTrf,
            pagiStaff: pStaff,
            siangStaff: sStaff,
            reasonPagi: reasonP,
            reasonSiang: reasonS,
            weekNumber: t.weekNumber,
            score,
          });
        }
      });
    });

    const cvMap = new Map<
      string,
      {
        sales: number;
        qty: number;
        nota: number;
        graphisSales: number;
        graphisPagiSales: number;
        graphisSiangSales: number;
        graphisQty: number;
        graphisNota: number;
        simplicitySales: number;
        simplicityPagiSales: number;
        simplicitySiangSales: number;
        simplicityQty: number;
        simplicityNota: number;
        reasonEvent: string;
        weekNumber: number;
      }
    >();

    effectiveCvRecords.forEach((r) => {
      const tgl = (r.tanggal || '').trim();
      if (!tgl) return;
      const gPagiS = Number(r.graphis?.pagi?.sales) || 0;
      const gSiangS = Number(r.graphis?.siang?.sales) || 0;
      const gSales = gPagiS + gSiangS;
      const gQty =
        (Number(r.graphis?.pagi?.qty) || 0) + (Number(r.graphis?.siang?.qty) || 0);
      const gNota =
        (Number(r.graphis?.pagi?.trx) || 0) + (Number(r.graphis?.siang?.trx) || 0);
      const sPagiS = Number(r.simplicity?.pagi?.sales) || 0;
      const sSiangS = Number(r.simplicity?.siang?.sales) || 0;
      const sSales = sPagiS + sSiangS;
      const sQty =
        (Number(r.simplicity?.pagi?.qty) || 0) + (Number(r.simplicity?.siang?.qty) || 0);
      const sNota =
        (Number(r.simplicity?.pagi?.trx) || 0) + (Number(r.simplicity?.siang?.trx) || 0);

      cvMap.set(tgl, {
        sales: gSales + sSales,
        qty: gQty + sQty,
        nota: gNota + sNota,
        graphisSales: gSales,
        graphisPagiSales: gPagiS,
        graphisSiangSales: gSiangS,
        graphisQty: gQty,
        graphisNota: gNota,
        simplicitySales: sSales,
        simplicityPagiSales: sPagiS,
        simplicitySiangSales: sSiangS,
        simplicityQty: sQty,
        simplicityNota: sNota,
        reasonEvent: (r.reasonEvent || '').trim(),
        weekNumber: r.weekNumber,
      });
    });

    return { dpDateMap: dpMap, cvDateMap: cvMap };
  }, [tables, effectiveCvRecords]);

  // Unified Date Builder
  const getDayConsolidated = (tanggal: string, fallbackWk = 45): DayConsolidatedMetrics => {
    const dp = dpDateMap.get(tanggal);
    const cv = cvDateMap.get(tanggal);
    const parts = tanggal.split('/');
    const mKey = parts.length === 3 ? parts[1].padStart(2, '0') : '11';

    const dpSales = dp?.sales ?? 0;
    const cvSales = cv?.sales ?? 0;
    const dpQty = dp?.qty ?? 0;
    const cvQty = cv?.qty ?? 0;
    const dpTrx = dp?.trx ?? 0;
    const cvNota = cv?.nota ?? 0;
    const trafficTotal = dp?.trafficTotal ?? 0;
    const grandSales = dpSales + cvSales;
    const totalQty = dpQty + cvQty;
    const totalTrxNota = dpTrx + cvNota;

    const streamTrx =
      reportStream === 'DP' ? dpTrx : reportStream === 'CV' ? cvNota : totalTrxNota;
    const convRatePct = trafficTotal > 0 ? (streamTrx / trafficTotal) * 100 : 0;

    return {
      tanggal,
      dayName: getIndoDayName(tanggal),
      isWeekend: isWeekendDate(tanggal),
      weekNumber: dp?.weekNumber || cv?.weekNumber || fallbackWk,
      monthKey: mKey,
      dpSales,
      dpPagiSales: dp?.pagiSales ?? 0,
      dpSiangSales: dp?.siangSales ?? 0,
      dpQty,
      dpPagiQty: dp?.pagiQty ?? 0,
      dpSiangQty: dp?.siangQty ?? 0,
      dpTrx,
      dpPagiTrx: dp?.pagiTrx ?? 0,
      dpSiangTrx: dp?.siangTrx ?? 0,
      pagiTraffic: dp?.pagiTraffic ?? 0,
      siangTraffic: dp?.siangTraffic ?? 0,
      trafficTotal,
      pagiStaff: dp?.pagiStaff || '',
      siangStaff: dp?.siangStaff || '',
      reasonPagi: dp?.reasonPagi || '',
      reasonSiang: dp?.reasonSiang || '',
      cvSales,
      cvQty,
      cvNota,
      graphisSales: cv?.graphisSales ?? 0,
      graphisPagiSales: cv?.graphisPagiSales ?? 0,
      graphisSiangSales: cv?.graphisSiangSales ?? 0,
      graphisQty: cv?.graphisQty ?? 0,
      graphisNota: cv?.graphisNota ?? 0,
      simplicitySales: cv?.simplicitySales ?? 0,
      simplicityPagiSales: cv?.simplicityPagiSales ?? 0,
      simplicitySiangSales: cv?.simplicitySiangSales ?? 0,
      simplicityQty: cv?.simplicityQty ?? 0,
      simplicityNota: cv?.simplicityNota ?? 0,
      cvReason: cv?.reasonEvent || '',
      grandSales,
      totalQty,
      totalTrxNota,
      convRatePct,
    };
  };

  // Pick initial day inside latest active week with data (or latest active day across DP/CV)
  useEffect(() => {
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === selectedWeekNumber);
    if (!sched) return;
    let bestIdx = -1;
    sched.slashDates.forEach((tgl, i) => {
      const m = getDayConsolidated(tgl, selectedWeekNumber);
      if (m.grandSales > 0 || m.trafficTotal > 0) bestIdx = i;
    });
    if (bestIdx >= 0) {
      setSelectedDayIdx(bestIdx);
    } else if (selectedWeekNumber === 46) {
      // Nov 13 2026 is Friday (index 4)
      setSelectedDayIdx(4);
    } else {
      setSelectedDayIdx(0);
    }
  }, [selectedWeekNumber]);

  // Active Week Table (for Screenshot modal & DP details)
  const activeWeekTable = useMemo(() => {
    const ex = tables.find((t) => t.weekNumber === selectedWeekNumber);
    if (ex) return ex;
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === selectedWeekNumber);
    if (sched) return createWeeklyTableFromAnnualSchedule(sched);
    return tables[0];
  }, [tables, selectedWeekNumber]);

  // Handle switching month cleanly across UI
  const handleSelectMonthKey = (mk: string) => {
    hasUserPickedPeriodRef.current = true;
    setSelectedMonthKey(mk);
  };

  // Handle switching week cleanly across UI
  const handleSelectWeekNumber = (wkNum: number) => {
    hasUserPickedPeriodRef.current = true;
    const clamped = Math.max(1, Math.min(53, wkNum));
    setSelectedWeekNumber(clamped);
    const ex = tables.find((t) => t.weekNumber === clamped);
    if (ex) {
      onSelectWeeklyFilter(ex.id);
    } else {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === clamped);
      if (sched) {
        const created = createWeeklyTableFromAnnualSchedule(sched);
        if (onChangeTables && !isPublicReadOnly) {
          const next = [...tables, created].sort((a, b) => a.weekNumber - b.weekNumber);
          onChangeTables(next);
        }
        onSelectWeeklyFilter(created.id);
      }
    }
  };

  // Handle jumping to an exact calendar date in Daily mode
  const handlePickExactDateIso = (isoVal: string) => {
    if (!isoVal) return;
    hasUserPickedPeriodRef.current = true;
    const slash = isoToSlashDate(isoVal);
    const sched = ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(slash));
    if (sched) {
      setSelectedWeekNumber(sched.weekNumber);
      const idx = sched.slashDates.indexOf(slash);
      if (idx >= 0) setTimeout(() => setSelectedDayIdx(idx), 0);
      const ex = tables.find((t) => t.weekNumber === sched.weekNumber);
      if (ex) onSelectWeeklyFilter(ex.id);
    }
  };

  // 7 Days of Selected Week with Consolidated Metrics
  const activeWeekDays = useMemo(() => {
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === selectedWeekNumber);
    const slashDates = sched
      ? sched.slashDates
      : (activeWeekTable?.days || []).map((d) => d.tanggal);
    return slashDates.map((tgl) => getDayConsolidated(tgl, selectedWeekNumber));
  }, [selectedWeekNumber, activeWeekTable, dpDateMap, cvDateMap, reportStream]);

  const safeDayIdx = selectedDayIdx < activeWeekDays.length ? selectedDayIdx : 0;
  const activeDay = activeWeekDays[safeDayIdx] || activeWeekDays[0];
  const prevDay =
    safeDayIdx > 0
      ? activeWeekDays[safeDayIdx - 1]
      : null;

  // Multi-Week Annual/Month Summaries (1..53)
  const weeklySummaries53 = useMemo(() => {
    return ANNUAL_2026_WEEKS.map((w) => {
      const days = w.slashDates.map((tgl) => getDayConsolidated(tgl, w.weekNumber));
      const dpSales = days.reduce((s, d) => s + d.dpSales, 0);
      const dpPagiSales = days.reduce((s, d) => s + d.dpPagiSales, 0);
      const dpSiangSales = days.reduce((s, d) => s + d.dpSiangSales, 0);
      const dpQty = days.reduce((s, d) => s + d.dpQty, 0);
      const dpTrx = days.reduce((s, d) => s + d.dpTrx, 0);
      const cvSales = days.reduce((s, d) => s + d.cvSales, 0);
      const cvQty = days.reduce((s, d) => s + d.cvQty, 0);
      const cvNota = days.reduce((s, d) => s + d.cvNota, 0);
      const graphisSales = days.reduce((s, d) => s + d.graphisSales, 0);
      const graphisPagiSales = days.reduce((s, d) => s + d.graphisPagiSales, 0);
      const graphisSiangSales = days.reduce((s, d) => s + d.graphisSiangSales, 0);
      const graphisQty = days.reduce((s, d) => s + d.graphisQty, 0);
      const graphisNota = days.reduce((s, d) => s + d.graphisNota, 0);
      const simplicitySales = days.reduce((s, d) => s + d.simplicitySales, 0);
      const simplicityPagiSales = days.reduce((s, d) => s + d.simplicityPagiSales, 0);
      const simplicitySiangSales = days.reduce((s, d) => s + d.simplicitySiangSales, 0);
      const simplicityQty = days.reduce((s, d) => s + d.simplicityQty, 0);
      const simplicityNota = days.reduce((s, d) => s + d.simplicityNota, 0);
      const pagiTraffic = days.reduce((s, d) => s + d.pagiTraffic, 0);
      const siangTraffic = days.reduce((s, d) => s + d.siangTraffic, 0);
      const trafficTotal = pagiTraffic + siangTraffic;
      const grandSales = dpSales + cvSales;
      const totalQty = dpQty + cvQty;
      const totalTrxNota = dpTrx + cvNota;
      const weekendSales = days
        .filter((d) => d.isWeekend)
        .reduce(
          (s, d) =>
            s +
            (reportStream === 'DP'
              ? d.dpSales
              : reportStream === 'CV'
                ? d.cvSales
                : d.grandSales),
          0
        );
      const activeDays = days.filter((d) => d.grandSales > 0 || d.trafficTotal > 0).length;

      return {
        weekNumber: w.weekNumber,
        monthKey: w.monthKey,
        monthName: w.monthName,
        periodeText: `${w.seninSlash} – ${w.mingguSlash}`,
        shortRange: `${w.seninSlash.slice(0, 5)}–${w.mingguSlash.slice(0, 5)}`,
        days,
        dpSales,
        dpPagiSales,
        dpSiangSales,
        dpQty,
        dpTrx,
        cvSales,
        cvQty,
        cvNota,
        graphisSales,
        graphisPagiSales,
        graphisSiangSales,
        graphisQty,
        graphisNota,
        simplicitySales,
        simplicityPagiSales,
        simplicitySiangSales,
        simplicityQty,
        simplicityNota,
        pagiTraffic,
        siangTraffic,
        trafficTotal,
        grandSales,
        totalQty,
        totalTrxNota,
        weekendSales,
        activeDays,
      };
    });
  }, [dpDateMap, cvDateMap, reportStream]);

  const activeWeekSummary =
    weeklySummaries53.find((w) => w.weekNumber === selectedWeekNumber) ||
    weeklySummaries53[44];
  const prevWeekSummary =
    selectedWeekNumber > 1
      ? weeklySummaries53.find((w) => w.weekNumber === selectedWeekNumber - 1) || null
      : null;

  // All 12 Calendar Months (01..12) Consolidated Summaries
  const all12MonthsConsolidated = useMemo(() => {
    return Object.keys(MONTH_NAMES_ID)
      .sort()
      .map((mKey) => {
        const lastDay = getMonthLastDayNum(mKey, 2026);
        const days: DayConsolidatedMetrics[] = [];
        for (let d = 1; d <= lastDay; d++) {
          const tgl = `${String(d).padStart(2, '0')}/${mKey}/2026`;
          days.push(getDayConsolidated(tgl, selectedWeekNumber));
        }
        const dpSales = days.reduce((s, d) => s + d.dpSales, 0);
        const dpPagiSales = days.reduce((s, d) => s + d.dpPagiSales, 0);
        const dpSiangSales = days.reduce((s, d) => s + d.dpSiangSales, 0);
        const dpQty = days.reduce((s, d) => s + d.dpQty, 0);
        const dpTrx = days.reduce((s, d) => s + d.dpTrx, 0);
        const cvSales = days.reduce((s, d) => s + d.cvSales, 0);
        const cvQty = days.reduce((s, d) => s + d.cvQty, 0);
        const cvNota = days.reduce((s, d) => s + d.cvNota, 0);
        const graphisSales = days.reduce((s, d) => s + d.graphisSales, 0);
        const graphisPagiSales = days.reduce((s, d) => s + d.graphisPagiSales, 0);
        const graphisSiangSales = days.reduce((s, d) => s + d.graphisSiangSales, 0);
        const graphisQty = days.reduce((s, d) => s + d.graphisQty, 0);
        const graphisNota = days.reduce((s, d) => s + d.graphisNota, 0);
        const simplicitySales = days.reduce((s, d) => s + d.simplicitySales, 0);
        const simplicityPagiSales = days.reduce((s, d) => s + d.simplicityPagiSales, 0);
        const simplicitySiangSales = days.reduce((s, d) => s + d.simplicitySiangSales, 0);
        const simplicityQty = days.reduce((s, d) => s + d.simplicityQty, 0);
        const simplicityNota = days.reduce((s, d) => s + d.simplicityNota, 0);
        const pagiTraffic = days.reduce((s, d) => s + d.pagiTraffic, 0);
        const siangTraffic = days.reduce((s, d) => s + d.siangTraffic, 0);
        const trafficTotal = pagiTraffic + siangTraffic;
        const grandSales = dpSales + cvSales;
        const totalQty = dpQty + cvQty;
        const totalTrxNota = dpTrx + cvNota;
        const activeDays = days.filter((d) => d.grandSales > 0 || d.trafficTotal > 0).length;
        const monthWeeks = weeklySummaries53.filter((w) => w.monthKey === mKey);

        return {
          monthKey: mKey,
          monthName: MONTH_NAMES_ID[mKey] || `Bulan ${mKey}`,
          lastDay,
          days,
          monthWeeks,
          dpSales,
          dpPagiSales,
          dpSiangSales,
          dpQty,
          dpTrx,
          cvSales,
          cvQty,
          cvNota,
          graphisSales,
          graphisPagiSales,
          graphisSiangSales,
          graphisQty,
          graphisNota,
          simplicitySales,
          simplicityPagiSales,
          simplicitySiangSales,
          simplicityQty,
          simplicityNota,
          pagiTraffic,
          siangTraffic,
          trafficTotal,
          grandSales,
          totalQty,
          totalTrxNota,
          activeDays,
        };
      });
  }, [dpDateMap, cvDateMap, weeklySummaries53]);

  const annualGrandTotals = useMemo(() => {
    return all12MonthsConsolidated.reduce(
      (acc, m) => ({
        dpSales: acc.dpSales + m.dpSales,
        dpPagiSales: acc.dpPagiSales + m.dpPagiSales,
        dpSiangSales: acc.dpSiangSales + m.dpSiangSales,
        dpQty: acc.dpQty + m.dpQty,
        dpTrx: acc.dpTrx + m.dpTrx,
        cvSales: acc.cvSales + m.cvSales,
        cvQty: acc.cvQty + m.cvQty,
        cvNota: acc.cvNota + m.cvNota,
        graphisSales: acc.graphisSales + m.graphisSales,
        graphisPagiSales: acc.graphisPagiSales + m.graphisPagiSales,
        graphisSiangSales: acc.graphisSiangSales + m.graphisSiangSales,
        graphisQty: acc.graphisQty + m.graphisQty,
        graphisNota: acc.graphisNota + m.graphisNota,
        simplicitySales: acc.simplicitySales + m.simplicitySales,
        simplicityPagiSales: acc.simplicityPagiSales + m.simplicityPagiSales,
        simplicitySiangSales: acc.simplicitySiangSales + m.simplicitySiangSales,
        simplicityQty: acc.simplicityQty + m.simplicityQty,
        simplicityNota: acc.simplicityNota + m.simplicityNota,
        pagiTraffic: acc.pagiTraffic + m.pagiTraffic,
        siangTraffic: acc.siangTraffic + m.siangTraffic,
        trafficTotal: acc.trafficTotal + m.trafficTotal,
        grandSales: acc.grandSales + m.grandSales,
        totalQty: acc.totalQty + m.totalQty,
        totalTrxNota: acc.totalTrxNota + m.totalTrxNota,
        activeDays: acc.activeDays + m.activeDays,
      }),
      {
        dpSales: 0,
        dpPagiSales: 0,
        dpSiangSales: 0,
        dpQty: 0,
        dpTrx: 0,
        cvSales: 0,
        cvQty: 0,
        cvNota: 0,
        graphisSales: 0,
        graphisPagiSales: 0,
        graphisSiangSales: 0,
        graphisQty: 0,
        graphisNota: 0,
        simplicitySales: 0,
        simplicityPagiSales: 0,
        simplicitySiangSales: 0,
        simplicityQty: 0,
        simplicityNota: 0,
        pagiTraffic: 0,
        siangTraffic: 0,
        trafficTotal: 0,
        grandSales: 0,
        totalQty: 0,
        totalTrxNota: 0,
        activeDays: 0,
      }
    );
  }, [all12MonthsConsolidated]);

  const activeMonthSummary =
    selectedMonthKey === 'ALL'
      ? null
      : all12MonthsConsolidated.find((m) => m.monthKey === selectedMonthKey) ||
        all12MonthsConsolidated[10];

  const prevMonthSummary =
    selectedMonthKey === 'ALL'
      ? null
      : (() => {
          const idx = all12MonthsConsolidated.findIndex(
            (m) => m.monthKey === selectedMonthKey
          );
          return idx > 0 ? all12MonthsConsolidated[idx - 1] : null;
        })();

  // Helper to project active stream values for KPI strip & growth
  const projectStream = (src: {
    dpSales: number;
    dpPagiSales: number;
    dpSiangSales: number;
    dpQty: number;
    dpTrx: number;
    cvSales: number;
    cvQty: number;
    cvNota: number;
    graphisSales: number;
    simplicitySales: number;
    pagiTraffic: number;
    siangTraffic: number;
    trafficTotal: number;
    grandSales: number;
    totalQty: number;
    totalTrxNota: number;
  }) => {
    if (reportStream === 'DP') {
      return {
        sales: src.dpSales,
        qty: src.dpQty,
        trx: src.dpTrx,
        trxLabel: 'Total Trx',
        primaryA: src.dpPagiSales,
        labelA: 'Shift Pagi',
        primaryB: src.dpSiangSales,
        labelB: 'Shift Siang',
        pagiTraffic: src.pagiTraffic,
        siangTraffic: src.siangTraffic,
        trafficTotal: src.trafficTotal,
      };
    }
    if (reportStream === 'CV') {
      return {
        sales: src.cvSales,
        qty: src.cvQty,
        trx: src.cvNota,
        trxLabel: 'Total Nota',
        primaryA: src.graphisSales,
        labelA: 'Graphis (6201)',
        primaryB: src.simplicitySales,
        labelB: 'Simplicity (6244)',
        pagiTraffic: src.pagiTraffic,
        siangTraffic: src.siangTraffic,
        trafficTotal: src.trafficTotal,
      };
    }
    return {
      sales: src.grandSales,
      qty: src.totalQty,
      trx: src.totalTrxNota,
      trxLabel: 'Total Trx+Nota',
      primaryA: src.dpSales,
      labelA: 'Report DP',
      primaryB: src.cvSales,
      labelB: 'Report CV',
      pagiTraffic: src.pagiTraffic,
      siangTraffic: src.siangTraffic,
      trafficTotal: src.trafficTotal,
    };
  };

  // Current Focused Scope Totals (depending on Horizon + Stream)
  const focusedScope = useMemo(() => {
    if (horizon === 'harian') {
      const curr = projectStream(activeDay);
      const prev = prevDay ? projectStream(prevDay) : null;
      return {
        periodTitle: `${activeDay.dayName}, ${activeDay.tanggal} (W${activeDay.weekNumber})`,
        prevLabel: prevDay ? `vs ${prevDay.dayName.slice(0, 3)}` : 'Hari pertama week',
        raw: activeDay,
        prevRaw: prevDay,
        curr,
        growthPct: calcGrowthPct(curr.sales, prev ? prev.sales : null),
        avgPerDay: curr.sales,
        activeDaysLabel: `1 hari fokus · ${activeDay.isWeekend ? 'Weekend' : 'Weekday'}`,
      };
    }
    if (horizon === 'weekly') {
      const curr = projectStream(activeWeekSummary);
      const prev = prevWeekSummary ? projectStream(prevWeekSummary) : null;
      return {
        periodTitle: `Weekly ${activeWeekSummary.weekNumber} (${activeWeekSummary.periodeText})`,
        prevLabel: prevWeekSummary ? `vs W${prevWeekSummary.weekNumber}` : 'Week awal',
        raw: activeWeekSummary,
        prevRaw: prevWeekSummary,
        curr,
        growthPct: calcGrowthPct(curr.sales, prev ? prev.sales : null),
        avgPerDay: Math.round(curr.sales / Math.max(1, activeWeekSummary.activeDays || 7)),
        activeDaysLabel: `${activeWeekSummary.activeDays}/7 hari terisi`,
      };
    }
    const mSource = activeMonthSummary || {
      ...annualGrandTotals,
      monthKey: 'ALL',
      monthName: 'Tahun 2026 (12 Bulan)',
      lastDay: 365,
    };
    const curr = projectStream(mSource);
    const prev = prevMonthSummary ? projectStream(prevMonthSummary) : null;
    return {
      periodTitle:
        selectedMonthKey === 'ALL'
          ? 'Kalender Penuh 2026 (Januari – Desember)'
          : `${mSource.monthName} 2026 (01–${mSource.lastDay} ${mSource.monthName})`,
      prevLabel: prevMonthSummary ? `vs ${prevMonthSummary.monthName.slice(0, 3)}` : 'Periode tahunan',
      raw: mSource,
      prevRaw: prevMonthSummary,
      curr,
      growthPct: calcGrowthPct(curr.sales, prev ? prev.sales : null),
      avgPerDay:
        mSource.activeDays > 0 ? Math.round(curr.sales / mSource.activeDays) : 0,
      activeDaysLabel: `${mSource.activeDays} hari aktif`,
    };
  }, [
    horizon,
    reportStream,
    activeDay,
    prevDay,
    activeWeekSummary,
    prevWeekSummary,
    activeMonthSummary,
    prevMonthSummary,
    annualGrandTotals,
    selectedMonthKey,
  ]);

  // Visible Weekly Matrix Table list (when in Weekly horizon)
  const visibleWeeklyTableRows = useMemo(() => {
    const base =
      weeklyMonthFilter === 'ALL'
        ? weeklySummaries53.filter(
            (w) =>
              w.grandSales > 0 ||
              w.trafficTotal > 0 ||
              w.monthKey === latestMeta.activeMonthKey ||
              w.weekNumber === selectedWeekNumber
          )
        : weeklySummaries53.filter((w) => w.monthKey === weeklyMonthFilter);
    return base.length > 0
      ? base
      : weeklySummaries53.filter((w) => w.monthKey === latestMeta.activeMonthKey);
  }, [weeklySummaries53, weeklyMonthFilter, latestMeta.activeMonthKey, selectedWeekNumber]);

  // Build Diagram Series matching Stream & Horizon
  const diagramData = useMemo(() => {
    const labelPrimary =
      reportStream === 'DP'
        ? 'Shift Pagi'
        : reportStream === 'CV'
          ? 'Graphis (6201)'
          : '1. Report DP';
    const labelSecondary =
      reportStream === 'DP'
        ? 'Shift Siang'
        : reportStream === 'CV'
          ? 'Simplicity (6244)'
          : '2. Report CV';

    if (horizon === 'harian' || (horizon === 'weekly' && weeklyMonthFilter === 'NONE')) {
      const bars: SalesDiagramBarItem[] = activeWeekDays.map((d, i) => {
        const p = projectStream(d);
        return {
          key: d.tanggal,
          label: `${d.dayName.slice(0, 3)}`,
          subLabel: d.tanggal.slice(0, 5),
          pagiSales: p.primaryA,
          siangSales: p.primaryB,
          totalSales: p.sales,
          isWeekend: d.isWeekend,
          isSelected: i === safeDayIdx,
          onClick: () => setSelectedDayIdx(i),
        };
      });
      const growthItems: GrowthContributionBarItem[] = activeWeekDays.map((d, i) => {
        const p = projectStream(d);
        return {
          key: d.tanggal,
          label: `${d.dayName.slice(0, 3)}`,
          subLabel: d.tanggal.slice(0, 5),
          totalSales: p.sales,
          isWeekend: d.isWeekend,
          isSelected: i === safeDayIdx,
          onClick: () => setSelectedDayIdx(i),
        };
      });
      return {
        title: `GRAFIK HARIAN (${labelPrimary.toUpperCase()} vs ${labelSecondary.toUpperCase()}) — W${activeWeekSummary.weekNumber}`,
        subtitle: `Periode ${activeWeekSummary.periodeText} · Klik batang tanggal untuk fokus hari`,
        badgeText:
          reportStream === 'ALL'
            ? 'KONSOLIDASI DP + CV'
            : reportStream === 'DP'
              ? 'REPORT DP'
              : 'REPORT CV',
        labelPrimary,
        labelSecondary,
        totalSales: projectStream(activeWeekSummary).sales,
        totalPrimary: projectStream(activeWeekSummary).primaryA,
        totalSecondary: projectStream(activeWeekSummary).primaryB,
        bars,
        growthItems,
        compactX: false,
      };
    }

    if (horizon === 'weekly') {
      const targetWeeks = visibleWeeklyTableRows.slice(-16);
      const bars: SalesDiagramBarItem[] = targetWeeks.map((w) => {
        const p = projectStream(w);
        return {
          key: `W${w.weekNumber}`,
          label: `W${w.weekNumber}`,
          subLabel: w.shortRange,
          pagiSales: p.primaryA,
          siangSales: p.primaryB,
          totalSales: p.sales,
          isSelected: w.weekNumber === selectedWeekNumber,
          onClick: () => handleSelectWeekNumber(w.weekNumber),
        };
      });
      const growthItems: GrowthContributionBarItem[] = targetWeeks.map((w) => {
        const p = projectStream(w);
        return {
          key: `W${w.weekNumber}`,
          label: `W${w.weekNumber}`,
          subLabel: w.shortRange,
          totalSales: p.sales,
          isSelected: w.weekNumber === selectedWeekNumber,
          onClick: () => handleSelectWeekNumber(w.weekNumber),
        };
      });
      const sumSales = bars.reduce((s, b) => s + b.totalSales, 0);
      const sumA = bars.reduce((s, b) => s + b.pagiSales, 0);
      const sumB = bars.reduce((s, b) => s + b.siangSales, 0);
      return {
        title: `PERBANDINGAN WEEKLY (${labelPrimary.toUpperCase()} vs ${labelSecondary.toUpperCase()})`,
        subtitle:
          weeklyMonthFilter === 'ALL'
            ? 'Seluruh Weekly Aktif 2026 · Klik batang untuk memfokuskan minggu'
            : `Bulan ${MONTH_NAMES_ID[weeklyMonthFilter] || ''} 2026 · Klik batang untuk memfokuskan minggu`,
        badgeText:
          reportStream === 'ALL'
            ? 'WEEKLY DP + CV'
            : reportStream === 'DP'
              ? 'WEEKLY REPORT DP'
              : 'WEEKLY REPORT CV',
        labelPrimary,
        labelSecondary,
        totalSales: sumSales,
        totalPrimary: sumA,
        totalSecondary: sumB,
        bars,
        growthItems,
        compactX: targetWeeks.length > 8,
      };
    }

    // Monthly horizon: if single month selected -> 1..30/31 calendar days; if 'ALL' -> 12 months
    if (selectedMonthKey !== 'ALL' && activeMonthSummary) {
      const bars: SalesDiagramBarItem[] = activeMonthSummary.days.map((d) => {
        const p = projectStream(d);
        return {
          key: d.tanggal,
          label: d.tanggal.slice(0, 2),
          subLabel: d.dayName.slice(0, 3),
          pagiSales: p.primaryA,
          siangSales: p.primaryB,
          totalSales: p.sales,
          isWeekend: d.isWeekend,
        };
      });
      const growthItems: GrowthContributionBarItem[] = activeMonthSummary.monthWeeks.map((w) => {
        const p = projectStream(w);
        return {
          key: `W${w.weekNumber}`,
          label: `W${w.weekNumber}`,
          subLabel: w.shortRange,
          totalSales: p.sales,
          isSelected: w.weekNumber === selectedWeekNumber,
          onClick: () => handleSelectWeekNumber(w.weekNumber),
        };
      });
      const pMonth = projectStream(activeMonthSummary);
      return {
        title: `TREN BULAN ${activeMonthSummary.monthName.toUpperCase()} 2026 (${labelPrimary.toUpperCase()} vs ${labelSecondary.toUpperCase()})`,
        subtitle: `Kalender 01–${activeMonthSummary.lastDay} ${activeMonthSummary.monthName} 2026`,
        badgeText:
          reportStream === 'ALL'
            ? 'MONTHLY DP + CV'
            : reportStream === 'DP'
              ? 'MONTHLY REPORT DP'
              : 'MONTHLY REPORT CV',
        labelPrimary,
        labelSecondary,
        totalSales: pMonth.sales,
        totalPrimary: pMonth.primaryA,
        totalSecondary: pMonth.primaryB,
        bars,
        growthItems,
        compactX: true,
      };
    }

    const bars: SalesDiagramBarItem[] = all12MonthsConsolidated.map((m) => {
      const p = projectStream(m);
      return {
        key: m.monthKey,
        label: m.monthName.slice(0, 3),
        subLabel: `${m.activeDays}hr`,
        pagiSales: p.primaryA,
        siangSales: p.primaryB,
        totalSales: p.sales,
        onClick: () => handleSelectMonthKey(m.monthKey),
      };
    });
    const growthItems: GrowthContributionBarItem[] = all12MonthsConsolidated.map((m) => {
      const p = projectStream(m);
      return {
        key: m.monthKey,
        label: m.monthName.slice(0, 3),
        subLabel: `${m.activeDays}hr`,
        totalSales: p.sales,
        onClick: () => handleSelectMonthKey(m.monthKey),
      };
    });
    const pAnn = projectStream(annualGrandTotals);
    return {
      title: `REKAP TAHUNAN 12 BULAN 2026 (${labelPrimary.toUpperCase()} vs ${labelSecondary.toUpperCase()})`,
      subtitle: 'Januari – Desember 2026 · Klik batang bulan untuk meninjau rincian harian bulan tsb',
      badgeText:
        reportStream === 'ALL'
          ? 'TAHUNAN DP + CV'
          : reportStream === 'DP'
            ? 'TAHUNAN REPORT DP'
            : 'TAHUNAN REPORT CV',
      labelPrimary,
      labelSecondary,
      totalSales: pAnn.sales,
      totalPrimary: pAnn.primaryA,
      totalSecondary: pAnn.primaryB,
      bars,
      growthItems,
      compactX: true,
    };
  }, [
    horizon,
    reportStream,
    activeWeekDays,
    activeWeekSummary,
    safeDayIdx,
    weeklyMonthFilter,
    visibleWeeklyTableRows,
    selectedWeekNumber,
    selectedMonthKey,
    activeMonthSummary,
    all12MonthsConsolidated,
    annualGrandTotals,
  ]);

  // Operational Snapshots (Data CV Ops + Review Monitoring)
  const cvLiveQtyMap = useMemo(
    () => calcCvLiveSalesQtyMap(effectiveCvRecords),
    [effectiveCvRecords]
  );

  // Per-brand Horizon-aware Inventory & Ops (Graphis 6201 + Simplicity 6244)
  const brandOpsComparison = useMemo(() => {
    const isMonthlyHorizon = horizon === 'bulanan';
    const opsPeriodLabel = isMonthlyHorizon
      ? selectedMonthKey === 'ALL'
        ? '12 Bulan 2026'
        : `${MONTH_NAMES_ID[selectedMonthKey] || selectedMonthKey} 2026`
      : `W${selectedWeekNumber}`;

    const matchEntries = effectiveCvOps.filter((e) => {
      if (isMonthlyHorizon) {
        if (e.periodType !== 'monthly') return false;
        return selectedMonthKey === 'ALL' ? true : e.monthKey === selectedMonthKey;
      }
      return e.periodType === 'weekly' && e.weekNumber === selectedWeekNumber;
    });

    const liveForPeriod = isMonthlyHorizon
      ? selectedMonthKey === 'ALL'
        ? {
            GRAPHIS: Array.from(cvLiveQtyMap.monthly.values()).reduce(
              (s, x) => s + x.GRAPHIS,
              0
            ),
            SIMPLICITY: Array.from(cvLiveQtyMap.monthly.values()).reduce(
              (s, x) => s + x.SIMPLICITY,
              0
            ),
          }
        : cvLiveQtyMap.monthly.get(selectedMonthKey) || { GRAPHIS: 0, SIMPLICITY: 0 }
      : cvLiveQtyMap.weekly.get(selectedWeekNumber) || { GRAPHIS: 0, SIMPLICITY: 0 };

    const summarizeBrandOps = (brandKey: CvInventoryBrandKey) => {
      const entries = matchEntries.filter((e) => e.brand === brandKey);
      const stockAwal = entries.reduce((s, e) => s + (e.stockAwal || 0), 0);
      const barangDatang = entries.reduce((s, e) => s + (e.barangDatang || 0), 0);
      const barangKeluar = entries.reduce((s, e) => s + (e.barangKeluar || 0), 0);
      const hasOverride = entries.some(
        (e) => e.salesQtyOverride !== null && e.salesQtyOverride !== undefined
      );
      const overrideQty = entries.reduce((s, e) => s + (e.salesQtyOverride || 0), 0);
      const liveQty = liveForPeriod[brandKey] || 0;
      const salesQty = hasOverride ? overrideQty : liveQty;
      const stockAkhir = stockAwal + barangDatang - salesQty - barangKeluar;
      const promo = entries
        .map((e) => (e.promo || '').trim())
        .filter(Boolean)
        .join(' · ');
      const reasonKendala = entries
        .map((e) => (e.reasonKendala || '').trim())
        .filter(Boolean)
        .join(' · ');
      const actionPlan = entries
        .map((e) => (e.actionPlan || '').trim())
        .filter(Boolean)
        .join(' · ');

      return {
        meta: CV_BRAND_CATALOG[brandKey],
        stockAwal,
        barangDatang,
        barangKeluar,
        salesQty,
        stockAkhir,
        promo: promo || '-',
        reasonKendala: reasonKendala || '-',
        actionPlan: actionPlan || '-',
      };
    };

    return {
      opsPeriodLabel,
      graphis: summarizeBrandOps('GRAPHIS'),
      simplicity: summarizeBrandOps('SIMPLICITY'),
    };
  }, [effectiveCvOps, cvLiveQtyMap, horizon, selectedMonthKey, selectedWeekNumber]);

  const activeCvOpsSnapshot = useMemo(() => {
    const g = brandOpsComparison.graphis;
    const s = brandOpsComparison.simplicity;
    const promo = [
      g.promo !== '-' ? `Graphis: ${g.promo}` : '',
      s.promo !== '-' ? `Simplicity: ${s.promo}` : '',
    ]
      .filter(Boolean)
      .join(' · ');
    const actionPlan = [
      g.actionPlan !== '-' ? `Graphis: ${g.actionPlan}` : '',
      s.actionPlan !== '-' ? `Simplicity: ${s.actionPlan}` : '',
    ]
      .filter(Boolean)
      .join(' · ');

    return {
      opsPeriodLabel: brandOpsComparison.opsPeriodLabel,
      totalStockAwal: g.stockAwal + s.stockAwal,
      totalMasuk: g.barangDatang + s.barangDatang,
      totalKeluar: g.barangKeluar + s.barangKeluar,
      totalSalesQty: g.salesQty + s.salesQty,
      stockAkhir: g.stockAkhir + s.stockAkhir,
      promo: promo || '-',
      actionPlan: actionPlan || '-',
    };
  }, [brandOpsComparison]);

  const latestFilledReviewDoc = useMemo(() => {
    const pool = reviewDocs && reviewDocs.length > 0 ? reviewDocs : [reviewDoc];
    const matching = pool.find(
      (d) => getReviewDocWeekNumber(d) === selectedWeekNumber && hasReviewDocContent(d)
    );
    if (matching) return matching;
    const filled = pool.filter(hasReviewDocContent);
    if (filled.length > 0) return filled[filled.length - 1];
    return reviewDoc;
  }, [reviewDocs, reviewDoc, selectedWeekNumber]);

  const reviewHighlights = useMemo(() => {
    const merch = latestFilledReviewDoc?.sections?.find((s) =>
      s.code.toUpperCase().includes('MERCHANDISE')
    );
    const promoSec = latestFilledReviewDoc?.sections?.find((s) =>
      s.code.toUpperCase().includes('PROMO')
    );
    const actionSec = latestFilledReviewDoc?.sections?.find((s) =>
      s.code.toUpperCase().includes('ACTION')
    );
    const fast =
      merch?.subRows
        ?.find((r) => r.label.toUpperCase().includes('FAST MOVING'))
        ?.items.map(formatReviewBulletItem)
        .filter(Boolean) || [];
    const slow =
      merch?.subRows
        ?.find((r) => r.label.toUpperCase().includes('SLOW MOVING'))
        ?.items.map(formatReviewBulletItem)
        .filter(Boolean) || [];
    const actions = (actionSec?.directItems || [])
      .map(formatReviewBulletItem)
      .filter(Boolean);

    const divisionNotes = ['Divisi 1', 'Divisi 2', 'Divisi 3', 'Divisi 4']
      .map((divLabel) => {
        const merchItems =
          merch?.subRows
            ?.find((r) => r.label.toLowerCase().includes(divLabel.toLowerCase()))
            ?.items.map(formatReviewBulletItem)
            .filter(Boolean) || [];
        const promoItems =
          promoSec?.subRows
            ?.find((r) => r.label.toLowerCase().includes(divLabel.toLowerCase()))
            ?.items.map(formatReviewBulletItem)
            .filter(Boolean) || [];
        return {
          divLabel,
          merchItems,
          promoItems,
        };
      })
      .filter((d) => d.merchItems.length > 0 || d.promoItems.length > 0);

    return {
      fast,
      slow,
      actions,
      divisionNotes,
      docTitle: latestFilledReviewDoc
        ? getReviewDocWeeklyLabel(latestFilledReviewDoc)
        : 'Review Bazar',
    };
  }, [latestFilledReviewDoc]);

  // Consolidated Per-Brand Metrics for Focused Scope
  const focusedBrandMetrics = useMemo(() => {
    const raw = focusedScope.raw;
    const prevRaw = focusedScope.prevRaw;
    const cvTotal = raw.cvSales || 0;
    const grandTotal = raw.grandSales || 0;

    const activeDaysDiv =
      horizon === 'harian'
        ? 1
        : horizon === 'weekly'
          ? Math.max(1, activeWeekSummary.activeDays || 7)
          : Math.max(
              1,
              (activeMonthSummary ? activeMonthSummary.activeDays : annualGrandTotals.activeDays) ||
                1
            );

    const buildBrandStat = (brandKey: CvInventoryBrandKey) => {
      const isGraphis = brandKey === 'GRAPHIS';
      const sales = isGraphis ? raw.graphisSales : raw.simplicitySales;
      const prevSales = prevRaw
        ? isGraphis
          ? prevRaw.graphisSales
          : prevRaw.simplicitySales
        : null;
      const qty = isGraphis ? raw.graphisQty : raw.simplicityQty;
      const nota = isGraphis ? raw.graphisNota : raw.simplicityNota;

      return {
        meta: CV_BRAND_CATALOG[brandKey],
        ops: isGraphis ? brandOpsComparison.graphis : brandOpsComparison.simplicity,
        sales,
        prevSales,
        growthPct: calcGrowthPct(sales, prevSales),
        avgDailySales: Math.round(sales / activeDaysDiv),
        qty,
        nota,
        pricePoint: qty > 0 ? Math.round(sales / qty) : 0,
        bigSize: nota > 0 ? Math.round(sales / nota) : 0,
        shareCvPct: cvTotal > 0 ? (sales / cvTotal) * 100 : 0,
        shareGrandPct: grandTotal > 0 ? (sales / grandTotal) * 100 : 0,
      };
    };

    const graphis = buildBrandStat('GRAPHIS');
    const simplicity = buildBrandStat('SIMPLICITY');
    const dominantBrand =
      graphis.sales === 0 && simplicity.sales === 0
        ? 'Belum ada data'
        : Math.abs(graphis.sales - simplicity.sales) < 1000
          ? 'Seimbang'
          : graphis.sales > simplicity.sales
            ? 'Graphis (6201)'
            : 'Simplicity (6244)';

    return { graphis, simplicity, dominantBrand };
  }, [
    focusedScope,
    brandOpsComparison,
    horizon,
    activeWeekSummary,
    activeMonthSummary,
    annualGrandTotals,
  ]);

  const activeMetricCount = Object.values(visibleMetrics).filter(Boolean).length;

  const toggleSummaryMetric = (key: SummaryMetricKey) => {
    setVisibleMetrics((prev) => {
      const next = { ...prev, [key]: !prev[key] };
      if (Object.values(next).filter(Boolean).length === 0) return prev;
      return next;
    });
  };

  // Filtered Table rows helper for Daily/Weekly/Monthly ledger
  const qLower = tableSearch.trim().toLowerCase();
  const filterDayRow = (d: DayConsolidatedMetrics) => {
    if (dayTypeFilter === 'WEEKEND' && !d.isWeekend) return false;
    if (dayTypeFilter === 'WEEKDAY' && d.isWeekend) return false;
    if (showOnlyActiveRows && d.grandSales <= 0 && d.trafficTotal <= 0) return false;
    if (qLower) {
      const hay =
        `${d.tanggal} ${d.dayName} ${d.pagiStaff} ${d.siangStaff} ${d.reasonPagi} ${d.reasonSiang} ${d.cvReason}`.toLowerCase();
      if (!hay.includes(qLower)) return false;
    }
    return true;
  };

  const currMetrics = focusedScope.curr;
  const pricePointVal =
    currMetrics.qty > 0 ? Math.round(currMetrics.sales / currMetrics.qty) : 0;
  const bigSizeVal =
    currMetrics.trx > 0 ? Math.round(currMetrics.sales / currMetrics.trx) : 0;
  const convRateVal =
    currMetrics.trafficTotal > 0
      ? (currMetrics.trx / currMetrics.trafficTotal) * 100
      : 0;
  const avgSpendPerVisitor =
    currMetrics.trafficTotal > 0
      ? Math.round(currMetrics.sales / currMetrics.trafficTotal)
      : 0;

  const streamDisplayName =
    reportStream === 'ALL'
      ? 'Total DP + CV'
      : reportStream === 'DP'
        ? 'Report DP'
        : 'Report CV Div. 1';

  const kpiColumnsClass =
    activeMetricCount <= 2
      ? 'grid-cols-1 md:grid-cols-2'
      : activeMetricCount === 3
        ? 'grid-cols-1 md:grid-cols-3'
        : 'grid-cols-1 sm:grid-cols-2 xl:grid-cols-4';

  return (
    <div
      className={`space-y-6 transition-all ${
        isPortrait3x4Frame
          ? 'max-w-[780px] mx-auto bg-white p-4 sm:p-6 rounded-xl border border-slate-300 shadow-sm'
          : ''
      }`}
    >
      {/* =====================================================================
          1. TOP EXECUTIVE WORKSPACE BAR (Clean 3-Zone Contract, Flat Elevation)
         ===================================================================== */}
      <section className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5 space-y-4">
        <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-4">
          <div className="min-w-0 space-y-1">
            <h1 className="text-xl sm:text-2xl font-extrabold text-slate-950 tracking-tight">
              Dashboard Penjualan Bazar Matahari
            </h1>
            <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-slate-500">
              <span className="font-semibold text-slate-700">Tanjung</span>
              <span aria-hidden="true">·</span>
              <span>PT MDS Retailing Tbk</span>
              <span aria-hidden="true">·</span>
              <span>
                Fokus:{' '}
                <strong className="text-slate-900">{focusedScope.periodTitle}</strong>
              </span>
              <span aria-hidden="true">·</span>
              <span>
                Bulan Aktif Terbaru:{' '}
                <strong className="text-emerald-700">
                  {MONTH_NAMES_ID[latestMeta.activeMonthKey] || 'November'} 2026 (W
                  {latestMeta.activeWeekNumber})
                </strong>
              </span>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            {/* Report Stream Switcher: DP+CV | Report DP | Report CV */}
            <div
              role="tablist"
              aria-label="Pilih Sumber Laporan"
              className="inline-flex items-center p-1 bg-slate-100 border border-slate-200 rounded-lg gap-1"
            >
              {(
                [
                  { id: 'ALL', label: 'DP + CV' },
                  { id: 'DP', label: 'Report DP' },
                  { id: 'CV', label: 'Report CV' },
                ] as { id: DashboardReportStream; label: string }[]
              ).map((tab) => {
                const isActive = reportStream === tab.id;
                return (
                  <button
                    key={tab.id}
                    type="button"
                    role="tab"
                    aria-selected={isActive}
                    onClick={() => setReportStream(tab.id)}
                    className={`px-3 py-1.5 text-xs font-bold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                      isActive
                        ? tab.id === 'ALL'
                          ? 'bg-emerald-700 text-white shadow-2xs'
                          : tab.id === 'DP'
                            ? 'bg-indigo-700 text-white shadow-2xs'
                            : 'bg-violet-700 text-white shadow-2xs'
                        : 'text-slate-600 hover:text-slate-950 hover:bg-white/60'
                    }`}
                  >
                    {tab.label}
                  </button>
                );
              })}
            </div>

            {/* Horizon Switcher: Daily | Weekly | Monthly */}
            <div
              role="tablist"
              aria-label="Pilih Horizon Waktu"
              className="inline-flex items-center p-1 bg-slate-100 border border-slate-200 rounded-lg gap-1"
            >
              {(
                [
                  { id: 'harian', label: 'Daily', icon: Calendar },
                  { id: 'weekly', label: 'Weekly', icon: CalendarRange },
                  { id: 'bulanan', label: 'Monthly', icon: CalendarDays },
                ] as { id: DashboardHorizon; label: string; icon: React.ComponentType<any> }[]
              ).map((h) => {
                const Icon = h.icon;
                const isActive = horizon === h.id;
                return (
                  <button
                    key={h.id}
                    type="button"
                    role="tab"
                    aria-selected={isActive}
                    onClick={() => onChangeHorizon(h.id)}
                    className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                      isActive
                        ? 'bg-slate-900 text-white shadow-2xs'
                        : 'text-slate-600 hover:text-slate-950 hover:bg-white/60'
                    }`}
                  >
                    <Icon className="w-3.5 h-3.5 shrink-0" />
                    <span>{h.label}</span>
                  </button>
                );
              })}
            </div>

            {/* Utility Actions */}
            <div className="inline-flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setMonthlyScreenshotFocus('ringkasan');
                  setWeeklyScreenshotFocus('ringkasan');
                  setIsScreenshotModalOpen(true);
                }}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-slate-800 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                title="Buka Studio Screenshot Rasio 3:4 (Harian, Weekly & Bulanan)"
              >
                <Camera className="w-3.5 h-3.5 text-slate-600 shrink-0" />
                <span>Screenshot 3:4</span>
              </button>

              {onOpenExportModal && (
                <button
                  type="button"
                  onClick={onOpenExportModal}
                  className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5 shrink-0" />
                  <span>Export Excel</span>
                </button>
              )}

              <button
                type="button"
                onClick={() => setIsPortrait3x4Frame((v) => !v)}
                className={`hidden sm:inline-flex items-center gap-1 px-2.5 py-2 text-xs font-semibold rounded-lg border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  isPortrait3x4Frame
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                }`}
                title="Sesuaikan lebar viewport untuk screenshot"
              >
                <Maximize2 className="w-3.5 h-3.5 shrink-0" />
              </button>
            </div>
          </div>
        </div>

        {/* ===================================================================
            CONTEXTUAL PERIOD NAVIGATOR (Adapts cleanly per Horizon)
           =================================================================== */}
        <div className="pt-3.5 border-t border-slate-200 space-y-3">
          {horizon === 'harian' && (
            <>
              <div className="flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex flex-wrap items-center gap-2 min-w-0">
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      disabled={selectedWeekNumber <= 1}
                      onClick={() => handleSelectWeekNumber(selectedWeekNumber - 1)}
                      className="p-2 border border-slate-200 rounded-lg text-slate-700 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
                      title="Weekly Sebelumnya"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <select
                      aria-label="Pilih Minggu untuk Tampilan Harian"
                      value={selectedWeekNumber}
                      onChange={(e) => handleSelectWeekNumber(Number(e.target.value))}
                      className="px-3 py-2 text-xs font-bold bg-slate-50 border border-slate-200 text-slate-900 rounded-lg focus:outline-none focus:border-slate-400 cursor-pointer"
                    >
                      {ANNUAL_2026_WEEKS.map((w) => {
                        const wkSum = weeklySummaries53.find(
                          (x) => x.weekNumber === w.weekNumber
                        );
                        const val = wkSum
                          ? reportStream === 'DP'
                            ? wkSum.dpSales
                            : reportStream === 'CV'
                              ? wkSum.cvSales
                              : wkSum.grandSales
                          : 0;
                        return (
                          <option key={w.weekNumber} value={w.weekNumber}>
                            W{w.weekNumber} ({w.monthName}): {w.seninSlash}–{w.mingguSlash}
                            {val > 0 ? ` — Rp ${formatShortRp(val)}` : ''}
                          </option>
                        );
                      })}
                    </select>
                    <button
                      type="button"
                      disabled={selectedWeekNumber >= 53}
                      onClick={() => handleSelectWeekNumber(selectedWeekNumber + 1)}
                      className="p-2 border border-slate-200 rounded-lg text-slate-700 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
                      title="Weekly Berikutnya"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>

                  <div className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs">
                    <span className="text-slate-500 font-medium">Lompat Tgl:</span>
                    <input
                      type="date"
                      value={slashDateToIso(activeDay.tanggal)}
                      min="2026-01-01"
                      max="2026-12-31"
                      onChange={(e) => handlePickExactDateIso(e.target.value)}
                      className="font-mono-num text-xs font-bold text-slate-900 bg-transparent focus:outline-none cursor-pointer"
                    />
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  {!isPublicReadOnly && (
                    <button
                      type="button"
                      onClick={() => onNavigate('daily_input')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-900 bg-amber-100 hover:bg-amber-200 border border-amber-300 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <PlusCircle className="w-3.5 h-3.5" />
                      <span>Input Harian</span>
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => {
                      if (onNavigateToTotalDpCv) onNavigateToTotalDpCv('harian');
                      else onNavigate('total_dp_cv');
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <BarChart3 className="w-3.5 h-3.5" />
                    <span>Tabel Total DP+CV</span>
                  </button>
                </div>
              </div>

              {/* 7-Day Selector Strip */}
              <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2">
                {activeWeekDays.map((d, idx) => {
                  const p = projectStream(d);
                  const isSel = idx === safeDayIdx;
                  return (
                    <button
                      key={d.tanggal}
                      type="button"
                      onClick={() => setSelectedDayIdx(idx)}
                      className={`p-2.5 rounded-lg border text-left transition-colors cursor-pointer ${
                        isSel
                          ? 'bg-slate-900 text-white border-slate-900'
                          : d.isWeekend
                            ? 'bg-rose-50/40 hover:bg-rose-50 text-slate-800 border-rose-200/80'
                            : 'bg-slate-50/60 hover:bg-slate-100 text-slate-800 border-slate-200'
                      }`}
                    >
                      <div className="flex items-center justify-between text-[11px] font-semibold">
                        <span className={isSel ? 'text-slate-300' : 'text-slate-500'}>
                          {d.dayName.slice(0, 3)}
                        </span>
                        <span className="font-mono-num">
                          {p.trx} {reportStream === 'DP' ? 'trx' : reportStream === 'CV' ? 'nt' : 't+n'}
                        </span>
                      </div>
                      <div className="font-mono-num text-xs font-bold mt-0.5">
                        {d.tanggal}
                      </div>
                      <div
                        className={`font-mono-num text-xs font-extrabold mt-1 truncate ${
                          isSel
                            ? 'text-emerald-300'
                            : p.sales > 0
                              ? 'text-slate-950'
                              : 'text-slate-400'
                        }`}
                      >
                        Rp {formatNumberId(p.sales) || '0'}
                      </div>
                    </button>
                  );
                })}
              </div>
            </>
          )}

          {horizon === 'weekly' && (
            <div className="space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex flex-wrap items-center gap-2 min-w-0">
                  <select
                    aria-label="Filter Bulan untuk Weekly"
                    value={weeklyMonthFilter}
                    onChange={(e) => {
                      const mk = e.target.value;
                      setWeeklyMonthFilter(mk);
                      if (mk !== 'ALL') {
                        const firstInMonth = ANNUAL_2026_WEEKS.find(
                          (w) => w.monthKey === mk
                        );
                        if (firstInMonth) handleSelectWeekNumber(firstInMonth.weekNumber);
                      }
                    }}
                    className="px-3 py-2 text-xs font-bold bg-slate-50 border border-slate-200 text-slate-900 rounded-lg focus:outline-none cursor-pointer"
                  >
                    <option value="ALL">Semua Bulan Aktif (2026)</option>
                    {Object.entries(MONTH_NAMES_ID).map(([k, v]) => (
                      <option key={k} value={k}>
                        Bulan {v} 2026
                      </option>
                    ))}
                  </select>

                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      disabled={selectedWeekNumber <= 1}
                      onClick={() => handleSelectWeekNumber(selectedWeekNumber - 1)}
                      className="p-2 border border-slate-200 rounded-lg text-slate-700 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <select
                      aria-label="Pilih Minggu 1 sampai 53"
                      value={selectedWeekNumber}
                      onChange={(e) => handleSelectWeekNumber(Number(e.target.value))}
                      className="px-3 py-2 text-xs font-bold bg-slate-50 border border-slate-200 text-slate-900 rounded-lg focus:outline-none cursor-pointer"
                    >
                      {ANNUAL_2026_WEEKS.filter(
                        (w) =>
                          weeklyMonthFilter === 'ALL' || w.monthKey === weeklyMonthFilter
                      ).map((w) => {
                        const wkSum = weeklySummaries53.find(
                          (x) => x.weekNumber === w.weekNumber
                        );
                        const val = wkSum
                          ? reportStream === 'DP'
                            ? wkSum.dpSales
                            : reportStream === 'CV'
                              ? wkSum.cvSales
                              : wkSum.grandSales
                          : 0;
                        return (
                          <option key={w.weekNumber} value={w.weekNumber}>
                            W{w.weekNumber} ({w.monthName}): {w.seninSlash}–{w.mingguSlash}
                            {val > 0 ? ` — Rp ${formatShortRp(val)}` : ''}
                          </option>
                        );
                      })}
                    </select>
                    <button
                      type="button"
                      disabled={selectedWeekNumber >= 53}
                      onClick={() => handleSelectWeekNumber(selectedWeekNumber + 1)}
                      className="p-2 border border-slate-200 rounded-lg text-slate-700 hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setWeeklyScreenshotFocus('growth_weekly');
                      setIsScreenshotModalOpen(true);
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Camera className="w-3.5 h-3.5" />
                    <span>Screenshot Growth Weekly</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => onNavigate('sales_matrix')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" />
                    <span>Tabel W{selectedWeekNumber}</span>
                  </button>
                </div>
              </div>

              {/* Quick Active Weeks Ribbon */}
              <div className="flex items-center gap-1.5 overflow-x-auto pb-1">
                {visibleWeeklyTableRows.slice(-8).map((w) => {
                  const isSel = w.weekNumber === selectedWeekNumber;
                  const p = projectStream(w);
                  return (
                    <button
                      key={w.weekNumber}
                      type="button"
                      onClick={() => handleSelectWeekNumber(w.weekNumber)}
                      className={`px-3 py-1.5 rounded-lg border text-left transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                        isSel
                          ? 'bg-slate-900 text-white border-slate-900'
                          : 'bg-slate-50 hover:bg-slate-100 text-slate-800 border-slate-200'
                      }`}
                    >
                      <div className="text-[11px] font-bold">
                        W{w.weekNumber}{' '}
                        <span className={isSel ? 'text-slate-300 font-normal' : 'text-slate-500 font-normal'}>
                          ({w.shortRange})
                        </span>
                      </div>
                      <div
                        className={`font-mono-num text-xs font-extrabold ${
                          isSel ? 'text-emerald-300' : 'text-slate-900'
                        }`}
                      >
                        Rp {formatShortRp(p.sales)}
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {horizon === 'bulanan' && (
            <div className="space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex flex-wrap items-center gap-2 min-w-0">
                  <select
                    aria-label="Pilih Bulan Kalender 2026"
                    value={selectedMonthKey}
                    onChange={(e) => handleSelectMonthKey(e.target.value)}
                    className="px-3 py-2 text-xs font-bold bg-slate-50 border border-slate-200 text-slate-900 rounded-lg focus:outline-none cursor-pointer"
                  >
                    <option value="ALL">
                      Semua 12 Bulan (Jan – Des 2026) — Rp{' '}
                      {formatNumberId(projectStream(annualGrandTotals).sales) || '0'}
                    </option>
                    {all12MonthsConsolidated.map((m) => (
                      <option key={m.monthKey} value={m.monthKey}>
                        {m.monthName} 2026 (01–{m.lastDay}) — Rp{' '}
                        {formatNumberId(projectStream(m).sales) || '0'}
                      </option>
                    ))}
                  </select>

                  <button
                    type="button"
                    onClick={() => handleSelectMonthKey('ALL')}
                    className={`px-3 py-2 text-xs font-bold rounded-lg border transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                      selectedMonthKey === 'ALL'
                        ? 'bg-slate-900 text-white border-slate-900'
                        : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    12 Bulan Penuh
                  </button>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setMonthlyScreenshotFocus('kontribusi_growth');
                      setIsScreenshotModalOpen(true);
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <Camera className="w-3.5 h-3.5" />
                    <span>Screenshot Kontribusi Bulanan</span>
                  </button>
                </div>
              </div>

              {/* 12-Month Symmetrical Selector Grid */}
              <div className="grid grid-cols-4 sm:grid-cols-6 lg:grid-cols-12 gap-1.5">
                {all12MonthsConsolidated.map((m) => {
                  const isSel = selectedMonthKey === m.monthKey;
                  const val = projectStream(m).sales;
                  return (
                    <button
                      key={m.monthKey}
                      type="button"
                      onClick={() => handleSelectMonthKey(m.monthKey)}
                      className={`px-2.5 py-2 rounded-lg border text-left transition-colors cursor-pointer min-w-0 ${
                        isSel
                          ? 'bg-slate-900 text-white border-slate-900'
                          : val > 0
                            ? 'bg-emerald-50/60 hover:bg-emerald-50 text-slate-900 border-emerald-200'
                            : 'bg-slate-50/70 hover:bg-slate-100 text-slate-600 border-slate-200'
                      }`}
                    >
                      <div className="flex items-center justify-between text-[11px] font-bold">
                        <span className="truncate">{m.monthName.slice(0, 3)}</span>
                        <span
                          className={`font-mono-num text-[10px] ${
                            isSel ? 'text-slate-300' : 'text-slate-400'
                          }`}
                        >
                          {m.activeDays}h
                        </span>
                      </div>
                      <div
                        className={`font-mono-num text-[11px] font-extrabold truncate mt-0.5 ${
                          isSel
                            ? 'text-emerald-300'
                            : val > 0
                              ? 'text-slate-950'
                              : 'text-slate-400'
                        }`}
                      >
                        Rp {formatShortRp(val)}
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          )}
        </div>
      </section>

      {/* =====================================================================
          2. KPI SUMMARY STRIP (Clean 4-Column Hairline Cards + Customizer)
         ===================================================================== */}
      <section className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="text-xs font-semibold text-slate-500">
            Ringkasan Kinerja <strong className="text-slate-900">{streamDisplayName}</strong>{' '}
            <span aria-hidden="true">·</span> {focusedScope.activeDaysLabel}
          </div>
          <button
            type="button"
            onClick={() => setIsMetricCustomizerOpen((v) => !v)}
            className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold text-slate-600 hover:text-slate-950 bg-white border border-slate-200 rounded-lg cursor-pointer"
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            <span>Kustomisasi Metrik</span>
          </button>
        </div>

        {isMetricCustomizerOpen && (
          <div className="bg-white border border-slate-200 rounded-xl p-3 flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-1.5">
              {SUMMARY_METRIC_OPTIONS.map((opt) => {
                const isVis = visibleMetrics[opt.key];
                return (
                  <button
                    key={opt.key}
                    type="button"
                    onClick={() => toggleSummaryMetric(opt.key)}
                    className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-bold border transition-colors cursor-pointer ${
                      isVis
                        ? 'bg-slate-900 text-white border-slate-900'
                        : 'bg-slate-50 text-slate-500 border-slate-200 hover:bg-slate-100'
                    }`}
                  >
                    {isVis ? (
                      <Eye className="w-3.5 h-3.5 text-emerald-300" />
                    ) : (
                      <EyeOff className="w-3.5 h-3.5" />
                    )}
                    <span>{opt.label}</span>
                  </button>
                );
              })}
            </div>
            <button
              type="button"
              onClick={() => setVisibleMetrics(DEFAULT_VISIBLE_METRICS)}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-700 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Reset (5)</span>
            </button>
          </div>
        )}

        <div className={`grid ${kpiColumnsClass} gap-4`}>
          {visibleMetrics.sales && (
            <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-2.5">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs sm:text-sm font-semibold text-slate-600">
                  Total Sales ({streamDisplayName})
                </span>
                {focusedScope.growthPct !== null && (
                  <span
                    className={`inline-flex items-center gap-0.5 text-sm sm:text-base font-extrabold font-mono-num ${
                      focusedScope.growthPct >= 0 ? 'text-emerald-700' : 'text-rose-700'
                    }`}
                  >
                    {focusedScope.growthPct >= 0 ? (
                      <ArrowUpRight className="w-4 h-4 shrink-0" />
                    ) : (
                      <ArrowDownRight className="w-4 h-4 shrink-0" />
                    )}
                    {focusedScope.growthPct >= 0 ? '+' : ''}
                    {focusedScope.growthPct.toFixed(1)}%{' '}
                    <span className="text-xs font-bold">{focusedScope.prevLabel}</span>
                  </span>
                )}
              </div>
              <div className="font-mono-num text-2xl font-extrabold text-slate-950 tracking-tight">
                Rp {formatNumberId(currMetrics.sales) || '0'}
              </div>
              <div className="text-xs sm:text-sm text-slate-600 flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
                <span className="truncate font-medium">
                  {currMetrics.labelA}:{' '}
                  <strong className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950">
                    {currMetrics.sales > 0
                      ? `${Math.round((currMetrics.primaryA / currMetrics.sales) * 100)}%`
                      : '0%'}
                  </strong>
                </span>
                <span className="truncate font-medium">
                  {currMetrics.labelB}:{' '}
                  <strong className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950">
                    {currMetrics.sales > 0
                      ? `${Math.round((currMetrics.primaryB / currMetrics.sales) * 100)}%`
                      : '0%'}
                  </strong>
                </span>
              </div>
            </div>
          )}

          {visibleMetrics.qty && (
            <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-2.5">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs sm:text-sm font-semibold text-slate-600">
                  Total Quantity (Pcs)
                </span>
                <span className="text-xs text-slate-500 font-mono-num">
                  DP {formatNumberId(focusedScope.raw.dpQty) || '0'} · CV{' '}
                  {formatNumberId(focusedScope.raw.cvQty) || '0'}
                </span>
              </div>
              <div className="font-mono-num text-2xl font-extrabold text-slate-950 tracking-tight">
                {formatNumberId(currMetrics.qty) || '0'}{' '}
                <span className="text-sm font-semibold text-slate-500">pcs</span>
              </div>
              {visibleMetrics.pricePoint ? (
                <div className="flex items-baseline justify-between gap-2 bg-indigo-50/75 border border-indigo-100 rounded-lg px-2.5 py-1.5">
                  <span className="text-xs sm:text-sm font-bold text-indigo-800">
                    Price Point
                  </span>
                  <strong className="font-mono-num text-base sm:text-lg font-extrabold text-indigo-950">
                    Rp {formatNumberId(pricePointVal) || '0'}
                  </strong>
                </div>
              ) : (
                <div className="text-xs sm:text-sm text-slate-500 flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
                  <span>
                    DP:{' '}
                    <strong className="font-mono-num text-slate-800">
                      {formatNumberId(focusedScope.raw.dpQty) || '0'} pcs
                    </strong>
                  </span>
                  <span>
                    CV:{' '}
                    <strong className="font-mono-num text-slate-800">
                      {formatNumberId(focusedScope.raw.cvQty) || '0'} pcs
                    </strong>
                  </span>
                </div>
              )}
            </div>
          )}

          {visibleMetrics.trx && (
            <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-2.5">
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs sm:text-sm font-semibold text-slate-600">
                  {currMetrics.trxLabel}
                </span>
                <span className="text-xs text-slate-500 font-mono-num">
                  DP {formatNumberId(focusedScope.raw.dpTrx) || '0'} · Nota{' '}
                  {formatNumberId(focusedScope.raw.cvNota) || '0'}
                </span>
              </div>
              <div className="font-mono-num text-2xl font-extrabold text-slate-950 tracking-tight">
                {formatNumberId(currMetrics.trx) || '0'}{' '}
                <span className="text-sm font-semibold text-slate-500">
                  {reportStream === 'DP' ? 'trx' : reportStream === 'CV' ? 'nota' : 'trx+nota'}
                </span>
              </div>
              {visibleMetrics.bigSize ? (
                <div className="flex items-baseline justify-between gap-2 bg-amber-50/75 border border-amber-100 rounded-lg px-2.5 py-1.5">
                  <span className="text-xs sm:text-sm font-bold text-amber-800">
                    Big Size
                  </span>
                  <strong className="font-mono-num text-base sm:text-lg font-extrabold text-amber-950">
                    Rp {formatNumberId(bigSizeVal) || '0'}
                  </strong>
                </div>
              ) : (
                <div className="text-xs sm:text-sm text-slate-500 flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
                  <span>
                    Trx DP:{' '}
                    <strong className="font-mono-num text-slate-800">
                      {formatNumberId(focusedScope.raw.dpTrx) || '0'}
                    </strong>
                  </span>
                  <span>
                    Nota CV:{' '}
                    <strong className="font-mono-num text-slate-800">
                      {formatNumberId(focusedScope.raw.cvNota) || '0'}
                    </strong>
                  </span>
                </div>
              )}
            </div>
          )}

          <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-2.5">
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs sm:text-sm font-bold text-slate-700">
                Conversion Rate
              </span>
              <span className="text-xs sm:text-sm font-semibold text-slate-600">
                Avg/Org:{' '}
                <strong className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950">
                  Rp {formatNumberId(avgSpendPerVisitor) || '0'}
                </strong>
              </span>
            </div>
            <div className="flex items-baseline justify-between gap-2">
              <div className="font-mono-num text-3xl font-extrabold text-emerald-700 tracking-tight leading-none">
                {convRateVal.toFixed(1)}%
              </div>
              <div className="font-mono-num text-sm sm:text-base font-bold text-slate-800">
                {formatNumberId(currMetrics.trafficTotal) || '0'}{' '}
                <span className="text-xs font-medium text-slate-500">org</span>
              </div>
            </div>
            <div className="text-xs sm:text-sm text-slate-600 flex items-center justify-between gap-2 pt-2 border-t border-slate-100">
              <span>
                Pagi:{' '}
                <strong className="font-mono-num font-bold text-slate-900">
                  {formatNumberId(currMetrics.pagiTraffic) || '0'}
                </strong>
              </span>
              <span>
                Siang:{' '}
                <strong className="font-mono-num font-bold text-slate-900">
                  {formatNumberId(currMetrics.siangTraffic) || '0'}
                </strong>
              </span>
            </div>
          </div>
        </div>
      </section>

      {/* =====================================================================
          3. RINCIAN 4 PILAR DASHBOARD: REPORT DP | REPORT CV | BRAND | TOTAL DP+CV
         ===================================================================== */}
      <section className="space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <span>
            Rincian 4 Pilar Dashboard ({horizon === 'harian' ? 'Daily' : horizon === 'weekly' ? 'Weekly' : 'Monthly'}):{' '}
            <strong className="text-slate-900">Report DP</strong> ·{' '}
            <strong className="text-slate-900">Report CV</strong> ·{' '}
            <strong className="text-slate-900">Brand Div. 1</strong> ·{' '}
            <strong className="text-slate-900">Total DP + CV</strong>
          </span>
          <span className="font-mono-num text-slate-600">{focusedScope.periodTitle}</span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4 items-stretch">
          {/* Pilar 1: Report DP */}
          <div
            className={`bg-white rounded-xl p-4 border flex flex-col justify-between space-y-3 ${
              reportStream === 'DP' ? 'border-indigo-500 ring-1 ring-indigo-200' : 'border-slate-200'
            }`}
          >
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-xs font-bold text-indigo-700">1. Report DP</div>
                <div className="text-sm font-extrabold text-slate-950">Direct Purchase</div>
              </div>
              <button
                type="button"
                onClick={() => onNavigate('sales_matrix')}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-indigo-700 hover:bg-indigo-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Buka DP</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </div>

            <div className="space-y-1">
              <div className="font-mono-num text-xl font-extrabold text-slate-950">
                Rp {formatNumberId(focusedScope.raw.dpSales) || '0'}
              </div>
              <div className="flex flex-wrap items-center gap-x-2.5 gap-y-1 text-xs text-slate-600 font-mono-num">
                <span>{formatNumberId(focusedScope.raw.dpQty) || '0'} pcs</span>
                <span aria-hidden="true">·</span>
                <span>{formatNumberId(focusedScope.raw.dpTrx) || '0'} Trx</span>
                <span aria-hidden="true">·</span>
                <span className="text-sm sm:text-base font-extrabold text-indigo-700">
                  {focusedScope.raw.grandSales > 0
                    ? `${((focusedScope.raw.dpSales / focusedScope.raw.grandSales) * 100).toFixed(1)}%`
                    : '0.0%'}{' '}
                  Total
                </span>
              </div>
              <div className="text-sm text-slate-700 font-mono-num flex flex-wrap items-center justify-between gap-2 pt-1.5">
                <span>
                  PP:{' '}
                  <strong className="font-extrabold text-slate-950">
                    Rp{' '}
                    {formatNumberId(
                      focusedScope.raw.dpQty > 0
                        ? Math.round(focusedScope.raw.dpSales / focusedScope.raw.dpQty)
                        : 0
                    ) || '0'}
                  </strong>
                </span>
                <span>
                  BS:{' '}
                  <strong className="font-extrabold text-slate-950">
                    Rp{' '}
                    {formatNumberId(
                      focusedScope.raw.dpTrx > 0
                        ? Math.round(focusedScope.raw.dpSales / focusedScope.raw.dpTrx)
                        : 0
                    ) || '0'}
                  </strong>
                </span>
              </div>
            </div>

            <div className="pt-2.5 border-t border-slate-100 grid grid-cols-2 gap-3 text-xs">
              <div>
                <div className="text-slate-500 font-medium">Shift Pagi</div>
                <div className="font-mono-num font-bold text-slate-900 mt-0.5">
                  Rp {formatShortRp(focusedScope.raw.dpPagiSales)}
                </div>
                <div className="font-mono-num text-[11px] text-slate-500 truncate">
                  {horizon === 'harian' && 'pagiStaff' in focusedScope.raw
                    ? `${focusedScope.raw.pagiStaff} · ${focusedScope.raw.pagiTraffic} org`
                    : `Trf: ${formatNumberId(focusedScope.raw.pagiTraffic) || '0'} org`}
                </div>
              </div>
              <div>
                <div className="text-slate-500 font-medium">Shift Siang</div>
                <div className="font-mono-num font-bold text-slate-900 mt-0.5">
                  Rp {formatShortRp(focusedScope.raw.dpSiangSales)}
                </div>
                <div className="font-mono-num text-[11px] text-slate-500 truncate">
                  {horizon === 'harian' && 'siangStaff' in focusedScope.raw
                    ? `${focusedScope.raw.siangStaff} · ${focusedScope.raw.siangTraffic} org`
                    : `Trf: ${formatNumberId(focusedScope.raw.siangTraffic) || '0'} org`}
                </div>
              </div>
            </div>
          </div>

          {/* Pilar 2: Report CV Div. 1 */}
          <div
            className={`bg-white rounded-xl p-4 border flex flex-col justify-between space-y-3 ${
              reportStream === 'CV' ? 'border-violet-500 ring-1 ring-violet-200' : 'border-slate-200'
            }`}
          >
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-xs font-bold text-violet-700">2. Report CV Div. 1</div>
                <div className="text-sm font-extrabold text-slate-950">Konsinyasi Fashion</div>
              </div>
              <button
                type="button"
                onClick={() => {
                  if (onNavigateToCv) onNavigateToCv(horizon);
                  else onNavigate('report_cv_div1');
                }}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-violet-700 hover:bg-violet-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Buka CV</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </div>

            <div className="space-y-1">
              <div className="font-mono-num text-xl font-extrabold text-slate-950">
                Rp {formatNumberId(focusedScope.raw.cvSales) || '0'}
              </div>
              <div className="flex flex-wrap items-center gap-x-2.5 gap-y-1 text-xs text-slate-600 font-mono-num">
                <span>{formatNumberId(focusedScope.raw.cvQty) || '0'} pcs</span>
                <span aria-hidden="true">·</span>
                <span>{formatNumberId(focusedScope.raw.cvNota) || '0'} Nota</span>
                <span aria-hidden="true">·</span>
                <span className="text-sm sm:text-base font-extrabold text-violet-700">
                  {focusedScope.raw.grandSales > 0
                    ? `${((focusedScope.raw.cvSales / focusedScope.raw.grandSales) * 100).toFixed(1)}%`
                    : '0.0%'}{' '}
                  Total
                </span>
              </div>
              <div className="text-sm text-slate-700 font-mono-num flex flex-wrap items-center justify-between gap-2 pt-1.5">
                <span>
                  PP:{' '}
                  <strong className="font-extrabold text-slate-950">
                    Rp{' '}
                    {formatNumberId(
                      focusedScope.raw.cvQty > 0
                        ? Math.round(focusedScope.raw.cvSales / focusedScope.raw.cvQty)
                        : 0
                    ) || '0'}
                  </strong>
                </span>
                <span>
                  BS:{' '}
                  <strong className="font-extrabold text-slate-950">
                    Rp{' '}
                    {formatNumberId(
                      focusedScope.raw.cvNota > 0
                        ? Math.round(focusedScope.raw.cvSales / focusedScope.raw.cvNota)
                        : 0
                    ) || '0'}
                  </strong>
                </span>
              </div>
            </div>

            <div className="pt-2.5 border-t border-slate-100 space-y-1.5 text-xs sm:text-sm">
              <div className="flex items-center justify-between gap-2">
                <span className="text-slate-600 font-semibold">
                  Growth CV ({focusedScope.prevLabel})
                </span>
                {(() => {
                  const cvGr = calcGrowthPct(
                    focusedScope.raw.cvSales,
                    focusedScope.prevRaw ? focusedScope.prevRaw.cvSales : null
                  );
                  return (
                    <span
                      className={`font-mono-num text-sm sm:text-base font-extrabold ${
                        cvGr === null
                          ? 'text-slate-400'
                          : cvGr >= 0
                            ? 'text-emerald-700'
                            : 'text-rose-700'
                      }`}
                    >
                      {cvGr === null ? '-' : `${cvGr >= 0 ? '+' : ''}${cvGr.toFixed(1)}%`}
                    </span>
                  );
                })()}
              </div>
              <div className="flex items-center justify-between gap-2">
                <span className="text-slate-600 font-semibold">
                  {horizon === 'harian' ? 'Kontribusi 2 Brand' : 'Avg CV / Hari'}
                </span>
                <span className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950">
                  {horizon === 'harian'
                    ? `${focusedBrandMetrics.graphis.shareCvPct.toFixed(0)}% G · ${focusedBrandMetrics.simplicity.shareCvPct.toFixed(0)}% S`
                    : `Rp ${formatShortRp(
                        focusedBrandMetrics.graphis.avgDailySales +
                          focusedBrandMetrics.simplicity.avgDailySales
                      )}`}
                </span>
              </div>
            </div>
          </div>

          {/* Pilar 3: Rincian Brand Divisi 1 (Tanpa pemisahan Pagi/Siang) */}
          <div className="bg-white rounded-xl p-4 border border-slate-200 flex flex-col justify-between space-y-3">
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-xs font-bold text-sky-700">3. Rincian Brand</div>
                <div className="text-sm font-extrabold text-slate-950">
                  Graphis &amp; Simplicity
                </div>
              </div>
              <button
                type="button"
                onClick={() => setBrandSectionTab('trend')}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-sky-700 hover:bg-sky-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Tren Brand</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </div>

            <div className="space-y-2.5">
              {/* Graphis 6201 */}
              <div className="bg-slate-50 rounded-lg px-2.5 py-2 space-y-0.5">
                <div className="flex items-center justify-between text-xs sm:text-sm">
                  <span className="font-bold text-violet-800">6201 · Graphis</span>
                  <span className="font-mono-num text-sm sm:text-base font-extrabold text-violet-700">
                    {focusedBrandMetrics.graphis.shareCvPct.toFixed(1)}% CV
                  </span>
                </div>
                <div className="flex items-baseline justify-between gap-2">
                  <span className="font-mono-num text-sm font-extrabold text-slate-950">
                    Rp {formatNumberId(focusedBrandMetrics.graphis.sales) || '0'}
                  </span>
                  <span className="font-mono-num text-[11px] text-slate-600">
                    {focusedBrandMetrics.graphis.qty} pcs · {focusedBrandMetrics.graphis.nota} nt
                  </span>
                </div>
              </div>

              {/* Simplicity 6244 */}
              <div className="bg-slate-50 rounded-lg px-2.5 py-2 space-y-0.5">
                <div className="flex items-center justify-between text-xs sm:text-sm">
                  <span className="font-bold text-sky-800">6244 · Simplicity</span>
                  <span className="font-mono-num text-sm sm:text-base font-extrabold text-sky-700">
                    {focusedBrandMetrics.simplicity.shareCvPct.toFixed(1)}% CV
                  </span>
                </div>
                <div className="flex items-baseline justify-between gap-2">
                  <span className="font-mono-num text-sm font-extrabold text-slate-950">
                    Rp {formatNumberId(focusedBrandMetrics.simplicity.sales) || '0'}
                  </span>
                  <span className="font-mono-num text-[11px] text-slate-600">
                    {focusedBrandMetrics.simplicity.qty} pcs · {focusedBrandMetrics.simplicity.nota} nt
                  </span>
                </div>
              </div>
            </div>

            <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-[11px]">
              <span className="text-slate-500">Dominan Periode:</span>
              <span className="font-bold text-slate-900">{focusedBrandMetrics.dominantBrand}</span>
            </div>
          </div>

          {/* Pilar 4: Report Total DP + CV */}
          <div
            className={`bg-white rounded-xl p-4 border flex flex-col justify-between space-y-3 ${
              reportStream === 'ALL'
                ? 'border-emerald-500 ring-1 ring-emerald-200'
                : 'border-slate-200'
            }`}
          >
            <div className="flex items-start justify-between gap-2">
              <div>
                <div className="text-xs font-bold text-emerald-700">
                  4. Report Total DP + CV
                </div>
                <div className="text-sm font-extrabold text-slate-950">Grand Total &amp; Traffic</div>
              </div>
              <button
                type="button"
                onClick={() => {
                  if (onNavigateToTotalDpCv) onNavigateToTotalDpCv(horizon);
                  else onNavigate('total_dp_cv');
                }}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-emerald-700 hover:bg-emerald-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Buka DP+CV</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </div>

            <div className="space-y-1">
              <div className="font-mono-num text-xl font-extrabold text-slate-950">
                Rp {formatNumberId(focusedScope.raw.grandSales) || '0'}
              </div>
              <div className="flex flex-wrap items-center gap-x-2.5 gap-y-1 text-xs text-slate-600 font-mono-num">
                <span>{formatNumberId(focusedScope.raw.totalQty) || '0'} pcs</span>
                <span aria-hidden="true">·</span>
                <span>{formatNumberId(focusedScope.raw.totalTrxNota) || '0'} Trx+Nota</span>
              </div>
              <div className="text-xs sm:text-sm text-slate-600 font-mono-num flex items-center justify-between gap-2 pt-1.5">
                <span>
                  Traffic: {formatNumberId(focusedScope.raw.trafficTotal) || '0'} org
                </span>
                <span className="text-sm sm:text-base font-extrabold text-emerald-700">
                  Conv:{' '}
                  {focusedScope.raw.trafficTotal > 0
                    ? `${((focusedScope.raw.totalTrxNota / focusedScope.raw.trafficTotal) * 100).toFixed(1)}%`
                    : '0.0%'}
                </span>
              </div>
            </div>

            {/* Share Bar DP vs CV + Avg/Org */}
            <div className="pt-2.5 border-t border-slate-100 space-y-1.5 text-xs">
              <div className="flex items-center justify-between gap-1.5 text-sm font-extrabold">
                <span className="text-indigo-700 font-mono-num">
                  DP{' '}
                  {focusedScope.raw.grandSales > 0
                    ? `${Math.round((focusedScope.raw.dpSales / focusedScope.raw.grandSales) * 100)}%`
                    : '0%'}
                </span>
                <span className="text-xs sm:text-sm text-slate-700 font-bold font-mono-num">
                  Avg/Org Rp{' '}
                  {formatShortRp(
                    focusedScope.raw.trafficTotal > 0
                      ? Math.round(focusedScope.raw.grandSales / focusedScope.raw.trafficTotal)
                      : 0
                  )}
                </span>
                <span className="text-violet-700 font-mono-num">
                  CV{' '}
                  {focusedScope.raw.grandSales > 0
                    ? `${Math.round((focusedScope.raw.cvSales / focusedScope.raw.grandSales) * 100)}%`
                    : '0%'}
                </span>
              </div>
              <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden flex">
                <div
                  style={{
                    width: `${
                      focusedScope.raw.grandSales > 0
                        ? (focusedScope.raw.dpSales / focusedScope.raw.grandSales) * 100
                        : 50
                    }%`,
                  }}
                  className="bg-indigo-600 h-full"
                />
                <div
                  style={{
                    width: `${
                      focusedScope.raw.grandSales > 0
                        ? (focusedScope.raw.cvSales / focusedScope.raw.grandSales) * 100
                        : 50
                    }%`,
                  }}
                  className="bg-violet-600 h-full"
                />
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* =====================================================================
          4. VISUAL ANALYTICS CHARTS (Stacked Sales Trend + Growth & Contribution)
         ===================================================================== */}
      <section className="space-y-4">
        <PhotoStyleSalesDiagram
          headerTitle={diagramData.title}
          headerSubtitle={diagramData.subtitle}
          badgeText={diagramData.badgeText}
          cardTitle={`Grafik ${diagramData.labelPrimary} vs ${diagramData.labelSecondary} & Tren Total`}
          primaryLabel={diagramData.labelPrimary}
          secondaryLabel={diagramData.labelSecondary}
          totalSales={diagramData.totalSales}
          totalPagi={diagramData.totalPrimary}
          totalSiang={diagramData.totalSecondary}
          bars={diagramData.bars}
          compactX={diagramData.compactX}
        />

        <GrowthAndContributionBarDiagram
          headerTitle={`PERTUMBUHAN (GROWTH %) & KONTRIBUSI (%) — ${streamDisplayName.toUpperCase()}`}
          headerSubtitle="Perbandingan kontribusi & pertumbuhan penjualan antar periode"
          periodBadgeText={focusedScope.periodTitle}
          items={diagramData.growthItems}
          compactX={diagramData.compactX}
        />
      </section>

      {/* =====================================================================
          4B. DATA BRAND & DIVISI (GRAPHIS 6201 · SIMPLICITY 6244 · DIVISI 1–4)
         ===================================================================== */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <Tag className="w-4 h-4 text-violet-700 shrink-0" />
              <h2 className="text-sm sm:text-base font-extrabold text-slate-950 truncate">
                Data Brand Divisi 1 (Graphis 6201 &amp; Simplicity 6244)
              </h2>
            </div>
            <p className="text-xs text-slate-500 mt-0.5">
              Fokus: <strong className="text-slate-800">{focusedScope.periodTitle}</strong>{' '}
              <span aria-hidden="true">·</span> Dominan:{' '}
              <strong className="text-violet-700">{focusedBrandMetrics.dominantBrand}</strong>{' '}
              <span aria-hidden="true">·</span> Data Stok:{' '}
              <strong className="text-slate-800">{brandOpsComparison.opsPeriodLabel}</strong>
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <div
              role="tablist"
              aria-label="Tampilan Data Brand"
              className="inline-flex items-center p-0.5 bg-slate-100 border border-slate-200 rounded-lg text-xs"
            >
              <button
                type="button"
                role="tab"
                aria-selected={brandSectionTab === 'summary'}
                onClick={() => setBrandSectionTab('summary')}
                className={`px-2.5 py-1.5 rounded-md font-bold cursor-pointer whitespace-nowrap ${
                  brandSectionTab === 'summary'
                    ? 'bg-white text-slate-950 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-950'
                }`}
              >
                Ringkasan &amp; Stok Brand
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={brandSectionTab === 'trend'}
                onClick={() => setBrandSectionTab('trend')}
                className={`px-2.5 py-1.5 rounded-md font-bold cursor-pointer whitespace-nowrap ${
                  brandSectionTab === 'trend'
                    ? 'bg-white text-slate-950 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-950'
                }`}
              >
                Tabel Tren Brand
              </button>
            </div>

            <button
              type="button"
              onClick={() => {
                if (onNavigateToCv) onNavigateToCv(horizon);
                else onNavigate('report_cv_div1');
              }}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-violet-700 hover:bg-violet-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
            >
              <span>Report CV</span>
              <ArrowUpRight className="w-3.5 h-3.5" />
            </button>

            <button
              type="button"
              onClick={() => onNavigate('data_cv_ops')}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
            >
              <span>Data Stok CV</span>
              <ArrowUpRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        {brandSectionTab === 'summary' ? (
          <div className="p-4 space-y-4">
            {/* Horizontal Brand Share Distribution Bar */}
            <div className="space-y-1.5">
              <div className="flex flex-wrap items-center justify-between gap-2 text-xs sm:text-sm font-bold">
                <span className="text-violet-800">
                  Dept 6201 · Graphis:{' '}
                  <strong className="font-mono-num text-base font-extrabold text-violet-950">
                    {focusedBrandMetrics.graphis.shareCvPct.toFixed(1)}%
                  </strong>{' '}
                  CV ({focusedBrandMetrics.graphis.shareGrandPct.toFixed(1)}% Total DP+CV)
                </span>
                <span className="text-slate-600 font-mono-num">
                  Total CV: Rp {formatNumberId(focusedScope.raw.cvSales) || '0'}
                </span>
                <span className="text-sky-800">
                  Dept 6244 · Simplicity:{' '}
                  <strong className="font-mono-num text-base font-extrabold text-sky-950">
                    {focusedBrandMetrics.simplicity.shareCvPct.toFixed(1)}%
                  </strong>{' '}
                  CV ({focusedBrandMetrics.simplicity.shareGrandPct.toFixed(1)}% Total DP+CV)
                </span>
              </div>
              <div className="w-full h-2.5 bg-slate-100 rounded-full overflow-hidden flex">
                <div
                  style={{
                    width: `${
                      focusedScope.raw.cvSales > 0
                        ? focusedBrandMetrics.graphis.shareCvPct
                        : 50
                    }%`,
                  }}
                  className="bg-violet-600 h-full transition-all"
                />
                <div
                  style={{
                    width: `${
                      focusedScope.raw.cvSales > 0
                        ? focusedBrandMetrics.simplicity.shareCvPct
                        : 50
                    }%`,
                  }}
                  className="bg-sky-600 h-full transition-all"
                />
              </div>
            </div>

            {/* Side-by-Side Brand Performance & Inventory Cards */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              {([
                {
                  stat: focusedBrandMetrics.graphis,
                  accentText: 'text-violet-700',
                  accentDot: 'bg-violet-600',
                },
                {
                  stat: focusedBrandMetrics.simplicity,
                  accentText: 'text-sky-700',
                  accentDot: 'bg-sky-600',
                },
              ] as const).map(({ stat, accentText, accentDot }) => (
                <div
                  key={stat.meta.key}
                  className="border border-slate-200 rounded-xl p-4 space-y-3 bg-white"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div>
                      <div className={`text-xs font-bold ${accentText} flex items-center gap-1.5`}>
                        <span className={`w-2 h-2 rounded-xs ${accentDot}`} aria-hidden="true" />
                        <span>
                          Dept {stat.meta.deptCode} · {stat.meta.divisi}
                        </span>
                      </div>
                      <h3 className="text-base font-extrabold text-slate-950 mt-0.5">
                        Brand {stat.meta.label}
                      </h3>
                    </div>
                    <div className="text-right">
                      {stat.growthPct !== null ? (
                        <div
                          className={`inline-flex items-center gap-0.5 text-sm sm:text-base font-extrabold font-mono-num ${
                            stat.growthPct >= 0 ? 'text-emerald-700' : 'text-rose-700'
                          }`}
                        >
                          {stat.growthPct >= 0 ? (
                            <ArrowUpRight className="w-4 h-4" />
                          ) : (
                            <ArrowDownRight className="w-4 h-4" />
                          )}
                          {stat.growthPct >= 0 ? '+' : ''}
                          {stat.growthPct.toFixed(1)}%{' '}
                          <span className="text-xs font-bold">{focusedScope.prevLabel}</span>
                        </div>
                      ) : (
                        <div className="text-xs sm:text-sm font-semibold text-slate-400">
                          {focusedScope.prevLabel}
                        </div>
                      )}
                      <div className="font-mono-num text-sm font-bold text-slate-700 mt-0.5">
                        Kontribusi{' '}
                        <strong className="text-base font-extrabold text-slate-950">
                          {stat.shareCvPct.toFixed(1)}%
                        </strong>{' '}
                        CV
                      </div>
                    </div>
                  </div>

                  {/* Daily / Period Total Sales & Bazar Share */}
                  <div className="flex flex-wrap items-baseline justify-between gap-2 pb-2.5 border-b border-slate-100">
                    <div className="font-mono-num text-xl font-extrabold text-slate-950">
                      Rp {formatNumberId(stat.sales) || '0'}
                    </div>
                    <div className="flex items-center gap-3 text-xs sm:text-sm text-slate-600 font-mono-num">
                      {horizon !== 'harian' && (
                        <>
                          <span>
                            Avg/Hari:{' '}
                            <strong className="text-slate-900">
                              Rp {formatShortRp(stat.avgDailySales)}
                            </strong>
                          </span>
                          <span aria-hidden="true">·</span>
                        </>
                      )}
                      <span className="font-semibold text-slate-700">
                        Share DP+CV:{' '}
                        <strong className="text-base font-extrabold text-slate-950">
                          {stat.shareGrandPct.toFixed(1)}%
                        </strong>
                      </span>
                    </div>
                  </div>

                  {/* 4 Key Commercial Ratios */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
                    <div className="bg-slate-50 rounded-lg p-2.5">
                      <div className="text-slate-500 font-semibold">Qty Terjual</div>
                      <div className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950 mt-0.5">
                        {formatNumberId(stat.qty) || '0'} pcs
                      </div>
                    </div>
                    <div className="bg-slate-50 rounded-lg p-2.5">
                      <div className="text-slate-500 font-semibold">Jumlah Nota</div>
                      <div className="font-mono-num text-sm sm:text-base font-extrabold text-slate-950 mt-0.5">
                        {formatNumberId(stat.nota) || '0'} nota
                      </div>
                    </div>
                    <div className="bg-indigo-50/70 border border-indigo-100 rounded-lg p-2.5">
                      <div className="text-xs sm:text-sm font-bold text-indigo-800">
                        Price Point
                      </div>
                      <div className="font-mono-num text-base sm:text-lg font-extrabold text-indigo-950 mt-0.5">
                        Rp {formatNumberId(stat.pricePoint) || '0'}
                      </div>
                    </div>
                    <div className="bg-amber-50/70 border border-amber-100 rounded-lg p-2.5">
                      <div className="text-xs sm:text-sm font-bold text-amber-800">
                        Big Size
                      </div>
                      <div className="font-mono-num text-base sm:text-lg font-extrabold text-amber-950 mt-0.5">
                        Rp {formatNumberId(stat.bigSize) || '0'}
                      </div>
                    </div>
                  </div>

                  {/* Per-Brand Stock & Ops Flow */}
                  <div className="pt-2 border-t border-slate-100 space-y-2">
                    <div className="flex items-center justify-between text-[11px] font-bold text-slate-500">
                      <span>Mutasi Stok Brand ({brandOpsComparison.opsPeriodLabel})</span>
                      <span className="font-mono-num text-slate-700">
                        Est. Akhir: {formatNumberId(stat.ops.stockAkhir) || '0'} pcs
                      </span>
                    </div>
                    <div className="grid grid-cols-4 gap-2 text-xs font-mono-num">
                      <div className="bg-slate-50 rounded-md px-2 py-1.5">
                        <div className="text-[11px] font-sans text-slate-500">Awal</div>
                        <div className="font-bold text-slate-900">
                          {formatNumberId(stat.ops.stockAwal) || '0'}
                        </div>
                      </div>
                      <div className="bg-slate-50 rounded-md px-2 py-1.5">
                        <div className="text-[11px] font-sans text-slate-500">DO Masuk</div>
                        <div className="font-bold text-emerald-700">
                          +{formatNumberId(stat.ops.barangDatang) || '0'}
                        </div>
                      </div>
                      <div className="bg-slate-50 rounded-md px-2 py-1.5">
                        <div className="text-[11px] font-sans text-slate-500">Sales/Retur</div>
                        <div className="font-bold text-rose-700">
                          -{formatNumberId(stat.ops.salesQty + stat.ops.barangKeluar) || '0'}
                        </div>
                      </div>
                      <div className="bg-slate-50 rounded-md px-2 py-1.5">
                        <div className="text-[11px] font-sans text-slate-500">Akhir</div>
                        <div className="font-bold text-slate-950">
                          {formatNumberId(stat.ops.stockAkhir) || '0'}
                        </div>
                      </div>
                    </div>
                    <div className="text-xs text-slate-600 space-y-0.5 pt-1">
                      <div className="truncate" title={stat.ops.promo}>
                        <strong className="text-slate-800">Promo:</strong> {stat.ops.promo}
                      </div>
                      <div className="truncate" title={stat.ops.reasonKendala}>
                        <strong className="text-slate-800">Reason/Kendala:</strong>{' '}
                        {stat.ops.reasonKendala}
                      </div>
                      <div className="truncate" title={stat.ops.actionPlan}>
                        <strong className="text-slate-800">Action Plan:</strong>{' '}
                        {stat.ops.actionPlan}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>

            {/* Divisi 1–4 Merchandise & Promo Notes from Review Monitoring (when populated) */}
            {reviewHighlights.divisionNotes.length > 0 && (
              <div className="pt-3 border-t border-slate-200 space-y-2">
                <div className="flex items-center justify-between gap-2 text-xs">
                  <span className="font-bold text-slate-800">
                    Catatan Merchandise &amp; Promo Per Divisi ({reviewHighlights.docTitle})
                  </span>
                  <button
                    type="button"
                    onClick={() => onNavigate('review_monitoring')}
                    className="text-xs font-semibold text-slate-600 hover:text-slate-950 cursor-pointer"
                  >
                    Kelola di Monitoring &rarr;
                  </button>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-2.5">
                  {reviewHighlights.divisionNotes.map((dn) => (
                    <div
                      key={dn.divLabel}
                      className="bg-slate-50 rounded-lg p-2.5 text-xs space-y-1"
                    >
                      <div className="font-bold text-slate-900">{dn.divLabel}</div>
                      {dn.merchItems.length > 0 && (
                        <div className="text-slate-600">
                          <span className="font-semibold text-slate-700">Merch:</span>{' '}
                          {dn.merchItems.join('; ')}
                        </div>
                      )}
                      {dn.promoItems.length > 0 && (
                        <div className="text-slate-600">
                          <span className="font-semibold text-slate-700">Promo:</span>{' '}
                          {dn.promoItems.join('; ')}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        ) : (
          <div className="overflow-x-auto spreadsheet-scroll">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Periode</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-800">
                    Graphis Sales (6201)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Qty
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Share
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-800">
                    Simplicity Sales (6244)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-700">
                    Qty
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-700">
                    Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-700">
                    Share
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-slate-900">
                    Total CV
                  </th>
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Dominan</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {(() => {
                  const brandRows: {
                    key: string;
                    label: string;
                    isSelected: boolean;
                    onClick: () => void;
                    graphisSales: number;
                    graphisQty: number;
                    graphisNota: number;
                    simplicitySales: number;
                    simplicityQty: number;
                    simplicityNota: number;
                    cvSales: number;
                  }[] =
                    horizon === 'harian'
                      ? activeWeekDays.map((d, idx) => ({
                          key: d.tanggal,
                          label: `${d.dayName.slice(0, 3)}, ${d.tanggal}`,
                          isSelected: idx === safeDayIdx,
                          onClick: () => setSelectedDayIdx(idx),
                          graphisSales: d.graphisSales,
                          graphisQty: d.graphisQty,
                          graphisNota: d.graphisNota,
                          simplicitySales: d.simplicitySales,
                          simplicityQty: d.simplicityQty,
                          simplicityNota: d.simplicityNota,
                          cvSales: d.cvSales,
                        }))
                      : horizon === 'weekly'
                        ? visibleWeeklyTableRows.map((w) => ({
                            key: `wk-${w.weekNumber}`,
                            label: `W${w.weekNumber} (${w.shortRange})`,
                            isSelected: w.weekNumber === selectedWeekNumber,
                            onClick: () => handleSelectWeekNumber(w.weekNumber),
                            graphisSales: w.graphisSales,
                            graphisQty: w.graphisQty,
                            graphisNota: w.graphisNota,
                            simplicitySales: w.simplicitySales,
                            simplicityQty: w.simplicityQty,
                            simplicityNota: w.simplicityNota,
                            cvSales: w.cvSales,
                          }))
                        : monthlyTableMode === '12_bulan' || !activeMonthSummary
                          ? all12MonthsConsolidated.map((m) => ({
                              key: `m-${m.monthKey}`,
                              label: `${m.monthName} 2026`,
                              isSelected: m.monthKey === selectedMonthKey,
                              onClick: () => handleSelectMonthKey(m.monthKey),
                              graphisSales: m.graphisSales,
                              graphisQty: m.graphisQty,
                              graphisNota: m.graphisNota,
                              simplicitySales: m.simplicitySales,
                              simplicityQty: m.simplicityQty,
                              simplicityNota: m.simplicityNota,
                              cvSales: m.cvSales,
                            }))
                          : monthlyTableMode === 'per_week'
                            ? activeMonthSummary.monthWeeks.map((w) => ({
                                key: `mwk-${w.weekNumber}`,
                                label: `W${w.weekNumber} (${w.periodeText})`,
                                isSelected: w.weekNumber === selectedWeekNumber,
                                onClick: () => handleSelectWeekNumber(w.weekNumber),
                                graphisSales: w.graphisSales,
                                graphisQty: w.graphisQty,
                                graphisNota: w.graphisNota,
                                simplicitySales: w.simplicitySales,
                                simplicityQty: w.simplicityQty,
                                simplicityNota: w.simplicityNota,
                                cvSales: w.cvSales,
                              }))
                            : activeMonthSummary.days.map((d) => ({
                                key: d.tanggal,
                                label: `${d.dayName.slice(0, 3)}, ${d.tanggal}`,
                                isSelected: false,
                                onClick: () => {},
                                graphisSales: d.graphisSales,
                                graphisQty: d.graphisQty,
                                graphisNota: d.graphisNota,
                                simplicitySales: d.simplicitySales,
                                simplicityQty: d.simplicityQty,
                                simplicityNota: d.simplicityNota,
                                cvSales: d.cvSales,
                              }));

                  const totG = brandRows.reduce((s, r) => s + r.graphisSales, 0);
                  const totGQ = brandRows.reduce((s, r) => s + r.graphisQty, 0);
                  const totGN = brandRows.reduce((s, r) => s + r.graphisNota, 0);
                  const totS = brandRows.reduce((s, r) => s + r.simplicitySales, 0);
                  const totSQ = brandRows.reduce((s, r) => s + r.simplicityQty, 0);
                  const totSN = brandRows.reduce((s, r) => s + r.simplicityNota, 0);
                  const totCv = brandRows.reduce((s, r) => s + r.cvSales, 0);

                  return (
                    <>
                      {brandRows.map((row) => {
                        const gShare =
                          row.cvSales > 0 ? (row.graphisSales / row.cvSales) * 100 : 0;
                        const sShare =
                          row.cvSales > 0 ? (row.simplicitySales / row.cvSales) * 100 : 0;
                        const dom =
                          row.cvSales <= 0
                            ? '-'
                            : Math.abs(row.graphisSales - row.simplicitySales) < 1000
                              ? 'Seimbang'
                              : row.graphisSales > row.simplicitySales
                                ? 'Graphis'
                                : 'Simplicity';
                        return (
                          <tr
                            key={row.key}
                            onClick={row.onClick}
                            className={`h-10 cursor-pointer transition-colors ${
                              row.isSelected ? 'bg-violet-50/60 font-semibold' : 'hover:bg-slate-50'
                            }`}
                          >
                            <td className="py-2 px-3 font-mono-num font-bold text-slate-900 whitespace-nowrap">
                              {row.label}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-900 whitespace-nowrap">
                              {row.graphisSales > 0
                                ? `Rp ${formatNumberId(row.graphisSales)}`
                                : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-600 whitespace-nowrap">
                              {row.graphisQty > 0 ? formatNumberId(row.graphisQty) : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-600 whitespace-nowrap">
                              {row.graphisNota > 0 ? formatNumberId(row.graphisNota) : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-sm font-extrabold text-violet-700 whitespace-nowrap">
                              {row.cvSales > 0 ? `${gShare.toFixed(1)}%` : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-900 whitespace-nowrap">
                              {row.simplicitySales > 0
                                ? `Rp ${formatNumberId(row.simplicitySales)}`
                                : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-600 whitespace-nowrap">
                              {row.simplicityQty > 0 ? formatNumberId(row.simplicityQty) : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-600 whitespace-nowrap">
                              {row.simplicityNota > 0 ? formatNumberId(row.simplicityNota) : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-sm font-extrabold text-sky-700 whitespace-nowrap">
                              {row.cvSales > 0 ? `${sShare.toFixed(1)}%` : '-'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num font-bold text-slate-950 whitespace-nowrap">
                              {row.cvSales > 0 ? `Rp ${formatNumberId(row.cvSales)}` : '-'}
                            </td>
                            <td className="py-2 px-3 text-left font-semibold text-slate-700 whitespace-nowrap">
                              {dom}
                            </td>
                          </tr>
                        );
                      })}
                      <tr className="bg-slate-900 text-white font-bold h-10">
                        <td className="py-2 px-3 whitespace-nowrap">
                          TOTAL ({brandRows.length} baris)
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totG) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totGQ) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totGN) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-sm font-extrabold whitespace-nowrap">
                          {totCv > 0 ? `${((totG / totCv) * 100).toFixed(1)}%` : '0.0%'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totS) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totSQ) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totSN) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-sm font-extrabold whitespace-nowrap">
                          {totCv > 0 ? `${((totS / totCv) * 100).toFixed(1)}%` : '0.0%'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                          Rp {formatNumberId(totCv) || '0'}
                        </td>
                        <td className="py-2 px-3 text-left whitespace-nowrap">
                          {totCv <= 0
                            ? '-'
                            : totG > totS
                              ? 'Graphis'
                              : totS > totG
                                ? 'Simplicity'
                                : 'Seimbang'}
                        </td>
                      </tr>
                    </>
                  );
                })()}
              </tbody>
            </table>
          </div>
        )}
      </section>

      {/* =====================================================================
          5. HIGH-DENSITY LEDGER TABLE (Daily, Weekly, or Monthly Multi-Tab)
         ===================================================================== */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div>
            <h2 className="text-sm sm:text-base font-extrabold text-slate-950">
              {horizon === 'harian' &&
                `Rincian Harian ${reportStream === 'CV' ? 'Brand CV (Graphis & Simplicity)' : 'DP + CV'} — Weekly ${selectedWeekNumber} (${activeWeekSummary.periodeText})`}
              {horizon === 'weekly' &&
                `Perbandingan Performa Weekly (${reportStream === 'CV' ? 'Graphis 6201 · Simplicity 6244' : 'DP + CV + Traffic'})`}
              {horizon === 'bulanan' &&
                (monthlyTableMode === '12_bulan'
                  ? `Rekap Kalender 12 Bulan 2026 (${reportStream === 'CV' ? 'Per Brand CV' : 'Januari – Desember'})`
                  : monthlyTableMode === 'per_week'
                    ? `Rekap Per Week — ${activeMonthSummary?.monthName || 'November'} 2026`
                    : `Rekap Harian — ${activeMonthSummary?.monthName || 'November'} 2026`)}
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Seluruh nominal dalam Rupiah · Angka rata kanan tabular-nums · Satuan: Trx (DP), Nota (CV), Trx+Nota (Total)
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {horizon === 'bulanan' && (
              <div className="inline-flex items-center p-0.5 bg-slate-100 border border-slate-200 rounded-lg text-xs">
                {(
                  [
                    { id: 'harian', label: 'Harian' },
                    { id: 'per_week', label: 'Per Week' },
                    { id: '12_bulan', label: '12 Bulan' },
                  ] as const
                ).map((t) => (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => setMonthlyTableMode(t.id)}
                    className={`px-2.5 py-1 rounded-md font-bold cursor-pointer whitespace-nowrap ${
                      monthlyTableMode === t.id
                        ? 'bg-white text-slate-950 shadow-2xs'
                        : 'text-slate-600 hover:text-slate-950'
                    }`}
                  >
                    {t.label}
                  </button>
                ))}
              </div>
            )}

            {(horizon === 'harian' ||
              (horizon === 'bulanan' && monthlyTableMode === 'harian')) && (
              <>
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={tableSearch}
                    onChange={(e) => setTableSearch(e.target.value)}
                    placeholder="Cari tgl / staff / catatan..."
                    className="pl-8 pr-2.5 py-1.5 text-xs border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:border-slate-400 w-44"
                  />
                </div>
                <select
                  aria-label="Filter Tipe Hari"
                  value={dayTypeFilter}
                  onChange={(e) =>
                    setDayTypeFilter(e.target.value as 'ALL' | 'WEEKEND' | 'WEEKDAY')
                  }
                  className="px-2.5 py-1.5 text-xs font-semibold border border-slate-200 rounded-lg bg-slate-50 text-slate-800 cursor-pointer"
                >
                  <option value="ALL">Semua Hari</option>
                  <option value="WEEKEND">Weekend (Jum–Min)</option>
                  <option value="WEEKDAY">Weekday (Sen–Kam)</option>
                </select>
                {horizon === 'bulanan' && (
                  <button
                    type="button"
                    onClick={() => setShowOnlyActiveRows((v) => !v)}
                    className={`px-2.5 py-1.5 text-xs font-semibold rounded-lg border cursor-pointer whitespace-nowrap ${
                      showOnlyActiveRows
                        ? 'bg-slate-900 text-white border-slate-900'
                        : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                    }`}
                  >
                    {showOnlyActiveRows ? 'Hanya Terisi' : 'Tampilkan Terisi'}
                  </button>
                )}
              </>
            )}
          </div>
        </div>

        {/* TABLE BODY */}
          <div className="overflow-x-auto spreadsheet-scroll">
          {(horizon === 'harian' ||
            (horizon === 'bulanan' && monthlyTableMode === 'harian')) && (
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Tanggal</th>
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Hari</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-800">
                    Sales DP
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-700">
                    Trx
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-800">
                    Graphis (6201)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-800">
                    Simplicity (6244)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-900">
                    Sales CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Total DP+CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Qty</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Trx+Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Traffic (P/S)</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Conv Rate</th>
                </tr>
              </thead>
              {(() => {
                const sourceDays =
                  horizon === 'harian'
                    ? activeWeekDays
                    : activeMonthSummary
                      ? activeMonthSummary.days
                      : all12MonthsConsolidated.flatMap((m) =>
                          m.days.filter((x) => x.grandSales > 0 || x.trafficTotal > 0)
                        );
                const rows = sourceDays.filter(filterDayRow);
                const totDp = rows.reduce((s, r) => s + r.dpSales, 0);
                const totTrx = rows.reduce((s, r) => s + r.dpTrx, 0);
                const totG = rows.reduce((s, r) => s + r.graphisSales, 0);
                const totS = rows.reduce((s, r) => s + r.simplicitySales, 0);
                const totCv = rows.reduce((s, r) => s + r.cvSales, 0);
                const totNota = rows.reduce((s, r) => s + r.cvNota, 0);
                const totGrand = rows.reduce((s, r) => s + r.grandSales, 0);
                const totQty = rows.reduce((s, r) => s + r.totalQty, 0);
                const totTN = rows.reduce((s, r) => s + r.totalTrxNota, 0);
                const totTrf = rows.reduce((s, r) => s + r.trafficTotal, 0);

                return (
                  <>
                    <tbody className="divide-y divide-slate-100">
                      {rows.length === 0 ? (
                        <tr>
                          <td colSpan={13} className="py-8 text-center text-slate-500">
                            Belum ada data pada filter yang dipilih.
                          </td>
                        </tr>
                      ) : (
                        rows.map((d) => {
                          const isCurrentFocusedDay =
                            horizon === 'harian' && d.tanggal === activeDay.tanggal;
                          return (
                            <tr
                              key={d.tanggal}
                              onClick={() => {
                                if (horizon === 'harian') {
                                  const i = activeWeekDays.findIndex(
                                    (x) => x.tanggal === d.tanggal
                                  );
                                  if (i >= 0) setSelectedDayIdx(i);
                                }
                              }}
                              className={`h-10 transition-colors ${
                                horizon === 'harian' ? 'cursor-pointer' : ''
                              } ${
                                isCurrentFocusedDay
                                  ? 'bg-emerald-50/70 font-semibold'
                                  : d.isWeekend
                                    ? 'bg-rose-50/25 hover:bg-slate-50'
                                    : 'hover:bg-slate-50'
                              }`}
                            >
                              <td className="py-2 px-3 font-mono-num font-bold text-slate-900 whitespace-nowrap">
                                {d.tanggal}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap text-slate-700">
                                {d.dayName}
                                {d.isWeekend && (
                                  <span className="ml-1.5 text-[11px] font-semibold text-rose-700">
                                    (Wknd)
                                  </span>
                                )}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-indigo-950 whitespace-nowrap">
                                {d.dpSales > 0 ? `Rp ${formatNumberId(d.dpSales)}` : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                                {d.dpTrx > 0 ? formatNumberId(d.dpTrx) : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-violet-900 whitespace-nowrap">
                                {d.graphisSales > 0 ? `Rp ${formatNumberId(d.graphisSales)}` : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-sky-900 whitespace-nowrap">
                                {d.simplicitySales > 0 ? `Rp ${formatNumberId(d.simplicitySales)}` : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num font-semibold text-violet-950 whitespace-nowrap">
                                {d.cvSales > 0 ? `Rp ${formatNumberId(d.cvSales)}` : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                                {d.cvNota > 0 ? formatNumberId(d.cvNota) : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num font-bold text-slate-950 whitespace-nowrap">
                                {d.grandSales > 0 ? `Rp ${formatNumberId(d.grandSales)}` : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                                {d.totalQty > 0 ? formatNumberId(d.totalQty) : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num font-bold text-emerald-800 whitespace-nowrap">
                                {d.totalTrxNota > 0 ? formatNumberId(d.totalTrxNota) : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                                {d.trafficTotal > 0
                                  ? `${formatNumberId(d.trafficTotal)} (${d.pagiTraffic}/${d.siangTraffic})`
                                  : '-'}
                              </td>
                              <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-emerald-800 whitespace-nowrap">
                                {d.trafficTotal > 0
                                  ? `${((d.totalTrxNota / d.trafficTotal) * 100).toFixed(1)}%`
                                  : '-'}
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                    {rows.length > 0 && (
                      <tfoot>
                        <tr className="bg-slate-900 text-white font-bold h-10">
                          <td className="py-2 px-3 whitespace-nowrap" colSpan={2}>
                            TOTAL ({rows.length} hari)
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            Rp {formatNumberId(totDp) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            {formatNumberId(totTrx) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            Rp {formatNumberId(totG) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            Rp {formatNumberId(totS) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            Rp {formatNumberId(totCv) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            {formatNumberId(totNota) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                            Rp {formatNumberId(totGrand) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            {formatNumberId(totQty) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                            {formatNumberId(totTN) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                            {formatNumberId(totTrf) || '0'}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-emerald-300 whitespace-nowrap">
                            {totTrf > 0 ? `${((totTN / totTrf) * 100).toFixed(1)}%` : '0.0%'}
                          </td>
                        </tr>
                      </tfoot>
                    )}
                  </>
                );
              })()}
            </table>
          )}

          {(horizon === 'weekly' ||
            (horizon === 'bulanan' && monthlyTableMode === 'per_week')) && (
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Week</th>
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Periode Tanggal</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-800">
                    Sales DP
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-700">
                    Trx
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-800">
                    Graphis (6201)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-800">
                    Simplicity (6244)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-900">
                    Sales CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Total DP+CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Trx+Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Traffic</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Conv Rate</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Growth</th>
                </tr>
              </thead>
              {(() => {
                const list =
                  horizon === 'bulanan' && activeMonthSummary
                    ? activeMonthSummary.monthWeeks
                    : visibleWeeklyTableRows;
                const totDp = list.reduce((s, r) => s + r.dpSales, 0);
                const totTrx = list.reduce((s, r) => s + r.dpTrx, 0);
                const totG = list.reduce((s, r) => s + r.graphisSales, 0);
                const totS = list.reduce((s, r) => s + r.simplicitySales, 0);
                const totCv = list.reduce((s, r) => s + r.cvSales, 0);
                const totNota = list.reduce((s, r) => s + r.cvNota, 0);
                const totGrand = list.reduce((s, r) => s + r.grandSales, 0);
                const totTN = list.reduce((s, r) => s + r.totalTrxNota, 0);
                const totTrf = list.reduce((s, r) => s + r.trafficTotal, 0);

                return (
                  <>
                    <tbody className="divide-y divide-slate-100">
                      {list.map((w, idx) => {
                        const prev = idx > 0 ? list[idx - 1] : null;
                        const currVal = projectStream(w).sales;
                        const prevVal = prev ? projectStream(prev).sales : null;
                        const growth = calcGrowthPct(currVal, prevVal);
                        const conv =
                          w.trafficTotal > 0 ? (w.totalTrxNota / w.trafficTotal) * 100 : 0;
                        const isSel = w.weekNumber === selectedWeekNumber;
                        return (
                          <tr
                            key={w.weekNumber}
                            onClick={() => handleSelectWeekNumber(w.weekNumber)}
                            className={`h-10 cursor-pointer transition-colors ${
                              isSel ? 'bg-emerald-50/70 font-semibold' : 'hover:bg-slate-50'
                            }`}
                          >
                            <td className="py-2 px-3 font-mono-num font-bold text-slate-950 whitespace-nowrap">
                              W{w.weekNumber} · {w.monthName.slice(0, 3)}
                            </td>
                            <td className="py-2 px-3 text-slate-600 whitespace-nowrap">
                              {w.periodeText}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-indigo-950 whitespace-nowrap">
                              Rp {formatNumberId(w.dpSales) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                              {formatNumberId(w.dpTrx) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-violet-900 whitespace-nowrap">
                              Rp {formatNumberId(w.graphisSales) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-sky-900 whitespace-nowrap">
                              Rp {formatNumberId(w.simplicitySales) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num font-semibold text-violet-950 whitespace-nowrap">
                              Rp {formatNumberId(w.cvSales) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                              {formatNumberId(w.cvNota) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num font-bold text-slate-950 whitespace-nowrap">
                              Rp {formatNumberId(w.grandSales) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num font-bold text-emerald-800 whitespace-nowrap">
                              {formatNumberId(w.totalTrxNota) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                              {formatNumberId(w.trafficTotal) || '0'}
                            </td>
                            <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-emerald-800 whitespace-nowrap">
                              {conv.toFixed(1)}%
                            </td>
                            <td
                              className={`py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold whitespace-nowrap ${
                                growth === null
                                  ? 'text-slate-400'
                                  : growth >= 0
                                    ? 'text-emerald-700'
                                    : 'text-rose-700'
                              }`}
                            >
                              {growth === null
                                ? '-'
                                : `${growth >= 0 ? '+' : ''}${growth.toFixed(1)}%`}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr className="bg-slate-900 text-white font-bold h-10">
                        <td className="py-2 px-3 whitespace-nowrap" colSpan={2}>
                          TOTAL ({list.length} weekly)
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totDp) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totTrx) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totG) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totS) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          Rp {formatNumberId(totCv) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totNota) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                          Rp {formatNumberId(totGrand) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                          {formatNumberId(totTN) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                          {formatNumberId(totTrf) || '0'}
                        </td>
                        <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-emerald-300 whitespace-nowrap">
                          {totTrf > 0 ? `${((totTN / totTrf) * 100).toFixed(1)}%` : '0.0%'}
                        </td>
                        <td className="py-2 px-3 text-right whitespace-nowrap">-</td>
                      </tr>
                    </tfoot>
                  </>
                );
              })()}
            </table>
          )}

          {horizon === 'bulanan' && monthlyTableMode === '12_bulan' && (
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-bold">
                  <th className="py-2.5 px-3 text-left whitespace-nowrap">Bulan</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Aktif</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-800">
                    Sales DP
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-indigo-700">
                    Trx
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-800">
                    Graphis (6201)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-sky-800">
                    Simplicity (6244)
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-900">
                    Sales CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-violet-700">
                    Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Total DP+CV
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap text-emerald-800">
                    Trx+Nota
                  </th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Traffic</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">Share %</th>
                  <th className="py-2.5 px-3 text-right whitespace-nowrap">MoM Growth</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {all12MonthsConsolidated.map((m, idx) => {
                  const prev = idx > 0 ? all12MonthsConsolidated[idx - 1] : null;
                  const currVal = projectStream(m).sales;
                  const prevVal = prev ? projectStream(prev).sales : null;
                  const totalAnnVal = projectStream(annualGrandTotals).sales;
                  const sharePct = totalAnnVal > 0 ? (currVal / totalAnnVal) * 100 : 0;
                  const growth = calcGrowthPct(currVal, prevVal);
                  const isSel = selectedMonthKey === m.monthKey;
                  return (
                    <tr
                      key={m.monthKey}
                      onClick={() => handleSelectMonthKey(m.monthKey)}
                      className={`h-10 cursor-pointer transition-colors ${
                        isSel ? 'bg-emerald-50/70 font-semibold' : 'hover:bg-slate-50'
                      }`}
                    >
                      <td className="py-2 px-3 font-bold text-slate-950 whitespace-nowrap">
                        {m.monthName} 2026
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-slate-600 whitespace-nowrap">
                        {m.activeDays}/{m.lastDay}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-indigo-950 whitespace-nowrap">
                        Rp {formatNumberId(m.dpSales) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                        {formatNumberId(m.dpTrx) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-violet-900 whitespace-nowrap">
                        Rp {formatNumberId(m.graphisSales) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-sky-900 whitespace-nowrap">
                        Rp {formatNumberId(m.simplicitySales) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num font-semibold text-violet-950 whitespace-nowrap">
                        Rp {formatNumberId(m.cvSales) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                        {formatNumberId(m.cvNota) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num font-bold text-slate-950 whitespace-nowrap">
                        Rp {formatNumberId(m.grandSales) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num font-bold text-emerald-800 whitespace-nowrap">
                        {formatNumberId(m.totalTrxNota) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-slate-700 whitespace-nowrap">
                        {formatNumberId(m.trafficTotal) || '0'}
                      </td>
                      <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-sky-800 whitespace-nowrap">
                        {sharePct.toFixed(1)}%
                      </td>
                      <td
                        className={`py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold whitespace-nowrap ${
                          growth === null
                            ? 'text-slate-400'
                            : growth >= 0
                              ? 'text-emerald-700'
                              : 'text-rose-700'
                        }`}
                      >
                        {growth === null
                          ? '-'
                          : `${growth >= 0 ? '+' : ''}${growth.toFixed(1)}%`}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="bg-slate-900 text-white font-bold h-10">
                  <td className="py-2 px-3 whitespace-nowrap" colSpan={2}>
                    GRAND TOTAL 2026 ({annualGrandTotals.activeDays} hari)
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    Rp {formatNumberId(annualGrandTotals.dpSales) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    {formatNumberId(annualGrandTotals.dpTrx) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    Rp {formatNumberId(annualGrandTotals.graphisSales) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    Rp {formatNumberId(annualGrandTotals.simplicitySales) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    Rp {formatNumberId(annualGrandTotals.cvSales) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    {formatNumberId(annualGrandTotals.cvNota) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                    Rp {formatNumberId(annualGrandTotals.grandSales) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num text-emerald-300 whitespace-nowrap">
                    {formatNumberId(annualGrandTotals.totalTrxNota) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num whitespace-nowrap">
                    {formatNumberId(annualGrandTotals.trafficTotal) || '0'}
                  </td>
                  <td className="py-2 px-3 text-right font-mono-num text-sm sm:text-base font-extrabold text-sky-300 whitespace-nowrap">
                    100.0%
                  </td>
                  <td className="py-2 px-3 text-right whitespace-nowrap">-</td>
                </tr>
              </tfoot>
            </table>
          )}
        </div>
      </section>

      {/* =====================================================================
          6. OPERATIONAL SNAPSHOT: DATA CV INVENTORY & MONITORING REVIEW
         ===================================================================== */}
      <section className="grid grid-cols-1 lg:grid-cols-2 gap-4 items-stretch">
        {/* Left: Data CV Inventory & Promo / Action Plan */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-3">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <Layers className="w-4 h-4 text-slate-700" />
              <h3 className="text-sm font-bold text-slate-950">
                Data CV — Inventory &amp; Action Plan ({activeCvOpsSnapshot.opsPeriodLabel})
              </h3>
            </div>
            <button
              type="button"
              onClick={() => onNavigate('data_cv_ops')}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
            >
              <span>Buka Data CV</span>
              <ArrowUpRight className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs pt-1">
            <div className="bg-slate-50 rounded-lg p-2.5">
              <div className="text-slate-500">Stock Awal</div>
              <div className="font-mono-num font-bold text-slate-950 text-sm mt-0.5">
                {formatNumberId(activeCvOpsSnapshot.totalStockAwal) || '0'} pcs
              </div>
            </div>
            <div className="bg-slate-50 rounded-lg p-2.5">
              <div className="text-slate-500">DO Masuk</div>
              <div className="font-mono-num font-bold text-emerald-700 text-sm mt-0.5">
                +{formatNumberId(activeCvOpsSnapshot.totalMasuk) || '0'} pcs
              </div>
            </div>
            <div className="bg-slate-50 rounded-lg p-2.5">
              <div className="text-slate-500">Sales + Retur</div>
              <div className="font-mono-num font-bold text-rose-700 text-sm mt-0.5">
                -{formatNumberId(activeCvOpsSnapshot.totalSalesQty + activeCvOpsSnapshot.totalKeluar) || '0'} pcs
              </div>
            </div>
            <div className="bg-slate-50 rounded-lg p-2.5">
              <div className="text-slate-500">Est. Stock Akhir</div>
              <div className="font-mono-num font-bold text-slate-950 text-sm mt-0.5">
                {formatNumberId(activeCvOpsSnapshot.stockAkhir) || '0'} pcs
              </div>
            </div>
          </div>

          <div className="text-xs text-slate-600 space-y-1 pt-2 border-t border-slate-100">
            <div>
              <strong className="text-slate-800">Promo Aktif:</strong> {activeCvOpsSnapshot.promo}
            </div>
            <div>
              <strong className="text-slate-800">Action Plan:</strong> {activeCvOpsSnapshot.actionPlan}
            </div>
          </div>
        </div>

        {/* Right: Review Monitoring Bazar & Evaluasi Reason */}
        <div className="bg-white border border-slate-200 rounded-xl p-4 flex flex-col justify-between space-y-3">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <ClipboardCheck className="w-4 h-4 text-slate-700" />
              <h3 className="text-sm font-bold text-slate-950">
                Monitoring &amp; Evaluasi ({reviewHighlights.docTitle})
              </h3>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => onNavigate('reason_evaluation')}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Reason</span>
              </button>
              <button
                type="button"
                onClick={() => onNavigate('review_monitoring')}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-700 hover:bg-slate-50 border border-slate-200 rounded-lg cursor-pointer whitespace-nowrap"
              >
                <span>Monitoring</span>
                <ArrowUpRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            <div className="bg-slate-50 rounded-lg p-2.5 space-y-1">
              <div className="font-bold text-slate-800">Fast Moving Items</div>
              {reviewHighlights.fast.length > 0 ? (
                <ul className="space-y-0.5 text-slate-600 list-disc list-inside">
                  {reviewHighlights.fast.slice(0, 3).map((item, i) => (
                    <li key={i} className="truncate">
                      {item}
                    </li>
                  ))}
                </ul>
              ) : (
                <div className="text-slate-400">Belum ada catatan item fast moving.</div>
              )}
            </div>
            <div className="bg-slate-50 rounded-lg p-2.5 space-y-1">
              <div className="font-bold text-slate-800">Slow Moving / Action Plan</div>
              {reviewHighlights.slow.length > 0 || reviewHighlights.actions.length > 0 ? (
                <ul className="space-y-0.5 text-slate-600 list-disc list-inside">
                  {[...reviewHighlights.slow, ...reviewHighlights.actions]
                    .slice(0, 3)
                    .map((item, i) => (
                      <li key={i} className="truncate">
                        {item}
                      </li>
                    ))}
                </ul>
              ) : (
                <div className="text-slate-400">Belum ada catatan slow moving / action plan.</div>
              )}
            </div>
          </div>

          <div className="text-xs text-slate-600 pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
            <span className="truncate">
              <strong className="text-slate-800">Catatan Reason ({activeDay.tanggal}):</strong>{' '}
              {activeDay.reasonPagi || activeDay.reasonSiang || activeDay.cvReason
                ? [
                    activeDay.reasonPagi ? `Pagi: ${activeDay.reasonPagi}` : '',
                    activeDay.reasonSiang ? `Siang: ${activeDay.reasonSiang}` : '',
                    activeDay.cvReason ? `CV: ${activeDay.cvReason}` : '',
                  ]
                    .filter(Boolean)
                    .join(' · ')
                : '-'}
            </span>
            <button
              type="button"
              onClick={() => onNavigate('staff_analytics')}
              className="inline-flex items-center gap-1 text-xs font-semibold text-slate-700 hover:text-slate-950 cursor-pointer shrink-0"
            >
              <Users className="w-3.5 h-3.5" />
              <span>Analitik Petugas</span>
            </button>
          </div>
        </div>
      </section>

      {/* =====================================================================
          7. INTEGRATED HD SCREENSHOT STUDIO MODAL
         ===================================================================== */}
      <DashboardScreenshotModal
        isOpen={isScreenshotModalOpen}
        onClose={() => setIsScreenshotModalOpen(false)}
        initialHorizon={horizon}
        initialStream={reportStream}
        initialMonthlyFocus={monthlyScreenshotFocus}
        initialWeeklyFocus={weeklyScreenshotFocus}
        tables={tables}
        cvRecords={effectiveCvRecords}
        cvInventoryOps={effectiveCvOps}
        initialWeekNumber={selectedWeekNumber}
        initialDayTanggal={activeDay.tanggal}
        initialMonthKey={selectedMonthKey}
        onSelectMonthKey={(mk) => handleSelectMonthKey(mk)}
      />
    </div>
  );
};
