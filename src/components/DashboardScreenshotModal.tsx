import React, { useState, useRef, useEffect, useMemo } from 'react';
import { createPortal } from 'react-dom';
import {
  Camera,
  Download,
  Share2,
  X,
  Calendar,
  CalendarRange,
  CalendarDays,
  Check,
  TrendingUp,
  LayoutDashboard,
  Copy,
  ZoomIn,
  ZoomOut,
  Sparkles,
} from 'lucide-react';
import {
  WeeklySalesTable,
  ANNUAL_2026_WEEKS,
  formatNumberId,
  getIndoDayName,
  isWeekendDate,
} from '../data/initialData';
import {
  CvDailyRecord,
  CvInventoryBrandKey,
  CvInventoryOpsEntry,
  CV_BRAND_CATALOG,
  MONTH_NAMES_ID,
  calcCvLiveSalesQtyMap,
} from '../data/cvDiv1Data';
import { DashboardHorizon, DashboardReportStream } from './DashboardOverview';

export type MonthlyScreenshotFocus = 'ringkasan' | 'kontribusi_growth';
export type WeeklyScreenshotFocus = 'ringkasan' | 'growth_weekly';

export interface DashboardScreenshotModalProps {
  isOpen: boolean;
  onClose: () => void;
  initialHorizon: DashboardHorizon;
  initialStream?: DashboardReportStream;
  initialMonthlyFocus?: MonthlyScreenshotFocus;
  initialWeeklyFocus?: WeeklyScreenshotFocus;
  tables: WeeklySalesTable[];
  cvRecords: CvDailyRecord[];
  cvInventoryOps: CvInventoryOpsEntry[];
  initialWeekNumber: number;
  initialDayTanggal: string;
  initialMonthKey: string; // '01'..'12' or 'ALL'
  onSelectMonthKey?: (monthKey: string) => void;
}

type ThemeStyle = 'navy' | 'emerald' | 'light';

interface ConsolidatedDay {
  tanggal: string;
  dayName: string;
  isWeekend: boolean;
  weekNumber: number;
  monthKey: string;
  dpSales: number;
  dpPagiSales: number;
  dpSiangSales: number;
  dpQty: number;
  dpTrx: number;
  cvSales: number;
  cvQty: number;
  cvNota: number;
  graphisSales: number;
  graphisQty: number;
  graphisNota: number;
  simplicitySales: number;
  simplicityQty: number;
  simplicityNota: number;
  pagiTraffic: number;
  siangTraffic: number;
  trafficTotal: number;
  pagiStaff: string;
  siangStaff: string;
  reasonPagi: string;
  reasonSiang: string;
  cvReason: string;
  grandSales: number;
  totalQty: number;
  totalTrxNota: number;
}

function getMonthLastDayNum(monthKey: string, year = 2026): number {
  const m = parseInt(monthKey, 10) || 11;
  return new Date(year, m, 0).getDate();
}

function fmtShortRp(v: number): string {
  if (v >= 1000000000) return `${(v / 1000000000).toFixed(2)}M`;
  if (v >= 1000000) return `${(v / 1000000).toFixed(1)}jt`;
  if (v >= 1000) return `${Math.round(v / 1000)}rb`;
  return String(Math.round(v || 0));
}

function clipText(ctx: CanvasRenderingContext2D, text: string, maxW: number): string {
  if (!text) return '';
  if (ctx.measureText(text).width <= maxW) return text;
  let s = text;
  while (s.length > 3 && ctx.measureText(s + '...').width > maxW) {
    s = s.slice(0, -1);
  }
  return s + '...';
}

export const DashboardScreenshotModal: React.FC<DashboardScreenshotModalProps> = ({
  isOpen,
  onClose,
  initialHorizon,
  initialStream = 'ALL',
  initialMonthlyFocus = 'ringkasan',
  initialWeeklyFocus = 'ringkasan',
  tables,
  cvRecords,
  cvInventoryOps,
  initialWeekNumber,
  initialDayTanggal,
  initialMonthKey,
  onSelectMonthKey,
}) => {
  const [mode, setMode] = useState<DashboardHorizon>(initialHorizon);
  const [stream, setStream] = useState<DashboardReportStream>(initialStream);
  const [monthlyFocus, setMonthlyFocus] = useState<MonthlyScreenshotFocus>(initialMonthlyFocus);
  const [weeklyFocus, setWeeklyFocus] = useState<WeeklyScreenshotFocus>(initialWeeklyFocus);
  const [theme, setTheme] = useState<ThemeStyle>('navy');
  const [weekNumber, setWeekNumber] = useState<number>(initialWeekNumber || 46);
  const [dayTanggal, setDayTanggal] = useState<string>(initialDayTanggal || '');
  const [monthKey, setMonthKey] = useState<string>(initialMonthKey || '11');
  const [zoomLevel, setZoomLevel] = useState<'fit' | 'full'>('fit');
  const [sharing, setSharing] = useState(false);
  const [copied, setCopied] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  useEffect(() => {
    if (isOpen) {
      setMode(initialHorizon);
      setStream(initialStream);
      setMonthlyFocus(initialMonthlyFocus);
      setWeeklyFocus(initialWeeklyFocus);
      setWeekNumber(initialWeekNumber || 46);
      setDayTanggal(initialDayTanggal || '');
      setMonthKey(initialMonthKey || '11');
    }
  }, [
    isOpen,
    initialHorizon,
    initialStream,
    initialMonthlyFocus,
    initialWeeklyFocus,
    initialWeekNumber,
    initialDayTanggal,
    initialMonthKey,
  ]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  // Unified Lookup Maps across DP + CV + Brands
  const { dpDateMap, cvDateMap, cvLiveQtyMap } = useMemo(() => {
    const dpMap = new Map<
      string,
      {
        sales: number;
        pagiSales: number;
        siangSales: number;
        qty: number;
        trx: number;
        pagiTraffic: number;
        siangTraffic: number;
        pagiStaff: string;
        siangStaff: string;
        reasonPagi: string;
        reasonSiang: string;
        score: number;
      }
    >();

    (tables || []).forEach((t) => {
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
            trx: pT + sT,
            pagiTraffic: pTrf,
            siangTraffic: sTrf,
            pagiStaff: pStaff,
            siangStaff: sStaff,
            reasonPagi: reasonP,
            reasonSiang: reasonS,
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
        graphisQty: number;
        graphisNota: number;
        simplicitySales: number;
        simplicityQty: number;
        simplicityNota: number;
        reasonEvent: string;
      }
    >();

    (cvRecords || []).forEach((r) => {
      const tgl = (r.tanggal || '').trim();
      if (!tgl) return;
      const gSales =
        (Number(r.graphis?.pagi?.sales) || 0) + (Number(r.graphis?.siang?.sales) || 0);
      const gQty =
        (Number(r.graphis?.pagi?.qty) || 0) + (Number(r.graphis?.siang?.qty) || 0);
      const gNota =
        (Number(r.graphis?.pagi?.trx) || 0) + (Number(r.graphis?.siang?.trx) || 0);
      const sSales =
        (Number(r.simplicity?.pagi?.sales) || 0) + (Number(r.simplicity?.siang?.sales) || 0);
      const sQty =
        (Number(r.simplicity?.pagi?.qty) || 0) + (Number(r.simplicity?.siang?.qty) || 0);
      const sNota =
        (Number(r.simplicity?.pagi?.trx) || 0) + (Number(r.simplicity?.siang?.trx) || 0);

      cvMap.set(tgl, {
        sales: gSales + sSales,
        qty: gQty + sQty,
        nota: gNota + sNota,
        graphisSales: gSales,
        graphisQty: gQty,
        graphisNota: gNota,
        simplicitySales: sSales,
        simplicityQty: sQty,
        simplicityNota: sNota,
        reasonEvent: (r.reasonEvent || '').trim(),
      });
    });

    return {
      dpDateMap: dpMap,
      cvDateMap: cvMap,
      cvLiveQtyMap: calcCvLiveSalesQtyMap(cvRecords || []),
    };
  }, [tables, cvRecords]);

  const buildDay = (tanggal: string, wkNum: number): ConsolidatedDay => {
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
    const pTrf = dp?.pagiTraffic ?? 0;
    const sTrf = dp?.siangTraffic ?? 0;

    return {
      tanggal,
      dayName: getIndoDayName(tanggal),
      isWeekend: isWeekendDate(tanggal),
      weekNumber: wkNum,
      monthKey: mKey,
      dpSales,
      dpPagiSales: dp?.pagiSales ?? 0,
      dpSiangSales: dp?.siangSales ?? 0,
      dpQty,
      dpTrx,
      cvSales,
      cvQty,
      cvNota,
      graphisSales: cv?.graphisSales ?? 0,
      graphisQty: cv?.graphisQty ?? 0,
      graphisNota: cv?.graphisNota ?? 0,
      simplicitySales: cv?.simplicitySales ?? 0,
      simplicityQty: cv?.simplicityQty ?? 0,
      simplicityNota: cv?.simplicityNota ?? 0,
      pagiTraffic: pTrf,
      siangTraffic: sTrf,
      trafficTotal: pTrf + sTrf,
      pagiStaff: dp?.pagiStaff || '-',
      siangStaff: dp?.siangStaff || '-',
      reasonPagi: dp?.reasonPagi || '',
      reasonSiang: dp?.reasonSiang || '',
      cvReason: cv?.reasonEvent || '',
      grandSales: dpSales + cvSales,
      totalQty: dpQty + cvQty,
      totalTrxNota: dpTrx + cvNota,
    };
  };

  const weekly53 = useMemo(() => {
    return ANNUAL_2026_WEEKS.map((w) => {
      const days = w.slashDates.map((tgl) => buildDay(tgl, w.weekNumber));
      const sum = (fn: (d: ConsolidatedDay) => number) => days.reduce((s, d) => s + fn(d), 0);
      const dpSales = sum((d) => d.dpSales);
      const cvSales = sum((d) => d.cvSales);
      const dpQty = sum((d) => d.dpQty);
      const cvQty = sum((d) => d.cvQty);
      const dpTrx = sum((d) => d.dpTrx);
      const cvNota = sum((d) => d.cvNota);
      const pagiTraffic = sum((d) => d.pagiTraffic);
      const siangTraffic = sum((d) => d.siangTraffic);
      return {
        weekNumber: w.weekNumber,
        monthKey: w.monthKey,
        monthName: w.monthName,
        periodeText: `${w.seninSlash} – ${w.mingguSlash}`,
        shortRange: `${w.seninSlash.slice(0, 5)}–${w.mingguSlash.slice(0, 5)}`,
        days,
        dpSales,
        dpPagiSales: sum((d) => d.dpPagiSales),
        dpSiangSales: sum((d) => d.dpSiangSales),
        dpQty,
        dpTrx,
        cvSales,
        cvQty,
        cvNota,
        graphisSales: sum((d) => d.graphisSales),
        graphisQty: sum((d) => d.graphisQty),
        graphisNota: sum((d) => d.graphisNota),
        simplicitySales: sum((d) => d.simplicitySales),
        simplicityQty: sum((d) => d.simplicityQty),
        simplicityNota: sum((d) => d.simplicityNota),
        pagiTraffic,
        siangTraffic,
        trafficTotal: pagiTraffic + siangTraffic,
        grandSales: dpSales + cvSales,
        totalQty: dpQty + cvQty,
        totalTrxNota: dpTrx + cvNota,
        activeDays: days.filter((d) => d.grandSales > 0 || d.trafficTotal > 0).length,
      };
    });
  }, [dpDateMap, cvDateMap]);

  const months12 = useMemo(() => {
    return Object.keys(MONTH_NAMES_ID)
      .sort()
      .map((mk) => {
        const lastDay = getMonthLastDayNum(mk, 2026);
        const days: ConsolidatedDay[] = [];
        for (let d = 1; d <= lastDay; d++) {
          const tgl = `${String(d).padStart(2, '0')}/${mk}/2026`;
          days.push(buildDay(tgl, 0));
        }
        const sum = (fn: (d: ConsolidatedDay) => number) => days.reduce((s, d) => s + fn(d), 0);
        const dpSales = sum((d) => d.dpSales);
        const cvSales = sum((d) => d.cvSales);
        const dpQty = sum((d) => d.dpQty);
        const cvQty = sum((d) => d.cvQty);
        const dpTrx = sum((d) => d.dpTrx);
        const cvNota = sum((d) => d.cvNota);
        const pagiTraffic = sum((d) => d.pagiTraffic);
        const siangTraffic = sum((d) => d.siangTraffic);
        return {
          monthKey: mk,
          monthName: MONTH_NAMES_ID[mk] || `Bulan ${mk}`,
          lastDay,
          days,
          dpSales,
          dpPagiSales: sum((d) => d.dpPagiSales),
          dpSiangSales: sum((d) => d.dpSiangSales),
          dpQty,
          dpTrx,
          cvSales,
          cvQty,
          cvNota,
          graphisSales: sum((d) => d.graphisSales),
          graphisQty: sum((d) => d.graphisQty),
          graphisNota: sum((d) => d.graphisNota),
          simplicitySales: sum((d) => d.simplicitySales),
          simplicityQty: sum((d) => d.simplicityQty),
          simplicityNota: sum((d) => d.simplicityNota),
          pagiTraffic,
          siangTraffic,
          trafficTotal: pagiTraffic + siangTraffic,
          grandSales: dpSales + cvSales,
          totalQty: dpQty + cvQty,
          totalTrxNota: dpTrx + cvNota,
          activeDays: days.filter((d) => d.grandSales > 0 || d.trafficTotal > 0).length,
        };
      });
  }, [dpDateMap, cvDateMap]);

  const activeWeekObj =
    weekly53.find((w) => w.weekNumber === weekNumber) || weekly53[45] || weekly53[0];
  const prevWeekObj =
    weekNumber > 1 ? weekly53.find((w) => w.weekNumber === weekNumber - 1) || null : null;

  const activeDayIdx = Math.max(
    0,
    activeWeekObj.days.findIndex((d) => d.tanggal === dayTanggal)
  );
  const activeDayObj = activeWeekObj.days[activeDayIdx] || activeWeekObj.days[0];
  const prevDayObj = activeDayIdx > 0 ? activeWeekObj.days[activeDayIdx - 1] : null;

  const activeMonthObj =
    monthKey === 'ALL'
      ? null
      : months12.find((m) => m.monthKey === monthKey) || months12[10] || months12[0];
  const prevMonthObj = useMemo(() => {
    if (monthKey === 'ALL') return null;
    const idx = months12.findIndex((m) => m.monthKey === monthKey);
    return idx > 0 ? months12[idx - 1] : null;
  }, [months12, monthKey]);

  const annualTotals = useMemo(() => {
    return months12.reduce(
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
        graphisQty: acc.graphisQty + m.graphisQty,
        graphisNota: acc.graphisNota + m.graphisNota,
        simplicitySales: acc.simplicitySales + m.simplicitySales,
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
        graphisQty: 0,
        graphisNota: 0,
        simplicitySales: 0,
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
  }, [months12]);

  // Horizon-matched Brand Inventory Snapshot (Graphis 6201 + Simplicity 6244)
  const brandOps = useMemo(() => {
    const isMonthly = mode === 'bulanan';
    const opsPeriodLabel = isMonthly
      ? monthKey === 'ALL'
        ? '12 Bulan 2026'
        : `${MONTH_NAMES_ID[monthKey] || monthKey} 2026`
      : `W${weekNumber}`;

    const matched = (cvInventoryOps || []).filter((e) => {
      if (isMonthly) {
        if (e.periodType !== 'monthly') return false;
        return monthKey === 'ALL' ? true : e.monthKey === monthKey;
      }
      return e.periodType === 'weekly' && e.weekNumber === weekNumber;
    });

    const liveMap = isMonthly
      ? monthKey === 'ALL'
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
        : cvLiveQtyMap.monthly.get(monthKey) || { GRAPHIS: 0, SIMPLICITY: 0 }
      : cvLiveQtyMap.weekly.get(weekNumber) || { GRAPHIS: 0, SIMPLICITY: 0 };

    const getBrand = (b: CvInventoryBrandKey) => {
      const list = matched.filter((x) => x.brand === b);
      const stockAwal = list.reduce((s, x) => s + (x.stockAwal || 0), 0);
      const barangDatang = list.reduce((s, x) => s + (x.barangDatang || 0), 0);
      const barangKeluar = list.reduce((s, x) => s + (x.barangKeluar || 0), 0);
      const hasOverride = list.some(
        (x) => x.salesQtyOverride !== null && x.salesQtyOverride !== undefined
      );
      const salesQty = hasOverride
        ? list.reduce((s, x) => s + (x.salesQtyOverride || 0), 0)
        : liveMap[b] || 0;
      const stockAkhir = stockAwal + barangDatang - salesQty - barangKeluar;
      const promo = list.map((x) => x.promo?.trim()).filter(Boolean).join(' · ') || '-';
      return {
        meta: CV_BRAND_CATALOG[b],
        stockAwal,
        barangDatang,
        barangKeluar,
        salesQty,
        stockAkhir,
        promo,
      };
    };

    return {
      opsPeriodLabel,
      graphis: getBrand('GRAPHIS'),
      simplicity: getBrand('SIMPLICITY'),
    };
  }, [cvInventoryOps, cvLiveQtyMap, mode, monthKey, weekNumber]);

  const isPortrait3x4 = mode === 'weekly' || mode === 'bulanan';

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
    if (stream === 'DP') {
      return {
        sales: src.dpSales,
        qty: src.dpQty,
        trx: src.dpTrx,
        trxUnit: 'Trx',
        primaryA: src.dpPagiSales,
        labelA: 'Shift Pagi',
        primaryB: src.dpSiangSales,
        labelB: 'Shift Siang',
        colorA: '#f59e0b',
        colorB: '#4f46e5',
        trafficTotal: src.trafficTotal,
        pagiTraffic: src.pagiTraffic,
        siangTraffic: src.siangTraffic,
      };
    }
    if (stream === 'CV') {
      return {
        sales: src.cvSales,
        qty: src.cvQty,
        trx: src.cvNota,
        trxUnit: 'Nota',
        primaryA: src.graphisSales,
        labelA: 'Graphis (6201)',
        primaryB: src.simplicitySales,
        labelB: 'Simplicity (6244)',
        colorA: '#7c3aed',
        colorB: '#0284c7',
        trafficTotal: src.trafficTotal,
        pagiTraffic: src.pagiTraffic,
        siangTraffic: src.siangTraffic,
      };
    }
    return {
      sales: src.grandSales,
      qty: src.totalQty,
      trx: src.totalTrxNota,
      trxUnit: 'Trx+Nota',
      primaryA: src.dpSales,
      labelA: '1. Report DP',
      primaryB: src.cvSales,
      labelB: '2. Report CV',
      colorA: '#4f46e5',
      colorB: '#059669',
      trafficTotal: src.trafficTotal,
      pagiTraffic: src.pagiTraffic,
      siangTraffic: src.siangTraffic,
    };
  };

  const selectedMonthLabel =
    monthKey === 'ALL'
      ? '12 Bulan 2026'
      : `${MONTH_NAMES_ID[monthKey] || 'November'} 2026`;

  // =========================================================================
  // CANVAS RENDER ENGINE (1:1 Harian · 3:4 Weekly · 3:4 Bulanan)
  // =========================================================================
  useEffect(() => {
    if (!isOpen) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    try {
      const W = isPortrait3x4 ? 1350 : 1400;
      const H = isPortrait3x4 ? 1800 : 1400;
      canvas.width = W;
      canvas.height = H;

      const isLight = theme === 'light';
      const roundRect = (
        x: number,
        y: number,
        w: number,
        h: number,
        r: number,
        fill?: string,
        stroke?: string,
        lineWidth = 1.5
      ) => {
        ctx.beginPath();
        ctx.moveTo(x + r, y);
        ctx.arcTo(x + w, y, x + w, y + h, r);
        ctx.arcTo(x + w, y + h, x, y + h, r);
        ctx.arcTo(x, y + h, x, y, r);
        ctx.arcTo(x, y, x + w, y, r);
        ctx.closePath();
        if (fill) {
          ctx.fillStyle = fill;
          ctx.fill();
        }
        if (stroke) {
          ctx.strokeStyle = stroke;
          ctx.lineWidth = lineWidth;
          ctx.stroke();
        }
      };

      // Background
      const bgGrad = ctx.createLinearGradient(0, 0, 0, H);
      if (theme === 'navy') {
        bgGrad.addColorStop(0, '#0f172a');
        bgGrad.addColorStop(1, '#1e293b');
      } else if (theme === 'emerald') {
        bgGrad.addColorStop(0, '#064e3b');
        bgGrad.addColorStop(1, '#0f172a');
      } else {
        bgGrad.addColorStop(0, '#f8fafc');
        bgGrad.addColorStop(1, '#f1f5f9');
      }
      ctx.fillStyle = bgGrad;
      ctx.fillRect(0, 0, W, H);

      const cardBg = '#ffffff';
      const cardBorder = isLight ? '#cbd5e1' : 'rgba(255,255,255,0.25)';

      // Top Executive Header
      const hdrX = 44;
      const hdrY = 26;
      const hdrW = W - 88;
      const hdrH = isPortrait3x4 ? 142 : 150;
      const hdrFill =
        stream === 'CV'
          ? '#4c1d95'
          : stream === 'DP'
            ? '#312e81'
            : '#064e3b';
      roundRect(hdrX, hdrY, hdrW, hdrH, 20, hdrFill, 'rgba(255,255,255,0.22)', 1.5);

      ctx.fillStyle = '#e2e8f0';
      ctx.font = 'bold 13.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
      ctx.fillText(
        'REPORT BAZAR MATAHARI · TANJUNG · PT MDS RETAILING TBK',
        hdrX + 26,
        hdrY + 34
      );

      const streamBadge =
        stream === 'ALL'
          ? 'KONSOLIDASI DP + CV'
          : stream === 'DP'
            ? 'REPORT DP'
            : 'REPORT CV DIV. 1';
      const ratioBadge = isPortrait3x4 ? '3:4 PORTRAIT HD' : '1:1 SQUARE HD';
      const rightTag = `${streamBadge} · ${ratioBadge}`;
      ctx.font = 'bold 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
      const tagW = ctx.measureText(rightTag).width + 28;
      roundRect(
        hdrX + hdrW - tagW - 24,
        hdrY + 18,
        tagW,
        28,
        8,
        'rgba(255,255,255,0.16)',
        'rgba(255,255,255,0.35)',
        1
      );
      ctx.fillStyle = '#ffffff';
      ctx.textAlign = 'center';
      ctx.fillText(rightTag, hdrX + hdrW - 24 - tagW / 2, hdrY + 36);
      ctx.textAlign = 'left';

      ctx.fillStyle = '#ffffff';
      ctx.font = isPortrait3x4
        ? '800 30px "Plus Jakarta Sans", Inter, system-ui, sans-serif'
        : '800 32px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
      const mainTitle =
        mode === 'harian'
          ? `DASHBOARD HARIAN — ${activeDayObj.dayName.toUpperCase()}, ${activeDayObj.tanggal}`
          : mode === 'weekly'
            ? weeklyFocus === 'growth_weekly'
              ? `DASHBOARD WEEKLY — GROWTH & KONTRIBUSI (W${activeWeekObj.weekNumber})`
              : `DASHBOARD WEEKLY ${activeWeekObj.weekNumber} (${activeWeekObj.periodeText})`
            : monthlyFocus === 'kontribusi_growth'
              ? `DASHBOARD BULANAN — KONTRIBUSI & GROWTH (${selectedMonthLabel.toUpperCase()})`
              : `DASHBOARD BULANAN — ${selectedMonthLabel.toUpperCase()}`;
      ctx.fillText(
        clipText(ctx, mainTitle, hdrW - 52),
        hdrX + 26,
        hdrY + (isPortrait3x4 ? 82 : 88)
      );

      ctx.fillStyle = '#cbd5e1';
      ctx.font = '600 15.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
      const subTitle =
        mode === 'harian'
          ? `Weekly ${activeWeekObj.weekNumber} (${activeWeekObj.periodeText}) · DP + CV Div. 1 (Graphis 6201 · Simplicity 6244) · Traffic & Conversion Rate`
          : mode === 'weekly'
            ? `Rekap Senin–Minggu (W${activeWeekObj.weekNumber}) · Growth WoW %, Kontribusi, Traffic & Data Brand Divisi 1`
            : `Rekap Kalender 2026 (${selectedMonthLabel}) · 12 Bulan, MoM Growth %, Kontribusi & Data Brand Divisi 1`;
      ctx.fillText(
        clipText(ctx, subTitle, hdrW - 52),
        hdrX + 26,
        hdrY + (isPortrait3x4 ? 116 : 122)
      );

      // Shared 4-Card KPI Strip
      const draw4KpiStrip = (
        y: number,
        h: number,
        cards: { label: string; value: string; sub: string; accent: string }[]
      ) => {
        const gap = 14;
        const cw = (W - 88 - gap * 3) / 4;
        cards.forEach((c, i) => {
          const x = 44 + i * (cw + gap);
          roundRect(x, y, cw, h, 14, cardBg, cardBorder, 1.5);
          ctx.fillStyle = '#64748b';
          ctx.font = '700 14px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(clipText(ctx, c.label, cw - 32), x + 16, y + 26);

          ctx.fillStyle = '#0f172a';
          ctx.font = '800 26.5px "JetBrains Mono", "Plus Jakarta Sans", monospace';
          ctx.fillText(clipText(ctx, c.value, cw - 32), x + 16, y + 60);

          ctx.fillStyle = c.accent;
          ctx.font = '800 15.5px "JetBrains Mono", "Plus Jakarta Sans", monospace';
          ctx.fillText(clipText(ctx, c.sub, cw - 32), x + 16, y + h - 14);
        });
      };

      // Shared Stacked Sales Trend Diagram
      const drawStackedDiagram = (
        x: number,
        y: number,
        w: number,
        h: number,
        title: string,
        labelA: string,
        labelB: string,
        colorA: string,
        colorB: string,
        totA: number,
        totB: number,
        bars: {
          label: string;
          sub?: string;
          valA: number;
          valB: number;
          total: number;
          isWeekend?: boolean;
          isSelected?: boolean;
        }[]
      ) => {
        roundRect(x, y, w, h, 14, cardBg, cardBorder, 1.5);
        roundRect(x, y, w, 46, 14, '#0f172a');
        ctx.fillStyle = '#ffffff';
        ctx.font = '700 13.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(clipText(ctx, title, w - 360), x + 18, y + 28);

        const totAll = totA + totB;
        const pctA = totAll > 0 ? Math.round((totA / totAll) * 100) : 0;
        const pctB = totAll > 0 ? 100 - pctA : 0;

        // Legend right
        ctx.textAlign = 'right';
        ctx.font = '700 11.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillStyle = '#e2e8f0';
        const legText = `${labelA} ${pctA}% · ${labelB} ${pctB}% · Total: Rp ${fmtShortRp(totAll)}`;
        ctx.fillText(legText, x + w - 18, y + 28);
        ctx.textAlign = 'left';

        const plotX = x + 52;
        const plotY = y + 72;
        const plotW = w - 72;
        const plotH = h - 116;
        const maxVal = Math.max(1000000, ...bars.map((b) => b.total));

        [1, 0.5, 0].forEach((r) => {
          const gy = plotY + plotH - r * plotH;
          ctx.fillStyle = '#64748b';
          ctx.font = '600 10.5px "JetBrains Mono", monospace';
          ctx.textAlign = 'right';
          ctx.fillText(fmtShortRp(Math.round(maxVal * r)), plotX - 8, gy + 4);
          ctx.textAlign = 'left';
          ctx.beginPath();
          ctx.strokeStyle = r === 0 ? '#cbd5e1' : '#e2e8f0';
          ctx.lineWidth = 1;
          if (r !== 0) ctx.setLineDash([4, 4]);
          ctx.moveTo(plotX, gy);
          ctx.lineTo(plotX + plotW, gy);
          ctx.stroke();
          ctx.setLineDash([]);
        });

        const n = Math.max(1, bars.length);
        const slotW = plotW / n;
        const barW = Math.min(34, Math.max(12, slotW * 0.48));
        const pts: { x: number; y: number }[] = [];

        bars.forEach((b, idx) => {
          const bx = plotX + idx * slotW + (slotW - barW) / 2;
          const totalH =
            b.total > 0 ? Math.max(6, (b.total / maxVal) * plotH * 0.88) : 3;
          const hA = b.total > 0 ? (b.valA / b.total) * totalH : 0;
          const hB = totalH - hA;
          const byTop = plotY + plotH - totalH;

          if (b.total > 0) {
            ctx.fillStyle = colorB;
            ctx.fillRect(bx, byTop, barW, hB);
            ctx.fillStyle = colorA;
            ctx.fillRect(bx, byTop + hB, barW, hA);
          } else {
            ctx.fillStyle = '#e2e8f0';
            ctx.fillRect(bx, plotY + plotH - 3, barW, 3);
          }

          if (b.isSelected) {
            ctx.strokeStyle = '#0f172a';
            ctx.lineWidth = 2;
            ctx.strokeRect(bx - 2, byTop - 2, barW + 4, totalH + 4);
          }

          ctx.textAlign = 'center';
          ctx.fillStyle = '#0f172a';
          ctx.font = '700 10px "JetBrains Mono", monospace';
          if (b.total > 0) {
            ctx.fillText(fmtShortRp(b.total), bx + barW / 2, byTop - 6);
          }
          ctx.fillStyle = b.isWeekend ? '#be123c' : '#334155';
          ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(b.label, bx + barW / 2, plotY + plotH + 16);
          if (b.sub) {
            ctx.fillStyle = '#64748b';
            ctx.font = '600 9.5px "JetBrains Mono", monospace';
            ctx.fillText(b.sub, bx + barW / 2, plotY + plotH + 30);
          }
          ctx.textAlign = 'left';
          pts.push({ x: bx + barW / 2, y: byTop });
        });

        if (pts.length > 1) {
          ctx.beginPath();
          ctx.strokeStyle = '#0f172a';
          ctx.lineWidth = 1.8;
          pts.forEach((p, i) => {
            if (i === 0) ctx.moveTo(p.x, p.y);
            else ctx.lineTo(p.x, p.y);
          });
          ctx.stroke();
          pts.forEach((p) => {
            ctx.beginPath();
            ctx.fillStyle = '#ffffff';
            ctx.strokeStyle = '#0f172a';
            ctx.lineWidth = 1.8;
            ctx.arc(p.x, p.y, 3, 0, Math.PI * 2);
            ctx.fill();
            ctx.stroke();
          });
        }
      };

      // Shared Growth/Contribution Bar Chart
      const drawGrowthOrContribChart = (
        x: number,
        y: number,
        w: number,
        h: number,
        modeKind: 'growth' | 'contrib',
        title: string,
        items: {
          label: string;
          sub?: string;
          val: number;
          base?: boolean;
          isHighlight?: boolean;
        }[]
      ) => {
        roundRect(x, y, w, h, 14, cardBg, cardBorder, 1.5);
        roundRect(x, y, w, 40, 14, '#f8fafc', '#e2e8f0', 1);
        ctx.fillStyle = '#0f172a';
        ctx.font = '800 14.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(clipText(ctx, title, w - 32), x + 16, y + 25);

        const plotX = x + 52;
        const plotY = y + 66;
        const plotW = w - 68;
        const plotH = h - 108;
        const maxVal =
          modeKind === 'contrib'
            ? Math.max(15, ...items.map((it) => it.val))
            : Math.max(20, ...items.map((it) => Math.min(200, Math.abs(it.val))));

        [1, 0.5, 0].forEach((r) => {
          const gy = plotY + plotH - r * plotH;
          ctx.fillStyle = '#64748b';
          ctx.font = '700 11.5px "JetBrains Mono", monospace';
          ctx.textAlign = 'right';
          ctx.fillText(
            modeKind === 'contrib'
              ? `${Math.round(maxVal * r)}%`
              : r === 0
                ? '0%'
                : `±${Math.round(maxVal * r)}%`,
            plotX - 6,
            gy + 4
          );
          ctx.textAlign = 'left';
          ctx.beginPath();
          ctx.strokeStyle = r === 0 ? '#cbd5e1' : '#e2e8f0';
          ctx.lineWidth = 1;
          if (r !== 0) ctx.setLineDash([3, 3]);
          ctx.moveTo(plotX, gy);
          ctx.lineTo(plotX + plotW, gy);
          ctx.stroke();
          ctx.setLineDash([]);
        });

        const n = Math.max(1, items.length);
        const slotW = plotW / n;
        const barW = Math.min(30, Math.max(10, slotW * 0.45));
        const valFontSize = n > 10 && w < 700 ? 11.5 : 13.5;

        items.forEach((it, idx) => {
          const bx = plotX + idx * slotW + (slotW - barW) / 2;
          const mag =
            modeKind === 'contrib'
              ? it.val
              : it.base
                ? maxVal * 0.14
                : Math.min(200, Math.abs(it.val));
          const bh = mag > 0 ? Math.max(5, (mag / maxVal) * plotH * 0.84) : 3;
          const byTop = plotY + plotH - bh;

          if (modeKind === 'contrib') {
            ctx.fillStyle = it.isHighlight ? '#0f172a' : '#0284c7';
          } else {
            ctx.fillStyle = it.base
              ? '#94a3b8'
              : it.val > 0
                ? '#059669'
                : it.val < 0
                  ? '#e11d48'
                  : '#cbd5e1';
          }
          ctx.fillRect(bx, byTop, barW, bh);

          ctx.textAlign = 'center';
          ctx.font = `800 ${valFontSize}px "JetBrains Mono", monospace`;
          ctx.fillStyle =
            modeKind === 'contrib'
              ? '#0369a1'
              : it.base
                ? '#64748b'
                : it.val >= 0
                  ? '#047857'
                  : '#be123c';
          const lbl =
            modeKind === 'contrib'
              ? `${it.val.toFixed(1)}%`
              : it.base
                ? 'Basis'
                : `${it.val >= 0 ? '+' : ''}${it.val.toFixed(0)}%`;
          ctx.fillText(lbl, bx + barW / 2, byTop - 6);

          ctx.fillStyle = it.isHighlight ? '#0f172a' : '#475569';
          ctx.font = '700 11.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(it.label, bx + barW / 2, plotY + plotH + 15);
          if (it.sub) {
            ctx.fillStyle = '#64748b';
            ctx.font = '600 9.5px "JetBrains Mono", monospace';
            ctx.fillText(it.sub, bx + barW / 2, plotY + plotH + 28);
          }
          ctx.textAlign = 'left';
        });
      };

      // Shared Brand Comparison Card Panel (NO Pagi/Siang shift split — daily/period totals only)
      const drawBrandPanel = (
        x: number,
        y: number,
        w: number,
        h: number,
        title: string,
        raw: {
          cvSales: number;
          grandSales: number;
          graphisSales: number;
          graphisQty: number;
          graphisNota: number;
          simplicitySales: number;
          simplicityQty: number;
          simplicityNota: number;
        }
      ) => {
        roundRect(x, y, w, h, 14, cardBg, cardBorder, 1.5);
        roundRect(x, y, w, 40, 14, '#f8fafc', '#e2e8f0', 1);
        ctx.fillStyle = '#0f172a';
        ctx.font = '700 13.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(title, x + 16, y + 25);

        const dom =
          raw.cvSales <= 0
            ? 'Belum ada data'
            : Math.abs(raw.graphisSales - raw.simplicitySales) < 1000
              ? 'Seimbang'
              : raw.graphisSales > raw.simplicitySales
                ? 'Dominan: Graphis (6201)'
                : 'Dominan: Simplicity (6244)';
        ctx.textAlign = 'right';
        ctx.fillStyle = '#6d28d9';
        ctx.font = '700 12px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(dom, x + w - 16, y + 25);
        ctx.textAlign = 'left';

        const cards = [
          {
            name: 'Brand Graphis',
            dept: 'Dept 6201 · Divisi 1',
            color: '#7c3aed',
            sales: raw.graphisSales,
            qty: raw.graphisQty,
            nota: raw.graphisNota,
            ops: brandOps.graphis,
          },
          {
            name: 'Brand Simplicity',
            dept: 'Dept 6244 · Divisi 1',
            color: '#0284c7',
            sales: raw.simplicitySales,
            qty: raw.simplicityQty,
            nota: raw.simplicityNota,
            ops: brandOps.simplicity,
          },
        ];

        const gap = 14;
        const cw = (w - 32 - gap) / 2;
        const ch = h - 56;
        const cy = y + 48;

        cards.forEach((b, i) => {
          const cx = x + 16 + i * (cw + gap);
          roundRect(cx, cy, cw, ch, 12, '#f8fafc', '#e2e8f0', 1);

          ctx.fillStyle = b.color;
          ctx.font = '700 11.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(b.dept, cx + 14, cy + 20);

          const shareCv = raw.cvSales > 0 ? (b.sales / raw.cvSales) * 100 : 0;
          const shareGrand = raw.grandSales > 0 ? (b.sales / raw.grandSales) * 100 : 0;
          ctx.textAlign = 'right';
          ctx.fillStyle = '#0f172a';
          ctx.font = '800 13.5px "JetBrains Mono", monospace';
          ctx.fillText(
            `${shareCv.toFixed(1)}% CV · ${shareGrand.toFixed(1)}% DP+CV`,
            cx + cw - 14,
            cy + 20
          );
          ctx.textAlign = 'left';

          ctx.fillStyle = '#0f172a';
          ctx.font = '800 16px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(b.name, cx + 14, cy + 43);

          ctx.textAlign = 'right';
          ctx.font = '800 19px "JetBrains Mono", monospace';
          ctx.fillText(`Rp ${formatNumberId(b.sales) || '0'}`, cx + cw - 14, cy + 43);
          ctx.textAlign = 'left';

          const pp = b.qty > 0 ? Math.round(b.sales / b.qty) : 0;
          const bs = b.nota > 0 ? Math.round(b.sales / b.nota) : 0;

          // Commercial row: Qty & Nota left, enlarged PP & BS right
          ctx.fillStyle = '#475569';
          ctx.font = '600 12.5px "JetBrains Mono", monospace';
          ctx.fillText(
            `Qty: ${formatNumberId(b.qty) || '0'} pcs · Nota: ${formatNumberId(b.nota) || '0'}`,
            cx + 14,
            cy + 66
          );
          ctx.textAlign = 'right';
          ctx.fillStyle = '#1e1b4b';
          ctx.font = '800 14.5px "JetBrains Mono", monospace';
          ctx.fillText(
            `PP: Rp ${fmtShortRp(pp)}  ·  BS: Rp ${fmtShortRp(bs)}`,
            cx + cw - 14,
            cy + 66
          );
          ctx.textAlign = 'left';

          // Stock row if height allows
          if (ch >= 115) {
            roundRect(cx + 12, cy + 78, cw - 24, ch - 88, 8, '#ffffff', '#e2e8f0', 1);
            ctx.fillStyle = '#64748b';
            ctx.font = '700 11px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
            ctx.fillText(
              `Mutasi Stok (${brandOps.opsPeriodLabel}):`,
              cx + 22,
              cy + 96
            );
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 11.5px "JetBrains Mono", monospace';
            ctx.fillText(
              `Awal ${b.ops.stockAwal} · DO +${b.ops.barangDatang} · Sls/Ret -${b.ops.salesQty + b.ops.barangKeluar} · Akhir ${b.ops.stockAkhir} pcs`,
              cx + 22,
              cy + 116
            );
            if (ch >= 145) {
              ctx.fillStyle = '#475569';
              ctx.font = '600 11px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
              ctx.fillText(
                clipText(ctx, `Promo: ${b.ops.promo}`, cw - 44),
                cx + 22,
                cy + 136
              );
            }
          }
        });
      };

      const calcGrowth = (curr: number, prev: number | null) => {
        if (prev === null || prev === undefined) return null;
        if (prev <= 0 && curr <= 0) return 0;
        if (prev <= 0 && curr > 0) return 100;
        return ((curr - prev) / prev) * 100;
      };

      // Shared 4-Pillar Comparison Strip: Report DP | Report CV | Brand Div. 1 | Total DP+CV
      const draw4PillarBreakdown = (
        y: number,
        h: number,
        src: {
          dpSales: number;
          dpPagiSales: number;
          dpSiangSales: number;
          dpQty: number;
          dpTrx: number;
          cvSales: number;
          cvQty: number;
          cvNota: number;
          graphisSales: number;
          graphisQty: number;
          graphisNota: number;
          simplicitySales: number;
          simplicityQty: number;
          simplicityNota: number;
          grandSales: number;
          totalQty: number;
          totalTrxNota: number;
          trafficTotal: number;
          pagiTraffic: number;
          siangTraffic: number;
        }
      ) => {
        const gap = 12;
        const cw = (W - 88 - gap * 3) / 4;
        const dpShare = src.grandSales > 0 ? ((src.dpSales / src.grandSales) * 100).toFixed(1) : '0.0';
        const cvShare = src.grandSales > 0 ? ((src.cvSales / src.grandSales) * 100).toFixed(1) : '0.0';
        const gShareCv = src.cvSales > 0 ? ((src.graphisSales / src.cvSales) * 100).toFixed(0) : '0';
        const sShareCv = src.cvSales > 0 ? ((src.simplicitySales / src.cvSales) * 100).toFixed(0) : '0';
        const convPct = src.trafficTotal > 0 ? ((src.totalTrxNota / src.trafficTotal) * 100).toFixed(1) : '0.0';
        const avgOrg = src.trafficTotal > 0 ? Math.round(src.grandSales / src.trafficTotal) : 0;

        const cards = [
          {
            title: '1. REPORT DP (DIRECT)',
            value: `Rp ${formatNumberId(src.dpSales) || '0'}`,
            line1: `${src.dpQty} pcs · ${src.dpTrx} Trx · ${dpShare}% Total`,
            line2: `Pagi ${fmtShortRp(src.dpPagiSales)} · Siang ${fmtShortRp(src.dpSiangSales)}`,
            accent: '#4f46e5',
            active: stream === 'DP',
          },
          {
            title: '2. REPORT CV DIV. 1',
            value: `Rp ${formatNumberId(src.cvSales) || '0'}`,
            line1: `${src.cvQty} pcs · ${src.cvNota} Nota · ${cvShare}% Total`,
            line2: `PP Rp ${fmtShortRp(src.cvQty > 0 ? Math.round(src.cvSales / src.cvQty) : 0)} · BS Rp ${fmtShortRp(src.cvNota > 0 ? Math.round(src.cvSales / src.cvNota) : 0)}`,
            accent: '#7c3aed',
            active: stream === 'CV',
          },
          {
            title: '3. BRAND (6201 & 6244)',
            value: `G ${fmtShortRp(src.graphisSales)} · S ${fmtShortRp(src.simplicitySales)}`,
            line1: `Graphis: ${src.graphisQty} pcs (${gShareCv}% CV)`,
            line2: `Simplicity: ${src.simplicityQty} pcs (${sShareCv}% CV)`,
            accent: '#0284c7',
            active: false,
          },
          {
            title: '4. REPORT TOTAL DP + CV',
            value: `Rp ${formatNumberId(src.grandSales) || '0'}`,
            line1: `${src.totalQty} pcs · ${src.totalTrxNota} Trx+Nota`,
            line2: `Conv ${convPct}% · Trf ${src.trafficTotal} org · Avg ${fmtShortRp(avgOrg)}`,
            accent: '#059669',
            active: stream === 'ALL',
          },
        ];

        cards.forEach((c, idx) => {
          const cx = 44 + idx * (cw + gap);
          roundRect(
            cx,
            y,
            cw,
            h,
            12,
            cardBg,
            c.active ? c.accent : cardBorder,
            c.active ? 2.2 : 1.5
          );
          ctx.fillStyle = c.accent;
          ctx.font = '700 12px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(clipText(ctx, c.title, cw - 28), cx + 14, y + 22);

          ctx.fillStyle = '#0f172a';
          ctx.font = '800 18.5px "JetBrains Mono", monospace';
          ctx.fillText(clipText(ctx, c.value, cw - 28), cx + 14, y + 48);

          ctx.fillStyle = '#1e293b';
          ctx.font = '700 12.5px "JetBrains Mono", monospace';
          ctx.fillText(clipText(ctx, c.line1, cw - 28), cx + 14, y + 73);

          ctx.fillStyle = '#0f172a';
          ctx.font = '700 12.5px "JetBrains Mono", "Plus Jakarta Sans", monospace';
          ctx.fillText(clipText(ctx, c.line2, cw - 28), cx + 14, y + h - 12);
        });
      };

      // =====================================================================
      // LAYOUT 1: HARIAN 1:1 (1400 × 1400 px)
      // =====================================================================
      if (mode === 'harian') {
        const curr = projectStream(activeDayObj);
        const prev = prevDayObj ? projectStream(prevDayObj) : null;
        const gr = calcGrowth(curr.sales, prev ? prev.sales : null);
        const pp = curr.qty > 0 ? Math.round(curr.sales / curr.qty) : 0;
        const bs = curr.trx > 0 ? Math.round(curr.sales / curr.trx) : 0;
        const conv =
          curr.trafficTotal > 0 ? (curr.trx / curr.trafficTotal) * 100 : 0;

        draw4KpiStrip(192, 118, [
          {
            label: `Sales Harian (${streamBadge})`,
            value: `Rp ${formatNumberId(curr.sales) || '0'}`,
            sub:
              gr === null
                ? `H${activeDayIdx + 1} Minggu W${activeWeekObj.weekNumber}`
                : `${gr >= 0 ? '+' : ''}${gr.toFixed(1)}% vs ${prevDayObj?.dayName.slice(0, 3)}`,
            accent: gr !== null && gr < 0 ? '#be123c' : '#047857',
          },
          {
            label: 'Total Qty Terjual',
            value: `${formatNumberId(curr.qty) || '0'} pcs`,
            sub: `Price Point: Rp ${formatNumberId(pp) || '0'}`,
            accent: '#4338ca',
          },
          {
            label: `Total ${curr.trxUnit}`,
            value: `${formatNumberId(curr.trx) || '0'} ${curr.trxUnit.toLowerCase()}`,
            sub: `Big Size: Rp ${formatNumberId(bs) || '0'}`,
            accent: '#b45309',
          },
          {
            label: 'Traffic & Conversion Rate',
            value: `${conv.toFixed(1)}% (${formatNumberId(curr.trafficTotal) || '0'} org)`,
            sub: `Pagi ${curr.pagiTraffic} · Siang ${curr.siangTraffic} org`,
            accent: '#0369a1',
          },
        ]);

        // Stacked Daily Bar Chart (7 Days of W)
        const totA = activeWeekObj.days.reduce(
          (s, d) => s + projectStream(d).primaryA,
          0
        );
        const totB = activeWeekObj.days.reduce(
          (s, d) => s + projectStream(d).primaryB,
          0
        );
        drawStackedDiagram(
          44,
          326,
          W - 88,
          340,
          `TREN PENJUALAN HARIAN W${activeWeekObj.weekNumber} (${curr.labelA.toUpperCase()} vs ${curr.labelB.toUpperCase()})`,
          curr.labelA,
          curr.labelB,
          curr.colorA,
          curr.colorB,
          totA,
          totB,
          activeWeekObj.days.map((d) => {
            const p = projectStream(d);
            return {
              label: d.dayName.slice(0, 3),
              sub: d.tanggal.slice(0, 5),
              valA: p.primaryA,
              valB: p.primaryB,
              total: p.sales,
              isWeekend: d.isWeekend,
              isSelected: d.tanggal === activeDayObj.tanggal,
            };
          })
        );

        // 4-Pillar Comparison Row: Report DP | Report CV | Brand | Total DP+CV
        draw4PillarBreakdown(682, 128, activeDayObj);

        // Data Brand Harian Panel (NO Pagi/Siang split)
        drawBrandPanel(
          44,
          826,
          W - 88,
          224,
          `DATA BRAND DIVISI 1 — HARIAN (${activeDayObj.tanggal})`,
          activeDayObj
        );

        // Mini 7-Day Table + Reason Footer (all 4 pillars per row)
        const tblY = 1066;
        const tblH = 282;
        roundRect(44, tblY, W - 88, tblH, 14, cardBg, cardBorder, 1.5);
        roundRect(44, tblY, W - 88, 36, 14, '#0f172a');
        ctx.fillStyle = '#ffffff';
        ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(
          `TABEL RINCIAN HARIAN W${activeWeekObj.weekNumber} — REPORT DP · BRAND · REPORT CV · TOTAL DP+CV`,
          60,
          tblY + 23
        );

        ctx.fillStyle = '#475569';
        ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText('HARI / TGL', 60, tblY + 56);
        ctx.fillText('REPORT DP', 225, tblY + 56);
        ctx.fillText('GRAPHIS (6201)', 380, tblY + 56);
        ctx.fillText('SIMPLICITY (6244)', 535, tblY + 56);
        ctx.fillText('REPORT CV', 705, tblY + 56);
        ctx.fillText('TOTAL DP+CV', 865, tblY + 56);
        ctx.fillText('TRX+NT', 1035, tblY + 56);
        ctx.fillText('TRAFFIC', 1130, tblY + 56);
        ctx.fillStyle = '#047857';
        ctx.font = '800 12px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText('CONV %', 1245, tblY + 56);

        let ry = tblY + 66;
        activeWeekObj.days.forEach((d, i) => {
          const isSel = d.tanggal === activeDayObj.tanggal;
          roundRect(
            54,
            ry,
            W - 108,
            26,
            6,
            isSel ? '#ecfdf5' : i % 2 === 0 ? '#f8fafc' : '#ffffff',
            isSel ? '#059669' : undefined,
            isSel ? 1.5 : 0
          );
          ctx.fillStyle = d.isWeekend ? '#be123c' : '#0f172a';
          ctx.font = `${isSel ? '700' : '600'} 11px "JetBrains Mono", monospace`;
          ctx.fillText(`${d.dayName.slice(0, 3)}, ${d.tanggal}`, 60, ry + 17);
          ctx.fillStyle = '#312e81';
          ctx.fillText(`Rp ${formatNumberId(d.dpSales) || '0'}`, 225, ry + 17);
          ctx.fillStyle = '#6d28d9';
          ctx.fillText(`Rp ${formatNumberId(d.graphisSales) || '0'}`, 380, ry + 17);
          ctx.fillStyle = '#0369a1';
          ctx.fillText(`Rp ${formatNumberId(d.simplicitySales) || '0'}`, 535, ry + 17);
          ctx.fillStyle = '#5b21b6';
          ctx.fillText(`Rp ${formatNumberId(d.cvSales) || '0'}`, 705, ry + 17);
          ctx.fillStyle = '#0f172a';
          ctx.font = '700 11px "JetBrains Mono", monospace';
          ctx.fillText(`Rp ${formatNumberId(d.grandSales) || '0'}`, 865, ry + 17);
          ctx.fillStyle = '#047857';
          ctx.fillText(`${d.totalTrxNota}`, 1035, ry + 17);
          ctx.fillStyle = '#334155';
          ctx.fillText(`${d.trafficTotal} org`, 1130, ry + 17);
          ctx.fillStyle = '#047857';
          ctx.font = '800 13.5px "JetBrains Mono", monospace';
          ctx.fillText(
            d.trafficTotal > 0
              ? `${((d.totalTrxNota / d.trafficTotal) * 100).toFixed(1)}%`
              : '-',
            1245,
            ry + 18
          );
          ry += 28;
        });

        const noteTxt = [
          activeDayObj.reasonPagi ? `Pagi: ${activeDayObj.reasonPagi}` : '',
          activeDayObj.reasonSiang ? `Siang: ${activeDayObj.reasonSiang}` : '',
          activeDayObj.cvReason ? `CV: ${activeDayObj.cvReason}` : '',
        ]
          .filter(Boolean)
          .join(' · ');
        ctx.fillStyle = '#64748b';
        ctx.font = '600 11px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
        ctx.fillText(
          clipText(ctx, `Catatan (${activeDayObj.tanggal}): ${noteTxt || '-'}`, W - 120),
          60,
          tblY + tblH - 10
        );
      } else if (mode === 'weekly') {
        // =====================================================================
        // LAYOUT 2: WEEKLY 3:4 (1350 × 1800 px)
        // =====================================================================
        const curr = projectStream(activeWeekObj);
        const prev = prevWeekObj ? projectStream(prevWeekObj) : null;
        const gr = calcGrowth(curr.sales, prev ? prev.sales : null);
        const pp = curr.qty > 0 ? Math.round(curr.sales / curr.qty) : 0;
        const bs = curr.trx > 0 ? Math.round(curr.sales / curr.trx) : 0;
        const conv =
          curr.trafficTotal > 0 ? (curr.trx / curr.trafficTotal) * 100 : 0;

        draw4KpiStrip(184, 110, [
          {
            label: `Sales W${activeWeekObj.weekNumber} (${streamBadge})`,
            value: `Rp ${formatNumberId(curr.sales) || '0'}`,
            sub:
              gr === null
                ? 'Minggu awal'
                : `${gr >= 0 ? '+' : ''}${gr.toFixed(1)}% vs W${prevWeekObj?.weekNumber}`,
            accent: gr !== null && gr < 0 ? '#be123c' : '#047857',
          },
          {
            label: 'Total Qty Mingguan',
            value: `${formatNumberId(curr.qty) || '0'} pcs`,
            sub: `Price Point: Rp ${formatNumberId(pp) || '0'}`,
            accent: '#4338ca',
          },
          {
            label: `Total ${curr.trxUnit}`,
            value: `${formatNumberId(curr.trx) || '0'}`,
            sub: `Big Size: Rp ${formatNumberId(bs) || '0'}`,
            accent: '#b45309',
          },
          {
            label: 'Traffic & Conv Rate',
            value: `${conv.toFixed(1)}% (${formatNumberId(curr.trafficTotal) || '0'})`,
            sub: `Pagi ${curr.pagiTraffic} · Siang ${curr.siangTraffic} org`,
            accent: '#0369a1',
          },
        ]);

        // 4-Pillar Comparison Strip (DP | CV | Brand | Total DP+CV)
        draw4PillarBreakdown(308, 108, activeWeekObj);

        // Active weeks window (up to 10 active/surrounding weeks)
        const activeWks = weekly53.filter(
          (w) =>
            w.grandSales > 0 ||
            w.trafficTotal > 0 ||
            Math.abs(w.weekNumber - activeWeekObj.weekNumber) <= 3
        );
        const targetWks = activeWks.slice(-10);
        const totalWksWindow = targetWks.reduce(
          (s, w) => s + projectStream(w).sales,
          0
        );

        const weekGrowthItems = targetWks.map((w) => {
          const prevW =
            w.weekNumber > 1
              ? weekly53.find((x) => x.weekNumber === w.weekNumber - 1) || null
              : null;
          const g = calcGrowth(
            projectStream(w).sales,
            prevW ? projectStream(prevW).sales : null
          );
          return {
            label: `W${w.weekNumber}`,
            sub: w.shortRange,
            val: g ?? 0,
            base: g === null,
            isHighlight: w.weekNumber === activeWeekObj.weekNumber,
          };
        });

        const weekContribItems = targetWks.map((w) => {
          const val = projectStream(w).sales;
          return {
            label: `W${w.weekNumber}`,
            sub: w.shortRange,
            val: totalWksWindow > 0 ? (val / totalWksWindow) * 100 : 0,
            isHighlight: w.weekNumber === activeWeekObj.weekNumber,
          };
        });

        if (weeklyFocus === 'ringkasan') {
          const halfW = (W - 88 - 16) / 2;

          // Left: Tabel 7 Hari W{n}
          const mY = 430;
          const mH = 420;
          roundRect(44, mY, halfW, mH, 14, cardBg, cardBorder, 1.5);
          roundRect(44, mY, halfW, 40, 14, '#0f172a');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(
            `RINCIAN 7 HARI W${activeWeekObj.weekNumber} (${activeWeekObj.shortRange})`,
            60,
            mY + 25
          );

          ctx.fillStyle = '#64748b';
          ctx.font = '700 10px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('HARI', 58, mY + 58);
          ctx.fillText('DP', 155, mY + 58);
          ctx.fillText('G / S', 240, mY + 58);
          ctx.fillText('CV', 345, mY + 58);
          ctx.fillText('DP+CV', 430, mY + 58);
          ctx.fillText('T+N', 530, mY + 58);
          ctx.fillText('TRF', 588, mY + 58);

          let dy = mY + 68;
          activeWeekObj.days.forEach((d, i) => {
            roundRect(
              54,
              dy,
              halfW - 20,
              39,
              7,
              d.isWeekend ? '#fff1f2' : i % 2 === 0 ? '#f8fafc' : '#ffffff',
              '#e2e8f0',
              1
            );
            ctx.fillStyle = d.isWeekend ? '#be123c' : '#0f172a';
            ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
            ctx.fillText(`${d.dayName.slice(0, 3)} ${d.tanggal.slice(0, 5)}`, 60, dy + 23);

            ctx.font = '600 10.5px "JetBrains Mono", monospace';
            ctx.fillStyle = '#312e81';
            ctx.fillText(fmtShortRp(d.dpSales), 155, dy + 23);
            ctx.fillStyle = '#6d28d9';
            ctx.fillText(`${fmtShortRp(d.graphisSales)}/${fmtShortRp(d.simplicitySales)}`, 240, dy + 23);
            ctx.fillStyle = '#5b21b6';
            ctx.fillText(fmtShortRp(d.cvSales), 345, dy + 23);
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 11px "JetBrains Mono", monospace';
            ctx.fillText(`Rp ${fmtShortRp(d.grandSales)}`, 430, dy + 23);
            ctx.fillStyle = '#047857';
            ctx.fillText(`${d.totalTrxNota}`, 530, dy + 23);
            ctx.fillStyle = '#475569';
            ctx.fillText(`${d.trafficTotal}`, 588, dy + 23);
            dy += 44;
          });

          roundRect(54, mY + 382, halfW - 20, 26, 6, '#064e3b');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 11px "JetBrains Mono", monospace';
          ctx.fillText(
            `TOTAL W${activeWeekObj.weekNumber}: DP Rp ${fmtShortRp(activeWeekObj.dpSales)} + CV Rp ${fmtShortRp(activeWeekObj.cvSales)} = Rp ${formatNumberId(activeWeekObj.grandSales)}`,
            64,
            mY + 399
          );

          // Right: Stacked Chart 7 Hari
          const rx = 44 + halfW + 16;
          drawStackedDiagram(
            rx,
            mY,
            halfW,
            mH,
            `DIAGRAM 7 HARI W${activeWeekObj.weekNumber} (${curr.labelA} vs ${curr.labelB})`,
            curr.labelA,
            curr.labelB,
            curr.colorA,
            curr.colorB,
            curr.primaryA,
            curr.primaryB,
            activeWeekObj.days.map((d) => {
              const p = projectStream(d);
              return {
                label: d.dayName.slice(0, 3),
                sub: d.tanggal.slice(0, 5),
                valA: p.primaryA,
                valB: p.primaryB,
                total: p.sales,
                isWeekend: d.isWeekend,
              };
            })
          );

          // Middle-Bottom: Data Brand Weekly (Graphis & Simplicity — NO Pagi/Siang)
          drawBrandPanel(
            44,
            864,
            W - 88,
            210,
            `DATA BRAND DIVISI 1 — WEEKLY ${activeWeekObj.weekNumber} (${activeWeekObj.periodeText})`,
            activeWeekObj
          );

          // Bottom Split Charts: Growth Weekly + Kontribusi Weekly
          drawGrowthOrContribChart(
            44,
            1088,
            halfW,
            308,
            'growth',
            'GROWTH SALES ANTAR WEEKLY (WoW %)',
            weekGrowthItems
          );
          drawGrowthOrContribChart(
            rx,
            1088,
            halfW,
            308,
            'contrib',
            'KONTRIBUSI SALES ANTAR WEEKLY (%)',
            weekContribItems
          );

          // Footer 4-Pillar Rekap Weekly Table
          const fY = 1410;
          const fH = 338;
          roundRect(44, fY, W - 88, fH, 14, cardBg, cardBorder, 1.5);
          roundRect(44, fY, W - 88, 38, 14, '#0f172a');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(
            `REKAP WEEKLY AKTIF (${Math.min(7, targetWks.length)} MINGGU TERAKHIR) — DP · BRAND · CV · TOTAL DP+CV`,
            60,
            fY + 24
          );

          ctx.fillStyle = '#64748b';
          ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('WEEK', 60, fY + 56);
          ctx.fillText('PERIODE', 135, fY + 56);
          ctx.fillText('REPORT DP', 275, fY + 56);
          ctx.fillText('GRAPHIS (6201)', 435, fY + 56);
          ctx.fillText('SIMPLICITY (6244)', 590, fY + 56);
          ctx.fillText('REPORT CV', 755, fY + 56);
          ctx.fillText('TOTAL DP+CV', 910, fY + 56);
          ctx.fillText('TRX+NT', 1085, fY + 56);
          ctx.fillText('TRAFFIC', 1190, fY + 56);

          let wy = fY + 66;
          const maxRows = Math.min(7, targetWks.length);
          targetWks.slice(-maxRows).forEach((w, idx) => {
            const isSel = w.weekNumber === activeWeekObj.weekNumber;
            roundRect(
              54,
              wy,
              W - 108,
              33,
              6,
              isSel ? '#ecfdf5' : idx % 2 === 0 ? '#f8fafc' : '#ffffff',
              isSel ? '#059669' : undefined,
              isSel ? 1.5 : 0
            );
            ctx.font = `${isSel ? '700' : '600'} 11px "JetBrains Mono", monospace`;
            ctx.fillStyle = '#0f172a';
            ctx.fillText(`W${w.weekNumber}`, 62, wy + 21);
            ctx.fillStyle = '#475569';
            ctx.fillText(w.shortRange, 135, wy + 21);
            ctx.fillStyle = '#312e81';
            ctx.fillText(`Rp ${formatNumberId(w.dpSales) || '0'}`, 275, wy + 21);
            ctx.fillStyle = '#6d28d9';
            ctx.fillText(`Rp ${formatNumberId(w.graphisSales) || '0'}`, 435, wy + 21);
            ctx.fillStyle = '#0369a1';
            ctx.fillText(`Rp ${formatNumberId(w.simplicitySales) || '0'}`, 590, wy + 21);
            ctx.fillStyle = '#5b21b6';
            ctx.fillText(`Rp ${formatNumberId(w.cvSales) || '0'}`, 755, wy + 21);
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 11px "JetBrains Mono", monospace';
            ctx.fillText(`Rp ${formatNumberId(w.grandSales) || '0'}`, 910, wy + 21);
            ctx.fillStyle = '#047857';
            ctx.fillText(`${w.totalTrxNota}`, 1085, wy + 21);
            ctx.fillStyle = '#334155';
            ctx.fillText(`${w.trafficTotal}`, 1190, wy + 21);
            wy += 37;
          });
        } else {
          // WEEKLY GROWTH FOCUS (3:4)
          drawGrowthOrContribChart(
            44,
            430,
            W - 88,
            300,
            'growth',
            `GRAFIK GROWTH SALES PER WEEKLY (WoW %) — ${streamBadge}`,
            weekGrowthItems
          );
          const halfW = (W - 88 - 16) / 2;
          drawGrowthOrContribChart(
            44,
            744,
            halfW,
            280,
            'contrib',
            'KONTRIBUSI PER WEEKLY (%)',
            weekContribItems
          );
          drawStackedDiagram(
            44 + halfW + 16,
            744,
            halfW,
            280,
            `TREN KOMPOSISI (${curr.labelA} vs ${curr.labelB})`,
            curr.labelA,
            curr.labelB,
            curr.colorA,
            curr.colorB,
            targetWks.reduce((s, w) => s + projectStream(w).primaryA, 0),
            targetWks.reduce((s, w) => s + projectStream(w).primaryB, 0),
            targetWks.map((w) => {
              const p = projectStream(w);
              return {
                label: `W${w.weekNumber}`,
                valA: p.primaryA,
                valB: p.primaryB,
                total: p.sales,
                isSelected: w.weekNumber === activeWeekObj.weekNumber,
              };
            })
          );

          drawBrandPanel(
            44,
            1038,
            W - 88,
            196,
            `DATA BRAND DIVISI 1 — WEEKLY ${activeWeekObj.weekNumber}`,
            activeWeekObj
          );

          // Full Weekly Growth Table bottom (with all 4 pillars)
          const tY = 1248;
          const tH = 500;
          roundRect(44, tY, W - 88, tH, 14, cardBg, cardBorder, 1.5);
          roundRect(44, tY, W - 88, 40, 14, '#0f172a');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(
            `TABEL PERTUMBUHAN (WoW %) & KONTRIBUSI PER WEEKLY — DP · BRAND · CV · TOTAL DP+CV`,
            60,
            tY + 25
          );
          ctx.fillStyle = '#64748b';
          ctx.font = '700 10.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('WEEK', 60, tY + 60);
          ctx.fillText('PERIODE', 130, tY + 60);
          ctx.fillText('REPORT DP', 265, tY + 60);
          ctx.fillText('GRAPHIS', 415, tY + 60);
          ctx.fillText('SIMPLICITY', 540, tY + 60);
          ctx.fillText('REPORT CV', 680, tY + 60);
          ctx.fillText('TOTAL DP+CV', 830, tY + 60);
          ctx.fillText('TRX+NT', 980, tY + 60);
          ctx.fillStyle = '#0369a1';
          ctx.font = '800 12px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('SHARE', 1072, tY + 60);
          ctx.fillStyle = '#047857';
          ctx.fillText('WoW %', 1178, tY + 60);

          let wy = tY + 72;
          targetWks.forEach((w, idx) => {
            const isSel = w.weekNumber === activeWeekObj.weekNumber;
            const gi = weekGrowthItems[idx]?.val ?? 0;
            const isBase = weekGrowthItems[idx]?.base;
            const ci = weekContribItems[idx]?.val ?? 0;
            roundRect(
              54,
              wy,
              W - 108,
              38,
              8,
              isSel ? '#ecfdf5' : idx % 2 === 0 ? '#f8fafc' : '#ffffff',
              isSel ? '#059669' : '#e2e8f0',
              isSel ? 1.5 : 1
            );
            ctx.font = `${isSel ? '700' : '600'} 11px "JetBrains Mono", monospace`;
            ctx.fillStyle = '#0f172a';
            ctx.fillText(`W${w.weekNumber}`, 62, wy + 24);
            ctx.fillStyle = '#475569';
            ctx.fillText(w.shortRange, 130, wy + 24);
            ctx.fillStyle = '#312e81';
            ctx.fillText(`Rp ${fmtShortRp(w.dpSales)}`, 265, wy + 24);
            ctx.fillStyle = '#6d28d9';
            ctx.fillText(`Rp ${fmtShortRp(w.graphisSales)}`, 415, wy + 24);
            ctx.fillStyle = '#0369a1';
            ctx.fillText(`Rp ${fmtShortRp(w.simplicitySales)}`, 540, wy + 24);
            ctx.fillStyle = '#5b21b6';
            ctx.fillText(`Rp ${fmtShortRp(w.cvSales)}`, 680, wy + 24);
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 11px "JetBrains Mono", monospace';
            ctx.fillText(`Rp ${formatNumberId(w.grandSales) || '0'}`, 830, wy + 24);
            ctx.fillStyle = '#047857';
            ctx.fillText(`${w.totalTrxNota}`, 980, wy + 24);
            ctx.fillStyle = '#0369a1';
            ctx.font = '800 13.5px "JetBrains Mono", monospace';
            ctx.fillText(`${ci.toFixed(1)}%`, 1072, wy + 24);
            ctx.fillStyle = isBase ? '#64748b' : gi >= 0 ? '#047857' : '#be123c';
            ctx.fillText(
              isBase ? 'Basis' : `${gi >= 0 ? '+' : ''}${gi.toFixed(1)}%`,
              1178,
              wy + 24
            );
            wy += 42;
          });
        }
      } else {
        // =====================================================================
        // LAYOUT 3: BULANAN 3:4 (1350 × 1800 px)
        // =====================================================================
        const mSrc = activeMonthObj || {
          ...annualTotals,
          monthKey: 'ALL',
          monthName: 'Tahun 2026',
          lastDay: 365,
        };
        const curr = projectStream(mSrc);
        const prev = prevMonthObj ? projectStream(prevMonthObj) : null;
        const gr = calcGrowth(curr.sales, prev ? prev.sales : null);
        const pp = curr.qty > 0 ? Math.round(curr.sales / curr.qty) : 0;
        const bs = curr.trx > 0 ? Math.round(curr.sales / curr.trx) : 0;
        const conv =
          curr.trafficTotal > 0 ? (curr.trx / curr.trafficTotal) * 100 : 0;

        draw4KpiStrip(184, 110, [
          {
            label: `Sales ${selectedMonthLabel} (${streamBadge})`,
            value: `Rp ${formatNumberId(curr.sales) || '0'}`,
            sub:
              gr === null
                ? `${mSrc.activeDays} hari aktif`
                : `${gr >= 0 ? '+' : ''}${gr.toFixed(1)}% vs ${prevMonthObj?.monthName.slice(0, 3)}`,
            accent: gr !== null && gr < 0 ? '#be123c' : '#047857',
          },
          {
            label: 'Total Qty Terjual',
            value: `${formatNumberId(curr.qty) || '0'} pcs`,
            sub: `Price Point: Rp ${formatNumberId(pp) || '0'}`,
            accent: '#4338ca',
          },
          {
            label: `Total ${curr.trxUnit}`,
            value: `${formatNumberId(curr.trx) || '0'}`,
            sub: `Big Size: Rp ${formatNumberId(bs) || '0'}`,
            accent: '#b45309',
          },
          {
            label: 'Traffic & Conv Rate',
            value: `${conv.toFixed(1)}% (${formatNumberId(curr.trafficTotal) || '0'})`,
            sub: `Pagi ${curr.pagiTraffic} · Siang ${curr.siangTraffic} org`,
            accent: '#0369a1',
          },
        ]);

        // 4-Pillar Comparison Strip (DP | CV | Brand | Total DP+CV)
        draw4PillarBreakdown(308, 108, mSrc);

        const annStreamTotal = projectStream(annualTotals).sales;
        const mContribItems = months12.map((m) => {
          const v = projectStream(m).sales;
          return {
            label: m.monthName.slice(0, 3),
            sub: `${m.activeDays}h`,
            val: annStreamTotal > 0 ? (v / annStreamTotal) * 100 : 0,
            isHighlight: m.monthKey === monthKey,
          };
        });

        const mGrowthItems = months12.map((m, idx) => {
          const prevM = idx > 0 ? months12[idx - 1] : null;
          const g = calcGrowth(
            projectStream(m).sales,
            prevM ? projectStream(prevM).sales : null
          );
          return {
            label: m.monthName.slice(0, 3),
            sub: `${m.activeDays}h`,
            val: g ?? 0,
            base: g === null,
            isHighlight: m.monthKey === monthKey,
          };
        });

        if (monthlyFocus === 'ringkasan') {
          const halfW = (W - 88 - 16) / 2;
          const midY = 430;
          const midH = 610;

          // Left: 12 Bulan Table (including Brand G/S + DP + CV + Total + Share + MoM)
          roundRect(44, midY, halfW, midH, 14, cardBg, cardBorder, 1.5);
          roundRect(44, midY, halfW, 40, 14, '#0f172a');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('REKAP 12 BULAN 2026 (DP · BRAND G/S · CV · TOTAL)', 60, midY + 25);

          ctx.fillStyle = '#64748b';
          ctx.font = '700 10px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('BULAN', 58, midY + 58);
          ctx.fillText('DP', 142, midY + 58);
          ctx.fillText('G / S', 225, midY + 58);
          ctx.fillText('CV', 338, midY + 58);
          ctx.fillText('TOTAL DP+CV', 415, midY + 58);
          ctx.fillStyle = '#0369a1';
          ctx.font = '800 11.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText('SHR', 532, midY + 58);
          ctx.fillStyle = '#047857';
          ctx.fillText('MoM', 588, midY + 58);

          let my = midY + 68;
          months12.forEach((m, idx) => {
            const isSel = m.monthKey === monthKey;
            const cVal = mContribItems[idx].val;
            const gVal = mGrowthItems[idx].val;
            const isBase = mGrowthItems[idx].base;
            roundRect(
              54,
              my,
              halfW - 20,
              38,
              7,
              isSel ? '#ecfdf5' : idx % 2 === 0 ? '#f8fafc' : '#ffffff',
              isSel ? '#059669' : '#e2e8f0',
              isSel ? 1.5 : 1
            );
            ctx.font = `${isSel ? '700' : '600'} 11px "Plus Jakarta Sans", Inter, system-ui, sans-serif`;
            ctx.fillStyle = '#0f172a';
            ctx.fillText(`${idx + 1}. ${m.monthName.slice(0, 3)}`, 60, my + 23);

            ctx.font = '600 10.5px "JetBrains Mono", monospace';
            ctx.fillStyle = '#312e81';
            ctx.fillText(fmtShortRp(m.dpSales), 142, my + 23);
            ctx.fillStyle = '#6d28d9';
            ctx.fillText(`${fmtShortRp(m.graphisSales)}/${fmtShortRp(m.simplicitySales)}`, 225, my + 23);
            ctx.fillStyle = '#5b21b6';
            ctx.fillText(fmtShortRp(m.cvSales), 338, my + 23);
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 11px "JetBrains Mono", monospace';
            ctx.fillText(`Rp ${fmtShortRp(m.grandSales)}`, 415, my + 23);
            ctx.fillStyle = '#0369a1';
            ctx.font = '800 12.5px "JetBrains Mono", monospace';
            ctx.fillText(`${cVal.toFixed(1)}%`, 532, my + 24);
            ctx.fillStyle = isBase ? '#64748b' : gVal >= 0 ? '#047857' : '#be123c';
            ctx.fillText(
              isBase ? 'Basis' : `${gVal >= 0 ? '+' : ''}${gVal.toFixed(0)}%`,
              588,
              my + 24
            );
            my += 41;
          });

          roundRect(54, midY + 570, halfW - 20, 28, 7, '#064e3b');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 11px "JetBrains Mono", monospace';
          ctx.fillText(
            `GRAND TOTAL 2026: Rp ${formatNumberId(annualTotals.grandSales) || '0'} (DP ${fmtShortRp(annualTotals.dpSales)} + CV ${fmtShortRp(annualTotals.cvSales)})`,
            64,
            midY + 588
          );

          // Right top: Stacked 12 Months Bar Chart
          const rx = 44 + halfW + 16;
          const pAnn = projectStream(annualTotals);
          drawStackedDiagram(
            rx,
            midY,
            halfW,
            298,
            `TREN 12 BULAN (${pAnn.labelA} vs ${pAnn.labelB})`,
            pAnn.labelA,
            pAnn.labelB,
            pAnn.colorA,
            pAnn.colorB,
            pAnn.primaryA,
            pAnn.primaryB,
            months12.map((m) => {
              const p = projectStream(m);
              return {
                label: m.monthName.slice(0, 3),
                valA: p.primaryA,
                valB: p.primaryB,
                total: p.sales,
                isSelected: m.monthKey === monthKey,
              };
            })
          );

          // Right bottom: Kontribusi % 12 Bulan
          drawGrowthOrContribChart(
            rx,
            midY + 312,
            halfW,
            298,
            'contrib',
            'KONTRIBUSI PENJUALAN PER BULAN (%)',
            mContribItems
          );

          // Full-width Data Brand Bulanan Panel (NO Pagi/Siang split)
          drawBrandPanel(
            44,
            1054,
            W - 88,
            204,
            `DATA BRAND DIVISI 1 — ${selectedMonthLabel.toUpperCase()}`,
            mSrc
          );

          // Bottom: MoM Growth Chart
          drawGrowthOrContribChart(
            44,
            1272,
            W - 88,
            284,
            'growth',
            `PERTUMBUHAN BULANAN (MoM GROWTH %) — ${streamBadge}`,
            mGrowthItems
          );

          // Annual 4-Box Footer Strip
          const fY = 1570;
          const fW = (W - 88 - 36) / 4;
          const fCards = [
            {
              label: 'TOTAL REPORT DP (2026)',
              val: `Rp ${formatNumberId(annualTotals.dpSales) || '0'}`,
              sub: `${formatNumberId(annualTotals.dpTrx) || '0'} Trx · ${formatNumberId(annualTotals.dpQty) || '0'} pcs`,
              c: '#4f46e5',
            },
            {
              label: 'TOTAL REPORT CV (2026)',
              val: `Rp ${formatNumberId(annualTotals.cvSales) || '0'}`,
              sub: `${formatNumberId(annualTotals.cvNota) || '0'} Nota · ${formatNumberId(annualTotals.cvQty) || '0'} pcs`,
              c: '#7c3aed',
            },
            {
              label: 'BRAND GRAPHIS vs SIMPLICITY',
              val: `G: ${fmtShortRp(annualTotals.graphisSales)} · S: ${fmtShortRp(annualTotals.simplicitySales)}`,
              sub: `Share: ${annualTotals.cvSales > 0 ? Math.round((annualTotals.graphisSales / annualTotals.cvSales) * 100) : 0}% vs ${annualTotals.cvSales > 0 ? Math.round((annualTotals.simplicitySales / annualTotals.cvSales) * 100) : 0}%`,
              c: '#0284c7',
            },
            {
              label: 'TOTAL DP+CV & TRAFFIC (2026)',
              val: `Rp ${fmtShortRp(annualTotals.grandSales)}`,
              sub: `Trf ${formatNumberId(annualTotals.trafficTotal) || '0'} · Conv ${annualTotals.trafficTotal > 0 ? ((annualTotals.totalTrxNota / annualTotals.trafficTotal) * 100).toFixed(1) : '0.0'}%`,
              c: '#059669',
            },
          ];
          fCards.forEach((fc, i) => {
            const fx = 44 + i * (fW + 12);
            roundRect(fx, fY, fW, 174, 12, cardBg, cardBorder, 1.5);
            ctx.fillStyle = fc.c;
            ctx.font = '700 11px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
            ctx.fillText(fc.label, fx + 16, fY + 32);
            ctx.fillStyle = '#0f172a';
            ctx.font = '800 19px "JetBrains Mono", monospace';
            ctx.fillText(clipText(ctx, fc.val, fW - 32), fx + 16, fY + 80);
            ctx.fillStyle = '#475569';
            ctx.font = '600 11.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
            ctx.fillText(clipText(ctx, fc.sub, fW - 32), fx + 16, fY + 128);
          });
        } else {
          // BULANAN KONTRIBUSI & GROWTH FOCUS (3:4)
          drawGrowthOrContribChart(
            44,
            430,
            W - 88,
            265,
            'contrib',
            `DIAGRAM KONTRIBUSI BULANAN (%) — ${streamBadge}`,
            mContribItems
          );
          drawGrowthOrContribChart(
            44,
            709,
            W - 88,
            265,
            'growth',
            `DIAGRAM PERTUMBUHAN BULANAN (MoM GROWTH %) — ${streamBadge}`,
            mGrowthItems
          );
          drawBrandPanel(
            44,
            988,
            W - 88,
            198,
            `DATA BRAND DIVISI 1 — ${selectedMonthLabel.toUpperCase()}`,
            mSrc
          );

          // 12-Month Matrix Cards (4 cols x 3 rows — each shows DP, CV, Brand G/S, Total DP+CV)
          const matY = 1200;
          const matH = 544;
          roundRect(44, matY, W - 88, matH, 14, cardBg, cardBorder, 1.5);
          roundRect(44, matY, W - 88, 40, 14, '#0f172a');
          ctx.fillStyle = '#ffffff';
          ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
          ctx.fillText(
            `MATRIKS KINERJA 12 BULAN 2026 (${streamBadge}) — DP · CV · BRAND · TOTAL DP+CV`,
            60,
            matY + 25
          );

          const cols = 4;
          const gX = 12;
          const gY = 12;
          const cW = (W - 116 - gX * 3) / cols;
          const cH = 150;
          months12.forEach((m, idx) => {
            const r = Math.floor(idx / cols);
            const c = idx % cols;
            const cx = 58 + c * (cW + gX);
            const cy = matY + 52 + r * (cH + gY);
            const isSel = m.monthKey === monthKey;
            const val = projectStream(m).sales;
            const contrib = mContribItems[idx].val;
            const grVal = mGrowthItems[idx].val;
            const isBase = mGrowthItems[idx].base;

            roundRect(
              cx,
              cy,
              cW,
              cH,
              10,
              isSel ? '#ecfdf5' : val > 0 ? '#f8fafc' : '#ffffff',
              isSel ? '#059669' : '#cbd5e1',
              isSel ? 2 : 1
            );
            ctx.fillStyle = '#0f172a';
            ctx.font = '700 12.5px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
            ctx.fillText(`${idx + 1}. ${m.monthName}`, cx + 12, cy + 23);

            ctx.font = '800 15.5px "JetBrains Mono", monospace';
            ctx.fillText(`Rp ${formatNumberId(val) || '0'}`, cx + 12, cy + 48);

            ctx.fillStyle = '#312e81';
            ctx.font = '600 10.5px "JetBrains Mono", monospace';
            ctx.fillText(
              `DP ${fmtShortRp(m.dpSales)} · CV ${fmtShortRp(m.cvSales)}`,
              cx + 12,
              cy + 72
            );

            ctx.fillStyle = '#6d28d9';
            ctx.fillText(
              `G ${fmtShortRp(m.graphisSales)} · S ${fmtShortRp(m.simplicitySales)}`,
              cx + 12,
              cy + 94
            );

            ctx.fillStyle = '#0369a1';
            ctx.font = '800 13.5px "JetBrains Mono", monospace';
            ctx.fillText(`Share: ${contrib.toFixed(1)}%`, cx + 12, cy + 118);

            ctx.textAlign = 'right';
            ctx.fillStyle = isBase ? '#64748b' : grVal >= 0 ? '#047857' : '#be123c';
            ctx.fillText(
              isBase ? 'Basis' : `${grVal >= 0 ? '+' : ''}${grVal.toFixed(1)}%`,
              cx + cW - 12,
              cy + 118
            );
            ctx.textAlign = 'left';

            ctx.fillStyle = '#475569';
            ctx.font = '600 10px "JetBrains Mono", monospace';
            ctx.fillText(
              `Total DP+CV: Rp ${fmtShortRp(m.grandSales)}`,
              cx + 12,
              cy + 138
            );
          });
        }
      }

      // Footer Watermark
      ctx.fillStyle = isLight ? '#475569' : '#94a3b8';
      ctx.font = '600 13px "Plus Jakarta Sans", Inter, system-ui, sans-serif';
      ctx.fillText(
        `REPORT BAZAR MATAHARI · TANJUNG · PT MDS RETAILING TBK (${ratioBadge})`,
        48,
        H - 18
      );
      ctx.textAlign = 'right';
      ctx.fillText(
        `Dicetak: ${new Date().toLocaleDateString('id-ID', {
          day: '2-digit',
          month: 'long',
          year: 'numeric',
        })}`,
        W - 48,
        H - 18
      );
      ctx.textAlign = 'left';
    } catch (err) {
      console.error('Error drawing dashboard screenshot canvas:', err);
    }
  }, [
    isOpen,
    mode,
    stream,
    monthlyFocus,
    weeklyFocus,
    isPortrait3x4,
    theme,
    weekNumber,
    dayTanggal,
    monthKey,
    activeWeekObj,
    prevWeekObj,
    activeDayObj,
    prevDayObj,
    activeDayIdx,
    activeMonthObj,
    prevMonthObj,
    annualTotals,
    weekly53,
    months12,
    brandOps,
    selectedMonthLabel,
  ]);

  if (!isOpen) return null;

  const getFileName = () => {
    const sSlug = stream === 'ALL' ? 'DP_CV' : stream;
    if (mode === 'harian') {
      return `Screenshot_${sSlug}_Harian_${activeDayObj.tanggal.replace(/\//g, '-')}_1x1.png`;
    }
    if (mode === 'weekly') {
      return `Screenshot_${sSlug}_Weekly_W${activeWeekObj.weekNumber}_${weeklyFocus}_3x4.png`;
    }
    return `Screenshot_${sSlug}_Bulanan_${selectedMonthLabel.replace(/\s+/g, '_')}_${monthlyFocus}_3x4.png`;
  };

  const handleDownloadPng = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    setDownloading(true);
    canvas.toBlob((blob) => {
      if (!blob) {
        const link = document.createElement('a');
        link.download = getFileName();
        link.href = canvas.toDataURL('image/png', 1.0);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => setDownloading(false), 400);
        return;
      }
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.download = getFileName();
      link.href = url;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => {
        URL.revokeObjectURL(url);
        setDownloading(false);
      }, 400);
    }, 'image/png', 1.0);
  };

  const handleNativeShare = async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    setSharing(true);
    try {
      const blob = await new Promise<Blob | null>((resolve) =>
        canvas.toBlob((b) => resolve(b), 'image/png', 1.0)
      );
      if (!blob) {
        handleDownloadPng();
        return;
      }
      const file = new File([blob], getFileName(), { type: 'image/png' });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({
          files: [file],
          title: getFileName().replace('.png', '').replace(/_/g, ' '),
          text: `Report Bazar Matahari Tanjung — W${activeWeekObj.weekNumber}`,
        });
        return;
      }
      if (navigator.clipboard && 'ClipboardItem' in window) {
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        setCopied(true);
        setTimeout(() => setCopied(false), 2400);
      } else {
        handleDownloadPng();
      }
    } catch (err: any) {
      if (err?.name !== 'AbortError') handleDownloadPng();
    } finally {
      setSharing(false);
    }
  };

  const handleCopyClipboard = async () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    try {
      const blob = await new Promise<Blob | null>((resolve) =>
        canvas.toBlob((b) => resolve(b), 'image/png', 1.0)
      );
      if (!blob) return;
      if (navigator.clipboard && 'ClipboardItem' in window) {
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        setCopied(true);
        setTimeout(() => setCopied(false), 2400);
        return;
      }
      handleDownloadPng();
    } catch {
      handleDownloadPng();
    }
  };

  return createPortal(
    <div
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
      className="fixed inset-0 z-[9999] bg-slate-950/85 backdrop-blur-sm flex flex-col items-center justify-start p-2 sm:p-4 overflow-y-auto overscroll-contain"
    >
      <div className="bg-white border border-slate-200 rounded-2xl shadow-2xl w-full max-w-5xl overflow-hidden my-2 sm:my-4 shrink-0">
        {/* Top Modal Header */}
        <div className="bg-slate-900 text-white px-4 py-3 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-500 text-slate-950 flex items-center justify-center font-bold">
              <Camera className="w-4 h-4" />
            </div>
            <div>
              <h3 className="text-sm font-bold">
                Studio Screenshot Semua Dashboard —{' '}
                {isPortrait3x4 ? 'Rasio 3:4 (1350 × 1800 px)' : 'Rasio 1:1 (1400 × 1400 px)'}
              </h3>
              <p className="text-[11px] text-slate-300">
                Mencakup DP + CV, Report DP, Report CV, Traffic &amp; Conv Rate, serta Data Brand Graphis (6201) &amp; Simplicity (6244)
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-slate-200 transition-colors cursor-pointer"
            title="Tutup"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Main Controls Bar */}
        <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2.5">
            {/* Report Stream Selector */}
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="text-xs font-bold text-slate-600">Laporan:</span>
              <div className="inline-flex items-center p-0.5 bg-slate-200/70 rounded-lg text-xs">
                {(
                  [
                    { id: 'ALL', label: 'DP + CV' },
                    { id: 'DP', label: 'Report DP' },
                    { id: 'CV', label: 'Report CV' },
                  ] as { id: DashboardReportStream; label: string }[]
                ).map((s) => (
                  <button
                    key={s.id}
                    type="button"
                    onClick={() => setStream(s.id)}
                    className={`px-2.5 py-1 rounded-md font-bold cursor-pointer whitespace-nowrap ${
                      stream === s.id
                        ? 'bg-slate-900 text-white shadow-2xs'
                        : 'text-slate-700 hover:bg-white/60'
                    }`}
                  >
                    {s.label}
                  </button>
                ))}
              </div>

              <span className="text-xs font-bold text-slate-600 ml-2">Mode:</span>
              <div className="flex flex-wrap items-center gap-1">
                <button
                  type="button"
                  onClick={() => setMode('harian')}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    mode === 'harian'
                      ? 'bg-amber-400 text-slate-950'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <Calendar className="w-3.5 h-3.5" />
                  <span>Harian 1:1</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setMode('weekly');
                    setWeeklyFocus('ringkasan');
                  }}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    mode === 'weekly' && weeklyFocus === 'ringkasan'
                      ? 'bg-emerald-700 text-white'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <CalendarRange className="w-3.5 h-3.5" />
                  <span>Weekly 3:4</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setMode('weekly');
                    setWeeklyFocus('growth_weekly');
                  }}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    mode === 'weekly' && weeklyFocus === 'growth_weekly'
                      ? 'bg-teal-700 text-white'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <TrendingUp className="w-3.5 h-3.5" />
                  <span>Growth Weekly 3:4</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setMode('bulanan');
                    setMonthlyFocus('ringkasan');
                  }}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    mode === 'bulanan' && monthlyFocus === 'ringkasan'
                      ? 'bg-sky-700 text-white'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <CalendarDays className="w-3.5 h-3.5" />
                  <span>Bulanan 3:4</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setMode('bulanan');
                    setMonthlyFocus('kontribusi_growth');
                  }}
                  className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-bold cursor-pointer ${
                    mode === 'bulanan' && monthlyFocus === 'kontribusi_growth'
                      ? 'bg-indigo-700 text-white'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  <LayoutDashboard className="w-3.5 h-3.5" />
                  <span>Bulanan Kontribusi 3:4</span>
                </button>
              </div>
            </div>

            {/* Theme + Export Buttons */}
            <div className="flex flex-wrap items-center gap-1.5">
              <div className="inline-flex items-center bg-white border border-slate-200 rounded-lg p-0.5 text-[11px] font-semibold">
                {(['navy', 'emerald', 'light'] as ThemeStyle[]).map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => setTheme(t)}
                    className={`px-2 py-1 rounded-md cursor-pointer capitalize ${
                      theme === t ? 'bg-slate-900 text-white' : 'text-slate-600'
                    }`}
                  >
                    {t === 'light' ? 'Terang' : t === 'navy' ? 'Navy' : 'Emerald'}
                  </button>
                ))}
              </div>

              <button
                type="button"
                onClick={() => setZoomLevel((z) => (z === 'fit' ? 'full' : 'fit'))}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 cursor-pointer"
              >
                {zoomLevel === 'fit' ? (
                  <>
                    <ZoomIn className="w-3.5 h-3.5" />
                    <span className="hidden sm:inline">100%</span>
                  </>
                ) : (
                  <>
                    <ZoomOut className="w-3.5 h-3.5" />
                    <span className="hidden sm:inline">Pas</span>
                  </>
                )}
              </button>

              <button
                type="button"
                onClick={handleNativeShare}
                disabled={sharing}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-bold bg-amber-400 hover:bg-amber-500 text-slate-950 cursor-pointer"
              >
                <Share2 className="w-3.5 h-3.5" />
                <span>Bagikan</span>
              </button>

              <button
                type="button"
                onClick={handleCopyClipboard}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-bold bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 cursor-pointer"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copied ? 'Tersalin' : 'Salin'}</span>
              </button>

              <button
                type="button"
                onClick={handleDownloadPng}
                disabled={downloading}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-700 hover:bg-emerald-800 text-white cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>{downloading ? 'Mengunduh...' : 'Download PNG'}</span>
              </button>
            </div>
          </div>

          {/* Contextual Period Pickers */}
          {mode === 'harian' && (
            <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <div className="flex flex-wrap items-center gap-1">
                {activeWeekObj.days.map((d) => {
                  const isSel = d.tanggal === activeDayObj.tanggal;
                  return (
                    <button
                      key={d.tanggal}
                      type="button"
                      onClick={() => setDayTanggal(d.tanggal)}
                      className={`px-2.5 py-1 rounded-md text-xs font-bold cursor-pointer ${
                        isSel
                          ? 'bg-slate-900 text-white'
                          : d.isWeekend
                            ? 'bg-rose-50 text-rose-800 border border-rose-200'
                            : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      {d.dayName.slice(0, 3)} ({d.tanggal.slice(0, 5)})
                    </button>
                  );
                })}
              </div>

              <select
                aria-label="Pilih Minggu"
                value={weekNumber}
                onChange={(e) => {
                  const w = Number(e.target.value);
                  setWeekNumber(w);
                  const found = weekly53.find((x) => x.weekNumber === w);
                  if (found && found.days[0]) setDayTanggal(found.days[0].tanggal);
                }}
                className="px-2.5 py-1 rounded-lg text-xs font-bold bg-white border border-slate-200 text-slate-900 cursor-pointer"
              >
                {weekly53.map((w) => (
                  <option key={w.weekNumber} value={w.weekNumber}>
                    W{w.weekNumber} ({w.shortRange}){w.grandSales > 0 ? ` — Rp ${fmtShortRp(w.grandSales)}` : ''}
                  </option>
                ))}
              </select>
            </div>
          )}

          {mode === 'weekly' && (
            <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <div className="text-xs font-semibold text-slate-600">
                Fokus: <strong className="text-slate-950">W{activeWeekObj.weekNumber} ({activeWeekObj.periodeText})</strong>
              </div>
              <select
                aria-label="Pilih Minggu Weekly"
                value={weekNumber}
                onChange={(e) => setWeekNumber(Number(e.target.value))}
                className="px-2.5 py-1 rounded-lg text-xs font-bold bg-white border border-slate-200 text-slate-900 cursor-pointer"
              >
                {weekly53.map((w) => (
                  <option key={w.weekNumber} value={w.weekNumber}>
                    W{w.weekNumber} ({w.monthName}): {w.shortRange}
                    {w.grandSales > 0 ? ` — Rp ${fmtShortRp(w.grandSales)}` : ''}
                  </option>
                ))}
              </select>
            </div>
          )}

          {mode === 'bulanan' && (
            <div className="pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
              <div className="text-xs font-semibold text-slate-600">
                Fokus: <strong className="text-slate-950">{selectedMonthLabel}</strong>
              </div>
              <select
                aria-label="Pilih Bulan"
                value={monthKey}
                onChange={(e) => {
                  const val = e.target.value;
                  setMonthKey(val);
                  if (onSelectMonthKey) onSelectMonthKey(val);
                }}
                className="px-2.5 py-1 rounded-lg text-xs font-bold bg-white border border-slate-200 text-slate-900 cursor-pointer"
              >
                <option value="ALL">Semua 12 Bulan (Januari – Desember 2026)</option>
                {months12.map((m) => (
                  <option key={m.monthKey} value={m.monthKey}>
                    {m.monthName} 2026{m.grandSales > 0 ? ` — Rp ${fmtShortRp(m.grandSales)}` : ''}
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Live Canvas Preview */}
        <div className="p-4 sm:p-6 bg-slate-900/95 flex flex-col items-center justify-center max-h-[82vh] overflow-auto">
          <div
            className={`w-full transition-all duration-200 ${
              zoomLevel === 'full'
                ? isPortrait3x4
                  ? 'min-w-[900px] aspect-[3/4]'
                  : 'min-w-[900px] aspect-square'
                : isPortrait3x4
                  ? 'max-w-[540px] aspect-[3/4]'
                  : 'max-w-[640px] aspect-square'
            } rounded-xl overflow-hidden shadow-2xl border border-white/20 bg-slate-950 flex items-center justify-center`}
          >
            <canvas ref={canvasRef} className="w-full h-full object-contain block" />
          </div>
          <div className="mt-3 flex items-center gap-2 text-xs text-slate-300 text-center">
            <Sparkles className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
            <span>
              Tampilan screenshot menyertakan <strong>{stream === 'ALL' ? 'DP + CV' : stream === 'DP' ? 'Report DP' : 'Report CV'}</strong>, Traffic &amp; Conversion Rate, Diagram Tren, serta <strong>Data Brand (Graphis 6201 &amp; Simplicity 6244)</strong>.
            </span>
          </div>
        </div>
      </div>
    </div>,
    document.body
  );
};
