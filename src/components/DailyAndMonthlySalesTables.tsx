import React, { useState, useMemo, useEffect, useRef, useCallback } from 'react';
import {
  Calendar,
  CalendarRange,
  CalendarDays,
  Eraser,
  FileSpreadsheet,
  RotateCcw,
  Trash2,
  Check,
  X,
  Save,
} from 'lucide-react';
import {
  WeeklySalesTable,
  DailySalesRecord,
  ShiftEntry,
  calculateDayMetrics,
  calculateWeekMetrics,
  formatNumberId,
  getIndoDayName,
  isWeekendDate,
  ANNUAL_2026_WEEKS,
  createWeeklyTableFromAnnualSchedule,
} from '../data/initialData';
import { PhotoStyleSalesDiagram } from './PhotoStyleSalesDiagram';
import { MonthlyAutoSummaryCard } from './MonthlyAutoSummaryCard';
import { GrowthAndContributionBarDiagram } from './GrowthAndContributionBarDiagram';

export type SalesTableHorizon = 'harian' | 'weekly' | 'bulanan';

const MONTH_LIST = [
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

function parseNumericInput(raw: string): number | null {
  const cleaned = raw.replace(/[^0-9]/g, '');
  if (!cleaned) return null;
  const num = Number(cleaned);
  return Number.isNaN(num) ? null : num;
}

interface CalendarDayRowItem {
  dayOfMonth: number; // 1 .. 31
  slashDate: string; // DD/MM/2026
  dayName: string; // Senin .. Minggu
  isWeekend: boolean;
  weekNumber: number; // W1 .. W53
  dayIndexInWeek: number; // 0 .. 6
  isWeekLoaded: boolean;
  record: DailySalesRecord;
  pagiSales: number;
  siangSales: number;
  totalSales: number;
  pagiQty: number;
  siangQty: number;
  totalQty: number;
  pagiTrx: number;
  siangTrx: number;
  totalTrx: number;
  basketSize: number | null;
  apc: number | null;
}

interface DailySalesTableProps {
  tables: WeeklySalesTable[];
  onChangeTables: (next: WeeklySalesTable[]) => void;
  selectedWeeklyFilter: string;
  onSelectWeeklyFilter: (id: string) => void;
  showDivZero: boolean;
  classicExcelColors: boolean;
  staffFilter: string;
  onChangeStaffFilter: (staff: string) => void;
  searchQuery: string;
  onChangeSearchQuery: (q: string) => void;
  onSelectHorizon?: (h: 'harian' | 'weekly' | 'bulanan') => void;
  onToggleTheme?: () => void;
  onToggleDivZero?: () => void;
  onManualSave?: () => void;
}

const STORAGE_KEY_SELECTED_MONTH = 'bazar_kapuas_selected_month_v1';

export const DailySalesSpreadsheetView: React.FC<DailySalesTableProps> = ({
  tables,
  onChangeTables,
  selectedWeeklyFilter,
  onSelectWeeklyFilter,
  showDivZero,
  classicExcelColors,
  staffFilter,
  onChangeStaffFilter,
  searchQuery,
  onSelectHorizon,
  onToggleTheme,
  onToggleDivZero,
  onManualSave,
}) => {
  // Default month based on saved localStorage or currently selected Weekly table (e.g., '10' for Oktober 2026)
  const initialMonth = useMemo(() => {
    try {
      const savedM = localStorage.getItem(STORAGE_KEY_SELECTED_MONTH);
      if (savedM && MONTH_LIST.some((m) => m.key === savedM)) {
        return savedM;
      }
    } catch {}
    const filledTables = tables.filter((t) =>
      t.days.some((d) => (d.pagi?.sales ?? 0) + (d.siang?.sales ?? 0) > 0)
    );
    const activeT =
      tables.find((t) => t.id === selectedWeeklyFilter) ||
      filledTables[filledTables.length - 1] ||
      tables[0];
    if (activeT) {
      const refSlash = activeT.days[3]?.tanggal || activeT.days[0]?.tanggal || '';
      const parts = refSlash.split('/');
      if (parts.length === 3 && parts[1] && parts[2] === '2026') {
        return parts[1].padStart(2, '0');
      }
    }
    return '11';
  }, []);

  const [selectedMonthKey, setSelectedMonthKey] = useState<string>(initialMonth);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_SELECTED_MONTH, selectedMonthKey);
    } catch {}
  }, [selectedMonthKey]);
  const [dateRangeFilter, setDateRangeFilter] = useState<
    'ALL' | '1_10' | '11_20' | '21_31' | 'FILLED'
  >('ALL');
  const [calendarExactDateFilter, setCalendarExactDateFilter] = useState<string>('');
  const [dayTypeFilter, setDayTypeFilter] = useState<'ALL' | 'WEEKDAY' | 'WEEKEND'>('ALL');
  const [viewFormat, setViewFormat] = useState<
    'TABLE_VERTICAL' | 'MATRIX_HORIZONTAL' | 'DETAIL_SHIFT'
  >('TABLE_VERTICAL');
  const [confirmDailyResetMode, setConfirmDailyResetMode] = useState<
    null | 'NUMBERS' | 'REASONS' | 'ALL_MONTH'
  >(null);
  const [modifiedDailyCells, setModifiedDailyCells] = useState<
    Record<string, { type: 'input' | 'calc'; seq: number }>
  >({});
  const [lastDailyEditBanner, setLastDailyEditBanner] = useState<{
    cellLabel: string;
    newValue: string;
  } | null>(null);
  const dailyCellTimersRef = useRef<Record<string, ReturnType<typeof setTimeout>>>({});

  const triggerDailyCellFlash = useCallback(
    (
      inputKeys: string[],
      calcKeys: string[] = [],
      bannerInfo?: { cellLabel: string; newValue: string }
    ) => {
      if (bannerInfo) {
        setLastDailyEditBanner(bannerInfo);
      }
      setModifiedDailyCells((prev) => {
        const next = { ...prev };
        inputKeys.forEach((k) => {
          next[k] = { type: 'input', seq: (prev[k]?.seq || 0) + 1 };
        });
        calcKeys.forEach((k) => {
          next[k] = { type: 'calc', seq: (prev[k]?.seq || 0) + 1 };
        });
        return next;
      });

      [...inputKeys, ...calcKeys].forEach((k) => {
        if (dailyCellTimersRef.current[k]) {
          clearTimeout(dailyCellTimersRef.current[k]);
        }
        dailyCellTimersRef.current[k] = setTimeout(() => {
          setModifiedDailyCells((prev) => {
            if (!prev[k]) return prev;
            const copy = { ...prev };
            delete copy[k];
            return copy;
          });
        }, 4200);
      });
    },
    []
  );

  const getDailyFlashClass = (cellKey: string): string => {
    const entry = modifiedDailyCells[cellKey];
    if (!entry) return '';
    return entry.type === 'input' ? 'cell-just-modified' : 'cell-calc-updated';
  };

  const currentMonthMeta =
    MONTH_LIST.find((m) => m.key === selectedMonthKey) || MONTH_LIST[9];
  const currentMonthIdx = MONTH_LIST.findIndex((m) => m.key === selectedMonthKey);

  // Build all calendar days from Tanggal 1 to Tanggal 28/30/31 for selectedMonthKey
  const calendarDays: CalendarDayRowItem[] = useMemo(() => {
    const list: CalendarDayRowItem[] = [];
    const daysCount = currentMonthMeta.daysCount;

    for (let d = 1; d <= daysCount; d++) {
      const dd = String(d).padStart(2, '0');
      const slashDate = `${dd}/${selectedMonthKey}/2026`;
      const dayName = getIndoDayName(slashDate);
      const wknd = isWeekendDate(slashDate);

      // Find in existing loaded tables first
      let foundWeekNum = 41;
      let foundDayIdx = 0;
      let isLoaded = false;
      let foundRecord: DailySalesRecord | null = null;

      for (const t of tables) {
        const idx = t.days.findIndex((dayRec) => dayRec.tanggal === slashDate);
        if (idx !== -1) {
          foundWeekNum = t.weekNumber;
          foundDayIdx = idx;
          isLoaded = true;
          foundRecord = t.days[idx];
          break;
        }
      }

      // If not in loaded tables, locate in ANNUAL_2026_WEEKS
      if (!foundRecord) {
        const sched = ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(slashDate));
        if (sched) {
          foundWeekNum = sched.weekNumber;
          foundDayIdx = sched.slashDates.indexOf(slashDate);
        }
        foundRecord = {
          id: `cal-${selectedMonthKey}-${dd}`,
          tanggal: slashDate,
          pagi: { nama: '', sales: null, qty: null, trx: null },
          siang: { nama: '', sales: null, qty: null, trx: null },
          reasonPagi: '',
          reasonSiang: '',
        };
      }

      const pagiSales = foundRecord.pagi.sales ?? 0;
      const siangSales = foundRecord.siang.sales ?? 0;
      const totalSales = pagiSales + siangSales;

      const pagiQty = foundRecord.pagi.qty ?? 0;
      const siangQty = foundRecord.siang.qty ?? 0;
      const totalQty = pagiQty + siangQty;

      const pagiTrx = foundRecord.pagi.trx ?? 0;
      const siangTrx = foundRecord.siang.trx ?? 0;
      const totalTrx = pagiTrx + siangTrx;

      const basketSize = totalTrx > 0 ? Math.round(totalSales / totalTrx) : null;
      const apc = totalQty > 0 ? Math.round(totalSales / totalQty) : null;

      list.push({
        dayOfMonth: d,
        slashDate,
        dayName,
        isWeekend: wknd,
        weekNumber: foundWeekNum,
        dayIndexInWeek: foundDayIdx,
        isWeekLoaded: isLoaded,
        record: foundRecord,
        pagiSales,
        siangSales,
        totalSales,
        pagiQty,
        siangQty,
        totalQty,
        pagiTrx,
        siangTrx,
        totalTrx,
        basketSize,
        apc,
      });
    }

    return list;
  }, [tables, selectedMonthKey, currentMonthMeta]);

  // Update Total Harian (Sales, Qty, or Trx) directly for a specific date (1..31)
  const handleUpdateDailyTotalMetric = (
    item: CalendarDayRowItem,
    metric: 'sales' | 'qty' | 'trx',
    rawVal: string
  ) => {
    const newTotal = parseNumericInput(rawVal);

    // Ensure the week table exists in `tables`
    let baseTables = [...tables];
    let targetTable = baseTables.find((t) => t.weekNumber === item.weekNumber);

    if (!targetTable) {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === item.weekNumber);
      if (!sched) return;
      targetTable = createWeeklyTableFromAnnualSchedule(sched);
      baseTables = [...baseTables, targetTable].sort((a, b) => a.weekNumber - b.weekNumber);
    }

    const nextTables = baseTables.map((t) => {
      if (t.weekNumber !== item.weekNumber) return t;
      const nextDays = t.days.map((d) => {
        if (d.tanggal !== item.slashDate) return d;

        if (newTotal === null) {
          return {
            ...d,
            pagi: { ...d.pagi, [metric]: null },
            siang: { ...d.siang, [metric]: null },
          };
        }

        const currentPagi = d.pagi[metric] ?? 0;
        const currentSiang = d.siang[metric] ?? 0;

        // If Siang is empty/zero, store directly in Pagi; if both exist, adjust Siang so Pagi + Siang === newTotal
        if (currentSiang === 0 && currentPagi === 0) {
          return {
            ...d,
            pagi: { ...d.pagi, [metric]: newTotal },
          };
        } else if (currentSiang === 0) {
          return {
            ...d,
            pagi: { ...d.pagi, [metric]: newTotal },
          };
        } else if (newTotal >= currentPagi) {
          return {
            ...d,
            siang: { ...d.siang, [metric]: newTotal - currentPagi },
          };
        } else {
          return {
            ...d,
            pagi: { ...d.pagi, [metric]: newTotal },
            siang: { ...d.siang, [metric]: 0 },
          };
        }
      });
      return { ...t, days: nextDays };
    });

    onChangeTables(nextTables);

    const calcKeys =
      metric === 'sales'
        ? [
            `${item.slashDate}-basketSize`,
            `${item.slashDate}-apc`,
            'footer-total-sales',
            'footer-basket-size',
            'footer-apc',
          ]
        : metric === 'qty'
          ? [`${item.slashDate}-apc`, 'footer-total-qty', 'footer-apc']
          : [`${item.slashDate}-basketSize`, 'footer-total-trx', 'footer-basket-size'];

    const metricLabel =
      metric === 'sales'
        ? 'Total Sales Harian'
        : metric === 'qty'
          ? 'Qty Harian'
          : 'Trx Harian';
    const formattedVal =
      newTotal !== null
        ? metric === 'sales'
          ? `Rp ${formatNumberId(newTotal)}`
          : formatNumberId(newTotal)
        : '0';

    triggerDailyCellFlash([`${item.slashDate}-total-${metric}`], calcKeys, {
      cellLabel: `${metricLabel} (${item.slashDate})`,
      newValue: formattedVal,
    });
  };

  // Update Shift Pagi or Siang field for a specific date (1..31)
  const handleUpdateShiftMetric = (
    item: CalendarDayRowItem,
    shift: 'pagi' | 'siang',
    field: keyof ShiftEntry,
    rawVal: string
  ) => {
    let baseTables = [...tables];
    let targetTable = baseTables.find((t) => t.weekNumber === item.weekNumber);

    if (!targetTable) {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === item.weekNumber);
      if (!sched) return;
      targetTable = createWeeklyTableFromAnnualSchedule(sched);
      baseTables = [...baseTables, targetTable].sort((a, b) => a.weekNumber - b.weekNumber);
    }

    const parsedNum = field === 'nama' ? null : parseNumericInput(rawVal);

    const nextTables = baseTables.map((t) => {
      if (t.weekNumber !== item.weekNumber) return t;
      const nextDays = t.days.map((d) => {
        if (d.tanggal !== item.slashDate) return d;
        const curShift = d[shift];
        const nextShift: ShiftEntry =
          field === 'nama'
            ? { ...curShift, nama: rawVal }
            : { ...curShift, [field]: parsedNum };
        return { ...d, [shift]: nextShift };
      });
      return { ...t, days: nextDays };
    });

    onChangeTables(nextTables);

    const calcKeys =
      field === 'sales'
        ? [
            `${item.slashDate}-total-sales`,
            `${item.slashDate}-basketSize`,
            `${item.slashDate}-apc`,
            `footer-${shift}-sales`,
            'footer-total-sales',
            'footer-basket-size',
            'footer-apc',
          ]
        : [];

    triggerDailyCellFlash([`${item.slashDate}-${shift}-${field}`], calcKeys, {
      cellLabel: `${shift === 'pagi' ? 'Sales Pagi' : 'Sales Siang'} (${item.slashDate})`,
      newValue:
        field === 'nama'
          ? rawVal
          : parsedNum !== null
            ? `Rp ${formatNumberId(parsedNum)}`
            : '0',
    });
  };

  const handleClearSingleDate = (item: CalendarDayRowItem) => {
    const nextTables = tables.map((t) => {
      if (t.weekNumber !== item.weekNumber) return t;
      return {
        ...t,
        days: t.days.map((d) =>
          d.tanggal === item.slashDate
            ? {
                ...d,
                pagi: { nama: '', sales: null, qty: null, trx: null },
                siang: { nama: '', sales: null, qty: null, trx: null },
                reasonPagi: '',
                reasonSiang: '',
              }
            : d
        ),
      };
    });
    onChangeTables(nextTables);
  };

  // Reset tools for Sales Daily (all dates 1..30/31 in selectedMonthKey)
  const handleResetDailyMonthNumbers = () => {
    const monthSuffix = `/${selectedMonthKey}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(monthSuffix)
          ? {
              ...d,
              pagi: { ...d.pagi, sales: null, qty: null, trx: null },
              siang: { ...d.siang, sales: null, qty: null, trx: null },
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmDailyResetMode(null);
  };

  const handleResetDailyMonthReasons = () => {
    const monthSuffix = `/${selectedMonthKey}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(monthSuffix)
          ? {
              ...d,
              pagi: { ...d.pagi, nama: '' },
              siang: { ...d.siang, nama: '' },
              reasonPagi: '',
              reasonSiang: '',
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmDailyResetMode(null);
  };

  const handleResetDailyMonthAll = () => {
    const monthSuffix = `/${selectedMonthKey}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(monthSuffix)
          ? {
              ...d,
              pagi: { nama: '', sales: null, qty: null, trx: null },
              siang: { nama: '', sales: null, qty: null, trx: null },
              reasonPagi: '',
              reasonSiang: '',
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmDailyResetMode(null);
  };

  // Filter days based on dateRangeFilter, dayTypeFilter, staffFilter, searchQuery
  const filteredCalendarDays = useMemo(() => {
    return calendarDays.filter((item) => {
      if (dateRangeFilter === '1_10' && (item.dayOfMonth < 1 || item.dayOfMonth > 10)) {
        return false;
      }
      if (dateRangeFilter === '11_20' && (item.dayOfMonth < 11 || item.dayOfMonth > 20)) {
        return false;
      }
      if (dateRangeFilter === '21_31' && item.dayOfMonth < 21) {
        return false;
      }
      if (dateRangeFilter === 'FILLED' && item.totalSales <= 0 && item.totalQty <= 0 && item.totalTrx <= 0) {
        return false;
      }
      if (calendarExactDateFilter) {
        const parts = calendarExactDateFilter.split('-');
        if (parts.length === 3) {
          const targetSlash = `${parts[2]}/${parts[1]}/${parts[0]}`;
          if (item.slashDate !== targetSlash) return false;
        }
      }
      if (dayTypeFilter === 'WEEKDAY' && item.isWeekend) return false;
      if (dayTypeFilter === 'WEEKEND' && !item.isWeekend) return false;

      if (staffFilter !== 'ALL') {
        const s = staffFilter.toLowerCase();
        const mp = (item.record.pagi.nama || '').toLowerCase().includes(s);
        const ms = (item.record.siang.nama || '').toLowerCase().includes(s);
        if (!mp && !ms) return false;
      }
      if (searchQuery.trim() !== '') {
        const q = searchQuery.toLowerCase();
        const hay = `${item.slashDate} ${item.dayName} W${item.weekNumber} ${item.record.pagi.nama} ${item.record.siang.nama} ${item.record.reasonPagi} ${item.record.reasonSiang}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [calendarDays, dateRangeFilter, dayTypeFilter, staffFilter, searchQuery]);

  // Month & Filtered Totals
  const monthTotals = useMemo(() => {
    let sales = 0;
    let qty = 0;
    let trx = 0;
    let pagiSales = 0;
    let siangSales = 0;
    let weekdaySales = 0;
    let weekendSales = 0;
    let activeDays = 0;

    calendarDays.forEach((c) => {
      sales += c.totalSales;
      qty += c.totalQty;
      trx += c.totalTrx;
      pagiSales += c.pagiSales;
      siangSales += c.siangSales;
      if (c.isWeekend) weekendSales += c.totalSales;
      else weekdaySales += c.totalSales;
      if (c.totalSales > 0) activeDays += 1;
    });

    const basketSize = trx > 0 ? Math.round(sales / trx) : null;
    const apc = qty > 0 ? Math.round(sales / qty) : null;

    return {
      sales,
      qty,
      trx,
      pagiSales,
      siangSales,
      weekdaySales,
      weekendSales,
      activeDays,
      basketSize,
      apc,
    };
  }, [calendarDays]);

  const filteredTotals = useMemo(() => {
    const sales = filteredCalendarDays.reduce((s, c) => s + c.totalSales, 0);
    const qty = filteredCalendarDays.reduce((s, c) => s + c.totalQty, 0);
    const trx = filteredCalendarDays.reduce((s, c) => s + c.totalTrx, 0);
    const pagiSales = filteredCalendarDays.reduce((s, c) => s + c.pagiSales, 0);
    const siangSales = filteredCalendarDays.reduce((s, c) => s + c.siangSales, 0);
    const pagiQty = filteredCalendarDays.reduce((s, c) => s + c.pagiQty, 0);
    const siangQty = filteredCalendarDays.reduce((s, c) => s + c.siangQty, 0);
    const pagiTrx = filteredCalendarDays.reduce((s, c) => s + c.pagiTrx, 0);
    const siangTrx = filteredCalendarDays.reduce((s, c) => s + c.siangTrx, 0);
    const basketSize = trx > 0 ? Math.round(sales / trx) : null;
    const apc = qty > 0 ? Math.round(sales / qty) : null;

    return {
      sales,
      qty,
      trx,
      pagiSales,
      siangSales,
      pagiQty,
      siangQty,
      pagiTrx,
      siangTrx,
      basketSize,
      apc,
    };
  }, [filteredCalendarDays]);

  const redHeaderClass = classicExcelColors
    ? 'bg-[#C00000] text-white border border-slate-900'
    : 'bg-rose-800 text-white border border-slate-300';

  const yellowTotalClass = classicExcelColors
    ? 'bg-[#FFFF00] text-slate-950 border border-slate-900 font-bold'
    : 'bg-amber-200/90 text-slate-950 border border-slate-300 font-bold';

  const goldGrandTotalClass = classicExcelColors
    ? 'bg-[#FFC000] text-slate-950 border border-slate-900 font-bold'
    : 'bg-amber-400 text-slate-950 border border-slate-300 font-bold';

  const renderCalculatedCell = (val: number | null | undefined) => {
    if (val === null || val === undefined || val === 0) {
      return showDivZero ? (
        <span className="text-slate-500 font-semibold">#DIV/0!</span>
      ) : (
        <span className="text-slate-400">0</span>
      );
    }
    return <span>{formatNumberId(val)}</span>;
  };

  return (
    <div className="space-y-3.5 w-full max-w-full min-w-0">
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
        {/* Clean Integrated Dark Banner Header: Mode Switcher + Pilih Bulan */}
        <div className="px-3.5 sm:px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-200 space-y-3 w-full max-w-full overflow-hidden">
          {/* Row 1: Title + Integrated Harian / Weekly / Bulanan Switcher */}
          <div className="flex flex-wrap items-center justify-between gap-2.5 w-full max-w-full min-w-0">
            <div className="min-w-0 max-w-full">
              <div className="flex flex-wrap items-center gap-1.5 text-[11px] font-bold text-amber-300 mb-0.5">
                <span>REPORT BAZAR MATAHARI</span>
                <span className="text-slate-400">·</span>
                <span className="text-white">Tanjung</span>
                <span className="text-slate-400">·</span>
                <span className="text-rose-300">PT MDS Retailing Tbk</span>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-sm sm:text-base font-extrabold text-white tracking-tight">
                  REPORT DP · SALES DAILY — {currentMonthMeta.name.toUpperCase()} 2026
                </span>
                <span className="px-2 py-0.5 text-[11px] font-semibold bg-emerald-500/30 text-emerald-200 border border-emerald-400/30 rounded-md">
                  {monthTotals.activeDays}/{currentMonthMeta.daysCount} Hari Terisi
                </span>
              </div>
              <div className="text-[11px] text-slate-300 font-mono-num mt-0.5">
                Sales ({currentMonthMeta.short}): Rp {formatNumberId(monthTotals.sales) || '0'} · Qty:{' '}
                {formatNumberId(monthTotals.qty) || '0'} pcs · Trx:{' '}
                {formatNumberId(monthTotals.trx) || '0'} trx · Basket Size:{' '}
                {monthTotals.basketSize ? `Rp ${formatNumberId(monthTotals.basketSize)}` : 'Rp 0'} · APC:{' '}
                {monthTotals.apc ? `Rp ${formatNumberId(monthTotals.apc)}` : 'Rp 0'}
              </div>
            </div>

            {onSelectHorizon && (
              <div className="no-print inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15 shrink-0">
                <button
                  type="button"
                  onClick={() => onSelectHorizon('harian')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all bg-amber-400 text-slate-950 shadow-2xs"
                >
                  Sales Daily
                </button>
                <button
                  type="button"
                  onClick={() => onSelectHorizon('weekly')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                >
                  Sales Weekly
                </button>
                <button
                  type="button"
                  onClick={() => onSelectHorizon('bulanan')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                >
                  Sales Monthly
                </button>
              </div>
            )}
          </div>

          {/* Row 2: Compact Month Navigation + Table Format & Style Controls */}
          <div className="no-print flex flex-wrap items-center justify-between gap-2 pt-2.5 border-t border-white/15 w-full max-w-full min-w-0">
            <div className="flex flex-wrap items-center gap-1.5 min-w-0 max-w-full">
              <div className="inline-flex items-center gap-1 bg-white/10 p-0.5 rounded-lg border border-white/15 max-w-full min-w-0">
                <button
                  type="button"
                  disabled={currentMonthIdx <= 0}
                  onClick={() => {
                    const prev = MONTH_LIST[Math.max(0, currentMonthIdx - 1)];
                    if (prev) setSelectedMonthKey(prev.key);
                  }}
                  className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                  title="Bulan Sebelumnya"
                >
                  ◀
                </button>

                <select
                  aria-label="Pilih Bulan untuk Sales Table Harian (Tanggal 1 - 30/31)"
                  value={selectedMonthKey}
                  onChange={(e) => setSelectedMonthKey(e.target.value)}
                  className="px-2.5 py-1 rounded-md bg-amber-400 text-slate-950 text-xs font-extrabold focus:outline-none cursor-pointer min-w-0 max-w-[210px] sm:max-w-[250px] truncate shadow-2xs"
                >
                  {MONTH_LIST.map((m) => (
                    <option key={m.key} value={m.key}>
                      {m.name.toUpperCase()} 2026 (Tgl 1 – {m.daysCount})
                    </option>
                  ))}
                </select>

                <button
                  type="button"
                  disabled={currentMonthIdx >= MONTH_LIST.length - 1}
                  onClick={() => {
                    const next = MONTH_LIST[Math.min(MONTH_LIST.length - 1, currentMonthIdx + 1)];
                    if (next) setSelectedMonthKey(next.key);
                  }}
                  className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                  title="Bulan Berikutnya"
                >
                  ▶
                </button>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-1.5">

              {(dateRangeFilter !== 'ALL' ||
                dayTypeFilter !== 'ALL' ||
                staffFilter !== 'ALL' ||
                Boolean(calendarExactDateFilter) ||
                viewFormat !== 'TABLE_VERTICAL') && (
                <button
                  type="button"
                  onClick={() => {
                    setDateRangeFilter('ALL');
                    setDayTypeFilter('ALL');
                    onChangeStaffFilter('ALL');
                    setCalendarExactDateFilter('');
                    setViewFormat('TABLE_VERTICAL');
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-bold bg-sky-500/25 hover:bg-sky-500/35 text-sky-200 border border-sky-400/30 rounded-lg transition-colors whitespace-nowrap"
                  title="Reset Semua Filter Tampilan Harian"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Reset Filter</span>
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Clean 5-Column Sub-Header Filter Bar with Calendar Date Filter */}
        <div className="no-print bg-slate-50 border-b border-slate-200 px-3.5 sm:px-4 py-2.5 space-y-2 w-full max-w-full overflow-hidden">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2 w-full">
            <select
              aria-label="Pilih Rentang Tanggal Harian"
              value={dateRangeFilter}
              onChange={(e) =>
                setDateRangeFilter(
                  e.target.value as 'ALL' | '1_10' | '11_20' | '21_31' | 'FILLED'
                )
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL">
                Rentang: Tgl 1 – {currentMonthMeta.daysCount} ({currentMonthMeta.daysCount} Hari)
              </option>
              <option value="1_10">Rentang: Tanggal 1 – 10</option>
              <option value="11_20">Rentang: Tanggal 11 – 20</option>
              <option value="21_31">
                Rentang: Tanggal 21 – {currentMonthMeta.daysCount}
              </option>
              <option value="FILLED">Filter: Hanya Tanggal Terisi</option>
            </select>

            <select
              aria-label="Filter Tipe Hari"
              value={dayTypeFilter}
              onChange={(e) =>
                setDayTypeFilter(e.target.value as 'ALL' | 'WEEKDAY' | 'WEEKEND')
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-indigo-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL">Hari: Semua (Senin–Minggu)</option>
              <option value="WEEKDAY">Hari: Weekday (Senin–Jumat)</option>
              <option value="WEEKEND">Hari: Weekend (Sabtu–Minggu)</option>
            </select>

            <select
              aria-label="Format Tabel Sales Harian"
              value={viewFormat}
              onChange={(e) =>
                setViewFormat(
                  e.target.value as
                    | 'TABLE_VERTICAL'
                    | 'MATRIX_HORIZONTAL'
                    | 'DETAIL_SHIFT'
                )
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-800 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="TABLE_VERTICAL">
                Format: Tabel Tgl 1–{currentMonthMeta.daysCount} (Vertikal)
              </option>
              <option value="MATRIX_HORIZONTAL">
                Format: Matriks Kesamping (WPS)
              </option>
              <option value="DETAIL_SHIFT">
                Format: Rincian Pagi &amp; Siang
              </option>
            </select>

            <select
              aria-label="Filter Petugas Sales Harian"
              value={staffFilter}
              onChange={(e) => onChangeStaffFilter(e.target.value)}
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-700 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL">Petugas: Semua ({filteredCalendarDays.length} Hari)</option>
              <option value="Gani">Petugas: Gani</option>
              <option value="Irfan">Petugas: Irfan</option>
              <option value="Feronika">Petugas: Feronika</option>
              <option value="Abduh">Petugas: Abduh</option>
              <option value="Fitri">Petugas: Fitri</option>
            </select>

            {/* Filter Kalender Tanggal Langsung */}
            <div className="flex items-center gap-1.5 px-2.5 py-1 bg-amber-50/90 border border-amber-300 rounded-lg shadow-2xs">
              <Calendar className="w-3.5 h-3.5 text-amber-700 shrink-0" />
              <label htmlFor="dp-daily-cal" className="text-xs font-bold text-amber-950 shrink-0">
                Kalender:
              </label>
              <input
                id="dp-daily-cal"
                type="date"
                value={calendarExactDateFilter}
                onChange={(e) => {
                  const val = e.target.value;
                  setCalendarExactDateFilter(val);
                  if (val) {
                    const parts = val.split('-');
                    if (parts.length === 3 && parts[1] !== selectedMonthKey) {
                      setSelectedMonthKey(parts[1]);
                    }
                  }
                }}
                className="w-full text-xs font-bold bg-white border border-amber-300 rounded px-1.5 py-0.5 text-slate-900 focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                title="Pilih tanggal dari kalender untuk memfilter baris tanggal ini di tabel"
              />
              {calendarExactDateFilter && (
                <button
                  type="button"
                  onClick={() => setCalendarExactDateFilter('')}
                  className="text-[10px] font-bold text-amber-800 hover:text-amber-950 px-1.5 py-0.5 rounded bg-amber-200/80 cursor-pointer shrink-0"
                  title="Hapus filter kalender tanggal"
                >
                  Semua
                </button>
              )}
            </div>
          </div>

          {/* Side-by-side Weekday vs Weekend Pills */}
          <div className="grid grid-cols-2 gap-2 text-[11px] w-full">
            <button
              type="button"
              onClick={() =>
                setDayTypeFilter((p) => (p === 'WEEKDAY' ? 'ALL' : 'WEEKDAY'))
              }
              className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                dayTypeFilter === 'WEEKDAY'
                  ? 'bg-indigo-600 text-white border-indigo-600'
                  : 'bg-indigo-50/80 text-indigo-900 border-indigo-200 hover:bg-indigo-100'
              }`}
            >
              Weekday (Sen–Jum): Rp {formatNumberId(monthTotals.weekdaySales) || '0'} (
              {monthTotals.sales > 0
                ? Math.round((monthTotals.weekdaySales / monthTotals.sales) * 100)
                : 0}
              %)
            </button>
            <button
              type="button"
              onClick={() =>
                setDayTypeFilter((p) => (p === 'WEEKEND' ? 'ALL' : 'WEEKEND'))
              }
              className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                dayTypeFilter === 'WEEKEND'
                  ? 'bg-rose-600 text-white border-rose-600'
                  : 'bg-rose-50/80 text-rose-900 border-rose-200 hover:bg-rose-100'
              }`}
            >
              Weekend (Sab–Min): Rp {formatNumberId(monthTotals.weekendSales) || '0'} (
              {monthTotals.sales > 0
                ? Math.round((monthTotals.weekendSales / monthTotals.sales) * 100)
                : 0}
              %)
            </button>
          </div>
        </div>

        {/* Ringkasan Kumulatif Bulanan Otomatis untuk Bulan Terpilih */}
        <MonthlyAutoSummaryCard
          tables={tables}
          selectedMonthKey={selectedMonthKey}
          onSelectMonthKey={setSelectedMonthKey}
          dayTypeFilter={dayTypeFilter}
        />

        {/* Live Visual Indicator Banner when daily cells are recently modified */}
        {lastDailyEditBanner && Object.keys(modifiedDailyCells).length > 0 && (
          <div className="no-print bg-emerald-50/90 border-b border-emerald-300 px-3.5 sm:px-4 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex items-center gap-2 text-emerald-950">
              <span className="inline-block w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
              <span className="font-bold">Sel Diubah:</span>
              <span className="px-1.5 py-0.5 rounded bg-emerald-100 border border-emerald-300 font-bold text-emerald-900">
                {lastDailyEditBanner.cellLabel}
              </span>
              <span className="text-slate-600">→</span>
              <span className="font-mono-num font-extrabold text-emerald-800">
                {lastDailyEditBanner.newValue}
              </span>
            </div>
            <div className="flex items-center gap-3 text-[11px] font-semibold text-slate-600">
              <span className="inline-flex items-center gap-1">
                <span className="w-2.5 h-2.5 rounded-xs bg-emerald-500/30 border-2 border-emerald-600 inline-block" />
                Sel Input Baru Diubah
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="w-2.5 h-2.5 rounded-xs bg-amber-400/30 border-2 border-amber-500 inline-block" />
                Total Otomatis Terhitung
              </span>
            </div>
          </div>
        )}

        {/* ===================================================================
            FORMAT 1 & 3: TABEL TANGGAL 1 S/D 30/31 (TOTAL SALES HARIAN, QTY & TRX)
           =================================================================== */}
        {viewFormat !== 'MATRIX_HORIZONTAL' ? (
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-center font-bold w-14`}>
                    Tgl
                  </th>
                  <th className={`${redHeaderClass} px-3 py-2.5 text-left font-bold`}>
                    Tanggal
                  </th>
                  <th className={`${redHeaderClass} px-3 py-2.5 text-left font-bold`}>
                    Hari
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-center font-bold`}>
                    Weekly
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    ☀ Sales Pagi (Rp)
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    🌙 Sales Siang (Rp)
                  </th>
                  <th
                    className={`${goldGrandTotalClass} px-3.5 py-2.5 text-right font-extrabold min-w-[145px]`}
                  >
                    Total Sales Harian (Rp)
                  </th>
                  <th
                    className={`${goldGrandTotalClass} px-3 py-2.5 text-right font-extrabold min-w-[100px]`}
                  >
                    Qty (pcs)
                  </th>
                  <th
                    className={`${goldGrandTotalClass} px-3 py-2.5 text-right font-extrabold min-w-[100px]`}
                  >
                    Trx (Struk)
                  </th>
                  <th className={`${redHeaderClass} px-3 py-2.5 text-right font-bold`}>
                    Basket Size
                  </th>
                  <th className={`${redHeaderClass} px-3 py-2.5 text-right font-bold`}>
                    APC (Rp/Qty)
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-center font-bold no-print`}>
                    Aksi
                  </th>
                </tr>
              </thead>
              <tbody>
                {filteredCalendarDays.map((item) => {
                  const hasData =
                    item.totalSales > 0 || item.totalQty > 0 || item.totalTrx > 0;
                  return (
                    <tr
                      key={item.slashDate}
                      className={`border-b border-slate-200 transition-colors ${
                        item.isWeekend
                          ? 'bg-rose-50/40 hover:bg-rose-50/70'
                          : hasData
                            ? 'bg-emerald-50/20 hover:bg-slate-50'
                            : 'bg-white hover:bg-slate-50'
                      }`}
                    >
                      {/* Nomor Tanggal (1 s/d 30/31) */}
                      <td className="px-2.5 py-1.5 border border-slate-200 text-center font-mono-num font-extrabold text-slate-900">
                        <span
                          className={`inline-flex items-center justify-center w-6 h-6 rounded-md text-xs ${
                            item.isWeekend
                              ? 'bg-rose-600 text-white'
                              : hasData
                                ? 'bg-amber-400 text-slate-950'
                                : 'bg-slate-100 text-slate-700'
                          }`}
                        >
                          {item.dayOfMonth}
                        </span>
                      </td>

                      {/* Tanggal DD/MM/2026 */}
                      <td className="px-3 py-1.5 border border-slate-200 font-mono-num font-semibold text-slate-800 whitespace-nowrap">
                        {item.slashDate}
                      </td>

                      {/* Nama Hari */}
                      <td className="px-3 py-1.5 border border-slate-200 whitespace-nowrap">
                        <span
                          className={`font-bold ${
                            item.isWeekend ? 'text-rose-700' : 'text-slate-900'
                          }`}
                        >
                          {item.dayName}
                        </span>
                        {item.isWeekend && (
                          <span className="ml-1.5 px-1.5 py-0.2 text-[10px] font-semibold bg-rose-100 text-rose-700 rounded">
                            Weekend
                          </span>
                        )}
                      </td>

                      {/* Weekly Badge */}
                      <td className="px-2.5 py-1.5 border border-slate-200 text-center whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => onSelectWeeklyFilter(`weekly-${item.weekNumber}`)}
                          className="px-2 py-0.5 rounded bg-slate-100 hover:bg-indigo-100 text-slate-700 hover:text-indigo-800 font-mono-num font-bold text-[11px] transition-colors"
                          title={`Pilih Weekly ${item.weekNumber}`}
                        >
                          W{item.weekNumber}
                        </button>
                      </td>

                      {/* Shift Pagi & Siang Columns (Always shown in Sales Harian) */}
                      <td
                        className={`p-1 border border-slate-200 bg-amber-50/20 ${getDailyFlashClass(`${item.slashDate}-pagi-sales`)}`}
                      >
                        <input
                          type="text"
                          inputMode="numeric"
                          aria-label={`Sales Pagi Tanggal ${item.dayOfMonth}`}
                          value={
                            item.pagiSales > 0 ? formatNumberId(item.pagiSales) : ''
                          }
                          onChange={(e) =>
                            handleUpdateShiftMetric(
                              item,
                              'pagi',
                              'sales',
                              e.target.value
                            )
                          }
                          placeholder="0"
                          className="w-full px-2 py-1 text-xs font-mono-num font-semibold text-right bg-transparent focus:bg-white rounded border border-transparent focus:border-amber-500 focus:outline-none text-amber-950"
                        />
                      </td>
                      <td
                        className={`p-1 border border-slate-200 bg-indigo-50/20 ${getDailyFlashClass(`${item.slashDate}-siang-sales`)}`}
                      >
                        <input
                          type="text"
                          inputMode="numeric"
                          aria-label={`Sales Siang Tanggal ${item.dayOfMonth}`}
                          value={
                            item.siangSales > 0 ? formatNumberId(item.siangSales) : ''
                          }
                          onChange={(e) =>
                            handleUpdateShiftMetric(
                              item,
                              'siang',
                              'sales',
                              e.target.value
                            )
                          }
                          placeholder="0"
                          className="w-full px-2 py-1 text-xs font-mono-num font-semibold text-right bg-transparent focus:bg-white rounded border border-transparent focus:border-indigo-500 focus:outline-none text-indigo-950"
                        />
                      </td>

                      {/* Editable Total Sales Harian (Rp) */}
                      <td
                        className={`${
                          hasData ? yellowTotalClass : 'border border-slate-200 bg-amber-50/30'
                        } ${getDailyFlashClass(`${item.slashDate}-total-sales`)} p-1`}
                      >
                        <input
                          type="text"
                          inputMode="numeric"
                          aria-label={`Total Sales Harian Tanggal ${item.dayOfMonth}`}
                          value={
                            item.totalSales > 0 ? formatNumberId(item.totalSales) : ''
                          }
                          onChange={(e) =>
                            handleUpdateDailyTotalMetric(item, 'sales', e.target.value)
                          }
                          placeholder="0"
                          className="w-full px-2.5 py-1 text-xs font-mono-num font-extrabold text-right bg-transparent focus:bg-white rounded border border-transparent focus:border-indigo-600 focus:outline-none text-slate-950"
                        />
                      </td>

                      {/* Editable Total Qty (pcs) */}
                      <td
                        className={`${
                          hasData ? yellowTotalClass : 'border border-slate-200 bg-amber-50/30'
                        } ${getDailyFlashClass(`${item.slashDate}-total-qty`)} p-1`}
                      >
                        <input
                          type="text"
                          inputMode="numeric"
                          aria-label={`Total Qty Harian Tanggal ${item.dayOfMonth}`}
                          value={item.totalQty > 0 ? formatNumberId(item.totalQty) : ''}
                          onChange={(e) =>
                            handleUpdateDailyTotalMetric(item, 'qty', e.target.value)
                          }
                          placeholder="0"
                          className="w-full px-2 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded border border-transparent focus:border-indigo-600 focus:outline-none text-slate-950"
                        />
                      </td>

                      {/* Editable Total Trx (Struk) */}
                      <td
                        className={`${
                          hasData ? yellowTotalClass : 'border border-slate-200 bg-amber-50/30'
                        } ${getDailyFlashClass(`${item.slashDate}-total-trx`)} p-1`}
                      >
                        <input
                          type="text"
                          inputMode="numeric"
                          aria-label={`Total Trx Harian Tanggal ${item.dayOfMonth}`}
                          value={item.totalTrx > 0 ? formatNumberId(item.totalTrx) : ''}
                          onChange={(e) =>
                            handleUpdateDailyTotalMetric(item, 'trx', e.target.value)
                          }
                          placeholder="0"
                          className="w-full px-2 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded border border-transparent focus:border-indigo-600 focus:outline-none text-slate-950"
                        />
                      </td>

                      {/* Basket Size (Sales / Trx) */}
                      <td
                        className={`px-3 py-1.5 border border-slate-200 text-right font-mono-num font-semibold bg-amber-50/40 ${getDailyFlashClass(`${item.slashDate}-basketSize`)}`}
                      >
                        {renderCalculatedCell(item.basketSize)}
                      </td>

                      {/* APC (Sales / Qty) */}
                      <td
                        className={`px-3 py-1.5 border border-slate-200 text-right font-mono-num font-semibold bg-amber-50/40 ${getDailyFlashClass(`${item.slashDate}-apc`)}`}
                      >
                        {renderCalculatedCell(item.apc)}
                      </td>

                      {/* Reset Single Date */}
                      <td className="px-2 py-1.5 border border-slate-200 text-center no-print">
                        {hasData ? (
                          <button
                            type="button"
                            onClick={() => handleClearSingleDate(item)}
                            className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-bold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded transition-colors"
                            title={`Reset / Kosongkan data tanggal ${item.slashDate}`}
                          >
                            <RotateCcw className="w-3 h-3" />
                            <span>Reset</span>
                          </button>
                        ) : (
                          <span className="text-slate-300">-</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className={goldGrandTotalClass}>
                  <td colSpan={4} className="px-3 py-2.5 text-left font-extrabold">
                    TOTAL HARIAN ({currentMonthMeta.name.toUpperCase()} 2026 —{' '}
                    {filteredCalendarDays.length} HARI)
                  </td>
                  <td
                    className={`px-2.5 py-2.5 text-right font-mono-num font-bold ${getDailyFlashClass('footer-pagi-sales')}`}
                  >
                    {formatNumberId(filteredTotals.pagiSales) || '0'}
                  </td>
                  <td
                    className={`px-2.5 py-2.5 text-right font-mono-num font-bold ${getDailyFlashClass('footer-siang-sales')}`}
                  >
                    {formatNumberId(filteredTotals.siangSales) || '0'}
                  </td>
                  <td
                    className={`px-3.5 py-2.5 text-right font-mono-num font-extrabold ${getDailyFlashClass('footer-total-sales')}`}
                  >
                    {formatNumberId(filteredTotals.sales) || '0'}
                  </td>
                  <td
                    className={`px-3 py-2.5 text-right font-mono-num font-extrabold ${getDailyFlashClass('footer-total-qty')}`}
                  >
                    {formatNumberId(filteredTotals.qty) || '0'}
                  </td>
                  <td
                    className={`px-3 py-2.5 text-right font-mono-num font-extrabold ${getDailyFlashClass('footer-total-trx')}`}
                  >
                    {formatNumberId(filteredTotals.trx) || '0'}
                  </td>
                  <td
                    className={`px-3 py-2.5 text-right font-mono-num ${getDailyFlashClass('footer-basket-size')}`}
                  >
                    {renderCalculatedCell(filteredTotals.basketSize)}
                  </td>
                  <td
                    className={`px-3 py-2.5 text-right font-mono-num ${getDailyFlashClass('footer-apc')}`}
                  >
                    {renderCalculatedCell(filteredTotals.apc)}
                  </td>
                  <td className="px-2 py-2.5 text-center no-print">-</td>
                </tr>
              </tfoot>
            </table>
          </div>
        ) : (
          /* ===================================================================
             FORMAT 2: MATRIKS KESAMPING TANGGAL 1 S/D 30/31 (FORMAT WPS MERAH-KUNING)
             =================================================================== */
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr>
                  <th
                    className={`${redHeaderClass} px-3 py-2.5 text-left font-bold min-w-[165px] sticky left-0 z-10`}
                  >
                    Tanggal (1 s/d {currentMonthMeta.daysCount})
                  </th>
                  {filteredCalendarDays.map((item) => (
                    <th
                      key={item.slashDate}
                      className={`${redHeaderClass} px-2.5 py-2 text-center font-bold min-w-[105px]`}
                    >
                      <div>Tgl {item.dayOfMonth}</div>
                      <div className="text-[10px] font-normal text-amber-200">
                        {item.dayName}
                      </div>
                    </th>
                  ))}
                  <th
                    className={`${goldGrandTotalClass} px-3.5 py-2.5 text-center font-extrabold min-w-[145px]`}
                  >
                    TOTAL {currentMonthMeta.name.toUpperCase()}
                  </th>
                </tr>
              </thead>
              <tbody>
                {/* Tanggal DD/MM/YYYY */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-1.5 font-semibold sticky left-0 z-10`}>
                    Tanggal Lengkap
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className="px-2 py-1.5 border border-slate-200 bg-slate-50 text-center font-mono-num text-[11px] font-semibold"
                    >
                      {item.slashDate}
                    </td>
                  ))}
                  <td className="px-3 py-1.5 border border-slate-300 bg-amber-50 text-center font-bold">
                    {filteredCalendarDays.length} Hari
                  </td>
                </tr>

                {/* Sales Pagi (Rp) */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-1.5 font-semibold sticky left-0 z-10`}>
                    ☀ Sales Pagi (Rp)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`p-1 border border-slate-200 bg-amber-50/25 ${getDailyFlashClass(`${item.slashDate}-pagi-sales`)}`}
                    >
                      <input
                        type="text"
                        inputMode="numeric"
                        value={item.pagiSales > 0 ? formatNumberId(item.pagiSales) : ''}
                        onChange={(e) =>
                          handleUpdateShiftMetric(item, 'pagi', 'sales', e.target.value)
                        }
                        placeholder="0"
                        className="w-full px-1.5 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded focus:outline-none text-amber-950"
                      />
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-pagi-sales')} px-3 py-1.5 text-right font-mono-num`}
                  >
                    {formatNumberId(filteredTotals.pagiSales) || '0'}
                  </td>
                </tr>

                {/* Sales Siang (Rp) */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-1.5 font-semibold sticky left-0 z-10`}>
                    🌙 Sales Siang (Rp)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`p-1 border border-slate-200 bg-indigo-50/25 ${getDailyFlashClass(`${item.slashDate}-siang-sales`)}`}
                    >
                      <input
                        type="text"
                        inputMode="numeric"
                        value={item.siangSales > 0 ? formatNumberId(item.siangSales) : ''}
                        onChange={(e) =>
                          handleUpdateShiftMetric(item, 'siang', 'sales', e.target.value)
                        }
                        placeholder="0"
                        className="w-full px-1.5 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded focus:outline-none text-indigo-950"
                      />
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-siang-sales')} px-3 py-1.5 text-right font-mono-num`}
                  >
                    {formatNumberId(filteredTotals.siangSales) || '0'}
                  </td>
                </tr>

                {/* Total Sales Harian (Editable) */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Total Sales Harian (Rp)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`${yellowTotalClass} ${getDailyFlashClass(`${item.slashDate}-total-sales`)} p-1`}
                    >
                      <input
                        type="text"
                        inputMode="numeric"
                        value={
                          item.totalSales > 0 ? formatNumberId(item.totalSales) : ''
                        }
                        onChange={(e) =>
                          handleUpdateDailyTotalMetric(item, 'sales', e.target.value)
                        }
                        placeholder="0"
                        className="w-full px-1.5 py-1 text-xs font-mono-num font-extrabold text-right bg-transparent focus:bg-white rounded focus:outline-none"
                      />
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-total-sales')} px-3 py-2 text-right font-mono-num`}
                  >
                    {formatNumberId(filteredTotals.sales) || '0'}
                  </td>
                </tr>

                {/* Total Qty Harian (Editable) */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Qty Harian (pcs)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`${yellowTotalClass} ${getDailyFlashClass(`${item.slashDate}-total-qty`)} p-1`}
                    >
                      <input
                        type="text"
                        inputMode="numeric"
                        value={item.totalQty > 0 ? formatNumberId(item.totalQty) : ''}
                        onChange={(e) =>
                          handleUpdateDailyTotalMetric(item, 'qty', e.target.value)
                        }
                        placeholder="0"
                        className="w-full px-1.5 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded focus:outline-none"
                      />
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-total-qty')} px-3 py-2 text-right font-mono-num`}
                  >
                    {formatNumberId(filteredTotals.qty) || '0'}
                  </td>
                </tr>

                {/* Total Trx Harian (Editable) */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Trx / Struk Harian
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`${yellowTotalClass} ${getDailyFlashClass(`${item.slashDate}-total-trx`)} p-1`}
                    >
                      <input
                        type="text"
                        inputMode="numeric"
                        value={item.totalTrx > 0 ? formatNumberId(item.totalTrx) : ''}
                        onChange={(e) =>
                          handleUpdateDailyTotalMetric(item, 'trx', e.target.value)
                        }
                        placeholder="0"
                        className="w-full px-1.5 py-1 text-xs font-mono-num font-bold text-right bg-transparent focus:bg-white rounded focus:outline-none"
                      />
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-total-trx')} px-3 py-2 text-right font-mono-num`}
                  >
                    {formatNumberId(filteredTotals.trx) || '0'}
                  </td>
                </tr>

                {/* Basket Size */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-1.5 font-semibold sticky left-0 z-10`}>
                    Basket Size (Rp/Trx)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`px-2 py-1.5 border border-slate-200 bg-amber-50/40 text-right font-mono-num font-semibold ${getDailyFlashClass(`${item.slashDate}-basketSize`)}`}
                    >
                      {renderCalculatedCell(item.basketSize)}
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-basket-size')} px-3 py-1.5 text-right font-mono-num`}
                  >
                    {renderCalculatedCell(filteredTotals.basketSize)}
                  </td>
                </tr>

                {/* APC */}
                <tr>
                  <td className={`${redHeaderClass} px-3 py-1.5 font-semibold sticky left-0 z-10`}>
                    APC (Rp/Qty)
                  </td>
                  {filteredCalendarDays.map((item) => (
                    <td
                      key={item.slashDate}
                      className={`px-2 py-1.5 border border-slate-200 bg-amber-50/40 text-right font-mono-num font-semibold ${getDailyFlashClass(`${item.slashDate}-apc`)}`}
                    >
                      {renderCalculatedCell(item.apc)}
                    </td>
                  ))}
                  <td
                    className={`${goldGrandTotalClass} ${getDailyFlashClass('footer-apc')} px-3 py-1.5 text-right font-mono-num`}
                  >
                    {renderCalculatedCell(filteredTotals.apc)}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* ===================================================================
            GRAFIK DIAGRAM SALES TABLE HARIAN (SESUAI FOTO)
           =================================================================== */}
        <div className="p-3.5 sm:p-4 bg-white border-t border-slate-200 space-y-4">
          <PhotoStyleSalesDiagram
            headerTitle={`DIAGRAM SALES HARIAN — ${currentMonthMeta.name.toUpperCase()} 2026 (TGL 1 – ${currentMonthMeta.daysCount} ${currentMonthMeta.short.toUpperCase()})`}
            headerSubtitle={`Diagram Batang Vertikal ${filteredCalendarDays.length} Hari & Garis Tren Penjualan Shift Pagi vs Shift Siang`}
            badgeText="DIAGRAM SALES (OMZET RUPIAH)"
            cardTitle="Grafik Batang Shift Pagi vs Siang & Garis Tren"
            totalSales={filteredTotals.sales}
            totalPagi={filteredTotals.pagiSales}
            totalSiang={filteredTotals.siangSales}
            compactX={true}
            bars={filteredCalendarDays.map((item) => ({
              key: item.slashDate,
              label: `${item.dayName.slice(0, 3)} (${item.dayOfMonth})`,
              subLabel: item.slashDate.slice(0, 5),
              pagiSales: item.pagiSales,
              siangSales: item.siangSales,
              totalSales: item.totalSales,
              isWeekend: item.isWeekend,
            }))}
          />

          <GrowthAndContributionBarDiagram
            headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI HARIAN — ${currentMonthMeta.name.toUpperCase()} 2026`}
            headerSubtitle={`Grafik Batang Pertumbuhan Harian (Growth %) & Porsi Kontribusi Sales ${filteredCalendarDays.length} Hari`}
            periodBadgeText={`${currentMonthMeta.name.toUpperCase()} 2026`}
            compactX={true}
            items={filteredCalendarDays.map((item, idx) => ({
              key: item.slashDate,
              label: `${item.dayName.slice(0, 3)} (${item.dayOfMonth})`,
              subLabel: item.slashDate.slice(0, 5),
              pagiSales: item.pagiSales,
              siangSales: item.siangSales,
              totalSales: item.totalSales,
              prevSales: idx > 0 ? filteredCalendarDays[idx - 1].totalSales : null,
              isWeekend: item.isWeekend,
            }))}
          />
        </div>
      </section>
    </div>
  );
};

/* ============================================================================
   2. SALES TABLE BULANAN (MONTHLY SALES SPREADSHEET VIEW)
   ============================================================================ */

interface MonthlySalesTableProps {
  tables: WeeklySalesTable[];
  onChangeTables: (next: WeeklySalesTable[]) => void;
  onSelectWeeklyAndSwitchToWeeklyMode: (weekId: string) => void;
  showDivZero: boolean;
  classicExcelColors: boolean;
  onSelectHorizon?: (h: 'harian' | 'weekly' | 'bulanan') => void;
  onToggleTheme?: () => void;
  onToggleDivZero?: () => void;
  onManualSave?: () => void;
}

export const MonthlySalesSpreadsheetView: React.FC<MonthlySalesTableProps> = ({
  tables,
  onChangeTables,
  onSelectWeeklyAndSwitchToWeeklyMode,
  showDivZero,
  classicExcelColors,
  onSelectHorizon,
  onToggleTheme,
  onToggleDivZero,
  onManualSave,
}) => {
  const [selectedMonthKey, setSelectedMonthKey] = useState<string>(() => {
    try {
      const savedM = localStorage.getItem(STORAGE_KEY_SELECTED_MONTH);
      if (savedM && MONTH_LIST.some((m) => m.key === savedM)) {
        return savedM;
      }
    } catch {}
    const filledTables = tables.filter((t) =>
      t.days.some((d) => (d.pagi?.sales ?? 0) + (d.siang?.sales ?? 0) > 0)
    );
    const latestT = filledTables[filledTables.length - 1];
    if (latestT) {
      const refSlash = latestT.days[3]?.tanggal || latestT.days[0]?.tanggal || '';
      const parts = refSlash.split('/');
      if (parts.length === 3 && parts[1] && parts[2] === '2026') {
        return parts[1].padStart(2, '0');
      }
    }
    return '11';
  });

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_SELECTED_MONTH, selectedMonthKey);
    } catch {}
  }, [selectedMonthKey]);
  const [quarterFilter, setQuarterFilter] = useState<'ALL' | 'Q1' | 'Q2' | 'Q3' | 'Q4'>('ALL');
  const [monthlyViewMode, setMonthlyViewMode] = useState<
    'ALL_MONTHLY' | 'REKAP_12_BULAN' | 'MATRIX_WPS'
  >('ALL_MONTHLY');
  const [dayTypeFilter, setDayTypeFilter] = useState<'ALL' | 'WEEKDAY' | 'WEEKEND'>('ALL');
  const [rowViewFilter, setRowViewFilter] = useState<'ALL' | 'SUMMARY_ONLY'>('ALL');
  const [confirmMonthlyResetTarget, setConfirmMonthlyResetTarget] = useState<
    null | 'CURRENT_MONTH' | 'ALL_12_MONTHS' | string
  >(null);

  // Reset handlers for Sales Monthly (Tanggal 1..30/31 per bulan or all 12 months)
  const handleResetMonthlyNumbers = (monthKeyToReset: string) => {
    const suffix = `/${monthKeyToReset}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(suffix)
          ? {
              ...d,
              pagi: { ...d.pagi, sales: null, qty: null, trx: null },
              siang: { ...d.siang, sales: null, qty: null, trx: null },
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmMonthlyResetTarget(null);
  };

  const handleResetMonthlyReasons = (monthKeyToReset: string) => {
    const suffix = `/${monthKeyToReset}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(suffix)
          ? {
              ...d,
              pagi: { ...d.pagi, nama: '' },
              siang: { ...d.siang, nama: '' },
              reasonPagi: '',
              reasonSiang: '',
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmMonthlyResetTarget(null);
  };

  const handleResetSingleMonthAll = (monthKeyToReset: string) => {
    const suffix = `/${monthKeyToReset}/2026`;
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        (d.tanggal || '').endsWith(suffix)
          ? {
              ...d,
              pagi: { nama: '', sales: null, qty: null, trx: null },
              siang: { nama: '', sales: null, qty: null, trx: null },
              reasonPagi: '',
              reasonSiang: '',
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
    setConfirmMonthlyResetTarget(null);
  };

  const handleResetAll12Months = () => {
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) => ({
        ...d,
        pagi: { nama: '', sales: null, qty: null, trx: null },
        siang: { nama: '', sales: null, qty: null, trx: null },
        reasonPagi: '',
        reasonSiang: '',
      })),
    }));
    onChangeTables(nextTables);
    setConfirmMonthlyResetTarget(null);
  };

  const handleResetMonthlySingleDate = (slashDate: string) => {
    const nextTables = tables.map((t) => ({
      ...t,
      days: t.days.map((d) =>
        d.tanggal === slashDate
          ? {
              ...d,
              pagi: { nama: '', sales: null, qty: null, trx: null },
              siang: { nama: '', sales: null, qty: null, trx: null },
              reasonPagi: '',
              reasonSiang: '',
            }
          : d
      ),
    }));
    onChangeTables(nextTables);
  };

  const redHeaderClass = classicExcelColors
    ? 'bg-[#C00000] text-white border border-slate-900'
    : 'bg-rose-800 text-white border border-slate-300';

  const yellowTotalClass = classicExcelColors
    ? 'bg-[#FFFF00] text-slate-950 border border-slate-900 font-bold'
    : 'bg-amber-200/90 text-slate-950 border border-slate-300 font-bold';

  const goldMonthlyTotalClass = classicExcelColors
    ? 'bg-[#FFC000] text-slate-950 border border-slate-900 font-bold'
    : 'bg-amber-400 text-slate-950 border border-slate-300 font-bold';

  const renderCalculatedCell = (val: number | null | undefined) => {
    if (val === null || val === undefined || val === 0) {
      return showDivZero ? (
        <span className="text-slate-500 font-semibold">#DIV/0!</span>
      ) : (
        <span className="text-slate-400">0</span>
      );
    }
    return <span>{formatNumberId(val)}</span>;
  };

  // Compute all 12 months summary strictly from Tanggal 1 to Tanggal 28/30/31 of each month
  const monthlySummaries = useMemo(() => {
    // Map all loaded daily records by slashDate (DD/MM/2026)
    const recordByDate = new Map<
      string,
      { day: DailySalesRecord; weekNumber: number }
    >();
    tables.forEach((t) => {
      t.days.forEach((d) => {
        const key = (d.tanggal || '').trim();
        if (!key) return;
        const existing = recordByDate.get(key);
        const dm = calculateDayMetrics(d);
        const existingSales = existing
          ? calculateDayMetrics(existing.day).totalSales
          : -1;
        if (!existing || dm.totalSales >= existingSales) {
          recordByDate.set(key, { day: d, weekNumber: t.weekNumber });
        }
      });
    });

    const schedWeekByDate = new Map<string, number>();
    ANNUAL_2026_WEEKS.forEach((w) => {
      w.slashDates.forEach((dt) => {
        schedWeekByDate.set(dt, w.weekNumber);
      });
    });

    return MONTH_LIST.map((m) => {
      // Build exact 1..daysCount (1-28/30/31) daily records for this month
      const monthDays: {
        dayOfMonth: number;
        slashDate: string;
        dayName: string;
        isWeekend: boolean;
        weekNumber: number;
        pagiNama: string;
        siangNama: string;
        pagiSales: number;
        siangSales: number;
        totalSales: number;
        pagiQty: number;
        siangQty: number;
        totalQty: number;
        pagiTrx: number;
        siangTrx: number;
        totalTrx: number;
        basketSize: number | null;
        apc: number | null;
      }[] = [];

      let monthPagiSales = 0;
      let monthSiangSales = 0;
      let monthPagiQty = 0;
      let monthSiangQty = 0;
      let monthPagiTrx = 0;
      let monthSiangTrx = 0;
      let monthWeekdaySales = 0;
      let monthWeekendSales = 0;
      let monthActiveDays = 0;

      for (let dNum = 1; dNum <= m.daysCount; dNum++) {
        const dd = String(dNum).padStart(2, '0');
        const slashDate = `${dd}/${m.key}/2026`;
        const dayName = getIndoDayName(slashDate);
        const wknd = isWeekendDate(slashDate);
        const found = recordByDate.get(slashDate);
        const wkNum =
          found?.weekNumber ?? schedWeekByDate.get(slashDate) ?? 1;

        const includeByDayType =
          dayTypeFilter === 'ALL' ||
          (dayTypeFilter === 'WEEKDAY' && !wknd) ||
          (dayTypeFilter === 'WEEKEND' && wknd);

        const ps = includeByDayType ? (found?.day.pagi.sales ?? 0) : 0;
        const ss = includeByDayType ? (found?.day.siang.sales ?? 0) : 0;
        const pq = includeByDayType ? (found?.day.pagi.qty ?? 0) : 0;
        const sq = includeByDayType ? (found?.day.siang.qty ?? 0) : 0;
        const pt = includeByDayType ? (found?.day.pagi.trx ?? 0) : 0;
        const st = includeByDayType ? (found?.day.siang.trx ?? 0) : 0;
        const totS = ps + ss;
        const totQ = pq + sq;
        const totT = pt + st;

        monthPagiSales += ps;
        monthSiangSales += ss;
        monthPagiQty += pq;
        monthSiangQty += sq;
        monthPagiTrx += pt;
        monthSiangTrx += st;
        if (wknd) monthWeekendSales += totS;
        else monthWeekdaySales += totS;
        if (totS > 0) monthActiveDays += 1;

        monthDays.push({
          dayOfMonth: dNum,
          slashDate,
          dayName,
          isWeekend: wknd,
          weekNumber: wkNum,
          pagiNama: found?.day.pagi.nama || '',
          siangNama: found?.day.siang.nama || '',
          pagiSales: ps,
          siangSales: ss,
          totalSales: totS,
          pagiQty: pq,
          siangQty: sq,
          totalQty: totQ,
          pagiTrx: pt,
          siangTrx: st,
          totalTrx: totT,
          basketSize: totT > 0 ? Math.round(totS / totT) : null,
          apc: totQ > 0 ? Math.round(totS / totQ) : null,
        });
      }

      // Group those exact month days (1..30/31) by the weeks that overlap this month
      const overlappingWeeks = ANNUAL_2026_WEEKS.filter((w) =>
        w.slashDates.some((dt) => dt.endsWith(`/${m.key}/2026`))
      );

      const weekDetails = overlappingWeeks.map((sched) => {
        const loaded = tables.find((t) => t.weekNumber === sched.weekNumber);
        const tableObj = loaded || createWeeklyTableFromAnnualSchedule(sched);
        // Only count days in this week that belong to month m.key (Tanggal 1..30/31 of month m)
        const daysInThisMonth = monthDays.filter(
          (md) => md.weekNumber === sched.weekNumber
        );

        let pagiSales = 0;
        let siangSales = 0;
        let pagiQty = 0;
        let siangQty = 0;
        let pagiTrx = 0;
        let siangTrx = 0;
        let weekdaySales = 0;
        let weekendSales = 0;
        let activeDays = 0;
        const reasons: string[] = [];

        daysInThisMonth.forEach((md) => {
          pagiSales += md.pagiSales;
          siangSales += md.siangSales;
          pagiQty += md.pagiQty;
          siangQty += md.siangQty;
          pagiTrx += md.pagiTrx;
          siangTrx += md.siangTrx;
          if (md.isWeekend) weekendSales += md.totalSales;
          else weekdaySales += md.totalSales;
          if (md.totalSales > 0) activeDays += 1;
        });

        const totalSales = pagiSales + siangSales;
        const totalQty = pagiQty + siangQty;
        const totalTrx = pagiTrx + siangTrx;
        const basketSize = totalTrx > 0 ? Math.round(totalSales / totalTrx) : null;
        const apc = totalQty > 0 ? Math.round(totalSales / totalQty) : null;

        const firstInMonth = daysInThisMonth[0]?.slashDate || sched.seninSlash;
        const lastInMonth =
          daysInThisMonth[daysInThisMonth.length - 1]?.slashDate ||
          sched.mingguSlash;

        return {
          sched,
          firstInMonth,
          lastInMonth,
          daysInMonthCount: daysInThisMonth.length,
          isLoaded: Boolean(loaded),
          tableObj,
          pagiSales,
          siangSales,
          totalSales,
          pagiQty,
          siangQty,
          totalQty,
          pagiTrx,
          siangTrx,
          totalTrx,
          weekdaySales,
          weekendSales,
          activeDays,
          basketSize,
          apc,
          reasons,
        };
      });

      const monthTotalSales = monthPagiSales + monthSiangSales;
      const monthTotalQty = monthPagiQty + monthSiangQty;
      const monthTotalTrx = monthPagiTrx + monthSiangTrx;
      const monthBasketSize =
        monthTotalTrx > 0 ? Math.round(monthTotalSales / monthTotalTrx) : null;
      const monthApc =
        monthTotalQty > 0 ? Math.round(monthTotalSales / monthTotalQty) : null;

      // Enrich weekDetails with Kontribusi (%) against monthTotalSales and Growth (%) against previous week
      const enrichedWeekDetails = weekDetails.map((w, wIdx) => {
        const contributionPct =
          monthTotalSales > 0 ? (w.totalSales / monthTotalSales) * 100 : 0;

        let prevSales = 0;
        let prevWeekNumber: number | null = null;
        if (wIdx > 0) {
          prevSales = weekDetails[wIdx - 1].totalSales;
          prevWeekNumber = weekDetails[wIdx - 1].sched.weekNumber;
        } else if (w.sched.weekNumber > 1) {
          prevWeekNumber = w.sched.weekNumber - 1;
          const prevTbl = tables.find((t) => t.weekNumber === prevWeekNumber);
          if (prevTbl) {
            prevSales = calculateWeekMetrics(prevTbl).weeklySales;
          }
        }

        const growthDiff = w.totalSales - prevSales;
        let growthPct: number | null = null;
        let growthStatus: 'UP' | 'DOWN' | 'FLAT' | 'NEW' | 'BASE' = 'FLAT';

        if (prevWeekNumber === null || (wIdx === 0 && prevSales === 0 && w.totalSales > 0)) {
          growthStatus = 'BASE';
          growthPct = null;
        } else if (w.totalSales === 0 && prevSales === 0) {
          growthStatus = 'FLAT';
          growthPct = 0;
        } else if (prevSales === 0 && w.totalSales > 0) {
          growthStatus = 'NEW';
          growthPct = 100;
        } else if (prevSales > 0) {
          growthPct = ((w.totalSales - prevSales) / prevSales) * 100;
          if (growthDiff > 0) growthStatus = 'UP';
          else if (growthDiff < 0) growthStatus = 'DOWN';
          else growthStatus = 'FLAT';
        }

        return {
          ...w,
          contributionPct,
          prevWeekNumber,
          prevSales,
          growthDiff,
          growthPct,
          growthStatus,
        };
      });

      return {
        ...m,
        startSlash: `01/${m.key}/2026`,
        endSlash: `${String(m.daysCount).padStart(2, '0')}/${m.key}/2026`,
        rangeLabel: `1–${m.daysCount} ${m.short} 2026`,
        monthDays,
        weekDetails: enrichedWeekDetails,
        monthPagiSales,
        monthSiangSales,
        monthTotalSales,
        monthPagiQty,
        monthSiangQty,
        monthTotalQty,
        monthPagiTrx,
        monthSiangTrx,
        monthTotalTrx,
        monthWeekdaySales,
        monthWeekendSales,
        monthActiveDays,
        monthBasketSize,
        monthApc,
      };
    });
  }, [tables, dayTypeFilter]);

  const currentMonthSummary =
    monthlySummaries.find((m) => m.key === selectedMonthKey) || monthlySummaries[9];

  const currentMonthIdx = MONTH_LIST.findIndex((m) => m.key === selectedMonthKey);

  const handleOpenAndSwitchToWeek = (weekNumber: number) => {
    const existing = tables.find((t) => t.weekNumber === weekNumber);
    if (existing) {
      onSelectWeeklyAndSwitchToWeeklyMode(existing.id);
      return;
    }
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNumber);
    if (!sched) return;
    const created = createWeeklyTableFromAnnualSchedule(sched);
    const next = [...tables, created].sort((a, b) => a.weekNumber - b.weekNumber);
    onChangeTables(next);
    onSelectWeeklyAndSwitchToWeeklyMode(created.id);
  };

  const grandAnnualSales = monthlySummaries.reduce((s, m) => s + m.monthTotalSales, 0);
  const grandAnnualQty = monthlySummaries.reduce((s, m) => s + m.monthTotalQty, 0);
  const grandAnnualTrx = monthlySummaries.reduce((s, m) => s + m.monthTotalTrx, 0);
  const grandAnnualPagi = monthlySummaries.reduce((s, m) => s + m.monthPagiSales, 0);
  const grandAnnualSiang = monthlySummaries.reduce((s, m) => s + m.monthSiangSales, 0);
  const grandAnnualWeekday = monthlySummaries.reduce((s, m) => s + m.monthWeekdaySales, 0);
  const grandAnnualWeekend = monthlySummaries.reduce((s, m) => s + m.monthWeekendSales, 0);
  const grandAnnualActiveDays = monthlySummaries.reduce((s, m) => s + m.monthActiveDays, 0);

  return (
    <div className="space-y-3.5 w-full max-w-full min-w-0">
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
        {/* Clean Integrated Dark Banner Header for Sales Table Bulanan */}
        <div className="px-3.5 sm:px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-200 space-y-3 w-full max-w-full overflow-hidden">
          {/* Row 1: Title + Integrated Harian / Weekly / Bulanan Switcher */}
          <div className="flex flex-wrap items-center justify-between gap-2.5 w-full max-w-full min-w-0">
            <div className="min-w-0 max-w-full">
              <div className="flex flex-wrap items-center gap-1.5 text-[11px] font-bold text-amber-300 mb-0.5">
                <span>REPORT BAZAR MATAHARI</span>
                <span className="text-slate-400">·</span>
                <span className="text-white">Tanjung</span>
                <span className="text-slate-400">·</span>
                <span className="text-rose-300">PT MDS Retailing Tbk</span>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-sm sm:text-base font-extrabold text-white tracking-tight">
                  REPORT DP · SALES MONTHLY — {currentMonthSummary.name.toUpperCase()} 2026 (TANGGAL 1–
                  {currentMonthSummary.daysCount} {currentMonthSummary.short.toUpperCase()})
                </span>
                <span className="px-2 py-0.5 text-[11px] font-semibold bg-sky-500/30 text-sky-200 border border-sky-400/30 rounded-md">
                  {currentMonthSummary.monthActiveDays}/{currentMonthSummary.daysCount} Hari Terisi (1–{currentMonthSummary.daysCount} {currentMonthSummary.short})
                </span>
              </div>
              <div className="text-[11px] text-slate-300 font-mono-num mt-0.5">
                Sales Bulan (1–{currentMonthSummary.daysCount} {currentMonthSummary.short}): Rp{' '}
                {formatNumberId(currentMonthSummary.monthTotalSales) || '0'} · Total 12 Bulan: Rp{' '}
                {formatNumberId(grandAnnualSales) || '0'} · Qty:{' '}
                {formatNumberId(currentMonthSummary.monthTotalQty) || '0'} pcs · Trx:{' '}
                {formatNumberId(currentMonthSummary.monthTotalTrx) || '0'} trx
              </div>
            </div>

            {onSelectHorizon && (
              <div className="no-print inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15 shrink-0">
                <button
                  type="button"
                  onClick={() => onSelectHorizon('harian')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                >
                  Sales Daily
                </button>
                <button
                  type="button"
                  onClick={() => onSelectHorizon('weekly')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                >
                  Sales Weekly
                </button>
                <button
                  type="button"
                  onClick={() => onSelectHorizon('bulanan')}
                  className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all bg-sky-500 text-white shadow-2xs"
                >
                  Sales Monthly
                </button>
              </div>
            )}
          </div>

          {/* Row 2: Compact Month & Quarter Selector + Style Controls */}
          <div className="no-print flex flex-wrap items-center justify-between gap-2 pt-2.5 border-t border-white/15 w-full max-w-full min-w-0">
            <div className="flex flex-wrap items-center gap-1.5 min-w-0 max-w-full">
              <div className="inline-flex items-center gap-1 bg-white/10 p-0.5 rounded-lg border border-white/15 max-w-full min-w-0">
                <button
                  type="button"
                  disabled={currentMonthIdx <= 0}
                  onClick={() => {
                    const prev = MONTH_LIST[Math.max(0, currentMonthIdx - 1)];
                    if (prev) setSelectedMonthKey(prev.key);
                  }}
                  className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                  title="Bulan Sebelumnya"
                >
                  ◀
                </button>
                <select
                  aria-label="Pilih Bulan pada Sales Table Bulanan"
                  value={selectedMonthKey}
                  onChange={(e) => setSelectedMonthKey(e.target.value)}
                  className="px-2.5 py-1 rounded-md bg-amber-400 text-slate-950 text-xs font-extrabold focus:outline-none cursor-pointer min-w-0 max-w-[220px] sm:max-w-[280px] truncate shadow-2xs"
                >
                  {monthlySummaries.map((m) => (
                    <option key={m.key} value={m.key}>
                      {m.key}. {m.name.toUpperCase()} 2026 (1–{m.daysCount} {m.short}){' '}
                      {m.monthTotalSales > 0
                        ? `[Rp ${formatNumberId(m.monthTotalSales)}]`
                        : '[Rp 0]'}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  disabled={currentMonthIdx >= MONTH_LIST.length - 1}
                  onClick={() => {
                    const next = MONTH_LIST[Math.min(MONTH_LIST.length - 1, currentMonthIdx + 1)];
                    if (next) setSelectedMonthKey(next.key);
                  }}
                  className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                  title="Bulan Berikutnya"
                >
                  ▶
                </button>
              </div>

              <select
                aria-label="Filter Kuartal Bulanan"
                value={quarterFilter}
                onChange={(e) => {
                  const q = e.target.value as 'ALL' | 'Q1' | 'Q2' | 'Q3' | 'Q4';
                  setQuarterFilter(q);
                  if (q !== 'ALL') {
                    const firstInQ = MONTH_LIST.find((m) => m.quarter === q);
                    if (firstInQ) setSelectedMonthKey(firstInQ.key);
                  }
                }}
                className="px-2.5 py-1.5 rounded-lg bg-white/10 hover:bg-white/15 border border-white/15 text-amber-200 text-xs font-bold focus:outline-none cursor-pointer"
              >
                <option value="ALL" className="text-slate-900">Semua Kuartal</option>
                <option value="Q1" className="text-slate-900">Q1 (Jan–Mar)</option>
                <option value="Q2" className="text-slate-900">Q2 (Apr–Jun)</option>
                <option value="Q3" className="text-slate-900">Q3 (Jul–Sep)</option>
                <option value="Q4" className="text-slate-900">Q4 (Okt–Des)</option>
              </select>

              <div className="inline-flex items-center gap-1.5 px-2 py-1 rounded-lg bg-white/10 border border-white/20 text-xs">
                <Calendar className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                <input
                  type="date"
                  aria-label="Pilih Tanggal Kalender Sales Monthly"
                  value={`2026-${selectedMonthKey}-01`}
                  onChange={(e) => {
                    const parts = e.target.value.split('-');
                    if (parts.length === 3 && MONTH_LIST.some((m) => m.key === parts[1])) {
                      setSelectedMonthKey(parts[1]);
                      setQuarterFilter('ALL');
                    }
                  }}
                  title="Pilih tanggal lewat kalender untuk langsung lompat ke bulan terkait"
                  className="bg-transparent text-white font-bold font-mono-num focus:outline-none cursor-pointer [color-scheme:dark]"
                />
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-1.5">

              {(quarterFilter !== 'ALL' ||
                dayTypeFilter !== 'ALL' ||
                rowViewFilter !== 'ALL' ||
                monthlyViewMode !== 'ALL_MONTHLY') && (
                <button
                  type="button"
                  onClick={() => {
                    setQuarterFilter('ALL');
                    setDayTypeFilter('ALL');
                    setRowViewFilter('ALL');
                    setMonthlyViewMode('ALL_MONTHLY');
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-bold bg-sky-500/25 hover:bg-sky-500/35 text-sky-200 border border-sky-400/30 rounded-lg transition-colors whitespace-nowrap"
                  title="Reset Filter Tampilan Bulanan"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Reset Filter</span>
                </button>
              )}
            </div>
          </div>

          {/* Row 3: Quick 12-Month Selector Strip (Jan 1-31 .. Des 1-31) */}
          <div className="no-print grid grid-cols-4 sm:grid-cols-6 lg:grid-cols-12 gap-1.5 pt-2 border-t border-white/10">
            {monthlySummaries.map((m) => {
              const isSel = m.key === selectedMonthKey;
              const hasData = m.monthTotalSales > 0;
              return (
                <button
                  key={m.key}
                  type="button"
                  onClick={() => setSelectedMonthKey(m.key)}
                  className={`px-2 py-1.5 rounded-lg text-left transition-all border ${
                    isSel
                      ? 'bg-sky-400 text-slate-950 border-sky-300 font-extrabold shadow-xs'
                      : hasData
                        ? 'bg-emerald-500/20 text-emerald-100 border-emerald-400/30 hover:bg-emerald-500/30 font-bold'
                        : 'bg-white/5 text-slate-300 border-white/10 hover:bg-white/10 font-medium'
                  }`}
                >
                  <div className="flex items-center justify-between gap-1 text-[10px]">
                    <span className="truncate font-extrabold">
                      {m.key}. {m.short}
                    </span>
                    <span
                      className={`text-[9px] font-mono-num ${
                        isSel ? 'text-slate-900 font-bold' : 'opacity-75'
                      }`}
                    >
                      1–{m.daysCount}
                    </span>
                  </div>
                  <div
                    className={`text-[9px] font-mono-num truncate mt-0.5 ${
                      isSel
                        ? 'text-slate-950 font-extrabold'
                        : hasData
                          ? 'text-amber-300 font-bold'
                          : 'text-slate-400'
                    }`}
                  >
                    {hasData ? `Rp ${formatNumberId(m.monthTotalSales)}` : 'Rp 0'}
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* Clean Aligned 4-Column Sub-Header Filter Bar */}
        <div className="no-print bg-slate-50 border-b border-slate-200 px-3.5 sm:px-4 py-2.5 space-y-2 w-full max-w-full overflow-hidden">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 w-full">
            <select
              aria-label="Format Tabel Sales Bulanan"
              value={monthlyViewMode}
              onChange={(e) =>
                setMonthlyViewMode(
                  e.target.value as
                    | 'ALL_MONTHLY'
                    | 'REKAP_12_BULAN'
                    | 'MATRIX_WPS'
                )
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL_MONTHLY">
                Format: Semua (Matriks Mingguan &amp; 12 Bulan)
              </option>
              <option value="REKAP_12_BULAN">
                Format: Sales dari 12 Bulan (Jan–Des 2026)
              </option>
              <option value="MATRIX_WPS">
                Format: Matriks Mingguan Bulan ({currentMonthSummary.name})
              </option>
            </select>

            <select
              aria-label="Filter Hari Bulanan"
              value={dayTypeFilter}
              onChange={(e) =>
                setDayTypeFilter(e.target.value as 'ALL' | 'WEEKDAY' | 'WEEKEND')
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-indigo-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL">Hari: Semua (Weekday &amp; Weekend)</option>
              <option value="WEEKDAY">Hari: Hanya Weekday (Sen–Jum)</option>
              <option value="WEEKEND">Hari: Hanya Weekend (Sab–Min)</option>
            </select>

            <select
              aria-label="Filter Baris Bulanan"
              value={rowViewFilter}
              onChange={(e) =>
                setRowViewFilter(e.target.value as 'ALL' | 'SUMMARY_ONLY')
              }
              className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-800 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="ALL">Baris: Lengkap (Pagi, Siang &amp; Total)</option>
              <option value="SUMMARY_ONLY">Baris: Ringkas (Hanya Total)</option>
            </select>

            <select
              aria-label="Buka Cepat Weekly di Bulan Ini"
              value=""
              onChange={(e) => {
                const wk = Number(e.target.value);
                if (wk) handleOpenAndSwitchToWeek(wk);
              }}
              className="w-full px-2.5 py-1.5 rounded-lg bg-amber-50 border border-amber-300 text-xs font-bold text-amber-950 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
            >
              <option value="">Buka &amp; Edit Weekly di {currentMonthSummary.short}...</option>
              {currentMonthSummary.weekDetails.map((w) => (
                <option key={w.sched.weekNumber} value={String(w.sched.weekNumber)}>
                  Buka Weekly {w.sched.weekNumber} ({w.sched.seninSlash.slice(0, 5)}–
                  {w.sched.mingguSlash.slice(0, 5)})
                </option>
              ))}
            </select>
          </div>

          {/* Weekday vs Weekend Monthly Summary Pills */}
          <div className="grid grid-cols-2 gap-2 text-[11px] w-full">
            <button
              type="button"
              onClick={() =>
                setDayTypeFilter((p) => (p === 'WEEKDAY' ? 'ALL' : 'WEEKDAY'))
              }
              className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                dayTypeFilter === 'WEEKDAY'
                  ? 'bg-indigo-600 text-white border-indigo-600'
                  : 'bg-indigo-50/80 text-indigo-900 border-indigo-200 hover:bg-indigo-100'
              }`}
            >
              Weekday {currentMonthSummary.short}: Rp{' '}
              {formatNumberId(currentMonthSummary.monthWeekdaySales) || '0'} (
              {currentMonthSummary.monthTotalSales > 0
                ? Math.round(
                    (currentMonthSummary.monthWeekdaySales /
                      currentMonthSummary.monthTotalSales) *
                      100
                  )
                : 0}
              %)
            </button>
            <button
              type="button"
              onClick={() =>
                setDayTypeFilter((p) => (p === 'WEEKEND' ? 'ALL' : 'WEEKEND'))
              }
              className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                dayTypeFilter === 'WEEKEND'
                  ? 'bg-rose-600 text-white border-rose-600'
                  : 'bg-rose-50/80 text-rose-900 border-rose-200 hover:bg-rose-100'
              }`}
            >
              Weekend {currentMonthSummary.short}: Rp{' '}
              {formatNumberId(currentMonthSummary.monthWeekendSales) || '0'} (
              {currentMonthSummary.monthTotalSales > 0
                ? Math.round(
                    (currentMonthSummary.monthWeekendSales /
                      currentMonthSummary.monthTotalSales) *
                      100
                  )
                : 0}
              %)
            </button>
          </div>
        </div>

        {/* Ringkasan Kumulatif Bulanan Otomatis untuk Bulan Terpilih */}
        <MonthlyAutoSummaryCard
          tables={tables}
          selectedMonthKey={selectedMonthKey}
          onSelectMonthKey={setSelectedMonthKey}
          dayTypeFilter={dayTypeFilter}
        />

        {/* ===================================================================
            TABEL 1: MATRIKS SALES BULANAN PER WEEKLY (FORMAT WPS MERAH-KUNING)
           =================================================================== */}
        {(monthlyViewMode === 'MATRIX_WPS' || monthlyViewMode === 'ALL_MONTHLY') && (
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr>
                  <th
                    className={`${redHeaderClass} px-3 py-2.5 text-left font-bold min-w-[180px] sticky left-0 z-10`}
                  >
                    Parameter Bulanan ({currentMonthSummary.name})
                  </th>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <th
                      key={w.sched.weekNumber}
                      className={`${redHeaderClass} px-3 py-2.5 text-center font-bold min-w-[165px]`}
                    >
                      <div className="flex items-center justify-center gap-1.5">
                        <span>WEEKLY {w.sched.weekNumber}</span>
                        <button
                          type="button"
                          onClick={() => handleOpenAndSwitchToWeek(w.sched.weekNumber)}
                          className="no-print px-1.5 py-0.5 text-[10px] font-bold bg-amber-400 text-slate-950 rounded hover:bg-amber-300 transition-colors"
                          title={`Buka dan Edit Tabel Weekly ${w.sched.weekNumber}`}
                        >
                          Buka
                        </button>
                      </div>
                    </th>
                  ))}
                  <th
                    className={`${goldMonthlyTotalClass} px-3.5 py-2.5 text-center font-extrabold min-w-[175px]`}
                  >
                    TOTAL BULAN {currentMonthSummary.name.toUpperCase()}
                  </th>
                </tr>
              </thead>
              <tbody>
                {/* Baris Rentang Tanggal */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-2 font-semibold sticky left-0 z-10`}>
                    Tanggal di Bulan {currentMonthSummary.short}
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className="px-3 py-2 border border-slate-200 bg-slate-50 text-center font-mono-num font-semibold text-slate-800"
                    >
                      {w.firstInMonth} – {w.lastInMonth} ({w.daysInMonthCount} hr)
                    </td>
                  ))}
                  <td className="px-3 py-2 border border-slate-300 bg-amber-50 text-center font-bold text-slate-900">
                    01/{currentMonthSummary.key}/2026 – {String(currentMonthSummary.daysCount).padStart(2, '0')}/{currentMonthSummary.key}/2026 ({currentMonthSummary.monthActiveDays}/{currentMonthSummary.daysCount} Hari)
                  </td>
                </tr>

                {rowViewFilter === 'ALL' && (
                  <>
                    {/* Sales Shift Pagi */}
                    <tr className="border-b border-slate-200">
                      <td className={`${redHeaderClass} px-3 py-1.5 font-medium sticky left-0 z-10`}>
                        Sales Shift Pagi (Rp)
                      </td>
                      {currentMonthSummary.weekDetails.map((w) => (
                        <td
                          key={w.sched.weekNumber}
                          className="px-3 py-1.5 border border-slate-200 text-right font-mono-num bg-white"
                        >
                          {w.pagiSales > 0 ? formatNumberId(w.pagiSales) : '0'}
                        </td>
                      ))}
                      <td className="px-3 py-1.5 border border-slate-300 text-right font-mono-num font-bold bg-amber-50/70">
                        {currentMonthSummary.monthPagiSales > 0
                          ? formatNumberId(currentMonthSummary.monthPagiSales)
                          : '0'}
                      </td>
                    </tr>

                    {/* Sales Shift Siang */}
                    <tr className="border-b border-slate-200">
                      <td className={`${redHeaderClass} px-3 py-1.5 font-medium sticky left-0 z-10`}>
                        Sales Shift Siang (Rp)
                      </td>
                      {currentMonthSummary.weekDetails.map((w) => (
                        <td
                          key={w.sched.weekNumber}
                          className="px-3 py-1.5 border border-slate-200 text-right font-mono-num bg-white"
                        >
                          {w.siangSales > 0 ? formatNumberId(w.siangSales) : '0'}
                        </td>
                      ))}
                      <td className="px-3 py-1.5 border border-slate-300 text-right font-mono-num font-bold bg-amber-50/70">
                        {currentMonthSummary.monthSiangSales > 0
                          ? formatNumberId(currentMonthSummary.monthSiangSales)
                          : '0'}
                      </td>
                    </tr>
                  </>
                )}

                {/* TOTAL SALES MINGGUAN & BULANAN */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Total Sales (Rp)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className={`${yellowTotalClass} px-3 py-2 text-right font-mono-num`}
                    >
                      {w.totalSales > 0 ? formatNumberId(w.totalSales) : '0'}
                    </td>
                  ))}
                  <td
                    className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num text-xs`}
                  >
                    {currentMonthSummary.monthTotalSales > 0
                      ? formatNumberId(currentMonthSummary.monthTotalSales)
                      : '0'}
                  </td>
                </tr>

                {rowViewFilter === 'ALL' && (
                  <>
                    {/* Sales Weekday */}
                    <tr className="border-b border-slate-200">
                      <td className={`${redHeaderClass} px-3 py-1.5 font-medium sticky left-0 z-10`}>
                        Sales Weekday (Sen–Jum)
                      </td>
                      {currentMonthSummary.weekDetails.map((w) => (
                        <td
                          key={w.sched.weekNumber}
                          className="px-3 py-1.5 border border-slate-200 text-right font-mono-num text-indigo-900 bg-indigo-50/30"
                        >
                          {w.weekdaySales > 0 ? formatNumberId(w.weekdaySales) : '0'}
                        </td>
                      ))}
                      <td className="px-3 py-1.5 border border-slate-300 text-right font-mono-num font-bold text-indigo-950 bg-indigo-50/70">
                        {currentMonthSummary.monthWeekdaySales > 0
                          ? formatNumberId(currentMonthSummary.monthWeekdaySales)
                          : '0'}
                      </td>
                    </tr>

                    {/* Sales Weekend */}
                    <tr className="border-b border-slate-200">
                      <td className={`${redHeaderClass} px-3 py-1.5 font-medium sticky left-0 z-10`}>
                        Sales Weekend (Sab–Min)
                      </td>
                      {currentMonthSummary.weekDetails.map((w) => (
                        <td
                          key={w.sched.weekNumber}
                          className="px-3 py-1.5 border border-slate-200 text-right font-mono-num text-rose-900 bg-rose-50/30"
                        >
                          {w.weekendSales > 0 ? formatNumberId(w.weekendSales) : '0'}
                        </td>
                      ))}
                      <td className="px-3 py-1.5 border border-slate-300 text-right font-mono-num font-bold text-rose-950 bg-rose-50/70">
                        {currentMonthSummary.monthWeekendSales > 0
                          ? formatNumberId(currentMonthSummary.monthWeekendSales)
                          : '0'}
                      </td>
                    </tr>

                    {/* Qty Pagi & Siang */}
                    <tr className="border-b border-slate-200">
                      <td className={`${redHeaderClass} px-3 py-1.5 font-medium sticky left-0 z-10`}>
                        Qty Pagi / Siang (pcs)
                      </td>
                      {currentMonthSummary.weekDetails.map((w) => (
                        <td
                          key={w.sched.weekNumber}
                          className="px-3 py-1.5 border border-slate-200 text-right font-mono-num text-slate-700 bg-white"
                        >
                          P: {w.pagiQty} · S: {w.siangQty}
                        </td>
                      ))}
                      <td className="px-3 py-1.5 border border-slate-300 text-right font-mono-num font-semibold bg-amber-50/60">
                        P: {currentMonthSummary.monthPagiQty} · S:{' '}
                        {currentMonthSummary.monthSiangQty}
                      </td>
                    </tr>
                  </>
                )}

                {/* TOTAL QTY BARANG */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Total Qty Barang (pcs)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className={`${yellowTotalClass} px-3 py-2 text-right font-mono-num`}
                    >
                      {w.totalQty > 0 ? formatNumberId(w.totalQty) : '0'}
                    </td>
                  ))}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num`}>
                    {currentMonthSummary.monthTotalQty > 0
                      ? formatNumberId(currentMonthSummary.monthTotalQty)
                      : '0'}
                  </td>
                </tr>

                {/* TOTAL TRANSAKSI / STRUK */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Total Transaksi (Struk)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className={`${yellowTotalClass} px-3 py-2 text-right font-mono-num`}
                    >
                      {w.totalTrx > 0 ? formatNumberId(w.totalTrx) : '0'}
                    </td>
                  ))}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num`}>
                    {currentMonthSummary.monthTotalTrx > 0
                      ? formatNumberId(currentMonthSummary.monthTotalTrx)
                      : '0'}
                  </td>
                </tr>

                {/* BASKET SIZE */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-2 font-semibold sticky left-0 z-10`}>
                    Basket Size (Sales / Struk)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className="px-3 py-2 border border-slate-200 bg-amber-50/50 text-right font-mono-num font-bold"
                    >
                      {renderCalculatedCell(w.basketSize)}
                    </td>
                  ))}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num`}>
                    {renderCalculatedCell(currentMonthSummary.monthBasketSize)}
                  </td>
                </tr>

                {/* APC / PRICE POINT */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-2 font-semibold sticky left-0 z-10`}>
                    APC / Price Point (Sales / Qty)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className="px-3 py-2 border border-slate-200 bg-amber-50/50 text-right font-mono-num font-bold"
                    >
                      {renderCalculatedCell(w.apc)}
                    </td>
                  ))}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num`}>
                    {renderCalculatedCell(currentMonthSummary.monthApc)}
                  </td>
                </tr>

                {/* KONTRIBUSI SALES MINGGUAN TERHADAP BULAN (%) */}
                <tr className="border-b border-slate-200">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Kontribusi Sales (%)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => (
                    <td
                      key={w.sched.weekNumber}
                      className="px-3 py-2 border border-slate-200 bg-sky-50/60 text-right font-mono-num font-extrabold text-sky-950"
                    >
                      <div className="flex items-center justify-end gap-1.5">
                        <span className="px-1.5 py-0.5 rounded bg-sky-100 border border-sky-300 text-sky-900 text-[11px] font-extrabold">
                          {w.contributionPct.toFixed(1)}%
                        </span>
                      </div>
                    </td>
                  ))}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num font-extrabold`}>
                    {currentMonthSummary.monthTotalSales > 0 ? '100.0%' : '0.0%'}
                  </td>
                </tr>

                {/* GROWTH PENJUALAN MINGGUAN (WoW %) */}
                <tr className="border-b border-slate-300">
                  <td className={`${redHeaderClass} px-3 py-2 font-bold sticky left-0 z-10`}>
                    Growth Mingguan (%)
                  </td>
                  {currentMonthSummary.weekDetails.map((w) => {
                    const isUp = w.growthStatus === 'UP' || w.growthStatus === 'NEW';
                    const isDown = w.growthStatus === 'DOWN';
                    const isBase = w.growthStatus === 'BASE';
                    return (
                      <td
                        key={w.sched.weekNumber}
                        className={`px-3 py-2 border border-slate-200 text-right font-mono-num font-bold ${
                          isUp
                            ? 'bg-emerald-50/70 text-emerald-900'
                            : isDown
                              ? 'bg-rose-50/70 text-rose-900'
                              : 'bg-slate-50/70 text-slate-700'
                        }`}
                      >
                        {isBase ? (
                          <div className="space-y-0.5">
                            <span className="inline-block px-1.5 py-0.5 rounded bg-amber-100 border border-amber-300 text-amber-900 text-[10px] font-extrabold">
                              Minggu Awal
                            </span>
                            {w.prevWeekNumber && (
                              <div className="text-[10px] text-slate-500">
                                Basis Bulan {currentMonthSummary.short}
                              </div>
                            )}
                          </div>
                        ) : (
                          <div className="space-y-0.5">
                            <div>
                              <span
                                className={`inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded text-[11px] font-extrabold border ${
                                  isUp
                                    ? 'bg-emerald-100 text-emerald-800 border-emerald-300'
                                    : isDown
                                      ? 'bg-rose-100 text-rose-800 border-rose-300'
                                      : 'bg-slate-100 text-slate-700 border-slate-300'
                                }`}
                              >
                                {isUp
                                  ? `▲ +${(w.growthPct ?? 100).toFixed(1)}%`
                                  : isDown
                                    ? `▼ ${(w.growthPct ?? 0).toFixed(1)}%`
                                    : '▬ 0.0%'}
                              </span>
                            </div>
                            <div
                              className={`text-[10px] font-semibold ${
                                w.growthDiff > 0
                                  ? 'text-emerald-700'
                                  : w.growthDiff < 0
                                    ? 'text-rose-700'
                                    : 'text-slate-500'
                              }`}
                            >
                              {w.growthDiff > 0
                                ? `+Rp ${formatNumberId(w.growthDiff)}`
                                : w.growthDiff < 0
                                  ? `-Rp ${formatNumberId(Math.abs(w.growthDiff))}`
                                  : 'Rp 0'}
                              {w.prevWeekNumber ? ` (vs W${w.prevWeekNumber})` : ''}
                            </div>
                          </div>
                        )}
                      </td>
                    );
                  })}
                  <td className={`${goldMonthlyTotalClass} px-3.5 py-2 text-right font-mono-num font-extrabold`}>
                    {(() => {
                      const prevMonthSummary =
                        currentMonthIdx > 0 ? monthlySummaries[currentMonthIdx - 1] : null;
                      const prevMonthSales = prevMonthSummary?.monthTotalSales ?? 0;
                      const curMonthSales = currentMonthSummary.monthTotalSales;
                      if (curMonthSales === 0 && prevMonthSales === 0) {
                        return <span>▬ 0.0%</span>;
                      }
                      if (prevMonthSales === 0 && curMonthSales > 0) {
                        // Compute intra-month growth from first active week to last active week in month
                        const activeWks = currentMonthSummary.weekDetails.filter(
                          (wk) => wk.totalSales > 0
                        );
                        if (activeWks.length >= 2) {
                          const firstW = activeWks[0];
                          const lastW = activeWks[activeWks.length - 1];
                          const diffW = lastW.totalSales - firstW.totalSales;
                          const pctW =
                            firstW.totalSales > 0
                              ? ((diffW / firstW.totalSales) * 100).toFixed(1)
                              : '0.0';
                          return (
                            <div className="space-y-0.5">
                              <div>
                                {diffW > 0
                                  ? `▲ +${pctW}%`
                                  : diffW < 0
                                    ? `▼ ${pctW}%`
                                    : '▬ 0.0%'}
                              </div>
                              <div className="text-[10px] font-bold opacity-85">
                                W{firstW.sched.weekNumber} → W{lastW.sched.weekNumber}
                              </div>
                            </div>
                          );
                        }
                        return <span>Bulan Aktif</span>;
                      }
                      const mDiff = curMonthSales - prevMonthSales;
                      const mPct = ((mDiff / prevMonthSales) * 100).toFixed(1);
                      return (
                        <div className="space-y-0.5">
                          <div>
                            {mDiff > 0
                              ? `▲ +${mPct}%`
                              : mDiff < 0
                                ? `▼ ${mPct}%`
                                : '▬ 0.0%'}
                          </div>
                          {prevMonthSummary && (
                            <div className="text-[10px] font-bold opacity-85">
                              vs {prevMonthSummary.short}
                            </div>
                          )}
                        </div>
                      );
                    })()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* ===================================================================
            TABEL 2: REKAPITULASI SALES 12 BULAN SETAHUN (JANUARI - DESEMBER 2026)
           =================================================================== */}
        <div className="border-t border-slate-200 p-3.5 sm:p-4 space-y-2.5 bg-white">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="text-xs font-extrabold text-slate-900 uppercase tracking-wide flex items-center gap-1.5">
              <CalendarDays className="w-4 h-4 text-rose-700" />
              Tabel Rekapitulasi Sales Bulanan Setahun (Januari – Desember 2026)
            </div>
            <span className="text-[11px] font-mono-num text-slate-600">
              Grand Total 2026: <strong>Rp {formatNumberId(grandAnnualSales) || '0'}</strong> (
              {formatNumberId(grandAnnualQty) || '0'} pcs ·{' '}
              {formatNumberId(grandAnnualTrx) || '0'} trx)
            </span>
          </div>

          <div className="overflow-x-auto border border-slate-300 rounded-lg">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr>
                  <th className={`${redHeaderClass} px-3 py-2.5 text-left font-bold`}>
                    Bulan (2026)
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-center font-bold`}>
                    Kuartal
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-center font-bold`}>
                    Tanggal Kalender (1–30/31)
                  </th>
                  <th className={`${redHeaderClass} px-2 py-2.5 text-center font-bold`}>
                    Hari Terisi
                  </th>
                  {rowViewFilter === 'ALL' && (
                    <>
                      <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                        Sales Pagi
                      </th>
                      <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                        Sales Siang
                      </th>
                      <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                        Sales Weekday
                      </th>
                      <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                        Sales Weekend
                      </th>
                    </>
                  )}
                  <th className={`${goldMonthlyTotalClass} px-3 py-2.5 text-right font-extrabold`}>
                    Total Sales Bulanan
                  </th>
                  <th className={`${goldMonthlyTotalClass} px-2.5 py-2.5 text-right font-extrabold`}>
                    Total Qty
                  </th>
                  <th className={`${goldMonthlyTotalClass} px-2.5 py-2.5 text-right font-extrabold`}>
                    Total Struk
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    Basket Size
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    APC
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    Kontribusi (%)
                  </th>
                  <th className={`${redHeaderClass} px-2.5 py-2.5 text-right font-bold`}>
                    Growth (%)
                  </th>
                  <th className={`${redHeaderClass} px-2 py-2.5 text-center font-bold no-print`}>
                    Aksi
                  </th>
                </tr>
              </thead>
              <tbody>
                {monthlySummaries
                  .filter((m) =>
                    quarterFilter === 'ALL' ? true : m.quarter === quarterFilter
                  )
                  .map((m) => {
                    const isSelected = m.key === selectedMonthKey;
                    const mFullIdx = monthlySummaries.findIndex((item) => item.key === m.key);
                    const prevMonth = mFullIdx > 0 ? monthlySummaries[mFullIdx - 1] : null;
                    const prevSales = prevMonth?.monthTotalSales ?? 0;
                    const contribPct =
                      grandAnnualSales > 0
                        ? (m.monthTotalSales / grandAnnualSales) * 100
                        : 0;
                    const mDiff = m.monthTotalSales - prevSales;
                    const mGrowthPct =
                      prevSales > 0 ? (mDiff / prevSales) * 100 : null;
                    return (
                      <tr
                        key={m.key}
                        onClick={() => setSelectedMonthKey(m.key)}
                        className={`border-b border-slate-200 cursor-pointer transition-colors ${
                          isSelected
                            ? 'bg-amber-50/90 font-semibold'
                            : 'bg-white hover:bg-slate-50'
                        }`}
                      >
                        <td className="px-3 py-2 border border-slate-200 font-bold text-slate-900 whitespace-nowrap">
                          {m.key}. {m.name} 2026 (1–{m.daysCount} {m.short})
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-center font-mono-num">
                          <span className="px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 font-bold text-[10px]">
                            {m.quarter}
                          </span>
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-center font-mono-num whitespace-nowrap">
                          {m.startSlash} – {m.endSlash}
                        </td>
                        <td className="px-2 py-2 border border-slate-200 text-center font-mono-num">
                          {m.monthActiveDays}/{m.daysCount} hari
                        </td>
                        {rowViewFilter === 'ALL' && (
                          <>
                            <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num">
                              {m.monthPagiSales > 0
                                ? formatNumberId(m.monthPagiSales)
                                : '0'}
                            </td>
                            <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num">
                              {m.monthSiangSales > 0
                                ? formatNumberId(m.monthSiangSales)
                                : '0'}
                            </td>
                            <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num text-indigo-900">
                              {m.monthWeekdaySales > 0
                                ? formatNumberId(m.monthWeekdaySales)
                                : '0'}
                            </td>
                            <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num text-rose-900">
                              {m.monthWeekendSales > 0
                                ? formatNumberId(m.monthWeekendSales)
                                : '0'}
                            </td>
                          </>
                        )}
                        <td className={`${yellowTotalClass} px-3 py-2 text-right font-mono-num`}>
                          {m.monthTotalSales > 0
                            ? formatNumberId(m.monthTotalSales)
                            : '0'}
                        </td>
                        <td className={`${yellowTotalClass} px-2.5 py-2 text-right font-mono-num`}>
                          {m.monthTotalQty > 0 ? formatNumberId(m.monthTotalQty) : '0'}
                        </td>
                        <td className={`${yellowTotalClass} px-2.5 py-2 text-right font-mono-num`}>
                          {m.monthTotalTrx > 0 ? formatNumberId(m.monthTotalTrx) : '0'}
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num font-semibold bg-amber-50/40">
                          {renderCalculatedCell(m.monthBasketSize)}
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num font-semibold bg-amber-50/40">
                          {renderCalculatedCell(m.monthApc)}
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num font-bold text-sky-900 bg-sky-50/40">
                          {contribPct.toFixed(1)}%
                        </td>
                        <td className="px-2.5 py-2 border border-slate-200 text-right font-mono-num font-bold whitespace-nowrap">
                          {m.monthTotalSales === 0 && prevSales === 0 ? (
                            <span className="text-slate-400">▬ 0.0%</span>
                          ) : mGrowthPct !== null ? (
                            <span
                              className={
                                mDiff > 0
                                  ? 'text-emerald-700'
                                  : mDiff < 0
                                    ? 'text-rose-700'
                                    : 'text-slate-600'
                              }
                            >
                              {mDiff > 0
                                ? `▲ +${mGrowthPct.toFixed(1)}%`
                                : mDiff < 0
                                  ? `▼ ${mGrowthPct.toFixed(1)}%`
                                  : '▬ 0.0%'}
                            </span>
                          ) : (
                            <span className="text-amber-800 text-[11px]">Bulan Awal</span>
                          )}
                        </td>
                        <td className="px-2 py-2 border border-slate-200 text-center no-print">
                          <div className="inline-flex items-center gap-1">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setSelectedMonthKey(m.key);
                                setMonthlyViewMode('MATRIX_WPS');
                              }}
                              className="px-2 py-1 rounded bg-slate-900 hover:bg-slate-800 text-amber-300 text-[11px] font-bold whitespace-nowrap"
                            >
                              Pilih {m.short}
                            </button>
                            {m.monthTotalSales > 0 ||
                            m.monthTotalQty > 0 ||
                            m.monthTotalTrx > 0 ? (
                              confirmMonthlyResetTarget === `ROW_${m.key}` ? (
                                <div
                                  onClick={(e) => e.stopPropagation()}
                                  className="inline-flex items-center gap-0.5 bg-red-950 border border-red-700 px-1.5 py-0.5 rounded"
                                >
                                  <button
                                    type="button"
                                    onClick={() => handleResetSingleMonthAll(m.key)}
                                    className="p-0.5 text-red-300 hover:bg-red-800 rounded"
                                    title={`Konfirmasi Reset Bulan ${m.name}`}
                                  >
                                    <Check className="w-3 h-3" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setConfirmMonthlyResetTarget(null)}
                                    className="p-0.5 text-slate-300 hover:bg-slate-800 rounded"
                                    title="Batal"
                                  >
                                    <X className="w-3 h-3" />
                                  </button>
                                </div>
                              ) : (
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setConfirmMonthlyResetTarget(`ROW_${m.key}`);
                                  }}
                                  className="inline-flex items-center gap-0.5 px-2 py-1 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-[11px] font-bold whitespace-nowrap transition-colors"
                                  title={`Reset Data Bulan ${m.name}`}
                                >
                                  <RotateCcw className="w-3 h-3" />
                                  Reset
                                </button>
                              )
                            ) : null}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
              <tfoot>
                <tr className={goldMonthlyTotalClass}>
                  <td colSpan={3} className="px-3 py-2.5 text-left font-extrabold">
                    GRAND TOTAL TAHUNAN (2026)
                  </td>
                  <td className="px-2 py-2.5 text-center font-mono-num font-extrabold">
                    {grandAnnualActiveDays} hari
                  </td>
                  {rowViewFilter === 'ALL' && (
                    <>
                      <td className="px-2.5 py-2.5 text-right font-mono-num">
                        {formatNumberId(grandAnnualPagi) || '0'}
                      </td>
                      <td className="px-2.5 py-2.5 text-right font-mono-num">
                        {formatNumberId(grandAnnualSiang) || '0'}
                      </td>
                      <td className="px-2.5 py-2.5 text-right font-mono-num">
                        {formatNumberId(grandAnnualWeekday) || '0'}
                      </td>
                      <td className="px-2.5 py-2.5 text-right font-mono-num">
                        {formatNumberId(grandAnnualWeekend) || '0'}
                      </td>
                    </>
                  )}
                  <td className="px-3 py-2.5 text-right font-mono-num font-extrabold">
                    {formatNumberId(grandAnnualSales) || '0'}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num font-extrabold">
                    {formatNumberId(grandAnnualQty) || '0'}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num font-extrabold">
                    {formatNumberId(grandAnnualTrx) || '0'}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num">
                    {renderCalculatedCell(
                      grandAnnualTrx > 0
                        ? Math.round(grandAnnualSales / grandAnnualTrx)
                        : null
                    )}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num">
                    {renderCalculatedCell(
                      grandAnnualQty > 0
                        ? Math.round(grandAnnualSales / grandAnnualQty)
                        : null
                    )}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num font-extrabold">
                    {grandAnnualSales > 0 ? '100.0%' : '0.0%'}
                  </td>
                  <td className="px-2.5 py-2.5 text-right font-mono-num font-extrabold">
                    Total 12 Bln
                  </td>
                  <td className="px-2 py-2.5 text-center no-print">-</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>

        {/* ===================================================================
            GRAFIK DIAGRAM SALES TABLE BULANAN (SESUAI FOTO)
           =================================================================== */}
        <div className="p-3.5 sm:p-4 bg-white border-t border-slate-200 space-y-4">
          <PhotoStyleSalesDiagram
            headerTitle={`DIAGRAM SALES BULANAN — ${currentMonthSummary.name.toUpperCase()} 2026 (PER WEEKLY)`}
            headerSubtitle="Diagram Batang Vertikal Mingguan & Garis Tren Penjualan Shift Pagi vs Shift Siang"
            badgeText="DIAGRAM SALES (OMZET RUPIAH)"
            cardTitle="Grafik Batang Shift Pagi vs Siang & Garis Tren"
            totalSales={currentMonthSummary.monthTotalSales}
            totalPagi={currentMonthSummary.monthPagiSales}
            totalSiang={currentMonthSummary.monthSiangSales}
            bars={currentMonthSummary.weekDetails.map((w) => ({
              key: `w-${w.sched.weekNumber}`,
              label: `W${w.sched.weekNumber}`,
              subLabel: `${w.sched.seninSlash.slice(0, 5)}–${w.sched.mingguSlash.slice(0, 5)}`,
              pagiSales: w.pagiSales,
              siangSales: w.siangSales,
              totalSales: w.totalSales,
              onClick: () => handleOpenAndSwitchToWeek(w.sched.weekNumber),
            }))}
          />

          <PhotoStyleSalesDiagram
            headerTitle="DIAGRAM SALES TAHUNAN — 12 BULAN (JANUARI - DESEMBER 2026)"
            headerSubtitle="Diagram Batang Vertikal 12 Bulan & Garis Tren Penjualan Shift Pagi vs Shift Siang"
            badgeText="DIAGRAM SALES (OMZET RUPIAH)"
            cardTitle="Grafik Batang Shift Pagi vs Siang & Garis Tren"
            totalSales={grandAnnualSales}
            totalPagi={grandAnnualPagi}
            totalSiang={grandAnnualSiang}
            bars={monthlySummaries.map((m) => ({
              key: m.key,
              label: m.short,
              subLabel: `1–${m.daysCount}`,
              pagiSales: m.monthPagiSales,
              siangSales: m.monthSiangSales,
              totalSales: m.monthTotalSales,
              isSelected: m.key === selectedMonthKey,
              onClick: () => setSelectedMonthKey(m.key),
            }))}
          />

          <GrowthAndContributionBarDiagram
            headerTitle={`DIAGRAM BATANG GROWTH & KONTRIBUSI MINGGUAN — ${currentMonthSummary.name.toUpperCase()} 2026`}
            headerSubtitle="Grafik Batang Pertumbuhan Mingguan (WoW Growth %) & Porsi Kontribusi Tiap Weekly terhadap Bulan"
            periodBadgeText={`BULAN ${currentMonthSummary.name.toUpperCase()} 2026`}
            items={currentMonthSummary.weekDetails.map((w, idx) => ({
              key: `w-gc-${w.sched.weekNumber}`,
              label: `W${w.sched.weekNumber}`,
              subLabel: `${w.sched.seninSlash.slice(0, 5)}–${w.sched.mingguSlash.slice(0, 5)}`,
              pagiSales: w.pagiSales,
              siangSales: w.siangSales,
              totalSales: w.totalSales,
              prevSales:
                idx > 0 ? currentMonthSummary.weekDetails[idx - 1].totalSales : null,
              onClick: () => handleOpenAndSwitchToWeek(w.sched.weekNumber),
            }))}
          />

          <GrowthAndContributionBarDiagram
            headerTitle="DIAGRAM BATANG GROWTH & KONTRIBUSI TAHUNAN — 12 BULAN (JANUARI - DESEMBER 2026)"
            headerSubtitle="Grafik Batang Pertumbuhan Bulanan (MoM Growth %) & Porsi Kontribusi Tiap Bulan terhadap Total Tahunan (100%)"
            periodBadgeText="12 BULAN (JAN – DES 2026)"
            items={monthlySummaries.map((m, idx) => ({
              key: `m-gc-${m.key}`,
              label: m.short,
              subLabel: `1–${m.daysCount}`,
              pagiSales: m.monthPagiSales,
              siangSales: m.monthSiangSales,
              totalSales: m.monthTotalSales,
              prevSales: idx > 0 ? monthlySummaries[idx - 1].monthTotalSales : null,
              isSelected: m.key === selectedMonthKey,
              onClick: () => setSelectedMonthKey(m.key),
            }))}
          />
        </div>
      </section>
    </div>
  );
};
