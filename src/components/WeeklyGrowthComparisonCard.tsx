import React, { useState, useMemo } from 'react';
import {
  WeeklySalesTable,
  ANNUAL_2026_WEEKS,
  calculateDayMetrics,
  calculateWeekMetrics,
  createWeeklyTableFromAnnualSchedule,
  formatNumberId,
  getIndoDayName,
  getWeeklyPeriodeText,
  isWeekendDate,
} from '../data/initialData';
import {
  TrendingUp,
  TrendingDown,
  Minus,
  ArrowUp,
  ArrowDown,
  ArrowRight,
  ArrowUpRight,
  ArrowDownRight,
  CalendarRange,
  Sun,
  Sunset,
  ShoppingBag,
  Receipt,
  Tag,
  BarChart3,
  RefreshCw,
} from 'lucide-react';

interface WeeklyGrowthComparisonCardProps {
  tables: WeeklySalesTable[];
  activeWeek: WeeklySalesTable;
  onSelectWeekNumber?: (weekNum: number) => void;
}

export interface GrowthDelta {
  current: number;
  previous: number;
  diff: number;
  pct: number | null; // null when previous === 0 and current > 0 (new baseline)
  status: 'UP' | 'DOWN' | 'FLAT' | 'NEW';
}

export function computeGrowthDelta(current: number, previous: number): GrowthDelta {
  const diff = current - previous;
  if (previous <= 0 && current <= 0) {
    return { current, previous, diff: 0, pct: 0, status: 'FLAT' };
  }
  if (previous <= 0 && current > 0) {
    return { current, previous, diff, pct: 100, status: 'NEW' };
  }
  const pct = ((current - previous) / previous) * 100;
  if (Math.abs(pct) < 0.05) {
    return { current, previous, diff, pct: 0, status: 'FLAT' };
  }
  return {
    current,
    previous,
    diff,
    pct,
    status: pct > 0 ? 'UP' : 'DOWN',
  };
}

export const WeeklyGrowthComparisonCard: React.FC<WeeklyGrowthComparisonCardProps> = ({
  tables,
  activeWeek,
  onSelectWeekNumber,
}) => {
  // Custom override for comparison week number (null = automatic W(N-1))
  const [customCompareWeekNum, setCustomCompareWeekNum] = useState<number | null>(null);
  const [showDayByDayComparison, setShowDayByDayComparison] = useState<boolean>(true);

  // Helper to resolve any week 1..53 from tables or ANNUAL_2026_WEEKS
  const resolveWeekTable = (wkNum: number): WeeklySalesTable => {
    const found = tables.find((t) => t.weekNumber === wkNum);
    if (found) return found;
    const sched =
      ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum) || ANNUAL_2026_WEEKS[0];
    return createWeeklyTableFromAnnualSchedule(sched);
  };

  const currentWeekNum = activeWeek.weekNumber;
  const defaultPrevWeekNum = Math.max(1, currentWeekNum - 1);
  const effectiveCompareWeekNum =
    customCompareWeekNum !== null ? customCompareWeekNum : defaultPrevWeekNum;

  const prevWeek = useMemo(
    () => resolveWeekTable(effectiveCompareWeekNum),
    [tables, effectiveCompareWeekNum]
  );

  // Compute detailed metrics for a WeeklySalesTable
  const extractDetailedWeekStats = (table: WeeklySalesTable) => {
    const wm = calculateWeekMetrics(table);
    let pagiSales = 0;
    let siangSales = 0;
    let pagiQty = 0;
    let siangQty = 0;
    let pagiTrx = 0;
    let siangTrx = 0;
    let weekdaySales = 0;
    let weekendSales = 0;
    let activeDays = 0;

    table.days.forEach((d) => {
      const pS = d.pagi.sales ?? 0;
      const sS = d.siang.sales ?? 0;
      const pQ = d.pagi.qty ?? 0;
      const sQ = d.siang.qty ?? 0;
      const pT = d.pagi.trx ?? 0;
      const sT = d.siang.trx ?? 0;
      const daySales = pS + sS;

      pagiSales += pS;
      siangSales += sS;
      pagiQty += pQ;
      siangQty += sQ;
      pagiTrx += pT;
      siangTrx += sT;

      if (daySales > 0 || pQ + sQ > 0 || pT + sT > 0) {
        activeDays += 1;
      }
      if (isWeekendDate(d.tanggal)) {
        weekendSales += daySales;
      } else {
        weekdaySales += daySales;
      }
    });

    return {
      totalSales: wm.weeklySales,
      totalQty: wm.weeklyQty,
      totalTrx: wm.weeklyTrx,
      pricePoint: wm.weeklyPricePoint ?? 0,
      bigSize: wm.weeklyBigSize ?? 0,
      pagiSales,
      siangSales,
      pagiQty,
      siangQty,
      pagiTrx,
      siangTrx,
      weekdaySales,
      weekendSales,
      activeDays,
      avgDailySales: Math.round(wm.weeklySales / 7),
    };
  };

  const currStats = useMemo(() => extractDetailedWeekStats(activeWeek), [activeWeek]);
  const prevStats = useMemo(() => extractDetailedWeekStats(prevWeek), [prevWeek]);

  const salesGrowth = computeGrowthDelta(currStats.totalSales, prevStats.totalSales);
  const pagiGrowth = computeGrowthDelta(currStats.pagiSales, prevStats.pagiSales);
  const siangGrowth = computeGrowthDelta(currStats.siangSales, prevStats.siangSales);
  const qtyGrowth = computeGrowthDelta(currStats.totalQty, prevStats.totalQty);
  const trxGrowth = computeGrowthDelta(currStats.totalTrx, prevStats.totalTrx);
  const ppGrowth = computeGrowthDelta(currStats.pricePoint, prevStats.pricePoint);
  const bsGrowth = computeGrowthDelta(currStats.bigSize, prevStats.bigSize);
  const weekdayGrowth = computeGrowthDelta(currStats.weekdaySales, prevStats.weekdaySales);
  const weekendGrowth = computeGrowthDelta(currStats.weekendSales, prevStats.weekendSales);

  // Build multi-week sequential trend list around active weeks
  const multiWeekTrendList = useMemo(() => {
    // Include all weeks in `tables` plus immediate neighbors of activeWeek
    const weekNumsSet = new Set<number>();
    tables.forEach((t) => weekNumsSet.add(t.weekNumber));
    for (
      let w = Math.max(1, currentWeekNum - 2);
      w <= Math.min(53, currentWeekNum + 2);
      w++
    ) {
      weekNumsSet.add(w);
    }

    const sortedWeekNums = Array.from(weekNumsSet).sort((a, b) => a - b);

    return sortedWeekNums.map((wkNum) => {
      const tbl = resolveWeekTable(wkNum);
      const st = extractDetailedWeekStats(tbl);
      const prevWkNum = wkNum > 1 ? wkNum - 1 : 1;
      const prevTbl = resolveWeekTable(prevWkNum);
      const prevSt = wkNum > 1 ? extractDetailedWeekStats(prevTbl) : null;

      const growth = prevSt
        ? computeGrowthDelta(st.totalSales, prevSt.totalSales)
        : {
            current: st.totalSales,
            previous: 0,
            diff: st.totalSales,
            pct: null,
            status: 'FLAT' as const,
          };
      const qtyG = prevSt
        ? computeGrowthDelta(st.totalQty, prevSt.totalQty)
        : null;
      const trxG = prevSt
        ? computeGrowthDelta(st.totalTrx, prevSt.totalTrx)
        : null;

      return {
        weekNumber: wkNum,
        table: tbl,
        stats: st,
        prevWeekNumber: wkNum > 1 ? prevWkNum : null,
        prevSales: prevSt ? prevSt.totalSales : 0,
        growth,
        qtyGrowth: qtyG,
        trxGrowth: trxG,
      };
    });
  }, [tables, currentWeekNum]);

  // Helper to render a directional arrow icon + status label + percentage badge
  const renderGrowthBadge = (
    g: GrowthDelta,
    size: 'sm' | 'md' | 'lg' = 'md',
    compareLabel?: string,
    showStatusWord: boolean = true
  ) => {
    const sizeClasses =
      size === 'lg'
        ? 'px-3 py-1.5 text-xs sm:text-sm font-extrabold gap-1.5 rounded-xl shadow-2xs'
        : size === 'md'
          ? 'px-2.5 py-1 text-[11px] font-extrabold gap-1 rounded-lg'
          : 'px-1.5 py-0.5 text-[10px] font-extrabold gap-1 rounded-md';

    const iconSize = size === 'lg' ? 'w-4 h-4 stroke-[2.75]' : 'w-3.5 h-3.5 stroke-[2.5]';

    if (g.status === 'UP') {
      return (
        <span
          title={`Tren Naik +${g.pct?.toFixed(1)}%`}
          className={`inline-flex items-center font-mono-num bg-emerald-100 text-emerald-900 border border-emerald-400 ${sizeClasses}`}
        >
          <span className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-emerald-600 text-white shrink-0">
            <ArrowUp className={size === 'lg' ? 'w-3 h-3 stroke-[3]' : 'w-2.5 h-2.5 stroke-[3]'} />
          </span>
          {showStatusWord && (
            <span className="font-sans font-extrabold uppercase tracking-wider text-emerald-800">
              NAIK
            </span>
          )}
          <span>+{g.pct?.toFixed(1)}%</span>
          {compareLabel && (
            <span className="font-sans font-semibold opacity-80">{compareLabel}</span>
          )}
        </span>
      );
    }

    if (g.status === 'DOWN') {
      return (
        <span
          title={`Tren Turun ${g.pct?.toFixed(1)}%`}
          className={`inline-flex items-center font-mono-num bg-rose-100 text-rose-900 border border-rose-400 ${sizeClasses}`}
        >
          <span className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-rose-600 text-white shrink-0">
            <ArrowDown className={size === 'lg' ? 'w-3 h-3 stroke-[3]' : 'w-2.5 h-2.5 stroke-[3]'} />
          </span>
          {showStatusWord && (
            <span className="font-sans font-extrabold uppercase tracking-wider text-rose-800">
              TURUN
            </span>
          )}
          <span>{g.pct?.toFixed(1)}%</span>
          {compareLabel && (
            <span className="font-sans font-semibold opacity-80">{compareLabel}</span>
          )}
        </span>
      );
    }

    if (g.status === 'NEW') {
      return (
        <span
          title="Tren Naik Baru (+100%)"
          className={`inline-flex items-center font-mono-num bg-emerald-100 text-emerald-900 border border-emerald-400 ${sizeClasses}`}
        >
          <span className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-emerald-600 text-white shrink-0">
            <ArrowUpRight className={size === 'lg' ? 'w-3 h-3 stroke-[3]' : 'w-2.5 h-2.5 stroke-[3]'} />
          </span>
          {showStatusWord && (
            <span className="font-sans font-extrabold uppercase tracking-wider text-emerald-800">
              NAIK
            </span>
          )}
          <span>+100.0% (Baru)</span>
          {compareLabel && (
            <span className="font-sans font-semibold opacity-80">{compareLabel}</span>
          )}
        </span>
      );
    }

    return (
      <span
        title="Tren Stabil (0.0%)"
        className={`inline-flex items-center font-mono-num bg-slate-100 text-slate-700 border border-slate-300 ${sizeClasses}`}
      >
        <span className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-slate-500 text-white shrink-0">
          <ArrowRight className={iconSize && size === 'lg' ? 'w-3 h-3 stroke-[3]' : 'w-2.5 h-2.5 stroke-[3]'} />
        </span>
        {showStatusWord && (
          <span className="font-sans font-extrabold uppercase tracking-wider text-slate-700">
            STABIL
          </span>
        )}
        <span>0.0%</span>
        {compareLabel && (
          <span className="font-sans font-semibold opacity-80">{compareLabel}</span>
        )}
      </span>
    );
  };

  // Helper to render compact inline arrow + delta text
  const renderDirectionalDelta = (diff: number, formattedText: string) => {
    if (diff > 0) {
      return (
        <span className="inline-flex items-center gap-1 text-emerald-700 font-extrabold">
          <ArrowUp className="w-3.5 h-3.5 stroke-[2.75] text-emerald-600 shrink-0" />
          <span>▲ {formattedText}</span>
        </span>
      );
    }
    if (diff < 0) {
      return (
        <span className="inline-flex items-center gap-1 text-rose-700 font-extrabold">
          <ArrowDown className="w-3.5 h-3.5 stroke-[2.75] text-rose-600 shrink-0" />
          <span>▼ {formattedText}</span>
        </span>
      );
    }
    return (
      <span className="inline-flex items-center gap-1 text-slate-500 font-bold">
        <ArrowRight className="w-3.5 h-3.5 stroke-[2.5] text-slate-400 shrink-0" />
        <span>▬ {formattedText}</span>
      </span>
    );
  };

  // Day-by-day summary count of Naik / Turun / Stabil across H1-H7
  const dailyTrendCounts = useMemo(() => {
    let upCount = 0;
    let downCount = 0;
    let flatCount = 0;
    activeWeek.days.forEach((currDay, idx) => {
      const prevDay = prevWeek.days[idx];
      const currDm = calculateDayMetrics(currDay);
      const prevDm = prevDay ? calculateDayMetrics(prevDay) : { totalSales: 0 };
      const dg = computeGrowthDelta(currDm.totalSales, prevDm.totalSales);
      if (dg.status === 'UP' || dg.status === 'NEW') upCount += 1;
      else if (dg.status === 'DOWN') downCount += 1;
      else flatCount += 1;
    });
    return { upCount, downCount, flatCount };
  }, [activeWeek, prevWeek]);

  const formatSignedRp = (diff: number) => {
    if (diff > 0) return `+Rp ${formatNumberId(diff)}`;
    if (diff < 0) return `-Rp ${formatNumberId(Math.abs(diff))}`;
    return 'Rp 0';
  };

  const formatSignedNum = (diff: number, unit: string) => {
    if (diff > 0) return `+${formatNumberId(diff)} ${unit}`;
    if (diff < 0) return `-${formatNumberId(Math.abs(diff))} ${unit}`;
    return `0 ${unit}`;
  };

  const maxHeroBarSales = Math.max(currStats.totalSales, prevStats.totalSales, 1);
  const currBarPct = Math.min(100, Math.round((currStats.totalSales / maxHeroBarSales) * 100));
  const prevBarPct = Math.min(100, Math.round((prevStats.totalSales / maxHeroBarSales) * 100));

  return (
    <div className="bg-white border-2 border-emerald-300 rounded-2xl overflow-hidden shadow-xs">
      {/* Header Bar */}
      <div className="bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 text-white px-4 sm:px-5 py-3.5 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
        <div className="flex items-start sm:items-center gap-3">
          <div
            className={`w-10 h-10 rounded-xl border flex items-center justify-center shrink-0 ${
              salesGrowth.status === 'UP' || salesGrowth.status === 'NEW'
                ? 'bg-emerald-500/20 border-emerald-400/50 text-emerald-300'
                : salesGrowth.status === 'DOWN'
                  ? 'bg-rose-500/20 border-rose-400/50 text-rose-300'
                  : 'bg-slate-500/20 border-slate-400/50 text-slate-300'
            }`}
          >
            {salesGrowth.status === 'UP' || salesGrowth.status === 'NEW' ? (
              <ArrowUp className="w-5 h-5 stroke-[2.75]" />
            ) : salesGrowth.status === 'DOWN' ? (
              <ArrowDown className="w-5 h-5 stroke-[2.75]" />
            ) : (
              <ArrowRight className="w-5 h-5 stroke-[2.75]" />
            )}
          </div>
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2 py-0.5 rounded bg-emerald-500 text-slate-950 text-[10px] font-extrabold uppercase tracking-wider">
                WEEKLY GROWTH (WoW)
              </span>
              <span className="text-xs font-bold text-amber-300 font-mono-num">
                W{currentWeekNum} vs W{effectiveCompareWeekNum}
              </span>
              {/* Quick Visual Legend Pill */}
              <div className="hidden sm:inline-flex items-center gap-2 bg-white/10 border border-white/15 px-2.5 py-0.5 rounded-full text-[10px] font-bold">
                <span className="inline-flex items-center gap-1 text-emerald-300">
                  <ArrowUp className="w-3 h-3 stroke-[3]" /> ▲ Naik
                </span>
                <span className="text-white/30">•</span>
                <span className="inline-flex items-center gap-1 text-rose-300">
                  <ArrowDown className="w-3 h-3 stroke-[3]" /> ▼ Turun
                </span>
                <span className="text-white/30">•</span>
                <span className="inline-flex items-center gap-1 text-slate-300">
                  <ArrowRight className="w-3 h-3 stroke-[3]" /> ▬ Stabil
                </span>
              </div>
            </div>
            <h2 className="text-sm sm:text-base font-extrabold tracking-tight text-white mt-0.5">
              Perbandingan Tren Penjualan Mingguan — Weekly {currentWeekNum} terhadap Weekly{' '}
              {effectiveCompareWeekNum}
            </h2>
            <p className="text-[11px] text-slate-300">
              Indikator panah tren cepat (▲ Naik / ▼ Turun / ▬ Stabil) untuk penjualan, shift pagi/siang, kuantitas, dan transaksi
            </p>
          </div>
        </div>

        {/* Comparison Week Selector */}
        <div className="flex flex-wrap items-center gap-2 text-xs">
          <div className="flex items-center gap-1.5 bg-white/10 border border-white/15 rounded-xl px-2.5 py-1.5">
            <span className="text-slate-300 font-semibold whitespace-nowrap">
              Bandingkan W{currentWeekNum} dengan:
            </span>
            <select
              aria-label="Pilih Minggu Pembanding Weekly Growth"
              value={effectiveCompareWeekNum}
              onChange={(e) => {
                const val = Number(e.target.value);
                setCustomCompareWeekNum(val === defaultPrevWeekNum ? null : val);
              }}
              className="bg-slate-800 text-amber-300 font-bold px-2 py-1 rounded-lg border border-white/20 focus:outline-none text-xs cursor-pointer"
            >
              {ANNUAL_2026_WEEKS.map((w) => {
                const ex = tables.find((t) => t.weekNumber === w.weekNumber);
                const sls = ex ? calculateWeekMetrics(ex).weeklySales : 0;
                const isDefaultPrev = w.weekNumber === defaultPrevWeekNum;
                return (
                  <option key={w.weekNumber} value={w.weekNumber}>
                    W{w.weekNumber} {isDefaultPrev ? '(Minggu Sebelumnya)' : ''} —{' '}
                    {w.seninSlash.slice(0, 5)}–{w.mingguSlash.slice(0, 5)}{' '}
                    {sls > 0 ? `[Rp ${formatNumberId(sls)}]` : '[Rp 0]'}
                  </option>
                );
              })}
            </select>
            {customCompareWeekNum !== null && (
              <button
                type="button"
                onClick={() => setCustomCompareWeekNum(null)}
                title="Kembalikan ke Minggu Sebelumnya (W-1)"
                className="p-1 rounded-lg bg-white/10 hover:bg-white/20 text-amber-300"
              >
                <RefreshCw className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          <button
            type="button"
            onClick={() => setShowDayByDayComparison((v) => !v)}
            className={`px-3 py-1.5 rounded-xl font-bold border transition-colors ${
              showDayByDayComparison
                ? 'bg-amber-400 text-slate-950 border-amber-300'
                : 'bg-white/10 text-slate-200 border-white/15 hover:bg-white/15'
            }`}
          >
            {showDayByDayComparison ? 'Sembunyikan Komparasi H1–H7' : 'Lihat Komparasi H1–H7'}
          </button>
        </div>
      </div>

      <div className="p-4 sm:p-5 space-y-5">
        {/* Main Hero Weekly Growth Banner: Minggu Ini vs Minggu Sebelumnya */}
        <div
          className={`rounded-2xl p-4 sm:p-5 border-2 transition-colors ${
            salesGrowth.status === 'UP' || salesGrowth.status === 'NEW'
              ? 'bg-gradient-to-br from-emerald-50/90 via-white to-teal-50/60 border-emerald-300'
              : salesGrowth.status === 'DOWN'
                ? 'bg-gradient-to-br from-rose-50/90 via-white to-orange-50/50 border-rose-300'
                : 'bg-slate-50/80 border-slate-300'
          }`}
        >
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 items-center">
            {/* Left 5 Cols: Large Directional Arrow Medallion + Headline Growth % & Delta Nominal */}
            <div className="lg:col-span-5 space-y-3 border-b lg:border-b-0 lg:border-r border-slate-200/80 pb-4 lg:pb-0 lg:pr-4">
              <div className="flex items-start gap-3">
                {/* Large Directional Trend Arrow Medallion */}
                <div
                  className={`w-14 h-14 rounded-2xl flex flex-col items-center justify-center shrink-0 border-2 shadow-xs ${
                    salesGrowth.status === 'UP' || salesGrowth.status === 'NEW'
                      ? 'bg-emerald-600 text-white border-emerald-700'
                      : salesGrowth.status === 'DOWN'
                        ? 'bg-rose-600 text-white border-rose-700'
                        : 'bg-slate-600 text-white border-slate-700'
                  }`}
                >
                  {salesGrowth.status === 'UP' || salesGrowth.status === 'NEW' ? (
                    <>
                      <ArrowUp className="w-6 h-6 stroke-[3]" />
                      <span className="text-[9px] font-extrabold uppercase tracking-wider leading-none mt-0.5">
                        NAIK
                      </span>
                    </>
                  ) : salesGrowth.status === 'DOWN' ? (
                    <>
                      <ArrowDown className="w-6 h-6 stroke-[3]" />
                      <span className="text-[9px] font-extrabold uppercase tracking-wider leading-none mt-0.5">
                        TURUN
                      </span>
                    </>
                  ) : (
                    <>
                      <ArrowRight className="w-6 h-6 stroke-[3]" />
                      <span className="text-[9px] font-extrabold uppercase tracking-wider leading-none mt-0.5">
                        STABIL
                      </span>
                    </>
                  )}
                </div>

                <div className="space-y-1.5 min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-xs font-extrabold uppercase tracking-wider text-slate-600">
                      Tren Total Penjualan Mingguan
                    </span>
                    {renderGrowthBadge(salesGrowth, 'lg')}
                  </div>

                  <div className="flex items-baseline gap-2.5 flex-wrap">
                    <div className="text-2xl sm:text-3xl font-extrabold font-mono-num text-slate-900">
                      Rp {formatNumberId(currStats.totalSales) || '0'}
                    </div>
                    <div
                      className={`inline-flex items-center gap-1 text-xs sm:text-sm font-extrabold font-mono-num px-2.5 py-0.5 rounded-lg border ${
                        salesGrowth.diff > 0
                          ? 'bg-emerald-100 text-emerald-900 border-emerald-300'
                          : salesGrowth.diff < 0
                            ? 'bg-rose-100 text-rose-900 border-rose-300'
                            : 'bg-slate-200/70 text-slate-700 border-slate-300'
                      }`}
                    >
                      {salesGrowth.diff > 0 ? (
                        <ArrowUp className="w-3.5 h-3.5 stroke-[2.75] text-emerald-700" />
                      ) : salesGrowth.diff < 0 ? (
                        <ArrowDown className="w-3.5 h-3.5 stroke-[2.75] text-rose-700" />
                      ) : (
                        <ArrowRight className="w-3.5 h-3.5 stroke-[2.75] text-slate-600" />
                      )}
                      <span>
                        {formatSignedRp(salesGrowth.diff)} vs W{effectiveCompareWeekNum}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Visual Bar Proportion W(N) vs W(N-1) */}
              <div className="bg-white/90 border border-slate-200/90 rounded-xl p-2.5 space-y-1.5">
                <div className="space-y-1">
                  <div className="flex items-center justify-between text-[11px] font-mono-num">
                    <span className="font-extrabold text-emerald-800">
                      Minggu Ini (W{currentWeekNum})
                    </span>
                    <span className="font-bold text-slate-900">
                      Rp {formatNumberId(currStats.totalSales) || '0'}
                    </span>
                  </div>
                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className={`h-full rounded-full ${
                        salesGrowth.status === 'DOWN' ? 'bg-rose-500' : 'bg-emerald-500'
                      }`}
                      style={{ width: `${currStats.totalSales > 0 ? Math.max(6, currBarPct) : 0}%` }}
                    />
                  </div>
                </div>
                <div className="space-y-1">
                  <div className="flex items-center justify-between text-[11px] font-mono-num">
                    <span className="font-bold text-slate-600">
                      Minggu Pembanding (W{effectiveCompareWeekNum})
                    </span>
                    <span className="font-semibold text-slate-700">
                      Rp {formatNumberId(prevStats.totalSales) || '0'}
                    </span>
                  </div>
                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div
                      className="h-full bg-slate-400 rounded-full"
                      style={{ width: `${prevStats.totalSales > 0 ? Math.max(6, prevBarPct) : 0}%` }}
                    />
                  </div>
                </div>
              </div>

              <p className="text-xs text-slate-600 leading-relaxed">
                {salesGrowth.status === 'UP' && (
                  <>
                    <strong className="text-emerald-700">▲ Tren NAIK:</strong> Penjualan{' '}
                    <strong className="text-slate-900">Weekly {currentWeekNum}</strong> mengalami{' '}
                    <strong className="text-emerald-700">
                      kenaikan +{salesGrowth.pct?.toFixed(1)}%
                    </strong>{' '}
                    (<strong className="text-emerald-700">{formatSignedRp(salesGrowth.diff)}</strong>
                    ) dibandingkan{' '}
                    <strong className="text-slate-900">Weekly {effectiveCompareWeekNum}</strong> (Rp{' '}
                    {formatNumberId(prevStats.totalSales) || '0'}).
                  </>
                )}
                {salesGrowth.status === 'DOWN' && (
                  <>
                    <strong className="text-rose-700">▼ Tren TURUN:</strong> Penjualan{' '}
                    <strong className="text-slate-900">Weekly {currentWeekNum}</strong> mengalami{' '}
                    <strong className="text-rose-700">
                      penurunan {salesGrowth.pct?.toFixed(1)}%
                    </strong>{' '}
                    (<strong className="text-rose-700">{formatSignedRp(salesGrowth.diff)}</strong>)
                    dibandingkan{' '}
                    <strong className="text-slate-900">Weekly {effectiveCompareWeekNum}</strong> (Rp{' '}
                    {formatNumberId(prevStats.totalSales) || '0'}).
                  </>
                )}
                {salesGrowth.status === 'NEW' && (
                  <>
                    <strong className="text-emerald-700">▲ Tren NAIK (Baru):</strong>{' '}
                    <strong className="text-slate-900">Weekly {currentWeekNum}</strong> mencatat
                    penjualan{' '}
                    <strong className="text-emerald-700">
                      Rp {formatNumberId(currStats.totalSales)}
                    </strong>{' '}
                    (Minggu pembanding{' '}
                    <strong className="text-slate-900">Weekly {effectiveCompareWeekNum}</strong>{' '}
                    masih Rp 0).
                  </>
                )}
                {salesGrowth.status === 'FLAT' && (
                  <>
                    <strong className="text-slate-700">▬ Tren STABIL:</strong> Penjualan{' '}
                    <strong className="text-slate-900">Weekly {currentWeekNum}</strong> stabil /
                    belum berubah dibandingkan{' '}
                    <strong className="text-slate-900">Weekly {effectiveCompareWeekNum}</strong>.
                  </>
                )}
              </p>
            </div>

            {/* Right 7 Cols: Side-by-Side Minggu Ini vs Minggu Sebelumnya */}
            <div className="lg:col-span-7 grid grid-cols-1 sm:grid-cols-2 gap-3">
              {/* Card Minggu Ini */}
              <div className="bg-white rounded-xl p-3.5 border-2 border-emerald-400 shadow-2xs">
                <div className="flex items-center justify-between gap-2">
                  <span className="px-2 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-extrabold">
                    MINGGU INI · W{currentWeekNum}
                  </span>
                  <span className="text-[11px] font-mono-num text-slate-500 font-semibold">
                    {currStats.activeDays}/7 Hari Aktif
                  </span>
                </div>
                <div className="text-[11px] font-semibold text-slate-600 mt-1 truncate">
                  {getWeeklyPeriodeText(activeWeek)}
                </div>
                <div className="mt-1 flex items-center justify-between gap-2">
                  <div className="text-lg font-extrabold font-mono-num text-slate-900">
                    Rp {formatNumberId(currStats.totalSales) || '0'}
                  </div>
                  {renderGrowthBadge(salesGrowth, 'sm', undefined, true)}
                </div>
                <div className="mt-2 pt-2 border-t border-slate-100 grid grid-cols-2 gap-1.5 text-[11px] font-mono-num">
                  <div className="text-amber-900 flex items-center justify-between pr-1">
                    <span>
                      <span className="text-slate-500">☀ Pagi:</span>{' '}
                      <span className="font-bold">Rp {formatNumberId(currStats.pagiSales) || '0'}</span>
                    </span>
                  </div>
                  <div className="text-indigo-900 flex items-center justify-between">
                    <span>
                      <span className="text-slate-500">🌙 Siang:</span>{' '}
                      <span className="font-bold">Rp {formatNumberId(currStats.siangSales) || '0'}</span>
                    </span>
                  </div>
                  <div className="text-slate-700">
                    <span className="text-slate-500">Qty:</span>{' '}
                    <span className="font-bold">{formatNumberId(currStats.totalQty) || '0'} pcs</span>
                  </div>
                  <div className="text-slate-700">
                    <span className="text-slate-500">Trx:</span>{' '}
                    <span className="font-bold">{formatNumberId(currStats.totalTrx) || '0'} trx</span>
                  </div>
                </div>
              </div>

              {/* Card Minggu Sebelumnya */}
              <div className="bg-white rounded-xl p-3.5 border border-slate-300">
                <div className="flex items-center justify-between gap-2">
                  <span className="px-2 py-0.5 rounded bg-slate-700 text-white text-[10px] font-extrabold">
                    PEMBANDING · W{effectiveCompareWeekNum}
                  </span>
                  <span className="text-[11px] font-mono-num text-slate-500 font-semibold">
                    {prevStats.activeDays}/7 Hari Aktif
                  </span>
                </div>
                <div className="text-[11px] font-semibold text-slate-600 mt-1 truncate">
                  {getWeeklyPeriodeText(prevWeek)}
                </div>
                <div className="text-lg font-extrabold font-mono-num text-slate-700 mt-1">
                  Rp {formatNumberId(prevStats.totalSales) || '0'}
                </div>
                <div className="mt-2 pt-2 border-t border-slate-100 grid grid-cols-2 gap-1.5 text-[11px] font-mono-num">
                  <div className="text-amber-900">
                    <span className="text-slate-500">☀ Pagi:</span>{' '}
                    <span className="font-bold">Rp {formatNumberId(prevStats.pagiSales) || '0'}</span>
                  </div>
                  <div className="text-indigo-900">
                    <span className="text-slate-500">🌙 Siang:</span>{' '}
                    <span className="font-bold">Rp {formatNumberId(prevStats.siangSales) || '0'}</span>
                  </div>
                  <div className="text-slate-700">
                    <span className="text-slate-500">Qty:</span>{' '}
                    <span className="font-bold">{formatNumberId(prevStats.totalQty) || '0'} pcs</span>
                  </div>
                  <div className="text-slate-700">
                    <span className="text-slate-500">Trx:</span>{' '}
                    <span className="font-bold">{formatNumberId(prevStats.totalTrx) || '0'} trx</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* 6 Parameter Breakdown Cards (Pagi, Siang, Weekday, Weekend, Qty & Trx, Price Point & Big Size) */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-6 gap-3">
          {/* 1. Shift Pagi Growth */}
          <div className="p-3.5 rounded-xl bg-amber-50/50 border border-amber-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-amber-950 flex items-center gap-1">
                <Sun className="w-3.5 h-3.5 text-amber-600" />
                ☀ Shift Pagi
              </span>
              {renderGrowthBadge(pagiGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              Rp {formatNumberId(currStats.pagiSales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num text-slate-600 flex items-center justify-between">
              <span>W{effectiveCompareWeekNum}: Rp {formatNumberId(prevStats.pagiSales) || '0'}</span>
            </div>
            <div className="text-[11px] font-mono-num pt-0.5">
              {renderDirectionalDelta(pagiGrowth.diff, formatSignedRp(pagiGrowth.diff))}
            </div>
          </div>

          {/* 2. Shift Siang Growth */}
          <div className="p-3.5 rounded-xl bg-indigo-50/50 border border-indigo-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-indigo-950 flex items-center gap-1">
                <Sunset className="w-3.5 h-3.5 text-indigo-600" />
                🌙 Shift Siang
              </span>
              {renderGrowthBadge(siangGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              Rp {formatNumberId(currStats.siangSales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num text-slate-600 flex items-center justify-between">
              <span>W{effectiveCompareWeekNum}: Rp {formatNumberId(prevStats.siangSales) || '0'}</span>
            </div>
            <div className="text-[11px] font-mono-num pt-0.5">
              {renderDirectionalDelta(siangGrowth.diff, formatSignedRp(siangGrowth.diff))}
            </div>
          </div>

          {/* 3. Weekday (Sen-Jum) Growth */}
          <div className="p-3.5 rounded-xl bg-sky-50/50 border border-sky-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-sky-950">Weekday (Sen–Jum)</span>
              {renderGrowthBadge(weekdayGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              Rp {formatNumberId(currStats.weekdaySales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num text-slate-600">
              W{effectiveCompareWeekNum}: Rp {formatNumberId(prevStats.weekdaySales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num pt-0.5">
              {renderDirectionalDelta(weekdayGrowth.diff, formatSignedRp(weekdayGrowth.diff))}
            </div>
          </div>

          {/* 4. Weekend (Sab-Min) Growth */}
          <div className="p-3.5 rounded-xl bg-rose-50/50 border border-rose-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-rose-950">Weekend (Sab–Min)</span>
              {renderGrowthBadge(weekendGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              Rp {formatNumberId(currStats.weekendSales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num text-slate-600">
              W{effectiveCompareWeekNum}: Rp {formatNumberId(prevStats.weekendSales) || '0'}
            </div>
            <div className="text-[11px] font-mono-num pt-0.5">
              {renderDirectionalDelta(weekendGrowth.diff, formatSignedRp(weekendGrowth.diff))}
            </div>
          </div>

          {/* 5. Total Qty & Trx Growth */}
          <div className="p-3.5 rounded-xl bg-emerald-50/50 border border-emerald-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-emerald-950 flex items-center gap-1">
                <ShoppingBag className="w-3.5 h-3.5 text-emerald-700" />
                Qty &amp; Trx
              </span>
              {renderGrowthBadge(qtyGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              {formatNumberId(currStats.totalQty) || '0'} pcs · {formatNumberId(currStats.totalTrx) || '0'} trx
            </div>
            <div className="text-[11px] font-mono-num text-slate-600">
              W{effectiveCompareWeekNum}: {formatNumberId(prevStats.totalQty) || '0'} pcs ·{' '}
              {formatNumberId(prevStats.totalTrx) || '0'} trx
            </div>
            <div className="text-[11px] font-mono-num pt-0.5 flex flex-wrap items-center gap-2">
              {renderDirectionalDelta(qtyGrowth.diff, formatSignedNum(qtyGrowth.diff, 'pcs'))}
            </div>
          </div>

          {/* 6. Price Point & Big Size Growth */}
          <div className="p-3.5 rounded-xl bg-violet-50/50 border border-violet-200 space-y-1.5">
            <div className="flex items-center justify-between gap-1">
              <span className="text-xs font-bold text-violet-950 flex items-center gap-1">
                <Tag className="w-3.5 h-3.5 text-violet-700" />
                Big Size (ATV)
              </span>
              {renderGrowthBadge(bsGrowth, 'sm', undefined, true)}
            </div>
            <div className="text-sm font-extrabold font-mono-num text-slate-900">
              Rp {formatNumberId(currStats.bigSize) || '0'}/trx
            </div>
            <div className="text-[11px] font-mono-num text-slate-600">
              Price Point: Rp {formatNumberId(currStats.pricePoint) || '0'}/pcs
            </div>
            <div className="text-[11px] font-mono-num pt-0.5">
              {renderDirectionalDelta(bsGrowth.diff, `ATV ${formatSignedRp(bsGrowth.diff)}`)}
            </div>
          </div>
        </div>

        {/* Day-by-Day (H1 Senin s/d H7 Minggu) Comparison Grid */}
        {showDayByDayComparison && (
          <div className="space-y-2.5 pt-1">
            <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-50 border border-slate-200 rounded-xl px-3.5 py-2">
              <h3 className="text-xs sm:text-sm font-extrabold text-slate-900 flex items-center gap-1.5">
                <CalendarRange className="w-4 h-4 text-emerald-600" />
                <span>
                  Perbandingan Harian Hari-ke-Hari (H1 Senin s/d H7 Minggu): W{currentWeekNum} vs W
                  {effectiveCompareWeekNum}
                </span>
              </h3>
              {/* Quick Summary of Daily Arrows */}
              <div className="flex flex-wrap items-center gap-2 text-[11px] font-bold">
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-100 text-emerald-800 border border-emerald-300">
                  <ArrowUp className="w-3 h-3 stroke-[3]" /> ▲ {dailyTrendCounts.upCount} Hari Naik
                </span>
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-rose-100 text-rose-800 border border-rose-300">
                  <ArrowDown className="w-3 h-3 stroke-[3]" /> ▼ {dailyTrendCounts.downCount} Hari Turun
                </span>
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-slate-200/80 text-slate-700 border border-slate-300">
                  <ArrowRight className="w-3 h-3 stroke-[3]" /> ▬ {dailyTrendCounts.flatCount} Hari Stabil
                </span>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-7 gap-2.5">
              {activeWeek.days.map((currDay, idx) => {
                const prevDay = prevWeek.days[idx];
                const currDm = calculateDayMetrics(currDay);
                const prevDm = prevDay
                  ? calculateDayMetrics(prevDay)
                  : { totalSales: 0, totalQty: 0, totalTrx: 0, pricePoint: null, bigSize: null };
                const dayGrowth = computeGrowthDelta(currDm.totalSales, prevDm.totalSales);
                const wknd = isWeekendDate(currDay.tanggal);
                const isUp = dayGrowth.status === 'UP' || dayGrowth.status === 'NEW';
                const isDown = dayGrowth.status === 'DOWN';

                return (
                  <div
                    key={currDay.id}
                    className={`p-3 rounded-xl border border-l-4 space-y-1.5 transition-colors ${
                      isUp
                        ? 'border-l-emerald-500 bg-emerald-50/20 border-emerald-200'
                        : isDown
                          ? 'border-l-rose-500 bg-rose-50/25 border-rose-200'
                          : wknd
                            ? 'border-l-slate-400 bg-rose-50/20 border-rose-200'
                            : 'border-l-slate-400 bg-slate-50/70 border-slate-200'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-1">
                      <span
                        className={`text-xs font-extrabold ${
                          wknd ? 'text-rose-900' : 'text-slate-900'
                        }`}
                      >
                        H{idx + 1} · {getIndoDayName(currDay.tanggal)}
                      </span>
                      {renderGrowthBadge(dayGrowth, 'sm', undefined, false)}
                    </div>

                    {/* Explicit Status Strip */}
                    <div className="flex items-center justify-between text-[10px] font-extrabold uppercase tracking-wider">
                      {isUp ? (
                        <span className="inline-flex items-center gap-0.5 text-emerald-700">
                          <ArrowUp className="w-3 h-3 stroke-[3]" /> ▲ NAIK
                        </span>
                      ) : isDown ? (
                        <span className="inline-flex items-center gap-0.5 text-rose-700">
                          <ArrowDown className="w-3 h-3 stroke-[3]" /> ▼ TURUN
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-0.5 text-slate-500">
                          <ArrowRight className="w-3 h-3 stroke-[3]" /> ▬ STABIL
                        </span>
                      )}
                      <span className="font-mono-num text-slate-500">
                        {currDm.totalQty} pcs · {currDm.totalTrx} trx
                      </span>
                    </div>

                    <div className="space-y-1 pt-0.5">
                      <div className="p-1.5 rounded-lg bg-white border border-slate-200/90">
                        <div className="flex items-center justify-between text-[10px] text-slate-500">
                          <span className="font-bold text-emerald-800">W{currentWeekNum}</span>
                          <span className="font-mono-num">{currDay.tanggal.slice(0, 5)}</span>
                        </div>
                        <div className="text-xs font-extrabold font-mono-num text-slate-900 mt-0.5">
                          Rp {formatNumberId(currDm.totalSales) || '0'}
                        </div>
                      </div>

                      <div className="p-1.5 rounded-lg bg-slate-100/80 border border-slate-200/70">
                        <div className="flex items-center justify-between text-[10px] text-slate-500">
                          <span className="font-bold text-slate-700">
                            W{effectiveCompareWeekNum}
                          </span>
                          <span className="font-mono-num">
                            {prevDay ? prevDay.tanggal.slice(0, 5) : '-'}
                          </span>
                        </div>
                        <div className="text-xs font-bold font-mono-num text-slate-600 mt-0.5">
                          Rp {formatNumberId(prevDm.totalSales) || '0'}
                        </div>
                      </div>
                    </div>

                    <div className="text-[10px] font-mono-num text-right pt-0.5">
                      {renderDirectionalDelta(dayGrowth.diff, formatSignedRp(dayGrowth.diff))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Multi-Week Sequential Growth Trend Table */}
        <div className="border border-slate-200 rounded-xl overflow-hidden">
          <div className="px-4 py-2.5 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <BarChart3 className="w-4 h-4 text-emerald-400" />
              <h3 className="text-xs sm:text-sm font-bold">
                Tabel Tren Pertumbuhan Antar Minggu (Multi-Week Growth Matrix)
              </h3>
            </div>
            <span className="text-[11px] text-slate-300">
              Klik baris Weekly untuk memilih minggu fokus
            </span>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs border-collapse min-w-[880px]">
              <thead>
                <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                  <th className="py-2.5 px-3 text-left font-bold">Weekly</th>
                  <th className="py-2.5 px-3 text-left font-bold">Periode (Senin – Minggu)</th>
                  <th className="py-2.5 px-3 text-right font-bold">Total Sales (Rp)</th>
                  <th className="py-2.5 px-3 text-right font-bold">Minggu Sebelumnya (W-1)</th>
                  <th className="py-2.5 px-3 text-center font-bold">Indikator Arah Tren</th>
                  <th className="py-2.5 px-3 text-right font-bold">Selisih (Δ Rp)</th>
                  <th className="py-2.5 px-3 text-center font-bold">% Weekly Growth</th>
                  <th className="py-2.5 px-3 text-right font-bold">Qty &amp; Growth</th>
                  <th className="py-2.5 px-3 text-right font-bold">Trx &amp; Growth</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {multiWeekTrendList.map((item) => {
                  const isCurrent = item.weekNumber === currentWeekNum;
                  const isUp = item.growth.status === 'UP' || item.growth.status === 'NEW';
                  const isDown = item.growth.status === 'DOWN';
                  return (
                    <tr
                      key={item.weekNumber}
                      onClick={() =>
                        onSelectWeekNumber && onSelectWeekNumber(item.weekNumber)
                      }
                      className={`cursor-pointer transition-colors ${
                        isCurrent
                          ? 'bg-emerald-50/70 font-semibold'
                          : 'hover:bg-slate-50'
                      }`}
                    >
                      <td className="py-2.5 px-3 font-bold">
                        <span
                          className={`inline-flex items-center px-2 py-0.5 rounded text-[11px] font-extrabold ${
                            isCurrent
                              ? 'bg-emerald-600 text-white'
                              : item.stats.totalSales > 0
                                ? 'bg-slate-800 text-white'
                                : 'bg-slate-200 text-slate-700'
                          }`}
                        >
                          WEEKLY {item.weekNumber}
                        </span>
                      </td>
                      <td className="py-2.5 px-3 text-slate-700">
                        {getWeeklyPeriodeText(item.table)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono-num font-extrabold text-slate-900">
                        Rp {formatNumberId(item.stats.totalSales) || '0'}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono-num text-slate-600">
                        {item.prevWeekNumber ? (
                          <span>
                            W{item.prevWeekNumber}: Rp {formatNumberId(item.prevSales) || '0'}
                          </span>
                        ) : (
                          '-'
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-center">
                        {isUp ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 border border-emerald-300 text-[10px] font-extrabold uppercase">
                            <ArrowUp className="w-3 h-3 stroke-[3]" /> ▲ NAIK
                          </span>
                        ) : isDown ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-rose-100 text-rose-800 border border-rose-300 text-[10px] font-extrabold uppercase">
                            <ArrowDown className="w-3 h-3 stroke-[3]" /> ▼ TURUN
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-slate-100 text-slate-700 border border-slate-300 text-[10px] font-extrabold uppercase">
                            <ArrowRight className="w-3 h-3 stroke-[3]" /> ▬ STABIL
                          </span>
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono-num">
                        {renderDirectionalDelta(item.growth.diff, formatSignedRp(item.growth.diff))}
                      </td>
                      <td className="py-2.5 px-3 text-center">
                        {renderGrowthBadge(item.growth, 'sm', undefined, false)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono-num">
                        <div className="font-bold text-slate-800">
                          {formatNumberId(item.stats.totalQty) || '0'} pcs
                        </div>
                        {item.qtyGrowth && (
                          <div
                            className={`text-[10px] font-bold inline-flex items-center justify-end gap-0.5 ${
                              item.qtyGrowth.diff > 0
                                ? 'text-emerald-700'
                                : item.qtyGrowth.diff < 0
                                  ? 'text-rose-700'
                                  : 'text-slate-400'
                            }`}
                          >
                            {item.qtyGrowth.status === 'UP'
                              ? `▲ +${item.qtyGrowth.pct?.toFixed(1)}%`
                              : item.qtyGrowth.status === 'DOWN'
                                ? `▼ ${item.qtyGrowth.pct?.toFixed(1)}%`
                                : item.qtyGrowth.status === 'NEW'
                                  ? '▲ +100% (Baru)'
                                  : '▬ 0%'}
                          </div>
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono-num">
                        <div className="font-bold text-slate-800">
                          {formatNumberId(item.stats.totalTrx) || '0'} trx
                        </div>
                        {item.trxGrowth && (
                          <div
                            className={`text-[10px] font-bold inline-flex items-center justify-end gap-0.5 ${
                              item.trxGrowth.diff > 0
                                ? 'text-emerald-700'
                                : item.trxGrowth.diff < 0
                                  ? 'text-rose-700'
                                  : 'text-slate-400'
                            }`}
                          >
                            {item.trxGrowth.status === 'UP'
                              ? `▲ +${item.trxGrowth.pct?.toFixed(1)}%`
                              : item.trxGrowth.status === 'DOWN'
                                ? `▼ ${item.trxGrowth.pct?.toFixed(1)}%`
                                : item.trxGrowth.status === 'NEW'
                                  ? '▲ +100% (Baru)'
                                  : '▬ 0%'}
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
