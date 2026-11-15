import React, { useMemo } from 'react';
import {
  WeeklySalesTable,
  formatNumberId,
  getIndoDayName,
  isWeekendDate,
} from '../data/initialData';

interface MonthlyAutoSummaryCardProps {
  tables: WeeklySalesTable[];
  selectedMonthKey: string; // '01' .. '12'
  onSelectMonthKey?: (monthKey: string) => void;
  dayTypeFilter?: 'ALL' | 'WEEKDAY' | 'WEEKEND';
}

const MONTH_META = [
  { key: '01', name: 'Januari', short: 'Jan', daysCount: 31 },
  { key: '02', name: 'Februari', short: 'Feb', daysCount: 28 },
  { key: '03', name: 'Maret', short: 'Mar', daysCount: 31 },
  { key: '04', name: 'April', short: 'Apr', daysCount: 30 },
  { key: '05', name: 'Mei', short: 'Mei', daysCount: 31 },
  { key: '06', name: 'Juni', short: 'Jun', daysCount: 30 },
  { key: '07', name: 'Juli', short: 'Jul', daysCount: 31 },
  { key: '08', name: 'Agustus', short: 'Agu', daysCount: 31 },
  { key: '09', name: 'September', short: 'Sep', daysCount: 30 },
  { key: '10', name: 'Oktober', short: 'Okt', daysCount: 31 },
  { key: '11', name: 'November', short: 'Nov', daysCount: 30 },
  { key: '12', name: 'Desember', short: 'Des', daysCount: 31 },
] as const;

export const MonthlyAutoSummaryCard: React.FC<MonthlyAutoSummaryCardProps> = ({
  tables,
  selectedMonthKey,
  onSelectMonthKey,
  dayTypeFilter = 'ALL',
}) => {
  const monthMeta =
    MONTH_META.find((m) => m.key === selectedMonthKey) || MONTH_META[9];

  const stats = useMemo(() => {
    const dayMap = new Map<
      string,
      {
        pagiSales: number;
        siangSales: number;
        pagiQty: number;
        siangQty: number;
        pagiTrx: number;
        siangTrx: number;
      }
    >();

    tables.forEach((t) => {
      t.days.forEach((d) => {
        const dt = (d.tanggal || '').trim();
        if (!dt) return;
        const ps = d.pagi.sales ?? 0;
        const ss = d.siang.sales ?? 0;
        const pq = d.pagi.qty ?? 0;
        const sq = d.siang.qty ?? 0;
        const pt = d.pagi.trx ?? 0;
        const st = d.siang.trx ?? 0;
        const prev = dayMap.get(dt);
        if (!prev || ps + ss >= prev.pagiSales + prev.siangSales) {
          dayMap.set(dt, {
            pagiSales: ps,
            siangSales: ss,
            pagiQty: pq,
            siangQty: sq,
            pagiTrx: pt,
            siangTrx: st,
          });
        }
      });
    });

    let cumulativeSales = 0;
    let pagiSales = 0;
    let siangSales = 0;
    let weekdaySales = 0;
    let weekendSales = 0;

    let totalTrx = 0;
    let pagiTrx = 0;
    let siangTrx = 0;

    let totalQty = 0;
    let pagiQty = 0;
    let siangQty = 0;

    let activeDays = 0;
    let filteredCalendarDays = 0;

    let decade1Sales = 0; // Tgl 1-10
    let decade2Sales = 0; // Tgl 11-20
    let decade3Sales = 0; // Tgl 21-30/31

    let peakDay: {
      dayNum: number;
      slashDate: string;
      dayName: string;
      sales: number;
      trx: number;
    } | null = null;

    for (let d = 1; d <= monthMeta.daysCount; d++) {
      const dd = String(d).padStart(2, '0');
      const slashDate = `${dd}/${monthMeta.key}/2026`;
      const wknd = isWeekendDate(slashDate);

      if (dayTypeFilter === 'WEEKDAY' && wknd) continue;
      if (dayTypeFilter === 'WEEKEND' && !wknd) continue;

      filteredCalendarDays += 1;

      const rec = dayMap.get(slashDate);
      const ps = rec?.pagiSales ?? 0;
      const ss = rec?.siangSales ?? 0;
      const daySales = ps + ss;

      const pt = rec?.pagiTrx ?? 0;
      const st = rec?.siangTrx ?? 0;
      const dayTrx = pt + st;

      const pq = rec?.pagiQty ?? 0;
      const sq = rec?.siangQty ?? 0;
      const dayQty = pq + sq;

      cumulativeSales += daySales;
      pagiSales += ps;
      siangSales += ss;
      if (wknd) weekendSales += daySales;
      else weekdaySales += daySales;

      totalTrx += dayTrx;
      pagiTrx += pt;
      siangTrx += st;

      totalQty += dayQty;
      pagiQty += pq;
      siangQty += sq;

      if (daySales > 0 || dayTrx > 0 || dayQty > 0) {
        activeDays += 1;
      }

      if (d <= 10) decade1Sales += daySales;
      else if (d <= 20) decade2Sales += daySales;
      else decade3Sales += daySales;

      if (daySales > 0 && (!peakDay || daySales > peakDay.sales)) {
        peakDay = {
          dayNum: d,
          slashDate,
          dayName: getIndoDayName(slashDate),
          sales: daySales,
          trx: dayTrx,
        };
      }
    }

    const effectiveCalendarDays =
      filteredCalendarDays > 0 ? filteredCalendarDays : monthMeta.daysCount;

    // Rata-rata Tiket (Transaksi) per Hari
    const avgTicketsPerActiveDay =
      activeDays > 0 ? totalTrx / activeDays : 0;
    const avgTicketsPerCalendarDay =
      effectiveCalendarDays > 0 ? totalTrx / effectiveCalendarDays : 0;

    // Rata-rata Sales per Hari
    const avgSalesPerActiveDay =
      activeDays > 0 ? Math.round(cumulativeSales / activeDays) : 0;
    const avgSalesPerCalendarDay =
      effectiveCalendarDays > 0
        ? Math.round(cumulativeSales / effectiveCalendarDays)
        : 0;

    // Rata-rata Nilai Tiket (Basket Size = Sales / Trx) & APC (Sales / Qty)
    const avgTicketValue =
      totalTrx > 0 ? Math.round(cumulativeSales / totalTrx) : 0;
    const avgItemPrice =
      totalQty > 0 ? Math.round(cumulativeSales / totalQty) : 0;

    const pagiSharePct =
      cumulativeSales > 0 ? Math.round((pagiSales / cumulativeSales) * 100) : 0;
    const siangSharePct =
      cumulativeSales > 0 ? Math.round((siangSales / cumulativeSales) * 100) : 0;

    return {
      cumulativeSales,
      pagiSales,
      siangSales,
      weekdaySales,
      weekendSales,
      totalTrx,
      pagiTrx,
      siangTrx,
      totalQty,
      pagiQty,
      siangQty,
      activeDays,
      effectiveCalendarDays,
      avgTicketsPerActiveDay,
      avgTicketsPerCalendarDay,
      avgSalesPerActiveDay,
      avgSalesPerCalendarDay,
      avgTicketValue,
      avgItemPrice,
      pagiSharePct,
      siangSharePct,
      decade1Sales,
      decade2Sales,
      decade3Sales,
      peakDay,
    };
  }, [tables, monthMeta, dayTypeFilter]);

  const formatDecimalId = (val: number) => {
    if (!val || val <= 0) return '0';
    return val.toLocaleString('id-ID', {
      minimumFractionDigits: val % 1 === 0 ? 0 : 1,
      maximumFractionDigits: 1,
    });
  };

  return (
    <div className="bg-slate-50/70 border-b border-slate-200 px-3.5 sm:px-5 py-3.5">
      {/* Section Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 pb-2.5 mb-3 border-b border-slate-200/80">
        <div>
          <h3 className="text-xs sm:text-sm font-bold text-slate-900 tracking-tight">
            Ringkasan Kumulatif Bulanan Otomatis — {monthMeta.name} 2026 (Tanggal 1–
            {monthMeta.daysCount} {monthMeta.short})
          </h3>
          <p className="text-[11px] text-slate-600 mt-0.5">
            Kalkulasi otomatis real-time dari seluruh transaksi harian pada bulan{' '}
            <span className="font-semibold text-slate-800">{monthMeta.name} 2026</span>
            {' · '}
            <span className="font-mono-num font-semibold text-emerald-700">
              {stats.activeDays} dari {stats.effectiveCalendarDays} hari terisi
            </span>
          </p>
        </div>

        {onSelectMonthKey && (
          <div className="no-print flex items-center gap-1.5">
            <label
              htmlFor="auto-summary-month-select"
              className="text-[11px] font-semibold text-slate-600"
            >
              Bulan:
            </label>
            <select
              id="auto-summary-month-select"
              value={selectedMonthKey}
              onChange={(e) => onSelectMonthKey(e.target.value)}
              className="px-2.5 py-1 rounded-md bg-white border border-slate-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-slate-900 cursor-pointer"
            >
              {MONTH_META.map((m) => (
                <option key={m.key} value={m.key}>
                  {m.name} 2026 (1–{m.daysCount} {m.short})
                </option>
              ))}
            </select>
          </div>
        )}
      </div>

      {/* Primary 4-Column Metric Grid (Divided by clean hairlines) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 bg-white border border-slate-200 rounded-lg divide-y sm:divide-y-0 sm:divide-x divide-slate-200">
        {/* 1. Total Kumulatif Sales */}
        <div className="p-3.5 flex flex-col justify-between">
          <div>
            <div className="text-[11px] font-semibold text-slate-500">
              1. Total Kumulatif Sales ({monthMeta.short} 2026)
            </div>
            <div className="text-lg sm:text-xl font-extrabold text-slate-950 font-mono-num tracking-tight mt-1">
              Rp {formatNumberId(stats.cumulativeSales) || '0'}
            </div>
          </div>
          <div className="mt-2.5 pt-2 border-t border-slate-100 space-y-1 text-[11px] text-slate-600 font-mono-num">
            <div className="flex items-center justify-between">
              <span className="text-amber-800 font-medium">Shift Pagi ({stats.pagiSharePct}%):</span>
              <span className="font-bold text-slate-900">
                Rp {formatNumberId(stats.pagiSales) || '0'}
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span className="text-indigo-800 font-medium">Shift Siang ({stats.siangSharePct}%):</span>
              <span className="font-bold text-slate-900">
                Rp {formatNumberId(stats.siangSales) || '0'}
              </span>
            </div>
          </div>
        </div>

        {/* 2. Total Transaksi (Struk / Tiket) */}
        <div className="p-3.5 flex flex-col justify-between">
          <div>
            <div className="text-[11px] font-semibold text-slate-500">
              2. Total Transaksi / Tiket ({monthMeta.short} 2026)
            </div>
            <div className="flex items-baseline gap-1.5 mt-1">
              <span className="text-lg sm:text-xl font-extrabold text-indigo-950 font-mono-num tracking-tight">
                {formatNumberId(stats.totalTrx) || '0'}
              </span>
              <span className="text-xs font-semibold text-slate-600">
                trx
              </span>
            </div>
          </div>
          <div className="mt-2.5 pt-2 border-t border-slate-100 space-y-1 text-[11px] text-slate-600 font-mono-num">
            <div className="flex items-center justify-between">
              <span>Trx Pagi / Siang:</span>
              <span className="font-bold text-slate-900">
                {formatNumberId(stats.pagiTrx) || '0'} / {formatNumberId(stats.siangTrx) || '0'} trx
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span>Total Kuantitas (Qty):</span>
              <span className="font-bold text-slate-900">
                {formatNumberId(stats.totalQty) || '0'} pcs
              </span>
            </div>
          </div>
        </div>

        {/* 3. Rata-rata Tiket per Hari */}
        <div className="p-3.5 flex flex-col justify-between">
          <div>
            <div className="text-[11px] font-semibold text-slate-500">
              3. Rata-rata Tiket per Hari
            </div>
            <div className="flex items-baseline gap-1.5 mt-1">
              <span className="text-lg sm:text-xl font-extrabold text-emerald-900 font-mono-num tracking-tight">
                {formatDecimalId(stats.avgTicketsPerActiveDay)}
              </span>
              <span className="text-xs font-semibold text-slate-600">
                tiket / hari aktif
              </span>
            </div>
          </div>
          <div className="mt-2.5 pt-2 border-t border-slate-100 space-y-1 text-[11px] text-slate-600 font-mono-num">
            <div className="flex items-center justify-between">
              <span>Per Hari Kalender ({stats.effectiveCalendarDays} hr):</span>
              <span className="font-bold text-slate-900">
                {formatDecimalId(stats.avgTicketsPerCalendarDay)} tiket/hr
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span>Nilai Rata-rata / Tiket:</span>
              <span className="font-bold text-emerald-800">
                Rp {formatNumberId(stats.avgTicketValue) || '0'}
              </span>
            </div>
          </div>
        </div>

        {/* 4. Rata-rata Sales Harian & Puncak Bulanan */}
        <div className="p-3.5 flex flex-col justify-between">
          <div>
            <div className="text-[11px] font-semibold text-slate-500">
              4. Rata-rata Sales per Hari Aktif
            </div>
            <div className="text-lg sm:text-xl font-extrabold text-amber-900 font-mono-num tracking-tight mt-1">
              Rp {formatNumberId(stats.avgSalesPerActiveDay) || '0'}
            </div>
          </div>
          <div className="mt-2.5 pt-2 border-t border-slate-100 space-y-1 text-[11px] text-slate-600 font-mono-num">
            <div className="flex items-center justify-between">
              <span>Rata-rata / Kalender ({stats.effectiveCalendarDays} hr):</span>
              <span className="font-bold text-slate-900">
                Rp {formatNumberId(stats.avgSalesPerCalendarDay) || '0'}
              </span>
            </div>
            <div className="flex items-center justify-between">
              <span>Puncak Harian:</span>
              <span className="font-bold text-rose-800 truncate max-w-[155px]">
                {stats.peakDay
                  ? `Tgl ${stats.peakDay.dayNum} (Rp ${formatNumberId(stats.peakDay.sales)})`
                  : 'Belum ada'}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Cumulative Decade Breakdown Strip (Tgl 1-10, 11-20, 21-30/31) */}
      <div className="mt-2.5 pt-2 flex flex-wrap items-center justify-between gap-x-4 gap-y-1 text-[11px] text-slate-600 font-mono-num">
        <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
          <span className="font-semibold text-slate-700">
            Akumulasi Periode Tanggal:
          </span>
          <span>
            Tgl 1–10: <strong className="text-slate-900">Rp {formatNumberId(stats.decade1Sales) || '0'}</strong>
          </span>
          <span aria-hidden="true">·</span>
          <span>
            Tgl 11–20: <strong className="text-slate-900">Rp {formatNumberId(stats.decade2Sales) || '0'}</strong>
          </span>
          <span aria-hidden="true">·</span>
          <span>
            Tgl 21–{monthMeta.daysCount}:{' '}
            <strong className="text-slate-900">Rp {formatNumberId(stats.decade3Sales) || '0'}</strong>
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
          <span>
            APC (Harga Rata-rata/Pcs):{' '}
            <strong className="text-slate-900">Rp {formatNumberId(stats.avgItemPrice) || '0'}</strong>
          </span>
          <span aria-hidden="true">·</span>
          <span>
            Weekday: <strong className="text-indigo-900">Rp {formatNumberId(stats.weekdaySales) || '0'}</strong>
          </span>
          <span aria-hidden="true">·</span>
          <span>
            Weekend: <strong className="text-rose-800">Rp {formatNumberId(stats.weekendSales) || '0'}</strong>
          </span>
        </div>
      </div>
    </div>
  );
};
