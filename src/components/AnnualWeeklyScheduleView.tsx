import React, { useState } from 'react';
import {
  WeeklySalesTable,
  AnnualWeekScheduleItem,
  generateAnnualWeeksMondayToSunday,
  createWeeklyTableFromAnnualSchedule,
  calculateWeekMetrics,
  formatNumberId,
} from '../data/initialData';
import {
  CalendarRange,
  Calendar,
  Search,
  FileSpreadsheet,
  CheckCircle2,
  PlusCircle,
  Sparkles,
  Download,
  Filter,
  ChevronLeft,
  ChevronRight,
  BarChart3,
  PenSquare,
  Trash2,
  Lock,
} from 'lucide-react';

interface AnnualWeeklyScheduleViewProps {
  tables: WeeklySalesTable[];
  onChangeTables: (updated: WeeklySalesTable[]) => void;
  onOpenWeekly?: (
    weekNumber: number,
    targetMenu?: 'sales_matrix' | 'dashboard' | 'daily_input' | 'reason_evaluation'
  ) => void;
  onSelectWeeklyAndOpenSpreadsheet: (weekId: string) => void;
  onSelectWeeklyAndOpenDailyInput: (weekIndex: number, weekId?: string) => void;
  onOpenWeeklyInMenu?: (
    weekId: string,
    menu: 'sales_matrix' | 'dashboard' | 'daily_input' | 'reason_evaluation'
  ) => void;
  isPublicReadOnly?: boolean;
}

const MONTHS_LIST = [
  { key: '10', label: 'Oktober' },
  { key: '01', label: 'Januari' },
  { key: '02', label: 'Februari' },
  { key: '03', label: 'Maret' },
  { key: '04', label: 'April' },
  { key: '05', label: 'Mei' },
  { key: '06', label: 'Juni' },
  { key: '07', label: 'Juli' },
  { key: '08', label: 'Agustus' },
  { key: '09', label: 'September' },
  { key: '11', label: 'November' },
  { key: '12', label: 'Desember' },
  { key: 'ALL', label: 'Semua Bulan (Terpaginasi)' },
];

const ORDERED_MONTHS = [
  { key: '01', label: 'Januari' },
  { key: '02', label: 'Februari' },
  { key: '03', label: 'Maret' },
  { key: '04', label: 'April' },
  { key: '05', label: 'Mei' },
  { key: '06', label: 'Juni' },
  { key: '07', label: 'Juli' },
  { key: '08', label: 'Agustus' },
  { key: '09', label: 'September' },
  { key: '10', label: 'Oktober' },
  { key: '11', label: 'November' },
  { key: '12', label: 'Desember' },
  { key: 'ALL', label: 'Semua (53 Weekly)' },
];

export const AnnualWeeklyScheduleView: React.FC<AnnualWeeklyScheduleViewProps> = ({
  tables,
  onChangeTables,
  onSelectWeeklyAndOpenSpreadsheet,
  onSelectWeeklyAndOpenDailyInput,
  onOpenWeeklyInMenu,
  isPublicReadOnly = false,
}) => {
  const [selectedYear, setSelectedYear] = useState<number>(2026);
  const [quarterFilter, setQuarterFilter] = useState<'ALL' | 'Q1' | 'Q2' | 'Q3' | 'Q4'>('ALL');
  // Default filter to Oktober ('10') so the table is compact (4 rows) and not long vertically
  const [monthFilter, setMonthFilter] = useState<string>('10');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'FILLED' | 'ACTIVE'>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [quickJumpWeekNum, setQuickJumpWeekNum] = useState<number>(41);
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(6);
  const [statusNotice, setStatusNotice] = useState<string | null>(null);

  const annualWeeks = generateAnnualWeeksMondayToSunday(selectedYear);

  const showNotice = (msg: string) => {
    setStatusNotice(msg);
    setTimeout(() => {
      setStatusNotice((prev) => (prev === msg ? null : prev));
    }, 2500);
  };

  // Map existing tables by weekNumber
  const tableByWeekNumber = new Map<number, WeeklySalesTable>();
  tables.forEach((t) => {
    tableByWeekNumber.set(t.weekNumber, t);
  });

  // Ensure a Weekly table exists and immediately open it in the target view
  const handleOpenOrAddWeek = (
    item: AnnualWeekScheduleItem,
    targetView: 'spreadsheet' | 'dashboard' | 'input' | 'reason'
  ) => {
    let targetTable = tables.find((t) => t.weekNumber === item.weekNumber);
    let updatedList = tables;

    if (!targetTable) {
      targetTable = createWeeklyTableFromAnnualSchedule(item);
      updatedList = [...tables, targetTable].sort((a, b) => a.weekNumber - b.weekNumber);
      onChangeTables(updatedList);
    }

    if (onOpenWeeklyInMenu) {
      const menuMap = {
        spreadsheet: 'sales_matrix',
        dashboard: 'dashboard',
        input: 'daily_input',
        reason: 'reason_evaluation',
      } as const;
      onOpenWeeklyInMenu(targetTable.id, menuMap[targetView]);
      return;
    }

    if (targetView === 'spreadsheet' || targetView === 'dashboard' || targetView === 'reason') {
      onSelectWeeklyAndOpenSpreadsheet(targetTable.id);
    } else {
      const idx = updatedList.findIndex((t) => t.id === targetTable!.id);
      onSelectWeeklyAndOpenDailyInput(Math.max(0, idx), targetTable.id);
    }
  };

  // Synchronize all existing tables' dates to Monday-Sunday of their respective weekNumber
  const handleSyncExistingTablesToMondaySunday = () => {
    const scheduleMap = new Map<number, AnnualWeekScheduleItem>();
    annualWeeks.forEach((w) => scheduleMap.set(w.weekNumber, w));

    const updated = tables.map((t) => {
      const sched = scheduleMap.get(t.weekNumber);
      if (!sched) return t;
      return {
        ...t,
        id: `weekly-${t.weekNumber}`,
        periode: sched.periodeText,
        startDateIso: sched.startDateIso,
        days: sched.slashDates.map((tgl, idx) => {
          const existingDay = t.days[idx];
          if (existingDay) {
            return {
              ...existingDay,
              tanggal: tgl,
            };
          }
          return {
            id: `w${t.weekNumber}-d${idx + 1}`,
            tanggal: tgl,
            pagi: { nama: '', sales: null, qty: null, trx: null },
            siang: { nama: '', sales: null, qty: null, trx: null },
            reasonPagi: '',
            reasonSiang: '',
          };
        }),
      };
    });

    onChangeTables(updated);
    showNotice('Seluruh tabel Weekly aktif telah diselaraskan ke siklus Senin s/d Minggu!');
  };

  // Remove empty tables (except Weekly 41 & 42) so the workspace stays compact
  const handleCleanupEmptyTables = () => {
    const cleaned = tables.filter((t) => {
      const wm = calculateWeekMetrics(t);
      const hasReason = t.days.some(
        (d) => d.reasonPagi.trim() !== '' || d.reasonSiang.trim() !== ''
      );
      return wm.weeklySales > 0 || hasReason || t.weekNumber === 41 || t.weekNumber === 42;
    });
    const finalTables = cleaned.length > 0 ? cleaned : tables.slice(0, 2);
    onChangeTables(finalTables);
    showNotice(
      `Daftar Weekly dirapikan! Menyisakan ${finalTables.length} tabel utama/terisi.`
    );
  };

  // Export Annual 52/53-Week Calendar to Excel (.xls)
  const handleExportAnnualCalendarExcel = () => {
    if (isPublicReadOnly) return;
    let html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="UTF-8">
        <style>
          body, table { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }
          table { border-collapse: collapse; }
          td, th { border: 1px solid #000000; padding: 5px 8px; vertical-align: middle; }
          .bg-header { background-color: #1E293B; color: #FFFFFF; font-weight: bold; text-align: center; }
          .bg-wknd { background-color: #FFE4E6; color: #9F1239; font-weight: bold; text-align: center; }
          .bg-yellow { background-color: #FFC000; color: #000000; font-weight: bold; }
          .text-center { text-align: center; }
          .text-right { text-align: right; }
          .font-bold { font-weight: bold; }
        </style>
      </head>
      <body>
        <table>
          <tr>
            <td colspan="12" style="font-size:14pt; font-weight:bold; background-color:#FEF3C7;">
              DATA KALENDER WEEKLY SETAHUN (${selectedYear}) — HARI SENIN S/D MINGGU (BAZAR MATAHARI TANJUNG)
            </td>
          </tr>
          <tr>
            <th class="bg-header">Weekly Ke-</th>
            <th class="bg-header">Kuartal</th>
            <th class="bg-header">Bulan</th>
            <th class="bg-header">Rentang Periode (Senin – Minggu)</th>
            <th class="bg-header">Senin (H1)</th>
            <th class="bg-header">Selasa (H2)</th>
            <th class="bg-header">Rabu (H3)</th>
            <th class="bg-header">Kamis (H4)</th>
            <th class="bg-header">Jumat (H5)</th>
            <th class="bg-wknd">Sabtu (H6)</th>
            <th class="bg-wknd">Minggu (H7)</th>
            <th class="bg-yellow">Total Sales Weekly</th>
          </tr>
    `;

    for (const w of annualWeeks) {
      const existing = tableByWeekNumber.get(w.weekNumber);
      const sales = existing ? calculateWeekMetrics(existing).weeklySales : 0;
      html += `
        <tr>
          <td class="text-center font-bold">WEEKLY ${w.weekNumber}</td>
          <td class="text-center">${w.quarter}</td>
          <td class="text-center font-bold">${w.monthName}</td>
          <td>${w.periodeText}</td>
          <td class="text-center">${w.slashDates[0]}</td>
          <td class="text-center">${w.slashDates[1]}</td>
          <td class="text-center">${w.slashDates[2]}</td>
          <td class="text-center">${w.slashDates[3]}</td>
          <td class="text-center">${w.slashDates[4]}</td>
          <td class="bg-wknd">${w.slashDates[5]}</td>
          <td class="bg-wknd">${w.slashDates[6]}</td>
          <td class="text-right font-bold">${sales > 0 ? formatNumberId(sales) : '-'}</td>
        </tr>
      `;
    }

    html += `</table></body></html>`;
    const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Kalender_Weekly_Setahun_${selectedYear}_Senin_Minggu.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const filteredWeeks = annualWeeks.filter((w) => {
    if (quarterFilter !== 'ALL' && w.quarter !== quarterFilter) return false;
    if (monthFilter !== 'ALL' && w.monthKey !== monthFilter) return false;

    const existing = tableByWeekNumber.get(w.weekNumber);
    const sales = existing ? calculateWeekMetrics(existing).weeklySales : 0;
    if (statusFilter === 'FILLED' && sales <= 0) return false;
    if (statusFilter === 'ACTIVE' && !existing) return false;

    if (searchQuery.trim() !== '') {
      const q = searchQuery.trim().toLowerCase();
      const matchWeek =
        `weekly ${w.weekNumber}`.includes(q) ||
        `w${w.weekNumber}`.includes(q) ||
        String(w.weekNumber) === q;
      const matchMonth = w.monthName.toLowerCase().includes(q);
      const matchPeriode = w.periodeText.toLowerCase().includes(q);
      const matchDate = w.slashDates.some((d) => d.includes(q));
      return matchWeek || matchMonth || matchPeriode || matchDate;
    }
    return true;
  });

  // Pagination so the table is never overly long vertically
  const totalPages = Math.max(1, Math.ceil(filteredWeeks.length / pageSize));
  const safePage = Math.min(currentPage, totalPages);
  const startIdx = (safePage - 1) * pageSize;
  const paginatedWeeks = filteredWeeks.slice(startIdx, startIdx + pageSize);

  // Calculate Annual Stats
  let activeWeeksCount = 0;
  let filledWeeksCount = 0;
  let totalAnnualSales = 0;
  tables.forEach((t) => {
    activeWeeksCount += 1;
    const wm = calculateWeekMetrics(t);
    if (wm.weeklySales > 0) {
      filledWeeksCount += 1;
      totalAnnualSales += wm.weeklySales;
    }
  });

  const selectedQuickWeekObj =
    annualWeeks.find((w) => w.weekNumber === quickJumpWeekNum) || annualWeeks[40] || annualWeeks[0];

  return (
    <div className="space-y-5">
      {/* Top Hero Banner + Direct "Buka Weekly" Quick Launcher */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 shadow-xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          <div>
            <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded-md bg-amber-400/20 text-amber-300 text-xs font-bold mb-1.5">
              <CalendarRange className="w-3.5 h-3.5" />
              <span>MASTER KALENDER WEEKLY SETAHUN (SENIN S/D MINGGU)</span>
            </div>
            <h1 className="text-lg sm:text-xl font-bold tracking-tight">
              Data Weekly Ke-1 s/d Ke-{annualWeeks.length} Tahun {selectedYear} (Senin – Minggu)
            </h1>
            <p className="text-xs text-slate-300 mt-0.5">
              Difilter per bulan &amp; halaman agar ringkas dan tidak panjang ke bawah. Pilih Weekly mana saja untuk langsung dibuka.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15">
              {[2025, 2026, 2027].map((yr) => (
                <button
                  key={yr}
                  type="button"
                  onClick={() => {
                    setSelectedYear(yr);
                    setCurrentPage(1);
                  }}
                  className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                    selectedYear === yr
                      ? 'bg-amber-400 text-slate-950'
                      : 'text-slate-200 hover:text-white'
                  }`}
                >
                  {yr}
                </button>
              ))}
            </div>

            {!isPublicReadOnly && (
              <>
                <button
                  type="button"
                  onClick={handleSyncExistingTablesToMondaySunday}
                  className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 rounded-xl transition-colors"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  Selaraskan Senin–Minggu
                </button>

                {tables.length > 3 && (
                  <button
                    type="button"
                    onClick={handleCleanupEmptyTables}
                    className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold bg-rose-500/20 hover:bg-rose-500/30 text-rose-200 border border-rose-400/30 rounded-xl transition-colors"
                    title="Hapus tabel mingguan kosong agar daftar tidak panjang"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    Rapikan Tabel Kosong ({tables.length})
                  </button>
                )}
              </>
            )}

            <button
              type="button"
              disabled={isPublicReadOnly}
              onClick={handleExportAnnualCalendarExcel}
              title={
                isPublicReadOnly
                  ? 'Mode Publik (Hanya Lihat): Tidak bisa klik Unduh / Export Excel'
                  : 'Unduh Kalender Weekly Setahun (.xls)'
              }
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-700 disabled:text-slate-400 disabled:opacity-60 disabled:cursor-not-allowed rounded-xl transition-colors shadow-2xs"
            >
              {isPublicReadOnly ? (
                <Lock className="w-3.5 h-3.5" />
              ) : (
                <Download className="w-3.5 h-3.5" />
              )}
              Unduh Excel (.xls)
            </button>
          </div>
        </div>

        {/* Direct "Buka Cepat Weekly" Action Bar */}
        <div className="bg-white/10 border border-white/15 rounded-xl p-3.5 flex flex-wrap items-center justify-between gap-2.5 w-full max-w-full min-w-0">
          <div className="flex flex-wrap items-center gap-2 flex-1 min-w-0 max-w-full">
            <span className="text-xs font-bold text-amber-300 whitespace-nowrap shrink-0">
              Buka Cepat Weekly:
            </span>
            <select
              aria-label="Pilih Cepat Weekly 1 sampai 53"
              value={quickJumpWeekNum}
              onChange={(e) => setQuickJumpWeekNum(Number(e.target.value))}
              className="px-2.5 py-1.5 text-xs font-bold bg-slate-900 text-white border border-slate-600 rounded-lg focus:outline-none focus:border-amber-400 flex-1 min-w-0 max-w-full sm:max-w-xs truncate"
            >
              {annualWeeks.map((w) => {
                const ex = tableByWeekNumber.get(w.weekNumber);
                const sls = ex ? calculateWeekMetrics(ex).weeklySales : 0;
                return (
                  <option key={w.weekNumber} value={w.weekNumber}>
                    W{w.weekNumber} ({w.monthName}): {w.seninSlash}–{w.mingguSlash}{' '}
                    {sls > 0
                      ? `[Rp ${formatNumberId(sls)}]`
                      : ex
                        ? '[Aktif]'
                        : ''}
                  </option>
                );
              })}
            </select>
            <div className="flex items-center gap-1.5 bg-slate-900 border border-slate-700 rounded-lg px-2 py-1 text-xs shrink-0">
              <Calendar className="w-3.5 h-3.5 text-amber-400 shrink-0" />
              <input
                aria-label="Pilih Tanggal Kalender"
                type="date"
                onChange={(e) => {
                  const val = e.target.value;
                  if (val) {
                    const found = annualWeeks.find((w) => {
                      const [y, m, d] = val.split('-').map(Number);
                      const t = new Date(y, m - 1, d).getTime();
                      const [sy, sm, sd] = w.startDateIso.split('-').map(Number);
                      const st = new Date(sy, sm - 1, sd).getTime();
                      const [my, mm, md] = w.endDateIso.split('-').map(Number);
                      const et = new Date(my, mm - 1, md).getTime();
                      return t >= st && t <= et;
                    });
                    if (found) {
                      setQuickJumpWeekNum(found.weekNumber);
                    }
                  }
                }}
                className="text-xs font-bold bg-transparent text-white focus:outline-none cursor-pointer"
                title="Pilih tanggal dari kalender"
              />
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 max-w-full">
            <button
              type="button"
              onClick={() => handleOpenOrAddWeek(selectedQuickWeekObj, 'spreadsheet')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 rounded-lg transition-colors shadow-2xs whitespace-nowrap"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Buka Tabel W{selectedQuickWeekObj.weekNumber}
            </button>
            <button
              type="button"
              onClick={() => handleOpenOrAddWeek(selectedQuickWeekObj, 'dashboard')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-emerald-500 hover:bg-emerald-400 text-slate-950 rounded-lg transition-colors whitespace-nowrap"
            >
              <BarChart3 className="w-3.5 h-3.5" />
              Buka di Dashboard
            </button>
            <button
              type="button"
              onClick={() => handleOpenOrAddWeek(selectedQuickWeekObj, 'input')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-white/15 hover:bg-white/25 text-white rounded-lg transition-colors whitespace-nowrap"
            >
              <PenSquare className="w-3.5 h-3.5" />
              Input Shift
            </button>
          </div>
        </div>
      </div>

      {statusNotice && (
        <div className="p-3 bg-emerald-50 border border-emerald-300 text-emerald-900 rounded-xl flex items-center gap-2 text-xs font-bold">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
          <span>{statusNotice}</span>
        </div>
      )}

      {/* Compact Filter Bar: Bulan, Kuartal, Status & Search */}
      <div className="bg-white border border-slate-200 rounded-xl p-3.5 sm:p-4 space-y-3 shadow-2xs w-full max-w-full overflow-hidden">
        {/* Row 1: 12 Month Pills */}
        <div className="flex flex-wrap items-center gap-1.5 w-full max-w-full">
          <span className="text-xs font-bold text-slate-700 mr-1 flex items-center gap-1 shrink-0">
            <Filter className="w-3.5 h-3.5 text-indigo-600" />
            Bulan:
          </span>
          {ORDERED_MONTHS.map((m) => (
            <button
              key={m.key}
              type="button"
              onClick={() => {
                setMonthFilter(m.key);
                setQuarterFilter('ALL');
                setCurrentPage(1);
              }}
              className={`px-2 py-1 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap ${
                monthFilter === m.key
                  ? 'bg-indigo-600 text-white shadow-2xs'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
              }`}
            >
              {m.label}
            </button>
          ))}
        </div>

        {/* Row 2: Kuartal, Status, Baris per Halaman & Search */}
        <div className="pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2.5 text-xs w-full max-w-full min-w-0">
          <div className="flex flex-wrap items-center gap-2.5 min-w-0 max-w-full">
            <div className="flex flex-wrap items-center gap-1 bg-slate-100 p-1 rounded-lg max-w-full">
              {[
                { id: 'ALL', label: 'Semua Kuartal' },
                { id: 'Q1', label: 'Q1 (Jan–Mar)' },
                { id: 'Q2', label: 'Q2 (Apr–Jun)' },
                { id: 'Q3', label: 'Q3 (Jul–Sep)' },
                { id: 'Q4', label: 'Q4 (Okt–Des)' },
              ].map((q) => (
                <button
                  key={q.id}
                  type="button"
                  onClick={() => {
                    setQuarterFilter(q.id as 'ALL' | 'Q1' | 'Q2' | 'Q3' | 'Q4');
                    if (q.id !== 'ALL') setMonthFilter('ALL');
                    setCurrentPage(1);
                  }}
                  className={`px-2 py-1 rounded-md font-semibold transition-colors whitespace-nowrap ${
                    quarterFilter === q.id
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {q.label}
                </button>
              ))}
            </div>

            <div className="flex flex-wrap items-center gap-1 bg-slate-100 p-1 rounded-lg max-w-full">
              {[
                { id: 'ALL', label: 'Semua Status' },
                { id: 'FILLED', label: `Ada Sales (${filledWeeksCount})` },
                { id: 'ACTIVE', label: `Tabel Aktif (${activeWeeksCount})` },
              ].map((st) => (
                <button
                  key={st.id}
                  type="button"
                  onClick={() => {
                    setStatusFilter(st.id as 'ALL' | 'FILLED' | 'ACTIVE');
                    setCurrentPage(1);
                  }}
                  className={`px-2 py-1 rounded-md font-semibold transition-colors whitespace-nowrap ${
                    statusFilter === st.id
                      ? 'bg-emerald-600 text-white'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {st.label}
                </button>
              ))}
            </div>

            <div className="flex items-center gap-1.5">
              <span className="text-slate-500 font-medium shrink-0">Tampil:</span>
              <select
                aria-label="Jumlah baris per halaman"
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="px-2 py-1 text-xs font-bold bg-slate-100 border border-slate-200 rounded-md text-slate-800"
              >
                <option value={5}>5 Weekly / hal</option>
                <option value={6}>6 Weekly / hal</option>
                <option value={10}>10 Weekly / hal</option>
                <option value={15}>15 Weekly / hal</option>
              </select>
            </div>
          </div>

          <div className="relative w-full sm:w-56 min-w-0">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                if (e.target.value.trim() !== '') {
                  setMonthFilter('ALL');
                }
                setCurrentPage(1);
              }}
              placeholder="Cari Weekly (mis: 41) / tanggal..."
              className="w-full pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-none focus:border-indigo-600"
            />
          </div>
        </div>
      </div>

      {/* Filtered & Paginated Weekly Table (Compact, Never Long Downward) */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
        <div className="px-5 py-3 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-amber-400" />
            <h2 className="text-sm font-bold">
              Data Weekly Senin s/d Minggu — Menampilkan {paginatedWeeks.length} dari{' '}
              {filteredWeeks.length} Minggu Terfilter
            </h2>
          </div>

          {/* Pagination Controls in Table Header */}
          {totalPages > 1 && (
            <div className="flex items-center gap-1.5 text-xs">
              <button
                type="button"
                disabled={safePage <= 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-40 text-white font-semibold"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
                Prev
              </button>
              <span className="px-2.5 py-1 rounded bg-amber-400 text-slate-950 font-bold font-mono-num">
                Hal {safePage} / {totalPages}
              </span>
              <button
                type="button"
                disabled={safePage >= totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-40 text-white font-semibold"
              >
                Next
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          )}
        </div>

        <div className="overflow-x-auto spreadsheet-scroll">
          <table className="w-full border-collapse text-xs min-w-[1180px]">
            <thead>
              <tr className="bg-slate-100 text-slate-800 border-b border-slate-300">
                <th className="py-2.5 px-3 text-left font-bold border-r border-slate-200 w-28">
                  Weekly Ke-
                </th>
                <th className="py-2.5 px-3 text-left font-bold border-r border-slate-200 w-28">
                  Bulan
                </th>
                <th className="py-2.5 px-3 text-left font-bold border-r border-slate-200 min-w-[190px]">
                  Periode (Senin – Minggu)
                </th>
                <th className="py-2 px-2 text-center font-bold bg-indigo-50/70 text-indigo-950 border-r border-slate-200">
                  <div>Senin (H1)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-indigo-50/70 text-indigo-950 border-r border-slate-200">
                  <div>Selasa (H2)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-indigo-50/70 text-indigo-950 border-r border-slate-200">
                  <div>Rabu (H3)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-indigo-50/70 text-indigo-950 border-r border-slate-200">
                  <div>Kamis (H4)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-indigo-50/70 text-indigo-950 border-r border-slate-200">
                  <div>Jumat (H5)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-rose-50 text-rose-900 border-r border-slate-200">
                  <div>Sabtu (H6)</div>
                </th>
                <th className="py-2 px-2 text-center font-bold bg-rose-50 text-rose-900 border-r border-slate-200">
                  <div>Minggu (H7)</div>
                </th>
                <th className="py-2.5 px-3 text-right font-bold bg-amber-50 text-amber-950 border-r border-slate-200 w-36">
                  Total Sales
                </th>
                <th className="py-2.5 px-3 text-center font-bold w-56">
                  Buka Weekly
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {paginatedWeeks.length === 0 ? (
                <tr>
                  <td colSpan={12} className="py-8 text-center text-slate-500">
                    Tidak ada periode Weekly pada filter ini. Klik{' '}
                    <button
                      type="button"
                      onClick={() => {
                        setMonthFilter('ALL');
                        setQuarterFilter('ALL');
                        setStatusFilter('ALL');
                        setSearchQuery('');
                        setCurrentPage(1);
                      }}
                      className="text-indigo-600 font-bold underline"
                    >
                      Reset Filter
                    </button>{' '}
                    untuk melihat semua minggu.
                  </td>
                </tr>
              ) : (
                paginatedWeeks.map((w) => {
                  const existingTable = tableByWeekNumber.get(w.weekNumber);
                  const metrics = existingTable ? calculateWeekMetrics(existingTable) : null;
                  const hasSales = metrics && metrics.weeklySales > 0;
                  const isOctHighlight = w.weekNumber === 41 || w.weekNumber === 42;

                  return (
                    <tr
                      key={w.weekNumber}
                      className={`transition-colors ${
                        hasSales
                          ? 'bg-emerald-50/35 hover:bg-emerald-50/70'
                          : isOctHighlight
                            ? 'bg-amber-50/30 hover:bg-amber-50/60'
                            : 'hover:bg-slate-50'
                      }`}
                    >
                      <td className="py-2.5 px-3 border-r border-slate-200 font-bold">
                        <button
                          type="button"
                          onClick={() => handleOpenOrAddWeek(w, 'spreadsheet')}
                          className={`inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold transition-transform hover:scale-105 ${
                            hasSales
                              ? 'bg-emerald-600 text-white'
                              : existingTable
                                ? 'bg-indigo-600 text-white'
                                : 'bg-slate-900 text-white'
                          }`}
                          title={`Klik untuk langsung membuka Weekly ${w.weekNumber}`}
                        >
                          WEEKLY {w.weekNumber}
                        </button>
                      </td>
                      <td className="py-2.5 px-3 border-r border-slate-200">
                        <div className="font-bold text-slate-900">{w.monthName}</div>
                        <div className="text-[10px] text-slate-500 font-semibold">
                          {w.quarter} · {w.year}
                        </div>
                      </td>
                      <td className="py-2.5 px-3 border-r border-slate-200">
                        <div className="font-semibold text-slate-900">{w.periodeText}</div>
                        <div className="text-[10px] text-slate-500">Senin s/d Minggu</div>
                      </td>
                      {w.slashDates.map((dt, dIdx) => {
                        const isWknd = dIdx >= 5;
                        return (
                          <td
                            key={dIdx}
                            className={`py-2.5 px-2 text-center font-mono-num border-r border-slate-200 ${
                              isWknd
                                ? 'bg-rose-50/40 font-semibold text-rose-900'
                                : 'text-slate-700'
                            }`}
                          >
                            {dt}
                          </td>
                        );
                      })}
                      <td className="py-2.5 px-3 text-right font-mono-num border-r border-slate-200">
                        {hasSales ? (
                          <div>
                            <div className="font-bold text-emerald-700">
                              Rp {formatNumberId(metrics!.weeklySales)}
                            </div>
                            <div className="text-[10px] text-slate-500">
                              {metrics!.weeklyQty} pcs · {metrics!.weeklyTrx} trx
                            </div>
                          </div>
                        ) : existingTable ? (
                          <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">
                            Siap Diisi (Rp 0)
                          </span>
                        ) : (
                          <span className="text-slate-400 text-[11px]">Siap Dibuka</span>
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-center">
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => handleOpenOrAddWeek(w, 'spreadsheet')}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white transition-colors shadow-2xs"
                          >
                            <FileSpreadsheet className="w-3.5 h-3.5" />
                            <span>Buka Tabel</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleOpenOrAddWeek(w, 'dashboard')}
                            className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-semibold bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 transition-colors"
                            title={`Buka Weekly ${w.weekNumber} di Dashboard Weekly`}
                          >
                            <BarChart3 className="w-3.5 h-3.5" />
                            <span>Dashboard</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => handleOpenOrAddWeek(w, 'input')}
                            className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 transition-colors"
                            title={`Input Shift Harian Weekly ${w.weekNumber}`}
                          >
                            <PenSquare className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Bottom Pagination Bar */}
        {totalPages > 1 && (
          <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            <span className="text-slate-600 font-medium">
              Menampilkan baris <strong>{startIdx + 1}</strong> –{' '}
              <strong>{Math.min(startIdx + pageSize, filteredWeeks.length)}</strong> dari total{' '}
              <strong>{filteredWeeks.length}</strong> periode Weekly
            </span>
            <div className="flex items-center gap-1">
              <button
                type="button"
                disabled={safePage <= 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                className="px-2.5 py-1 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 font-semibold text-slate-700"
              >
                Sebelumnya
              </button>
              {Array.from({ length: totalPages }, (_, i) => i + 1).map((pg) => (
                <button
                  key={pg}
                  type="button"
                  onClick={() => setCurrentPage(pg)}
                  className={`w-7 h-7 rounded font-bold font-mono-num ${
                    safePage === pg
                      ? 'bg-indigo-600 text-white'
                      : 'bg-white border border-slate-200 text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  {pg}
                </button>
              ))}
              <button
                type="button"
                disabled={safePage >= totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                className="px-2.5 py-1 rounded border border-slate-300 bg-white hover:bg-slate-100 disabled:opacity-40 font-semibold text-slate-700"
              >
                Berikutnya
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
