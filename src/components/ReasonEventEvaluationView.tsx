import React, { useState, useEffect } from 'react';
import {
  WeeklySalesTable,
  DailySalesRecord,
  calculateDayMetrics,
  calculateWeekMetrics,
  formatNumberId,
  getWeeklyPeriodeText,
  getIndoDayName,
  isWeekendDate,
  getWeeklyDayCycleText,
  ANNUAL_2026_WEEKS,
  createWeeklyTableFromAnnualSchedule,
  getTableMonthKey,
  INITIAL_REVIEW_DOCS,
} from '../data/initialData';
import { exportWeeklyTablesToWPSExcel } from '../utils/exportExcel';
import { isoToSlashDate, formatIndoDateLong } from '../utils/dateUtils';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  markAutoDraftClean,
  clearAutoDraftFromStorage,
} from '../utils/autoDraft';
import { AutoDraftStatusBadge } from './AutoDraftStatusBanner';
import {
  Sun,
  Sunset,
  Calendar,
  CalendarRange,
  Search,
  Filter,
  LayoutGrid,
  Table2,
  Download,
  Undo2,
  Redo2,
  MessageSquareText,
  UserCheck,
  TrendingUp,
  X,
  FileSpreadsheet,
  Eraser,
  Save,
  Lock,
} from 'lucide-react';

interface ReasonEvalDraftPayload {
  tableId: string;
  weekNumber: number;
  days: DailySalesRecord[];
}

interface ReasonEventEvaluationViewProps {
  tables: WeeklySalesTable[];
  selectedWeeklyFilter: string; // 'ALL' or table.id
  onSelectWeeklyFilter: (filterId: string) => void;
  onChangeTables: (updated: WeeklySalesTable[]) => void;
  onOpenSpreadsheet?: (weekId: string) => void;
  canUndo?: boolean;
  canRedo?: boolean;
  undoCount?: number;
  redoCount?: number;
  onUndo?: () => void;
  onRedo?: () => void;
  onManualSave?: () => void;
  isPublicReadOnly?: boolean;
}

const TOPIC_PRESETS = [
  { id: 'ALL', label: 'Semua Topik Catatan', keyword: '' },
  { id: 'POP', label: 'POP / Promo', keyword: 'pop' },
  { id: 'TRAFFIC', label: 'Traffic / Pengunjung', keyword: 'pengunjung' },
  { id: 'EXPO', label: 'Event Expo / XXI', keyword: 'expo' },
  { id: 'SHOES', label: 'Produk Shoes / Sandal', keyword: 'shoes' },
  { id: 'SIZE', label: 'Broken Size', keyword: 'size' },
];

export const ReasonEventEvaluationView: React.FC<ReasonEventEvaluationViewProps> = ({
  tables,
  selectedWeeklyFilter,
  onSelectWeeklyFilter,
  onChangeTables,
  onOpenSpreadsheet,
  canUndo = false,
  canRedo = false,
  undoCount = 0,
  redoCount = 0,
  onUndo,
  onRedo,
  onManualSave,
  isPublicReadOnly = false,
}) => {
  const [viewMode, setViewMode] = useState<'cards' | 'table'>('cards');
  const [dayOfWeekFilter, setDayOfWeekFilter] = useState<string>('ALL');
  const [staffFilter, setStaffFilter] = useState<string>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeTopic, setActiveTopic] = useState<string>('ALL');
  const [monthFilter, setMonthFilter] = useState<string>('ALL');
  const [calendarDateFilter, setCalendarDateFilter] = useState<string>('');
  const [tablePage, setTablePage] = useState<number>(1);
  const [isEvalDraftDirty, setIsEvalDraftDirty] = useState<boolean>(false);
  const [recoveredEvalDraft, setRecoveredEvalDraft] =
    useState<ReasonEvalDraftPayload | null>(null);

  const handleOpenWeekNumber = (weekNum: number) => {
    const existing = tables.find((t) => t.weekNumber === weekNum);
    if (existing) {
      onSelectWeeklyFilter(existing.id);
      setTablePage(1);
      return;
    }
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNum);
    if (sched) {
      const created = createWeeklyTableFromAnnualSchedule(sched);
      const next = [...tables, created].sort((a, b) => a.weekNumber - b.weekNumber);
      onChangeTables(next);
      onSelectWeeklyFilter(created.id);
      setTablePage(1);
    }
  };

  const handleSelectCalendarDate = (iso: string) => {
    setCalendarDateFilter(iso);
    if (!iso) return;
    const targetSlash = isoToSlashDate(iso);
    const existingWithDate = tables.find((t) =>
      t.days.some((d) => d.tanggal.trim() === targetSlash)
    );
    if (existingWithDate) {
      onSelectWeeklyFilter(existingWithDate.id);
      setTablePage(1);
      return;
    }
    const matchedSched =
      ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(targetSlash)) ||
      ANNUAL_2026_WEEKS.find(
        (w) => iso >= w.startDateIso && iso <= w.endDateIso
      );
    if (matchedSched) {
      handleOpenWeekNumber(matchedSched.weekNumber);
    }
  };

  const activeSingleTable =
    tables.find((t) => t.id === selectedWeeklyFilter) || tables[0];

  useEffect(() => {
    if (isPublicReadOnly) return;
    const savedEnv = loadAutoDraftFromStorage<ReasonEvalDraftPayload>(
      AUTO_DRAFT_KEYS.REASON_EVALUATION
    );
    if (savedEnv && savedEnv.isDirty && savedEnv.data?.days) {
      const target = tables.find((t) => t.id === savedEnv.data.tableId);
      if (target && JSON.stringify(target.days) !== JSON.stringify(savedEnv.data.days)) {
        setRecoveredEvalDraft(savedEnv.data);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const { lastDraftTimeStr, isOnline, flushDraftNow } =
    usePeriodicAutoDraft<ReasonEvalDraftPayload>({
      storageKey: AUTO_DRAFT_KEYS.REASON_EVALUATION,
      menuName: 'Evaluasi Reason & Event',
      data: {
        tableId: activeSingleTable?.id || 'weekly-41',
        weekNumber: activeSingleTable?.weekNumber || 41,
        days: activeSingleTable?.days || [],
      },
      isDirty: isEvalDraftDirty,
      enabled: !isPublicReadOnly,
      intervalMs: 2500,
    });

  const handleRestoreEvalDraft = () => {
    if (!recoveredEvalDraft) return;
    const next = tables.map((t) =>
      t.id === recoveredEvalDraft.tableId
        ? { ...t, days: recoveredEvalDraft.days }
        : t
    );
    onChangeTables(next);
    onSelectWeeklyFilter(recoveredEvalDraft.tableId);
    setRecoveredEvalDraft(null);
    setIsEvalDraftDirty(false);
    markAutoDraftClean(AUTO_DRAFT_KEYS.REASON_EVALUATION);
  };

  const filteredTablesByMonth = tables.filter((t) =>
    monthFilter === 'ALL' ? true : getTableMonthKey(t) === monthFilter
  );

  const allMatchedTables =
    selectedWeeklyFilter === 'ALL'
      ? filteredTablesByMonth.length > 0
        ? filteredTablesByMonth
        : tables
      : tables.filter((t) => t.id === selectedWeeklyFilter);

  const totalTablePages = Math.max(1, allMatchedTables.length);
  const safeTablePage = Math.min(tablePage, totalTablePages);
  const visibleTables =
    selectedWeeklyFilter === 'ALL'
      ? allMatchedTables.slice(safeTablePage - 1, safeTablePage)
      : allMatchedTables;

  const handleDayReasonChange = (
    tableId: string,
    dayIndex: number,
    updater: (day: DailySalesRecord) => DailySalesRecord
  ) => {
    const next = tables.map((t) => {
      if (t.id !== tableId) return t;
      const nextDays = t.days.map((d, idx) => (idx === dayIndex ? updater(d) : d));
      return { ...t, days: nextDays };
    });
    setIsEvalDraftDirty(true);
    flushDraftNow(true);
    onChangeTables(next);
  };

  const handleClearTableNamesAndReasons = (tableId: string) => {
    const next = tables.map((t) => {
      if (t.id !== tableId) return t;
      return {
        ...t,
        days: t.days.map((d) => ({
          ...d,
          pagi: { ...d.pagi, nama: '' },
          siang: { ...d.siang, nama: '' },
          reasonPagi: '',
          reasonSiang: '',
        })),
      };
    });
    onChangeTables(next);
  };

  const dayMatchesFilters = (day: DailySalesRecord): boolean => {
    if (calendarDateFilter) {
      const targetSlash = isoToSlashDate(calendarDateFilter);
      if (day.tanggal.trim() !== targetSlash) return false;
    }

    const dName = getIndoDayName(day.tanggal);
    const wknd = isWeekendDate(day.tanggal);

    if (dayOfWeekFilter !== 'ALL') {
      if (dayOfWeekFilter === 'WEEKDAY' && wknd) return false;
      if (dayOfWeekFilter === 'WEEKEND' && !wknd) return false;
      if (
        dayOfWeekFilter !== 'WEEKDAY' &&
        dayOfWeekFilter !== 'WEEKEND' &&
        dName.toLowerCase() !== dayOfWeekFilter.toLowerCase()
      ) {
        return false;
      }
    }

    if (staffFilter !== 'ALL') {
      const matchPagi = day.pagi.nama.toLowerCase() === staffFilter.toLowerCase();
      const matchSiang = day.siang.nama.toLowerCase() === staffFilter.toLowerCase();
      if (!matchPagi && !matchSiang) return false;
    }

    const q = searchQuery.trim().toLowerCase();
    if (q !== '') {
      const textBlob = `${day.tanggal} ${dName} ${day.pagi.nama} ${day.siang.nama} ${day.reasonPagi} ${day.reasonSiang}`.toLowerCase();
      if (!textBlob.includes(q)) return false;
    }

    return true;
  };

  // Compute aggregated Shift Pagi vs Shift Siang stats across visibleTables
  let totalPagiSales = 0;
  let totalPagiQty = 0;
  let totalPagiTrx = 0;
  let filledPagiReasonCount = 0;

  let totalSiangSales = 0;
  let totalSiangQty = 0;
  let totalSiangTrx = 0;
  let filledSiangReasonCount = 0;
  let totalDaysCount = 0;

  visibleTables.forEach((t) => {
    t.days.forEach((d) => {
      if (!dayMatchesFilters(d)) return;
      totalDaysCount += 1;
      totalPagiSales += d.pagi.sales ?? 0;
      totalPagiQty += d.pagi.qty ?? 0;
      totalPagiTrx += d.pagi.trx ?? 0;
      if (d.reasonPagi && d.reasonPagi.trim() !== '') filledPagiReasonCount += 1;

      totalSiangSales += d.siang.sales ?? 0;
      totalSiangQty += d.siang.qty ?? 0;
      totalSiangTrx += d.siang.trx ?? 0;
      if (d.reasonSiang && d.reasonSiang.trim() !== '') filledSiangReasonCount += 1;
    });
  });

  const combinedSales = totalPagiSales + totalSiangSales;
  const pagiSharePct =
    combinedSales > 0 ? Math.round((totalPagiSales / combinedSales) * 100) : 0;
  const siangSharePct =
    combinedSales > 0 ? Math.round((totalSiangSales / combinedSales) * 100) : 0;

  const handleExportReasonExcel = () => {
    if (isPublicReadOnly) return;
    exportWeeklyTablesToWPSExcel(tables, INITIAL_REVIEW_DOCS[0], false, {
      includeWeeklyMatrix: false,
      includeDailyTable: false,
      includeMonthlyTable: false,
      includeReasonEvaluation: true,
      includeStaffAndReview: false,
      sheetLayout: 'SEPARATE_SHEETS',
      selectedWeekIds: visibleTables.map((t) => t.id),
      selectedMonthKey: monthFilter,
      customFileName: `Evaluasi_Reason_Event_Bazar_Matahari_Tanjung_${new Date().toISOString().slice(0, 10)}`,
    });
  };

  return (
    <div className="space-y-6">
      {!isPublicReadOnly && (
        <AutoDraftStatusBadge
          variant="banner"
          menuLabel={`Evaluasi Reason/Event (W${activeSingleTable?.weekNumber || 41})`}
          lastDraftTimeStr={lastDraftTimeStr}
          isDirty={isEvalDraftDirty}
          isOnline={isOnline}
          hasRestorableDraft={Boolean(recoveredEvalDraft)}
          restorableDraftSummary={
            recoveredEvalDraft
              ? `Weekly ${recoveredEvalDraft.weekNumber}`
              : undefined
          }
          onRestoreDraft={handleRestoreEvalDraft}
          onClearDraft={() => {
            setRecoveredEvalDraft(null);
            clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REASON_EVALUATION);
          }}
          onManualSave={() => {
            if (onManualSave) onManualSave();
            setIsEvalDraftDirty(false);
            markAutoDraftClean(AUTO_DRAFT_KEYS.REASON_EVALUATION);
          }}
        />
      )}

      {/* Top Hero Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-2xl p-5 sm:p-6 shadow-xs">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
          <div>
            <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded-md bg-amber-400/20 text-amber-300 text-xs font-bold mb-2">
              <MessageSquareText className="w-3.5 h-3.5" />
              <span>MENU EVALUASI OPERASIONAL HARIAN SHIFT PAGI &amp; SIANG</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-bold tracking-tight">
              Rekapitulasi &amp; Evaluasi Harian: Reason/Event Pagi vs Reason/Event Siang
            </h1>
            <p className="text-xs text-slate-300 mt-1">
              Perbandingan langsung catatan kondisi operasional, traffic pengunjung, event bazar, dan kendala POP/promo antara Shift Pagi (10.00–16.00) dan Shift Siang (15.00–21.30).
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {!isPublicReadOnly && (
              <AutoDraftStatusBadge
                variant="dark"
                menuLabel="Evaluasi"
                lastDraftTimeStr={lastDraftTimeStr}
                isDirty={isEvalDraftDirty}
                isOnline={isOnline}
              />
            )}
            {onManualSave && (
              <button
                type="button"
                onClick={() => {
                  onManualSave();
                  setIsEvalDraftDirty(false);
                  markAutoDraftClean(AUTO_DRAFT_KEYS.REASON_EVALUATION);
                }}
                title="Simpan Catatan Reason/Event Sekarang (Ctrl+S)"
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 rounded-xl shadow-xs transition-colors"
              >
                <Save className="w-3.5 h-3.5" />
                <span>Simpan</span>
              </button>
            )}
            {(onUndo || onRedo) && (
              <div className="inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15">
                <button
                  type="button"
                  onClick={onUndo}
                  disabled={!canUndo}
                  title="Undo (Ctrl+Z)"
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-white hover:bg-white/15 disabled:opacity-35 rounded-lg transition-colors"
                >
                  <Undo2 className="w-3.5 h-3.5" />
                  <span>Undo</span>
                  {undoCount > 0 && (
                    <span className="px-1.5 py-0.2 bg-amber-400 text-slate-950 rounded text-[10px] font-bold font-mono-num">
                      {undoCount}
                    </span>
                  )}
                </button>
                <div className="w-px h-4 bg-white/20 mx-0.5" />
                <button
                  type="button"
                  onClick={onRedo}
                  disabled={!canRedo}
                  title="Redo (Ctrl+Y)"
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-white hover:bg-white/15 disabled:opacity-35 rounded-lg transition-colors"
                >
                  <Redo2 className="w-3.5 h-3.5" />
                  <span>Redo</span>
                  {redoCount > 0 && (
                    <span className="px-1.5 py-0.2 bg-emerald-400 text-slate-950 rounded text-[10px] font-bold font-mono-num">
                      {redoCount}
                    </span>
                  )}
                </button>
              </div>
            )}

            <div className="inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15">
              <button
                type="button"
                onClick={() => setViewMode('cards')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                  viewMode === 'cards'
                    ? 'bg-amber-400 text-slate-950'
                    : 'text-slate-200 hover:text-white'
                }`}
              >
                <LayoutGrid className="w-3.5 h-3.5" />
                Kartu Berdampingan
              </button>
              <button
                type="button"
                onClick={() => setViewMode('table')}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg transition-colors ${
                  viewMode === 'table'
                    ? 'bg-amber-400 text-slate-950'
                    : 'text-slate-200 hover:text-white'
                }`}
              >
                <Table2 className="w-3.5 h-3.5" />
                Tabel Matriks
              </button>
            </div>

            <button
              type="button"
              disabled={isPublicReadOnly}
              onClick={handleExportReasonExcel}
              title={
                isPublicReadOnly
                  ? 'Mode Publik (Hanya Lihat): Tidak bisa klik Unduh / Export'
                  : 'Unduh Rekap Reason/Event (.xls)'
              }
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-700 disabled:text-slate-400 disabled:opacity-60 disabled:cursor-not-allowed rounded-xl transition-colors shadow-2xs"
            >
              {isPublicReadOnly ? (
                <Lock className="w-3.5 h-3.5" />
              ) : (
                <Download className="w-3.5 h-3.5" />
              )}
              Unduh Rekap (.xls)
            </button>
          </div>
        </div>

        {/* 3 Comparison Summary Cards: Shift Pagi vs Shift Siang */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5 mt-5 pt-4 border-t border-white/10">
          <div className="bg-amber-500/15 border border-amber-400/30 rounded-xl p-3.5 flex items-center justify-between">
            <div>
              <div className="flex items-center gap-1.5 text-amber-300 text-xs font-bold">
                <Sun className="w-4 h-4" />
                <span>REASON / EVENT SHIFT PAGI (10.00–16.00)</span>
              </div>
              <div className="text-lg font-bold font-mono-num text-white mt-1">
                Rp {formatNumberId(totalPagiSales) || '0'}{' '}
                <span className="text-xs font-normal text-amber-200">({pagiSharePct}%)</span>
              </div>
              <div className="text-[11px] text-slate-300 mt-0.5 font-mono-num">
                {filledPagiReasonCount}/{totalDaysCount} hari dievaluasi · {totalPagiQty} pcs · {totalPagiTrx} trx
              </div>
            </div>
          </div>

          <div className="bg-indigo-500/20 border border-indigo-400/30 rounded-xl p-3.5 flex items-center justify-between">
            <div>
              <div className="flex items-center gap-1.5 text-indigo-300 text-xs font-bold">
                <Sunset className="w-4 h-4" />
                <span>REASON / EVENT SHIFT SIANG (15.00–21.30)</span>
              </div>
              <div className="text-lg font-bold font-mono-num text-white mt-1">
                Rp {formatNumberId(totalSiangSales) || '0'}{' '}
                <span className="text-xs font-normal text-indigo-200">({siangSharePct}%)</span>
              </div>
              <div className="text-[11px] text-slate-300 mt-0.5 font-mono-num">
                {filledSiangReasonCount}/{totalDaysCount} hari dievaluasi · {totalSiangQty} pcs · {totalSiangTrx} trx
              </div>
            </div>
          </div>

          <div className="bg-white/5 border border-white/10 rounded-xl p-3.5 flex flex-col justify-between">
            <div className="flex items-center justify-between text-xs">
              <span className="text-slate-300 font-semibold">Proporsi Omzet Pagi vs Siang</span>
              <span className="font-mono-num font-bold text-emerald-400">
                Total: Rp {formatNumberId(combinedSales) || '0'}
              </span>
            </div>
            <div className="w-full h-3 bg-slate-800 rounded-full overflow-hidden flex my-2">
              <div
                style={{ width: `${pagiSharePct}%` }}
                className="h-full bg-amber-400"
                title={`Shift Pagi: ${pagiSharePct}%`}
              />
              <div
                style={{ width: `${siangSharePct}%` }}
                className="h-full bg-indigo-500"
                title={`Shift Siang: ${siangSharePct}%`}
              />
            </div>
            <div className="flex items-center justify-between text-[11px] font-mono-num">
              <span className="text-amber-300">Pagi: {pagiSharePct}%</span>
              <span className="text-indigo-300">Siang: {siangSharePct}%</span>
            </div>
          </div>
        </div>
      </div>

      {/* Compact Filter Bar: Kalender Tanggal, Hari, Petugas, Topik Evaluasi & Pencarian */}
      <div className="no-print bg-white border border-slate-200 rounded-xl p-3 sm:p-3.5 space-y-2.5 shadow-2xs w-full max-w-full overflow-hidden">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2 w-full">
          {/* Filter Kalender Tanggal */}
          <div className="flex items-center gap-1.5 px-2 py-1 rounded-lg bg-amber-50 border border-amber-300 min-w-0">
            <Calendar className="w-3.5 h-3.5 text-amber-700 shrink-0" />
            <input
              type="date"
              aria-label="Filter Kalender Tanggal Evaluasi"
              value={calendarDateFilter}
              onChange={(e) => handleSelectCalendarDate(e.target.value)}
              title={
                calendarDateFilter
                  ? `Menampilkan ${formatIndoDateLong(calendarDateFilter)}`
                  : 'Filter tanggal evaluasi lewat kalender'
              }
              className="w-full min-w-0 text-xs font-bold bg-transparent text-amber-950 focus:outline-none cursor-pointer"
            />
            {calendarDateFilter && (
              <button
                type="button"
                onClick={() => setCalendarDateFilter('')}
                className="px-1.5 py-0.5 rounded bg-amber-200 hover:bg-amber-300 text-amber-950 text-[10px] font-bold shrink-0 cursor-pointer"
                title="Tampilkan Semua Tanggal"
              >
                Semua
              </button>
            )}
          </div>

          {/* Filter Hari Dropdown */}
          <select
            aria-label="Filter Hari Evaluasi"
            value={dayOfWeekFilter}
            onChange={(e) => setDayOfWeekFilter(e.target.value)}
            className="w-full px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:outline-none focus:border-indigo-500 cursor-pointer truncate"
          >
            <option value="ALL">Hari: Semua (7 Hari)</option>
            <option value="WEEKDAY">Hari: Weekday (Sen–Jum)</option>
            <option value="WEEKEND">Hari: Weekend (Sab–Min)</option>
            <option value="Senin">Hari: Senin (H1)</option>
            <option value="Selasa">Hari: Selasa (H2)</option>
            <option value="Rabu">Hari: Rabu (H3)</option>
            <option value="Kamis">Hari: Kamis (H4)</option>
            <option value="Jumat">Hari: Jumat (H5)</option>
            <option value="Sabtu">Hari: Sabtu (H6)</option>
            <option value="Minggu">Hari: Minggu (H7)</option>
          </select>

          {/* Filter Petugas Dropdown */}
          <select
            aria-label="Filter Petugas Evaluasi"
            value={staffFilter}
            onChange={(e) => setStaffFilter(e.target.value)}
            className="w-full px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:outline-none focus:border-indigo-500 cursor-pointer truncate"
          >
            <option value="ALL">Petugas: Semua</option>
            <option value="Gani">Petugas: Gani</option>
            <option value="Irfan">Petugas: Irfan</option>
            <option value="Feronika">Petugas: Feronika</option>
            <option value="Abduh">Petugas: Abduh</option>
            <option value="Fitri">Petugas: Fitri</option>
          </select>

          {/* Filter Topik Evaluasi Dropdown */}
          <select
            aria-label="Filter Topik Evaluasi"
            value={activeTopic}
            onChange={(e) => {
              const tp = TOPIC_PRESETS.find((t) => t.id === e.target.value);
              if (tp) {
                setActiveTopic(tp.id);
                setSearchQuery(tp.keyword);
              }
            }}
            className="w-full px-2.5 py-1.5 text-xs font-bold bg-indigo-50/70 border border-indigo-200 rounded-lg text-indigo-900 focus:outline-none focus:border-indigo-500 cursor-pointer truncate"
          >
            {TOPIC_PRESETS.map((tp) => (
              <option key={tp.id} value={tp.id}>
                Topik: {tp.label}
              </option>
            ))}
          </select>

          {/* Search Box */}
          <div className="relative w-full min-w-0">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setActiveTopic('ALL');
              }}
              placeholder="Cari evaluasi / POP..."
              className="w-full pl-8 pr-7 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-none focus:border-indigo-600"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => {
                  setSearchQuery('');
                  setActiveTopic('ALL');
                }}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Pagination bar when 'ALL' mode is active with multiple tables */}
      {selectedWeeklyFilter === 'ALL' && totalTablePages > 1 && (
        <div className="bg-orange-50/90 border border-orange-200 rounded-xl px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="font-semibold text-orange-950">
            Menampilkan <span className="font-bold">1 tabel per halaman</span> (Halaman{' '}
            {safeTablePage} dari {totalTablePages}) agar halaman tidak panjang ke bawah.
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() => setTablePage((p) => Math.max(1, p - 1))}
              disabled={safeTablePage <= 1}
              className="px-2.5 py-1 rounded-lg bg-white border border-orange-200 font-bold text-orange-800 disabled:opacity-40"
            >
              ◀ Sebelumnya
            </button>
            <span className="px-2.5 py-1 font-mono-num font-bold text-orange-950">
              {safeTablePage} / {totalTablePages}
            </span>
            <button
              type="button"
              onClick={() => setTablePage((p) => Math.min(totalTablePages, p + 1))}
              disabled={safeTablePage >= totalTablePages}
              className="px-2.5 py-1 rounded-lg bg-white border border-orange-200 font-bold text-orange-800 disabled:opacity-40"
            >
              Berikutnya ▶
            </button>
          </div>
        </div>
      )}

      {/* Render Each Weekly Section */}
      {visibleTables.map((table) => {
        const weekMetrics = calculateWeekMetrics(table);
        const periodeLabel = getWeeklyPeriodeText(table);
        const dayCycleLabel = getWeeklyDayCycleText(table);
        const dayMetricsList = table.days.map((d) => calculateDayMetrics(d));

        return (
          <section
            key={table.id}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs space-y-0"
          >
            {/* Section Header with Integrated Filter Weekly Controls */}
            <div className="px-3.5 sm:px-5 py-3 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white space-y-2.5 w-full max-w-full min-w-0">
              <div className="flex flex-wrap items-center justify-between gap-2 w-full max-w-full min-w-0">
                <div className="flex flex-wrap items-center gap-1.5 min-w-0 max-w-full">
                  {/* Grouped Prev + WEEKLY Dropdown + Next so they NEVER split onto different lines */}
                  <div className="inline-flex items-center gap-1 bg-white/10 p-0.5 rounded-lg border border-white/15 max-w-full min-w-0">
                    <button
                      type="button"
                      onClick={() => handleOpenWeekNumber(Math.max(1, table.weekNumber - 1))}
                      disabled={table.weekNumber <= 1}
                      title="Buka Weekly Sebelumnya"
                      className="px-2 py-1 rounded-md hover:bg-white/20 disabled:opacity-35 text-white text-xs font-bold transition-colors shrink-0"
                    >
                      ◀
                    </button>
                    <select
                      value={table.weekNumber}
                      onChange={(e) => {
                        const wk = Number(e.target.value);
                        if (wk) handleOpenWeekNumber(wk);
                      }}
                      className="px-2 py-1 rounded-md bg-amber-400 hover:bg-amber-300 text-slate-950 text-xs font-extrabold cursor-pointer focus:outline-none min-w-0 max-w-[165px] sm:max-w-[220px] truncate shadow-2xs"
                      title="Filter & Pilih Weekly (W1 - W53)"
                    >
                      {ANNUAL_2026_WEEKS.map((w) => {
                        const exists = tables.some((t) => t.weekNumber === w.weekNumber);
                        return (
                          <option
                            key={w.weekNumber}
                            value={w.weekNumber}
                            className="bg-slate-900 text-white font-semibold"
                          >
                            WEEKLY {w.weekNumber} ({w.periodeText.replace(/\/2026/g, '')}){' '}
                            {exists ? '✓' : ''}
                          </option>
                        );
                      })}
                    </select>
                    <button
                      type="button"
                      onClick={() => handleOpenWeekNumber(Math.min(53, table.weekNumber + 1))}
                      disabled={table.weekNumber >= 53}
                      title="Buka Weekly Berikutnya"
                      className="px-2 py-1 rounded-md hover:bg-white/20 disabled:opacity-35 text-white text-xs font-bold transition-colors shrink-0"
                    >
                      ▶
                    </button>
                  </div>

                  {/* Filter Bulan inside Header */}
                  <select
                    value={monthFilter}
                    onChange={(e) => {
                      const nextM = e.target.value;
                      setMonthFilter(nextM);
                      setTablePage(1);
                      if (nextM !== 'ALL') {
                        const matching = tables.filter((t) => getTableMonthKey(t) === nextM);
                        if (matching.length > 0) {
                          onSelectWeeklyFilter(matching[0].id);
                        } else {
                          const firstAnnualInMonth = ANNUAL_2026_WEEKS.find(
                            (w) => w.startDateIso.slice(0, 7) === `2026-${nextM}`
                          );
                          if (firstAnnualInMonth) {
                            handleOpenWeekNumber(firstAnnualInMonth.weekNumber);
                          }
                        }
                      }
                    }}
                    className="px-2.5 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 border border-white/20 text-amber-200 text-xs font-bold focus:outline-none cursor-pointer min-w-0 max-w-[125px] truncate"
                    title="Filter Bulan"
                  >
                    <option value="ALL" className="bg-slate-900 text-white">Semua Bulan</option>
                    <option value="01" className="bg-slate-900 text-white">Januari</option>
                    <option value="02" className="bg-slate-900 text-white">Februari</option>
                    <option value="03" className="bg-slate-900 text-white">Maret</option>
                    <option value="04" className="bg-slate-900 text-white">April</option>
                    <option value="05" className="bg-slate-900 text-white">Mei</option>
                    <option value="06" className="bg-slate-900 text-white">Juni</option>
                    <option value="07" className="bg-slate-900 text-white">Juli</option>
                    <option value="08" className="bg-slate-900 text-white">Agustus</option>
                    <option value="09" className="bg-slate-900 text-white">September</option>
                    <option value="10" className="bg-slate-900 text-white">Oktober</option>
                    <option value="11" className="bg-slate-900 text-white">November</option>
                    <option value="12" className="bg-slate-900 text-white">Desember</option>
                  </select>
                </div>

                {/* Quick Active Weekly Pills inside Header */}
                <div className="flex flex-wrap items-center gap-1 max-w-full">
                  {filteredTablesByMonth.map((t) => {
                    const isCur = t.id === table.id && selectedWeeklyFilter !== 'ALL';
                    return (
                      <button
                        key={t.id}
                        type="button"
                        onClick={() => onSelectWeeklyFilter(t.id)}
                        className={`px-2 py-1 rounded-md text-[11px] font-bold transition-colors ${
                          isCur
                            ? 'bg-amber-400 text-slate-950 shadow-2xs'
                            : 'bg-white/10 text-slate-200 hover:bg-white/20'
                        }`}
                      >
                        W{t.weekNumber}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-white/10">
                <div>
                  <div className="text-sm font-bold">
                    {table.title} — Periode : {periodeLabel}
                  </div>
                  <div className="text-[11px] text-slate-300 font-mono-num">
                    Siklus Hari: {dayCycleLabel} · Total Sales: Rp{' '}
                    {formatNumberId(weekMetrics.weeklySales) || '0'} ({weekMetrics.weeklyQty} pcs,{' '}
                    {weekMetrics.weeklyTrx} trx)
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => handleClearTableNamesAndReasons(table.id)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-amber-200 transition-colors"
                  >
                    <Eraser className="w-3.5 h-3.5" />
                    Kosongkan Nama &amp; Reason
                  </button>
                  {onOpenSpreadsheet && (
                    <button
                      type="button"
                      onClick={() => onOpenSpreadsheet(table.id)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white transition-colors"
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5" />
                      Buka Tabel Sales
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* MODE 1: SIDE-BY-SIDE CARDS VIEW */}
            {viewMode === 'cards' ? (
              <div className="p-4 sm:p-5 bg-slate-50/60 grid grid-cols-1 lg:grid-cols-2 gap-4">
                {table.days.map((day, dIdx) => {
                  if (!dayMatchesFilters(day)) return null;
                  const dName = getIndoDayName(day.tanggal);
                  const wknd = isWeekendDate(day.tanggal);
                  const dm = dayMetricsList[dIdx];

                  return (
                    <div
                      key={`${day.id}-eval-card`}
                      className="bg-white border border-slate-200 rounded-xl p-4 space-y-3 shadow-2xs"
                    >
                      {/* Card Day Header */}
                      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2.5">
                        <div className="flex items-center gap-2">
                          <span
                            className={`px-2.5 py-0.5 rounded-md text-xs font-bold ${
                              wknd
                                ? 'bg-rose-100 text-rose-800 border border-rose-200'
                                : 'bg-slate-900 text-white'
                            }`}
                          >
                            Hari {dIdx + 1} · {dName}
                          </span>
                          <span className="text-xs font-bold font-mono-num text-slate-900">
                            {day.tanggal}
                          </span>
                          {wknd && (
                            <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-rose-50 text-rose-700">
                              Weekend
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-2 text-xs font-mono-num">
                          <span className="font-bold text-emerald-700">
                            Total: Rp {formatNumberId(dm.totalSales) || '0'}
                          </span>
                          <span className="text-slate-400">
                            ({dm.totalQty} pcs · {dm.totalTrx} trx)
                          </span>
                        </div>
                      </div>

                      {/* Side-by-Side Reason/Event Pagi vs Reason/Event Siang */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {/* Shift Pagi Box */}
                        <div className="p-3 rounded-xl bg-amber-50/60 border border-amber-200/90 space-y-2 flex flex-col justify-between">
                          <div className="space-y-1.5">
                            <div className="flex items-center justify-between border-b border-amber-200/60 pb-1.5">
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold text-amber-950">
                                <Sun className="w-3.5 h-3.5 text-amber-600" />
                                REASON / EVENT PAGI
                              </span>
                              <span className="text-[11px] font-bold font-mono-num text-amber-900">
                                {day.pagi.sales !== null
                                  ? `Rp ${formatNumberId(day.pagi.sales)}`
                                  : 'Rp 0'}
                              </span>
                            </div>
                            <div className="flex items-center justify-between text-[11px] text-amber-900/90">
                              <span className="inline-flex items-center gap-1 font-semibold">
                                <UserCheck className="w-3 h-3" />
                                Petugas: {day.pagi.nama || '-'}
                              </span>
                              <span className="font-mono-num">
                                {day.pagi.qty ?? 0} pcs · {day.pagi.trx ?? 0} trx
                              </span>
                            </div>
                            <textarea
                              rows={4}
                              value={day.reasonPagi}
                              onChange={(e) =>
                                handleDayReasonChange(table.id, dIdx, (d) => ({
                                  ...d,
                                  reasonPagi: e.target.value,
                                }))
                              }
                              placeholder="Tulis evaluasi Reason / Event Shift Pagi..."
                              className="w-full text-xs leading-relaxed text-slate-800 bg-white border border-amber-300/80 rounded-lg p-2 focus:outline-1 focus:outline-amber-500"
                            />
                          </div>
                        </div>

                        {/* Shift Siang Box */}
                        <div className="p-3 rounded-xl bg-indigo-50/60 border border-indigo-200/90 space-y-2 flex flex-col justify-between">
                          <div className="space-y-1.5">
                            <div className="flex items-center justify-between border-b border-indigo-200/60 pb-1.5">
                              <span className="inline-flex items-center gap-1 text-[11px] font-bold text-indigo-950">
                                <Sunset className="w-3.5 h-3.5 text-indigo-600" />
                                REASON / EVENT SIANG
                              </span>
                              <span className="text-[11px] font-bold font-mono-num text-indigo-900">
                                {day.siang.sales !== null
                                  ? `Rp ${formatNumberId(day.siang.sales)}`
                                  : 'Rp 0'}
                              </span>
                            </div>
                            <div className="flex items-center justify-between text-[11px] text-indigo-900/90">
                              <span className="inline-flex items-center gap-1 font-semibold">
                                <UserCheck className="w-3 h-3" />
                                Petugas: {day.siang.nama || '-'}
                              </span>
                              <span className="font-mono-num">
                                {day.siang.qty ?? 0} pcs · {day.siang.trx ?? 0} trx
                              </span>
                            </div>
                            <textarea
                              rows={4}
                              value={day.reasonSiang}
                              onChange={(e) =>
                                handleDayReasonChange(table.id, dIdx, (d) => ({
                                  ...d,
                                  reasonSiang: e.target.value,
                                }))
                              }
                              placeholder="Tulis evaluasi Reason / Event Shift Siang..."
                              className="w-full text-xs leading-relaxed text-slate-800 bg-white border border-indigo-300/80 rounded-lg p-2 focus:outline-1 focus:outline-indigo-600"
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              /* MODE 2: FULL COMPARISON MATRIX TABLE VIEW */
              <div className="overflow-x-auto spreadsheet-scroll">
                <table className="w-full text-xs border-collapse min-w-[1100px]">
                  <thead>
                    <tr className="bg-slate-100 text-slate-800 border-b border-slate-300">
                      <th className="py-3 px-4 text-left font-bold w-40 border-r border-slate-200">
                        Hari &amp; Tanggal
                      </th>
                      <th className="py-3 px-4 text-left font-bold bg-amber-50 text-amber-950 border-r border-slate-200 w-[42%]">
                        <div className="flex items-center gap-1.5">
                          <Sun className="w-4 h-4 text-amber-600" />
                          <span>Reason / Event Shift Pagi (10.00 – 16.00)</span>
                        </div>
                      </th>
                      <th className="py-3 px-4 text-left font-bold bg-indigo-50 text-indigo-950 border-r border-slate-200 w-[42%]">
                        <div className="flex items-center gap-1.5">
                          <Sunset className="w-4 h-4 text-indigo-600" />
                          <span>Reason / Event Shift Siang (15.00 – 21.30)</span>
                        </div>
                      </th>
                      <th className="py-3 px-4 text-right font-bold bg-emerald-50 text-emerald-950">
                        Total Harian
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {table.days.map((day, dIdx) => {
                      if (!dayMatchesFilters(day)) return null;
                      const dName = getIndoDayName(day.tanggal);
                      const wknd = isWeekendDate(day.tanggal);
                      const dm = dayMetricsList[dIdx];

                      return (
                        <tr key={`${day.id}-eval-row`} className="align-top hover:bg-slate-50/70">
                          <td className="py-3 px-4 border-r border-slate-200">
                            <span
                              className={`inline-flex px-2 py-0.5 rounded text-[10px] font-bold ${
                                wknd
                                  ? 'bg-rose-100 text-rose-800'
                                  : 'bg-slate-900 text-white'
                              }`}
                            >
                              Hari {dIdx + 1} · {dName}
                            </span>
                            <div className="font-mono-num font-bold text-slate-900 mt-1">
                              {day.tanggal}
                            </div>
                          </td>
                          <td className="py-3 px-4 border-r border-slate-200 bg-amber-50/20">
                            <div className="flex items-center justify-between text-[11px] font-bold text-amber-900 mb-1.5">
                              <span>Petugas Pagi: {day.pagi.nama || '-'}</span>
                              <span className="font-mono-num">
                                {day.pagi.sales !== null
                                  ? `Rp ${formatNumberId(day.pagi.sales)}`
                                  : 'Rp 0'}{' '}
                                ({day.pagi.qty ?? 0} pcs · {day.pagi.trx ?? 0} trx)
                              </span>
                            </div>
                            <textarea
                              rows={3}
                              value={day.reasonPagi}
                              onChange={(e) =>
                                handleDayReasonChange(table.id, dIdx, (d) => ({
                                  ...d,
                                  reasonPagi: e.target.value,
                                }))
                              }
                              placeholder="Isi Reason / Event Shift Pagi..."
                              className="w-full text-xs leading-relaxed text-slate-800 bg-white border border-amber-200 rounded-lg p-2 focus:outline-1 focus:outline-amber-500"
                            />
                          </td>
                          <td className="py-3 px-4 border-r border-slate-200 bg-indigo-50/20">
                            <div className="flex items-center justify-between text-[11px] font-bold text-indigo-900 mb-1.5">
                              <span>Petugas Siang: {day.siang.nama || '-'}</span>
                              <span className="font-mono-num">
                                {day.siang.sales !== null
                                  ? `Rp ${formatNumberId(day.siang.sales)}`
                                  : 'Rp 0'}{' '}
                                ({day.siang.qty ?? 0} pcs · {day.siang.trx ?? 0} trx)
                              </span>
                            </div>
                            <textarea
                              rows={3}
                              value={day.reasonSiang}
                              onChange={(e) =>
                                handleDayReasonChange(table.id, dIdx, (d) => ({
                                  ...d,
                                  reasonSiang: e.target.value,
                                }))
                              }
                              placeholder="Isi Reason / Event Shift Siang..."
                              className="w-full text-xs leading-relaxed text-slate-800 bg-white border border-indigo-200 rounded-lg p-2 focus:outline-1 focus:outline-indigo-600"
                            />
                          </td>
                          <td className="py-3 px-4 text-right font-mono-num">
                            <div className="font-bold text-emerald-700 text-sm">
                              Rp {formatNumberId(dm.totalSales) || '0'}
                            </div>
                            <div className="text-[11px] text-slate-500 mt-0.5">
                              {dm.totalQty} pcs · {dm.totalTrx} trx
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        );
      })}
    </div>
  );
};
