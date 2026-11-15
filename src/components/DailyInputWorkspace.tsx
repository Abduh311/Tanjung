import React, { useState, useEffect } from 'react';
import {
  WeeklySalesTable,
  DailySalesRecord,
  calculateDayMetrics,
  formatNumberId,
  ANNUAL_2026_WEEKS,
  createWeeklyTableFromAnnualSchedule,
} from '../data/initialData';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  markAutoDraftClean,
  clearAutoDraftFromStorage,
} from '../utils/autoDraft';
import { AutoDraftStatusBadge } from './AutoDraftStatusBanner';
import {
  slashDateToIso,
  isoToSlashDate,
  formatIndoDateLong,
} from '../utils/dateUtils';
import {
  Calendar,
  Sun,
  Sunset,
  CheckCircle2,
  Calculator,
  ChevronLeft,
  ChevronRight,
  MessageSquareText,
  Undo2,
  Redo2,
  RotateCcw,
  Save,
} from 'lucide-react';

interface DailyInputDraftPayload {
  weekNumber: number;
  selectedDayIndex: number;
  dayRecord: DailySalesRecord;
}

interface DailyInputWorkspaceProps {
  tables: WeeklySalesTable[];
  activeWeekIndex: number;
  onSelectWeekIndex: (idx: number) => void;
  onChangeTables: (updated: WeeklySalesTable[]) => void;
  canUndo?: boolean;
  canRedo?: boolean;
  onUndo?: () => void;
  onRedo?: () => void;
  onManualSave?: () => void;
}

export const DailyInputWorkspace: React.FC<DailyInputWorkspaceProps> = ({
  tables,
  activeWeekIndex,
  onSelectWeekIndex,
  onChangeTables,
  canUndo = false,
  canRedo = false,
  onUndo,
  onRedo,
  onManualSave,
}) => {
  const [selectedDayIndex, setSelectedDayIndex] = useState<number>(0);
  const [savedNotice, setSavedNotice] = useState<boolean>(false);
  const [isDraftDirty, setIsDraftDirty] = useState<boolean>(false);
  const [recoveredDraftBanner, setRecoveredDraftBanner] =
    useState<DailyInputDraftPayload | null>(null);

  const currentTable = tables[activeWeekIndex] || tables[0];
  const currentDay = currentTable.days[selectedDayIndex] || currentTable.days[0];
  const dayMetrics = calculateDayMetrics(currentDay);

  // Check if there is an uncommitted Auto-Draft in localStorage on mount
  useEffect(() => {
    const savedEnv = loadAutoDraftFromStorage<DailyInputDraftPayload>(
      AUTO_DRAFT_KEYS.DAILY_INPUT
    );
    if (savedEnv && savedEnv.isDirty && savedEnv.data?.dayRecord) {
      const d = savedEnv.data;
      const targetTable = tables.find((t) => t.weekNumber === d.weekNumber);
      const existingDay = targetTable?.days[d.selectedDayIndex];
      if (
        existingDay &&
        JSON.stringify(existingDay) !== JSON.stringify(d.dayRecord)
      ) {
        setRecoveredDraftBanner(d);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const lastDraftTimeStr = '';
  const isOnline = true;

  const handleRestoreDraft = () => {
    if (!recoveredDraftBanner) return;
    const { weekNumber, selectedDayIndex: dIdx, dayRecord } = recoveredDraftBanner;
    const tblIdx = tables.findIndex((t) => t.weekNumber === weekNumber);
    if (tblIdx >= 0) {
      onSelectWeekIndex(tblIdx);
      setSelectedDayIndex(dIdx);
      const nextTables = tables.map((t, idx) => {
        if (idx !== tblIdx) return t;
        return {
          ...t,
          days: t.days.map((dayItem, dayI) =>
            dayI === dIdx ? dayRecord : dayItem
          ),
        };
      });
      onChangeTables(nextTables);
    }
    setRecoveredDraftBanner(null);
    setIsDraftDirty(false);
    markAutoDraftClean(AUTO_DRAFT_KEYS.DAILY_INPUT);
  };

  const handleOpenWeekNumber = (weekNum: number) => {
    const existingIdx = tables.findIndex((t) => t.weekNumber === weekNum);
    if (existingIdx >= 0) {
      onSelectWeekIndex(existingIdx);
      setSelectedDayIndex(0);
      return;
    }
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNum);
    if (sched) {
      const created = createWeeklyTableFromAnnualSchedule(sched);
      const next = [...tables, created].sort((a, b) => a.weekNumber - b.weekNumber);
      onChangeTables(next);
      const newIdx = next.findIndex((t) => t.id === created.id);
      if (newIdx >= 0) {
        onSelectWeekIndex(newIdx);
      }
      setSelectedDayIndex(0);
    }
  };

  const handleSelectCalendarDate = (isoDate: string) => {
    if (!isoDate) return;
    const slash = isoToSlashDate(isoDate);

    // 1. Search in existing tables
    for (let tIdx = 0; tIdx < tables.length; tIdx++) {
      const table = tables[tIdx];
      const dIdx = table.days.findIndex((d) => d.tanggal === slash);
      if (dIdx >= 0) {
        onSelectWeekIndex(tIdx);
        setSelectedDayIndex(dIdx);
        return;
      }
    }

    // 2. Search in ANNUAL_2026_WEEKS and create table if missing
    const sched = ANNUAL_2026_WEEKS.find((w) => {
      const [y, m, d] = isoDate.split('-').map(Number);
      const targetTime = new Date(y, m - 1, d).getTime();
      const [sy, sm, sd] = w.startDateIso.split('-').map(Number);
      const startTime = new Date(sy, sm - 1, sd).getTime();
      const [my, mm, md] = w.endDateIso.split('-').map(Number);
      const endTime = new Date(my, mm - 1, md).getTime();
      return targetTime >= startTime && targetTime <= endTime;
    });

    if (sched) {
      const created = createWeeklyTableFromAnnualSchedule(sched);
      const next = [...tables, created].sort((a, b) => a.weekNumber - b.weekNumber);
      onChangeTables(next);
      const newIdx = next.findIndex((t) => t.id === created.id);
      if (newIdx >= 0) {
        onSelectWeekIndex(newIdx);
        const dayIdx = created.days.findIndex((d) => d.tanggal === slash);
        setSelectedDayIndex(dayIdx >= 0 ? dayIdx : 0);
      }
    }
  };

  const updateCurrentDay = (updater: (d: DailySalesRecord) => DailySalesRecord) => {
    const updatedDay = updater(currentDay);
    const nextTables = tables.map((t, tIdx) => {
      if (tIdx !== activeWeekIndex) return t;
      return {
        ...t,
        days: t.days.map((d, dIdx) => (dIdx === selectedDayIndex ? updatedDay : d)),
      };
    });
    onChangeTables(nextTables);
  };

  return (
    <div className="space-y-6">
      {/* Top Selection Bar: Choose Week & Choose Day */}
      <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5 space-y-4 w-full max-w-full overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-3 w-full max-w-full min-w-0">
          <div className="min-w-0">
            <h2 className="text-base font-bold text-slate-900">
              Form Input Penjualan & Evaluasi Shift Harian
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Pilih minggu dan tanggal operasional untuk mengisi data Shift Pagi & Shift Siang secara terstruktur
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 min-w-0 max-w-full">
            {savedNotice && (
              <span className="inline-flex items-center gap-1 text-xs font-medium text-emerald-700">
                <CheckCircle2 className="w-3.5 h-3.5" />
                Tersimpan
              </span>
            )}
            {onManualSave && (
              <button
                type="button"
                onClick={() => {
                  onManualSave();
                  setIsDraftDirty(false);
                  markAutoDraftClean(AUTO_DRAFT_KEYS.DAILY_INPUT);
                  setSavedNotice(true);
                  setTimeout(() => setSavedNotice(false), 1800);
                }}
                title="Simpan Data Shift Harian Sekarang (Ctrl+S)"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg shadow-2xs transition-colors"
              >
                <Save className="w-3.5 h-3.5" />
                <span>Simpan Data</span>
              </button>
            )}
            {(onUndo || onRedo) && (
              <div className="inline-flex items-center bg-slate-100 p-1 rounded-lg border border-slate-200">
                <button
                  type="button"
                  onClick={onUndo}
                  disabled={!canUndo}
                  title="Undo (Ctrl+Z)"
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-white disabled:opacity-40 rounded-md transition-colors"
                >
                  <Undo2 className="w-3.5 h-3.5" />
                  <span>Undo</span>
                </button>
                <button
                  type="button"
                  onClick={onRedo}
                  disabled={!canRedo}
                  title="Redo (Ctrl+Y)"
                  className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-slate-700 hover:bg-white disabled:opacity-40 rounded-md transition-colors"
                >
                  <Redo2 className="w-3.5 h-3.5" />
                  <span>Redo</span>
                </button>
              </div>
            )}
            {/* Controlled Buka Weekly (W1-W53) with Prev/Next */}
            <div className="inline-flex items-center gap-1 bg-cyan-50 border border-cyan-200 rounded-lg p-1 min-w-0 max-w-full">
              <button
                type="button"
                onClick={() =>
                  handleOpenWeekNumber(Math.max(1, (currentTable?.weekNumber || 41) - 1))
                }
                disabled={(currentTable?.weekNumber || 1) <= 1}
                title="Buka Weekly Sebelumnya"
                className="px-2 py-1 text-xs font-bold text-cyan-900 hover:bg-white rounded disabled:opacity-40 transition-colors shrink-0"
              >
                ◀
              </button>
              <select
                aria-label="Buka Weekly"
                value={String(currentTable?.weekNumber || 41)}
                onChange={(e) => handleOpenWeekNumber(Number(e.target.value))}
                className="px-2 py-1 text-xs font-bold bg-white border border-cyan-200 text-slate-900 rounded focus:outline-none cursor-pointer min-w-0 max-w-[195px] sm:max-w-[230px] truncate"
              >
                {ANNUAL_2026_WEEKS.map((w) => {
                  const isLoaded = tables.some((t) => t.weekNumber === w.weekNumber);
                  return (
                    <option key={w.weekNumber} value={String(w.weekNumber)}>
                      W{w.weekNumber} ({w.seninSlash}–{w.mingguSlash}){isLoaded ? ' ✓' : ''}
                    </option>
                  );
                })}
              </select>
              <button
                type="button"
                onClick={() =>
                  handleOpenWeekNumber(Math.min(53, (currentTable?.weekNumber || 41) + 1))
                }
                disabled={(currentTable?.weekNumber || 53) >= 53}
                title="Buka Weekly Berikutnya"
                className="px-2 py-1 text-xs font-bold text-cyan-900 hover:bg-white rounded disabled:opacity-40 transition-colors shrink-0"
              >
                ▶
              </button>
            </div>

            {/* Filter Kalender Tanggal Langsung */}
            <div className="inline-flex items-center gap-1.5 bg-amber-50 border border-amber-300 rounded-lg px-2.5 py-1 min-w-0 shadow-2xs">
              <Calendar className="w-3.5 h-3.5 text-amber-700 shrink-0" />
              <label htmlFor="daily-input-cal-filter" className="text-xs font-bold text-amber-950 shrink-0">
                Filter Kalender:
              </label>
              <input
                id="daily-input-cal-filter"
                type="date"
                value={slashDateToIso(currentDay.tanggal)}
                onChange={(e) => handleSelectCalendarDate(e.target.value)}
                className="px-2 py-0.5 text-xs font-bold bg-white border border-amber-300 rounded text-slate-900 focus:outline-none focus:ring-1 focus:ring-amber-500 cursor-pointer"
                title="Pilih tanggal dari kalender untuk langsung membuka week dan hari operasional terkait"
              />
            </div>

            <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg max-w-full">
              {tables.slice(0, 6).map((t, idx) => (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => {
                    onSelectWeekIndex(idx);
                    setSelectedDayIndex(0);
                  }}
                  className={`px-2 py-1 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                    idx === activeWeekIndex
                      ? 'bg-white text-slate-900 shadow-xs font-bold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  W{t.weekNumber}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* 7-Day Selector Tabs */}
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2 pt-1">
          {currentTable.days.map((day, idx) => {
            const m = calculateDayMetrics(day);
            const isSelected = idx === selectedDayIndex;
            const hasSales = m.totalSales > 0;

            return (
              <button
                key={day.id}
                type="button"
                onClick={() => setSelectedDayIndex(idx)}
                className={`p-3 rounded-lg border text-left transition-all ${
                  isSelected
                    ? 'bg-slate-900 text-white border-slate-900 shadow-xs'
                    : 'bg-slate-50/70 hover:bg-slate-100 text-slate-800 border-slate-200'
                }`}
              >
                <div className="flex items-center justify-between text-[11px] opacity-80">
                  <span>Hari {idx + 1}</span>
                  <span>{hasSales ? 'Terisi' : 'Kosong'}</span>
                </div>
                <div className="font-mono-num font-bold text-xs mt-1 truncate">
                  {day.tanggal}
                </div>
                <div
                  className={`font-mono-num text-[11px] mt-1 truncate ${
                    isSelected ? 'text-slate-300' : 'text-slate-500'
                  }`}
                >
                  Rp {formatNumberId(m.totalSales) || '0'}
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* Main Input Grid: Left (Shift Pagi & Siang Forms) + Right (Live Daily Summary & Reasons) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left 7 Cols: Shift Pagi & Shift Siang Cards */}
        <div className="lg:col-span-7 space-y-6">
          {/* Date Header & Prev/Next Day Navigation */}
          <div className="bg-white border border-slate-200 rounded-xl px-4 sm:px-5 py-3.5 flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2.5">
              <div className="flex items-center gap-2 bg-indigo-50 border border-indigo-200 rounded-lg px-2.5 py-1">
                <Calendar className="w-4 h-4 text-indigo-600 shrink-0" />
                <label htmlFor="form-day-cal" className="text-xs font-bold text-indigo-950 shrink-0">
                  Tanggal:
                </label>
                <input
                  id="form-day-cal"
                  type="date"
                  value={slashDateToIso(currentDay.tanggal)}
                  onChange={(e) => handleSelectCalendarDate(e.target.value)}
                  className="px-2 py-0.5 text-xs font-bold bg-white border border-indigo-300 rounded text-indigo-950 focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                  title="Pilih tanggal kalender untuk berpindah hari & weekly"
                />
              </div>

              <div className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-xs font-bold text-slate-800 flex items-center gap-1.5 font-mono-num">
                <span>{currentDay.tanggal}</span>
                <span className="text-slate-400">·</span>
                <span className="text-indigo-700 font-semibold">{formatIndoDateLong(currentDay.tanggal)}</span>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() =>
                  updateCurrentDay((d) => ({
                    ...d,
                    pagi: { nama: '', sales: null, qty: null, trx: null, traffic: null },
                    siang: { nama: '', sales: null, qty: null, trx: null, traffic: null },
                    reasonPagi: '',
                    reasonSiang: '',
                  }))
                }
                title="Reset / Kosongkan Data Tanggal Ini"
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-md transition-colors"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                Reset Hari Ini
              </button>
              <button
                type="button"
                disabled={selectedDayIndex === 0}
                onClick={() => setSelectedDayIndex((i) => Math.max(0, i - 1))}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 rounded-md transition-colors"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
                Sebelumnya
              </button>
              <button
                type="button"
                disabled={selectedDayIndex === currentTable.days.length - 1}
                onClick={() =>
                  setSelectedDayIndex((i) =>
                    Math.min(currentTable.days.length - 1, i + 1)
                  )
                }
                className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 rounded-md transition-colors"
              >
                Berikutnya
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          {/* Shift Pagi & Shift Siang Side-by-Side */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {/* SHIFT PAGI */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <Sun className="w-4 h-4 text-amber-500" />
                  <h3 className="text-sm font-bold text-slate-900">Shift Pagi</h3>
                </div>
                <span className="text-xs text-slate-500">10.00 – 16.00</span>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Nama Petugas Shift Pagi
                </label>
                <input
                  type="text"
                  value={currentDay.pagi.nama}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({
                      ...d,
                      pagi: { ...d.pagi, nama: e.target.value },
                    }))
                  }
                  placeholder="Ketik nama petugas..."
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Sales Shift Pagi (Rp)
                </label>
                <input
                  type="number"
                  value={currentDay.pagi.sales ?? ''}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({
                      ...d,
                      pagi: {
                        ...d.pagi,
                        sales: e.target.value === '' ? null : Number(e.target.value),
                      },
                    }))
                  }
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Qty (Pcs)
                  </label>
                  <input
                    type="number"
                    value={currentDay.pagi.qty ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        pagi: {
                          ...d.pagi,
                          qty: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Trx (Struk)
                  </label>
                  <input
                    type="number"
                    value={currentDay.pagi.trx ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        pagi: {
                          ...d.pagi,
                          trx: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-indigo-800 mb-1.5">
                    Traffic (Org)
                  </label>
                  <input
                    type="number"
                    value={currentDay.pagi.traffic ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        pagi: {
                          ...d.pagi,
                          traffic: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num bg-indigo-50/40 border border-indigo-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>
            </div>

            {/* SHIFT SIANG */}
            <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <Sunset className="w-4 h-4 text-indigo-600" />
                  <h3 className="text-sm font-bold text-slate-900">Shift Siang</h3>
                </div>
                <span className="text-xs text-slate-500">15.00 – 21.30</span>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Nama Petugas Shift Siang
                </label>
                <input
                  type="text"
                  value={currentDay.siang.nama}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({
                      ...d,
                      siang: { ...d.siang, nama: e.target.value },
                    }))
                  }
                  placeholder="Ketik nama petugas..."
                  className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Sales Shift Siang (Rp)
                </label>
                <input
                  type="number"
                  value={currentDay.siang.sales ?? ''}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({
                      ...d,
                      siang: {
                        ...d.siang,
                        sales: e.target.value === '' ? null : Number(e.target.value),
                      },
                    }))
                  }
                  placeholder="0"
                  className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Qty (Pcs)
                  </label>
                  <input
                    type="number"
                    value={currentDay.siang.qty ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        siang: {
                          ...d.siang,
                          qty: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Trx (Struk)
                  </label>
                  <input
                    type="number"
                    value={currentDay.siang.trx ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        siang: {
                          ...d.siang,
                          trx: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-indigo-800 mb-1.5">
                    Traffic (Org)
                  </label>
                  <input
                    type="number"
                    value={currentDay.siang.traffic ?? ''}
                    onChange={(e) =>
                      updateCurrentDay((d) => ({
                        ...d,
                        siang: {
                          ...d.siang,
                          traffic: e.target.value === '' ? null : Number(e.target.value),
                        },
                      }))
                    }
                    placeholder="0"
                    className="w-full px-3 py-2 text-sm font-mono-num bg-indigo-50/40 border border-indigo-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Catatan Reason / Event Shift Pagi & Siang */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
              <MessageSquareText className="w-4 h-4 text-slate-600" />
              <h3 className="text-sm font-bold text-slate-900">
                Reason / Event Evaluasi Harian ({currentDay.tanggal})
              </h3>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Reason / Event Shift Pagi
                </label>
                <textarea
                  rows={4}
                  value={currentDay.reasonPagi}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({ ...d, reasonPagi: e.target.value }))
                  }
                  placeholder="Tulis kondisi traffic, barang dominan laku, atau kendala POP pada shift pagi..."
                  className="w-full p-3 text-xs leading-relaxed border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Reason / Event Shift Siang
                </label>
                <textarea
                  rows={4}
                  value={currentDay.reasonSiang}
                  onChange={(e) =>
                    updateCurrentDay((d) => ({ ...d, reasonSiang: e.target.value }))
                  }
                  placeholder="Tulis kondisi traffic sore/malam, event sekitar bazar, atau kendala broken size..."
                  className="w-full p-3 text-xs leading-relaxed border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Right 5 Cols: Auto-Calculated Daily Formula Card */}
        <div className="lg:col-span-5 space-y-5">
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4 sticky top-20">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <Calculator className="w-4 h-4 text-indigo-600" />
                <h3 className="text-sm font-bold text-slate-900">
                  Hasil Kalkulasi Otomatis ({currentDay.tanggal})
                </h3>
              </div>
              <span className="text-xs font-mono-num text-slate-500">
                {currentTable.title}
              </span>
            </div>

            <div className="space-y-3">
              <div className="p-3.5 bg-amber-50/70 border border-amber-200 rounded-lg flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold text-slate-900">Total Sales Harian</div>
                  <div className="text-[11px] text-slate-500">Sales Pagi + Sales Siang</div>
                </div>
                <div className="text-lg font-bold font-mono-num text-slate-950">
                  Rp {formatNumberId(dayMetrics.totalSales) || '0'}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg">
                  <div className="text-xs font-semibold text-slate-700">Total Qty</div>
                  <div className="text-base font-bold font-mono-num text-slate-900 mt-1">
                    {formatNumberId(dayMetrics.totalQty) || '0'} pcs
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    Pagi {currentDay.pagi.qty ?? 0} + Siang {currentDay.siang.qty ?? 0}
                  </div>
                </div>

                <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg">
                  <div className="text-xs font-semibold text-slate-700">Total Trx</div>
                  <div className="text-base font-bold font-mono-num text-slate-900 mt-1">
                    {formatNumberId(dayMetrics.totalTrx) || '0'} trx
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    Pagi {currentDay.pagi.trx ?? 0} + Siang {currentDay.siang.trx ?? 0}
                  </div>
                </div>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold text-slate-900">Price Point</div>
                  <div className="text-[11px] text-slate-500">Total Sales ÷ Total Qty</div>
                </div>
                <div className="text-base font-bold font-mono-num text-slate-900">
                  {dayMetrics.pricePoint !== null
                    ? `Rp ${formatNumberId(dayMetrics.pricePoint)}`
                    : '#DIV/0! (0)'}
                </div>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between">
                <div>
                  <div className="text-xs font-bold text-slate-900">Big Size (Basket Size)</div>
                  <div className="text-[11px] text-slate-500">Total Sales ÷ Total Trx</div>
                </div>
                <div className="text-base font-bold font-mono-num text-slate-900">
                  {dayMetrics.bigSize !== null
                    ? `Rp ${formatNumberId(dayMetrics.bigSize)}`
                    : '#DIV/0! (0)'}
                </div>
              </div>

              {(() => {
                const pTraf = currentDay.pagi.traffic ?? 0;
                const sTraf = currentDay.siang.traffic ?? 0;
                const totTraf = pTraf + sTraf;
                const dpConv = totTraf > 0 ? (dayMetrics.totalTrx / totTraf) * 100 : null;
                return (
                  <div className="p-3.5 bg-indigo-50/70 border border-indigo-200 rounded-lg flex items-center justify-between">
                    <div>
                      <div className="text-xs font-bold text-indigo-950">Traffic &amp; Conversion Rate (DP)</div>
                      <div className="text-[11px] text-indigo-700">
                        Pagi {formatNumberId(pTraf) || '0'} + Siang {formatNumberId(sTraf) || '0'} = {formatNumberId(totTraf) || '0'} org
                      </div>
                    </div>
                    <div className="text-base font-bold font-mono-num text-indigo-900">
                      {dpConv !== null ? `${dpConv.toFixed(1)}%` : '—'}
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
