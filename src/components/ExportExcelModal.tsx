import React, { useState, useEffect, useMemo } from 'react';
import {
  X,
  Download,
  FileSpreadsheet,
  CheckSquare,
  Square,
  Sparkles,
  Calendar,
  CalendarRange,
  CalendarDays,
  MessageSquareText,
  ClipboardList,
  Layers,
  Tag,
  Percent,
} from 'lucide-react';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  calculateWeekMetrics,
  formatNumberId,
  getWeeklyPeriodeText,
  getTableMonthKey,
  getReviewDocWeekNumber,
  getReviewDocWeeklyLabel,
  ANNUAL_2026_WEEKS,
} from '../data/initialData';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  MONTH_NAMES_ID,
  loadCvDiv1Records,
  loadCvInventoryOps,
} from '../data/cvDiv1Data';
import {
  ExportExcelOptions,
  DEFAULT_EXPORT_EXCEL_OPTIONS,
  exportWeeklyTablesToWPSExcel,
  resolveLatestActiveMonthKey,
  resolveLatestActiveWeekNumber,
  isWeeklyTableActive,
  hasReviewDocContent,
} from '../utils/exportExcel';

interface ExportExcelModalProps {
  isOpen: boolean;
  onClose: () => void;
  weeklyTables: WeeklySalesTable[];
  reviewDoc: ReviewMonitoringDoc;
  reviewDocs?: ReviewMonitoringDoc[];
  selectedWeeklyFilter: string;
  showDivZero: boolean;
  cvRecords?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  onExportSuccess?: (msg: string) => void;
}

const MONTH_OPTIONS = [
  { key: 'ALL', label: 'Semua Bulan (Januari – Desember 2026)' },
  { key: '01', label: 'Januari 2026' },
  { key: '02', label: 'Februari 2026' },
  { key: '03', label: 'Maret 2026' },
  { key: '04', label: 'April 2026' },
  { key: '05', label: 'Mei 2026' },
  { key: '06', label: 'Juni 2026' },
  { key: '07', label: 'Juli 2026' },
  { key: '08', label: 'Agustus 2026' },
  { key: '09', label: 'September 2026' },
  { key: '10', label: 'Oktober 2026' },
  { key: '11', label: 'November 2026' },
  { key: '12', label: 'Desember 2026' },
];

export const ExportExcelModal: React.FC<ExportExcelModalProps> = ({
  isOpen,
  onClose,
  weeklyTables,
  reviewDoc,
  reviewDocs = [],
  selectedWeeklyFilter,
  showDivZero,
  cvRecords,
  cvInventoryOps,
  onExportSuccess,
}) => {
  const effectiveCvRecords = useMemo(
    () => (cvRecords && cvRecords.length > 0 ? cvRecords : loadCvDiv1Records()),
    [cvRecords]
  );
  const effectiveCvOps = useMemo(
    () =>
      cvInventoryOps && cvInventoryOps.length > 0
        ? cvInventoryOps
        : loadCvInventoryOps(),
    [cvInventoryOps]
  );

  const latestActiveInfo = useMemo(() => {
    const latestMonthKey = resolveLatestActiveMonthKey(
      'ALL',
      weeklyTables,
      effectiveCvRecords,
      effectiveCvOps
    );
    const latestMonthName = MONTH_NAMES_ID[latestMonthKey] || 'November';
    const activeTables = [...weeklyTables]
      .filter((t) => isWeeklyTableActive(t))
      .sort((a, b) => a.weekNumber - b.weekNumber);
    const latestWeekNumber = resolveLatestActiveWeekNumber(
      undefined,
      weeklyTables,
      effectiveCvRecords,
      effectiveCvOps,
      latestMonthKey
    );
    const latestThreeLabels =
      activeTables.length > 0
        ? activeTables
            .slice(-3)
            .map((t) => `W${t.weekNumber}`)
            .join(', ')
        : `W${latestWeekNumber}`;

    const filledReviews = [...(reviewDocs || [])]
      .filter((d) => hasReviewDocContent(d))
      .sort(
        (a, b) =>
          (getReviewDocWeekNumber(a) || 0) - (getReviewDocWeekNumber(b) || 0)
      );
    const latestReview =
      filledReviews[filledReviews.length - 1] ||
      reviewDocs.find((d) => getReviewDocWeekNumber(d) === latestWeekNumber) ||
      reviewDoc;

    return {
      latestMonthKey,
      latestMonthName,
      latestWeekNumber,
      latestWeeksSummary: latestThreeLabels,
      latestReview,
    };
  }, [weeklyTables, effectiveCvRecords, effectiveCvOps, reviewDocs, reviewDoc]);

  // Helper: get default week IDs (empty array = all active latest weeks)
  const getDefaultWeekIds = (): string[] => {
    return [];
  };

  const [options, setOptions] = useState<ExportExcelOptions>({
    ...DEFAULT_EXPORT_EXCEL_OPTIONS,
    includeExecutiveCover: false,
    includeCvReport: true,
    includeDataCv: true,
    includeTotalConsolidated: true,
    compactMode: true,
    selectedWeekIds: getDefaultWeekIds(),
    selectedReviewWeekFilter: 'FOLLOW_WEEKLY',
  });
  const [periodMode, setPeriodMode] = useState<'ALL' | 'BY_MONTH' | 'CUSTOM_WEEKS'>('ALL');
  const [isExporting, setIsExporting] = useState(false);

  useEffect(() => {
    if (isOpen) {
      setPeriodMode('ALL');
      setOptions((prev) => ({
        ...prev,
        selectedWeekIds: [],
        selectedMonthKey: latestActiveInfo.latestMonthKey,
      }));
    }
  }, [isOpen, latestActiveInfo.latestMonthKey]);

  if (!isOpen) return null;

  const toggleSection = (key: keyof Pick<
    ExportExcelOptions,
    | 'includeWeeklyMatrix'
    | 'includeDailyTable'
    | 'includeMonthlyTable'
    | 'includeReasonEvaluation'
    | 'includeStaffAndReview'
    | 'includeChartDiagrams'
    | 'includeCvReport'
    | 'includeDataCv'
    | 'includeTotalConsolidated'
    | 'includeExecutiveCover'
  >) => {
    setOptions((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  };

  const selectAllSections = () => {
    setOptions((prev) => ({
      ...prev,
      includeExecutiveCover: !(prev.compactMode ?? true),
      includeWeeklyMatrix: true,
      includeDailyTable: true,
      includeMonthlyTable: true,
      includeReasonEvaluation: true,
      includeStaffAndReview: true,
      includeChartDiagrams: true,
      includeCvReport: true,
      includeDataCv: true,
      includeTotalConsolidated: true,
    }));
  };

  const selectOnlySalesTables = () => {
    setOptions((prev) => ({
      ...prev,
      includeExecutiveCover: false,
      includeWeeklyMatrix: true,
      includeDailyTable: true,
      includeMonthlyTable: true,
      includeReasonEvaluation: false,
      includeStaffAndReview: false,
      includeCvReport: false,
      includeDataCv: false,
      includeTotalConsolidated: false,
    }));
  };

  const toggleWeekId = (id: string) => {
    setOptions((prev) => {
      const current = prev.selectedWeekIds || [];
      const exists = current.includes(id);
      const next = exists ? current.filter((x) => x !== id) : [...current, id];
      return {
        ...prev,
        selectedWeekIds: next,
      };
    });
  };

  const resolveExportWeekIds = (): string[] => {
    if (periodMode === 'ALL') {
      return [];
    }
    if (periodMode === 'BY_MONTH') {
      return weeklyTables
        .filter((t) =>
          options.selectedMonthKey === 'ALL'
            ? true
            : getTableMonthKey(t) === options.selectedMonthKey
        )
        .map((t) => t.id);
    }
    return options.selectedWeekIds && options.selectedWeekIds.length > 0
      ? options.selectedWeekIds
      : [];
  };

  const isCompactMode = options.compactMode !== false;

  const handleInstantExportAllInOne = async () => {
    setIsExporting(true);
    try {
      const monthArg =
        periodMode === 'BY_MONTH' ? options.selectedMonthKey : 'ALL';
      await exportWeeklyTablesToWPSExcel(weeklyTables, reviewDoc, showDivZero, {
        includeExecutiveCover: !isCompactMode,
        includeWeeklyMatrix: true,
        includeDailyTable: true,
        includeMonthlyTable: true,
        includeReasonEvaluation: true,
        includeStaffAndReview: true,
        includeChartDiagrams: options.includeChartDiagrams !== false,
        includeCvReport: true,
        includeDataCv: true,
        includeTotalConsolidated: true,
        compactMode: isCompactMode,
        cvRecords: effectiveCvRecords,
        cvInventoryOps: effectiveCvOps,
        activeCvWeekNumber:
          periodMode === 'ALL' ? latestActiveInfo.latestWeekNumber : undefined,
        sheetLayout: 'SEPARATE_SHEETS',
        selectedWeekIds: resolveExportWeekIds(),
        selectedMonthKey: monthArg,
        allReviewDocs: reviewDocs,
        selectedReviewWeekFilter: options.selectedReviewWeekFilter || 'FOLLOW_WEEKLY',
      });
      onExportSuccess?.(
        isCompactMode
          ? `Workbook Ringkas 5 Tab (Data Terbaru: ${latestActiveInfo.latestMonthName} 2026) berhasil di-export ke Excel (.xlsx)!`
          : `Master Workbook Detail Lengkap (${latestActiveInfo.latestMonthName} 2026) berhasil di-export ke Excel (.xlsx)!`
      );
      onClose();
    } finally {
      setIsExporting(false);
    }
  };

  const handleQuickExportThreeSalesSheets = async () => {
    setIsExporting(true);
    try {
      const finalWeekIds = resolveExportWeekIds();
      const monthArg =
        periodMode === 'BY_MONTH' ? options.selectedMonthKey : 'ALL';

      await exportWeeklyTablesToWPSExcel(weeklyTables, reviewDoc, showDivZero, {
        ...options,
        includeExecutiveCover: false,
        includeDailyTable: true,
        includeWeeklyMatrix: true,
        includeMonthlyTable: true,
        includeReasonEvaluation: false,
        includeStaffAndReview: false,
        includeCvReport: false,
        includeDataCv: false,
        includeTotalConsolidated: false,
        compactMode: false,
        cvRecords: effectiveCvRecords,
        cvInventoryOps: effectiveCvOps,
        activeCvWeekNumber:
          periodMode === 'ALL' ? latestActiveInfo.latestWeekNumber : undefined,
        sheetLayout: 'SEPARATE_SHEETS',
        selectedWeekIds: finalWeekIds,
        selectedMonthKey: monthArg,
        allReviewDocs: reviewDocs,
        selectedReviewWeekFilter: options.selectedReviewWeekFilter || 'FOLLOW_WEEKLY',
      });
      onExportSuccess?.(
        `3 Sheet Sales Report DP (${latestActiveInfo.latestMonthName} 2026: Daily, Weekly ${latestActiveInfo.latestWeeksSummary}, Monthly) berhasil di-export!`
      );
      onClose();
    } finally {
      setIsExporting(false);
    }
  };

  const handleCustomExport = async () => {
    setIsExporting(true);
    try {
      const finalWeekIds = resolveExportWeekIds();
      const monthArg =
        periodMode === 'BY_MONTH' ? options.selectedMonthKey : 'ALL';

      await exportWeeklyTablesToWPSExcel(weeklyTables, reviewDoc, showDivZero, {
        ...options,
        compactMode: isCompactMode,
        cvRecords: effectiveCvRecords,
        cvInventoryOps: effectiveCvOps,
        activeCvWeekNumber:
          periodMode === 'ALL' ? latestActiveInfo.latestWeekNumber : undefined,
        selectedWeekIds: finalWeekIds,
        selectedMonthKey: monthArg,
        allReviewDocs: reviewDocs,
        selectedReviewWeekFilter: options.selectedReviewWeekFilter || 'FOLLOW_WEEKLY',
      });
      onExportSuccess?.(
        isCompactMode
          ? 'File Excel Ringkas (.xlsx) pilihan berhasil di-export!'
          : 'File Excel Detail (.xlsx) pilihan berhasil di-export!'
      );
      onClose();
    } finally {
      setIsExporting(false);
    }
  };

  const selectedSectionCount = [
    options.includeWeeklyMatrix,
    options.includeDailyTable,
    options.includeMonthlyTable,
    options.includeReasonEvaluation,
    options.includeStaffAndReview,
    Boolean(options.includeCvReport),
    Boolean(options.includeDataCv),
    Boolean(options.includeTotalConsolidated),
  ].filter(Boolean).length;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-5 bg-slate-950/70 backdrop-blur-xs">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-2xl w-full max-w-3xl max-h-[92vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-5 py-4 bg-gradient-to-r from-slate-900 via-slate-800 to-emerald-950 text-white flex items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-xl bg-emerald-500 text-slate-950 flex items-center justify-center shrink-0 shadow-xs">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h2 className="text-base sm:text-lg font-extrabold tracking-tight text-white truncate">
                Export WPS / Excel (.xlsx) Terformula Otomatis
              </h2>
              <p className="text-xs text-emerald-200 truncate">
                Pilih bagian yang ingin di-export atau gabungkan semua dalam 1 file Excel (.xlsx)
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-2 rounded-lg bg-white/10 hover:bg-white/20 text-slate-200 hover:text-white transition-colors"
            aria-label="Tutup Modal Export"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-4 sm:p-5 overflow-y-auto space-y-5 flex-1">
          {/* 1-CLICK ALL-IN-ONE EXPORT BANNER */}
          <div className="bg-gradient-to-r from-emerald-50 via-teal-50 to-amber-50 border-2 border-emerald-500/40 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3.5 shadow-2xs">
            <div className="space-y-1">
              <div className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md bg-emerald-600 text-white text-[11px] font-extrabold tracking-wide uppercase">
                <Sparkles className="w-3.5 h-3.5" />
                {isCompactMode ? 'Rekomendasi 1-Klik · Format Ringkas 5 Tab' : 'Rekomendasi 1-Klik · Format Detail Lengkap'}
              </div>
              <div className="text-sm sm:text-base font-extrabold text-slate-900">
                {isCompactMode
                  ? 'Export Ringkas 5 Tab Utama dalam 1 File Excel (.xlsx)'
                  : 'Export Semua Modul Detail dalam 1 File Excel (.xlsx)'}
              </div>
              <p className="text-xs text-slate-600 leading-relaxed">
                {isCompactMode ? (
                  <>
                    Langsung menghasilkan 5 tab ringkas &amp; rapi:{' '}
                    <strong>Report DP</strong>, <strong>Report CV</strong>,{' '}
                    <strong>Total DP+CV</strong>, <strong>Monitoring</strong>, dan{' '}
                    <strong>Data CV</strong> dengan rumus otomatis{' '}
                    <code className="px-1 py-0.5 bg-white rounded border border-slate-200 text-emerald-800 font-semibold">
                      =SUM
                    </code>
                    .
                  </>
                ) : (
                  <>
                    Menghasilkan semua sub-tab terpisah (Ringkasan Eksekutif, Daily, Weekly, Monthly, Evaluasi, Review, sub-tab CV, dan sub-tab Total DP+CV).
                  </>
                )}
              </p>
            </div>

            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 shrink-0">
              <button
                type="button"
                disabled={isExporting}
                onClick={handleQuickExportThreeSalesSheets}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-slate-950 text-xs font-extrabold shadow-xs transition-all cursor-pointer disabled:opacity-50"
              >
                <FileSpreadsheet className="w-4 h-4 shrink-0" />
                <span>Hanya Report DP (3 Sub-Tab)</span>
              </button>
              <button
                type="button"
                disabled={isExporting}
                onClick={handleInstantExportAllInOne}
                className="inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-extrabold shadow-md transition-all cursor-pointer disabled:opacity-50"
              >
                <Download className="w-4 h-4 shrink-0" />
                <span>
                  {isExporting
                    ? 'Memproses...'
                    : isCompactMode
                      ? 'Export Semua (5 Tab)'
                      : 'Export Semua (Detail)'}
                </span>
              </button>
            </div>
          </div>

          {/* SECTION 1: PILIH SHEET YANG MAU DI-EXPORT */}
          <div className="space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <label className="text-xs sm:text-sm font-extrabold text-slate-900 uppercase tracking-wide">
                1. Pilih Modul yang Mau Di-Export ({selectedSectionCount} dari 8 bagian dipilih)
              </label>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onClick={selectAllSections}
                  className="px-2.5 py-1 text-xs font-bold rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-colors"
                >
                  Pilih Semua Sheet
                </button>
                <button
                  type="button"
                  onClick={selectOnlySalesTables}
                  className="px-2.5 py-1 text-xs font-bold rounded-lg bg-amber-50 text-amber-900 hover:bg-amber-100 border border-amber-200 transition-colors"
                >
                  Hanya Sheet Sales Daily, Weekly &amp; Monthly
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              {/* Item 1: Sheet Sales Daily (Harian) */}
              <button
                type="button"
                onClick={() => toggleSection('includeDailyTable')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeDailyTable
                    ? 'bg-amber-50/70 border-amber-400 ring-1 ring-amber-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeDailyTable ? (
                  <CheckSquare className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <Calendar className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                    <span>Sheet Sales Daily (Harian)</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Tabel rincian harian tanggal demi tanggal (Pagi, Siang, Total Harian, Basket Size, APC).
                  </p>
                </div>
              </button>

              {/* Item 2: Sheet Sales Weekly (WPS) */}
              <button
                type="button"
                onClick={() => toggleSection('includeWeeklyMatrix')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeWeeklyMatrix
                    ? 'bg-rose-50/70 border-rose-400 ring-1 ring-rose-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeWeeklyMatrix ? (
                  <CheckSquare className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <CalendarRange className="w-3.5 h-3.5 text-rose-600 shrink-0" />
                    <span>Sheet Sales Weekly (Maks. 3 Weekly)</span>
                    <span className="text-[10px] bg-rose-100 text-rose-700 font-bold px-1.5 py-0.5 rounded">
                      Maks 3 Weekly
                    </span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Matriks mingguan dibatasi maksimal 3 weekly untuk tampilan rapi dan pas dicetak.
                  </p>
                </div>
              </button>

              {/* Item 3: Sheet Sales Monthly (Bulanan) */}
              <button
                type="button"
                onClick={() => toggleSection('includeMonthlyTable')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeMonthlyTable
                    ? 'bg-sky-50/70 border-sky-400 ring-1 ring-sky-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeMonthlyTable ? (
                  <CheckSquare className="w-5 h-5 text-sky-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <CalendarDays className="w-3.5 h-3.5 text-sky-600 shrink-0" />
                    <span>Sheet Sales Monthly (Bulanan &amp; 12 Bulan)</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Rekap matriks per Weekly, persentase kontribusi (%), Weekday vs Weekend &amp; rekap 12 bulan.
                  </p>
                </div>
              </button>

              {/* Item 4: Evaluasi Reason / Event */}
              <button
                type="button"
                onClick={() => toggleSection('includeReasonEvaluation')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeReasonEvaluation
                    ? 'bg-orange-50/70 border-orange-400 ring-1 ring-orange-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeReasonEvaluation ? (
                  <CheckSquare className="w-5 h-5 text-orange-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <MessageSquareText className="w-3.5 h-3.5 text-orange-600 shrink-0" />
                    <span>Evaluasi Reason / Event (Pagi & Siang)</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Catatan evaluasi kondisi bazar Shift Pagi dan Shift Siang setiap hari.
                  </p>
                </div>
              </button>

              {/* Item 5: Review Monitoring Bazar */}
              <button
                type="button"
                onClick={() => toggleSection('includeStaffAndReview')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeStaffAndReview
                    ? 'bg-indigo-50/70 border-indigo-400 ring-1 ring-indigo-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeStaffAndReview ? (
                  <CheckSquare className="w-5 h-5 text-indigo-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <ClipboardList className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
                    <span>Review Monitoring Bazar</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Laporan Merchandise (Divisi 1–4, Fast/Slow Moving), Promo (Divisi 1–4), Action Plan &amp; Need Support.
                  </p>
                </div>
              </button>

              {/* Item 6: Grafik Diagram Visual di dalam Excel */}
              <button
                type="button"
                onClick={() => toggleSection('includeChartDiagrams')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeChartDiagrams !== false
                    ? 'bg-emerald-50/70 border-emerald-400 ring-1 ring-emerald-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeChartDiagrams !== false ? (
                  <CheckSquare className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <Sparkles className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                    <span>Sertakan Grafik Diagram Batang di Excel</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Menampilkan Diagram Batang, Tren, Growth &amp; Kontribusi di bawah tabel <strong>Report DP</strong>, <strong>Report CV</strong>, dan <strong>Total DP+CV (+Traffic)</strong>.
                  </p>
                </div>
              </button>

              {/* Item 7: Report CV Divisi 1 */}
              <button
                type="button"
                onClick={() => toggleSection('includeCvReport')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeCvReport
                    ? 'bg-purple-50/70 border-purple-400 ring-1 ring-purple-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeCvReport ? (
                  <CheckSquare className="w-5 h-5 text-purple-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <Tag className="w-3.5 h-3.5 text-purple-600 shrink-0" />
                    <span>Report CV Div. 1 (Sales Harian, Weekly &amp; 12 Bln)</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Graphis (6201) &amp; Simplicity (6244): Rekap Nota CV, Qty &amp; Sales Harian, Per Week, dan Kalender 12 Bulan.
                  </p>
                </div>
              </button>

              {/* Item 8: Data CV (Inventory, Promo, Reason & Action Plan) */}
              <button
                type="button"
                onClick={() => toggleSection('includeDataCv')}
                className={`flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeDataCv
                    ? 'bg-fuchsia-50/70 border-fuchsia-400 ring-1 ring-fuchsia-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeDataCv ? (
                  <CheckSquare className="w-5 h-5 text-fuchsia-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <ClipboardList className="w-3.5 h-3.5 text-fuchsia-600 shrink-0" />
                    <span>Data CV (Inventory, Promo &amp; Action Plan)</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Per Monthly &amp; Per Week: Stock Awal/Datang/Sales Qty/Keluar/Stock Akhir (=E+F-G-H), Promo, Reason &amp; Action Plan.
                  </p>
                </div>
              </button>

              {/* Item 9: Total Konsolidasi DP+CV + Traffic & Conversion Rate */}
              <button
                type="button"
                onClick={() => toggleSection('includeTotalConsolidated')}
                className={`sm:col-span-2 flex items-start gap-3 p-3 rounded-xl border text-left transition-all ${
                  options.includeTotalConsolidated
                    ? 'bg-teal-50/70 border-teal-400 ring-1 ring-teal-400/40'
                    : 'bg-slate-50/60 border-slate-200 opacity-75 hover:opacity-100'
                }`}
              >
                {options.includeTotalConsolidated ? (
                  <CheckSquare className="w-5 h-5 text-teal-600 shrink-0 mt-0.5" />
                ) : (
                  <Square className="w-5 h-5 text-slate-400 shrink-0 mt-0.5" />
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs sm:text-sm font-extrabold text-slate-900">
                    <Percent className="w-3.5 h-3.5 text-teal-600 shrink-0" />
                    <span>Total DP+CV &amp; Traffic Kasir / Conversion Rate</span>
                  </div>
                  <p className="text-xs text-slate-600 mt-0.5">
                    Konsolidasi Harian, Weekly &amp; 12 Bulan (Trx DP + Nota CV = Trx+Nota), Traffic Kasir Pagi/Siang &amp; Target Benchmark.
                  </p>
                </div>
              </button>
            </div>
          </div>

          {/* SECTION 2: FORMAT TAB SHEET */}
          <div className="space-y-2">
            <label className="text-xs sm:text-sm font-extrabold text-slate-900 uppercase tracking-wide flex items-center gap-1.5">
              <Layers className="w-4 h-4 text-indigo-600" />
              2. Format Ringkas vs Detail Tab Sheet (.xlsx)
            </label>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
              <button
                type="button"
                onClick={() =>
                  setOptions((prev) => ({
                    ...prev,
                    compactMode: true,
                    includeExecutiveCover: false,
                  }))
                }
                className={`p-3.5 rounded-xl border text-left transition-all ${
                  isCompactMode
                    ? 'border-emerald-500 bg-emerald-50/90 ring-2 ring-emerald-500/30'
                    : 'border-slate-200 bg-white hover:bg-slate-50'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs sm:text-sm font-extrabold text-emerald-950">
                    ⚡ Format Ringkas (5 Tab Utama)
                  </span>
                  <span className="text-[10px] bg-emerald-600 text-white font-bold px-2 py-0.5 rounded-full">
                    Direkomendasikan
                  </span>
                </div>
                <p className="text-xs text-slate-700 mt-1.5 leading-relaxed">
                  Hanya 5 tab utama yang bersih: <strong>Report DP</strong>,{' '}
                  <strong>Report CV</strong>, <strong>Total DP+CV</strong>,{' '}
                  <strong>Monitoring</strong>, dan <strong>Data CV</strong>. Praktis dibuka &amp; siap cetak.
                </p>
              </button>

              <button
                type="button"
                onClick={() =>
                  setOptions((prev) => ({
                    ...prev,
                    compactMode: false,
                    includeExecutiveCover: true,
                  }))
                }
                className={`p-3.5 rounded-xl border text-left transition-all ${
                  !isCompactMode
                    ? 'border-indigo-500 bg-indigo-50/90 ring-2 ring-indigo-500/30'
                    : 'border-slate-200 bg-white hover:bg-slate-50'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs sm:text-sm font-extrabold text-slate-900">
                    📑 Format Detail (Semua Sub-Tab)
                  </span>
                  <span className="text-[10px] bg-slate-200 text-slate-800 font-bold px-2 py-0.5 rounded-full">
                    14+ Tab Terpisah
                  </span>
                </div>
                <p className="text-xs text-slate-600 mt-1.5 leading-relaxed">
                  Memisahkan setiap sub-bagian (Ringkasan Eksekutif, Daily, Weekly, Monthly, Evaluasi, Review, CV Harian/Weekly/Bulanan, dll).
                </p>
              </button>
            </div>
          </div>

          {/* SECTION 3: FILTER PERIODE DATA YANG DI-EXPORT */}
          <div className="space-y-2.5 bg-slate-50 border border-slate-200 rounded-xl p-3.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="space-y-0.5">
                <label className="text-xs sm:text-sm font-extrabold text-slate-900 uppercase tracking-wide block">
                  3. Pilih Cakupan Periode Data
                </label>
                <div className="text-[11px] font-bold text-emerald-700">
                  Data Terbaru Aktif: {latestActiveInfo.latestMonthName} 2026 · Weekly {latestActiveInfo.latestWeeksSummary}
                </div>
              </div>

              <div className="inline-flex items-center bg-white p-0.5 rounded-lg border border-slate-200">
                <button
                  type="button"
                  onClick={() => setPeriodMode('ALL')}
                  className={`px-2.5 py-1 rounded-md text-xs font-bold transition-colors ${
                    periodMode === 'ALL'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Data Terbaru &amp; Aktif ({weeklyTables.length})
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setPeriodMode('BY_MONTH');
                    if (!options.selectedMonthKey || options.selectedMonthKey === 'ALL') {
                      setOptions((prev) => ({
                        ...prev,
                        selectedMonthKey: latestActiveInfo.latestMonthKey,
                      }));
                    }
                  }}
                  className={`px-2.5 py-1 rounded-md text-xs font-bold transition-colors ${
                    periodMode === 'BY_MONTH'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Per Bulan
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setPeriodMode('CUSTOM_WEEKS');
                    if (
                      selectedWeeklyFilter !== 'ALL' &&
                      weeklyTables.some((t) => t.id === selectedWeeklyFilter)
                    ) {
                      setOptions((prev) => ({
                        ...prev,
                        selectedWeekIds: [selectedWeeklyFilter],
                      }));
                    }
                  }}
                  className={`px-2.5 py-1 rounded-md text-xs font-bold transition-colors ${
                    periodMode === 'CUSTOM_WEEKS'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Pilih Weekly Tertentu
                </button>
              </div>
            </div>

            {/* DROPDOWN LANGSUNG: FILTER TOTAL WEEKLY 1 SAMPAI 53 */}
            <div className="pt-1 space-y-1.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <label className="text-xs font-extrabold text-amber-950">
                  Filter Pilihan Total Weekly (Total Weekly 1 s/d 53):
                </label>
                <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                  Sel &quot;TOTAL WEEKLY&quot; di Excel juga bisa diklik pilih Total Weekly 1–53
                </span>
              </div>
              <select
                aria-label="Filter Pilihan Total Weekly 1 sampai 53 untuk Export Excel"
                value={
                  periodMode === 'CUSTOM_WEEKS' &&
                  (options.selectedWeekIds || []).length === 1
                    ? options.selectedWeekIds![0]
                    : 'ALL'
                }
                onChange={(e) => {
                  const val = e.target.value;
                  if (val === 'ALL') {
                    setPeriodMode('ALL');
                    setOptions((prev) => ({
                      ...prev,
                      selectedWeekIds: [],
                    }));
                  } else {
                    setPeriodMode('CUSTOM_WEEKS');
                    setOptions((prev) => ({
                      ...prev,
                      selectedWeekIds: [val],
                    }));
                  }
                }}
                className="w-full px-3 py-2 rounded-lg bg-amber-50/70 border border-amber-400 text-xs font-extrabold text-slate-900 focus:outline-none focus:border-amber-600 cursor-pointer"
              >
                <option value="ALL">
                  📋 Semua Total Weekly Aktif ({weeklyTables.length} Tabel + Database Lengkap Total Weekly 1–53 di Excel)
                </option>
                {ANNUAL_2026_WEEKS.map((w) => {
                  const existing = weeklyTables.find(
                    (t) => t.weekNumber === w.weekNumber
                  );
                  const wm = existing ? calculateWeekMetrics(existing) : null;
                  const weekKey = existing ? existing.id : `week-${w.weekNumber}`;
                  return (
                    <option key={w.weekNumber} value={weekKey}>
                      TOTAL WEEKLY {w.weekNumber} — {w.periodeText} ({w.monthName} 2026)
                      {wm && wm.weeklySales > 0
                        ? ` [Rp ${formatNumberId(wm.weeklySales)}]`
                        : existing
                          ? ' [Tersimpan]'
                          : ''}
                    </option>
                  );
                })}
              </select>
            </div>

            {periodMode === 'BY_MONTH' && (
              <div className="pt-1">
                <select
                  aria-label="Pilih Bulan Export"
                  value={options.selectedMonthKey || 'ALL'}
                  onChange={(e) =>
                    setOptions((prev) => ({
                      ...prev,
                      selectedMonthKey: e.target.value,
                    }))
                  }
                  className="w-full px-3 py-2 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-600"
                >
                  {MONTH_OPTIONS.map((m) => (
                    <option key={m.key} value={m.key}>
                      {m.label}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {periodMode === 'CUSTOM_WEEKS' && (
              <div className="space-y-2 pt-1">
                <div className="flex items-center justify-between text-xs">
                  <span className="text-slate-600 font-medium">
                    Centang Total Weekly (1 s/d 53) yang ingin dimasukkan ke Excel:
                  </span>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        setOptions((prev) => ({
                          ...prev,
                          selectedWeekIds: weeklyTables.map((t) => t.id),
                        }))
                      }
                      className="text-indigo-600 font-bold hover:underline"
                    >
                      Centang Weekly Aktif ({weeklyTables.length})
                    </button>
                    <span className="text-slate-300">|</span>
                    <button
                      type="button"
                      onClick={() =>
                        setOptions((prev) => ({
                          ...prev,
                          selectedWeekIds: ANNUAL_2026_WEEKS.map((w) => {
                            const ex = weeklyTables.find(
                              (t) => t.weekNumber === w.weekNumber
                            );
                            return ex ? ex.id : `week-${w.weekNumber}`;
                          }),
                        }))
                      }
                      className="text-emerald-700 font-bold hover:underline"
                    >
                      Centang Semua 1–53
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5 max-h-40 overflow-y-auto pr-1">
                  {ANNUAL_2026_WEEKS.map((w) => {
                    const existing = weeklyTables.find(
                      (t) => t.weekNumber === w.weekNumber
                    );
                    const wm = existing ? calculateWeekMetrics(existing) : null;
                    const itemId = existing ? existing.id : `week-${w.weekNumber}`;
                    const isChecked = (options.selectedWeekIds || []).includes(
                      itemId
                    );
                    return (
                      <button
                        key={w.weekNumber}
                        type="button"
                        onClick={() => toggleWeekId(itemId)}
                        className={`flex items-center justify-between gap-2 px-3 py-2 rounded-lg border text-left text-xs transition-colors ${
                          isChecked
                            ? 'bg-amber-50 border-amber-400 font-bold text-slate-900'
                            : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        <div className="flex items-center gap-2 min-w-0">
                          {isChecked ? (
                            <CheckSquare className="w-4 h-4 text-amber-600 shrink-0" />
                          ) : (
                            <Square className="w-4 h-4 text-slate-400 shrink-0" />
                          )}
                          <span className="truncate">
                            TOTAL WEEKLY {w.weekNumber} ({w.periodeText})
                          </span>
                        </div>
                        <span className="font-mono-num text-[11px] text-emerald-700 shrink-0">
                          Rp {wm ? formatNumberId(wm.weeklySales) || '0' : '0'}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* FILTER WEEKLY KHUSUS UNTUK HASIL EXPORT REVIEW MONITORING */}
            {options.includeStaffAndReview && (
              <div className="pt-3 border-t border-slate-200/90 space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <label className="text-xs font-extrabold text-indigo-950 flex items-center gap-1.5">
                    <ClipboardList className="w-3.5 h-3.5 text-indigo-600" />
                    <span>Filter Weekly untuk Hasil Export Review Monitoring Bazar:</span>
                  </label>
                  <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                    Dilengkapi Kolom &amp; AutoFilter Weekly di Excel
                  </span>
                </div>

                <select
                  aria-label="Filter Weekly Hasil Export Review"
                  value={options.selectedReviewWeekFilter || 'FOLLOW_WEEKLY'}
                  onChange={(e) =>
                    setOptions((prev) => ({
                      ...prev,
                      selectedReviewWeekFilter: e.target.value,
                    }))
                  }
                  className="w-full px-3 py-2 rounded-lg bg-white border border-indigo-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer"
                >
                  <option value="FOLLOW_WEEKLY">
                    🔄 Review Terbaru / Sesuai Filter Weekly ({getReviewDocWeeklyLabel(latestActiveInfo.latestReview)})
                  </option>
                  <option value="ALL">
                    📋 Semua Periode Weekly Review ({Math.max(1, reviewDocs.length)} Laporan Review)
                  </option>
                  {reviewDocs.map((d) => {
                    const wk = getReviewDocWeekNumber(d);
                    return (
                      <option key={d.id} value={d.id}>
                        ⭐ Dokumen Review Tersimpan: {getReviewDocWeeklyLabel(d)}
                        {wk ? ` [W${wk}]` : ''}
                      </option>
                    );
                  })}
                  {ANNUAL_2026_WEEKS.map((w) => {
                    const hasSaved = reviewDocs.some(
                      (d) => getReviewDocWeekNumber(d) === w.weekNumber
                    );
                    return (
                      <option key={`w-${w.weekNumber}`} value={`w-${w.weekNumber}`}>
                        📅 Filter Weekly {w.weekNumber} — {w.periodeText} ({w.monthName})
                        {hasSaved ? ' ✓ [Ada Review]' : ''}
                      </option>
                    );
                  })}
                </select>
              </div>
            )}
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-5 py-3.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="text-xs text-slate-600">
            Format: <strong>.xlsx (WPS Office / Microsoft Excel)</strong> · Otomatis Terformula & Berwarna
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-2 rounded-xl border border-slate-300 bg-white hover:bg-slate-100 text-xs font-bold text-slate-700 transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              disabled={isExporting || selectedSectionCount === 0}
              onClick={handleCustomExport}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-extrabold shadow-xs transition-all cursor-pointer disabled:opacity-45"
            >
              <Download className="w-4 h-4" />
              <span>
                {isExporting
                  ? 'Mengunduh...'
                  : `Export Pilihan (${selectedSectionCount} Bagian) ke 1 Excel`}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
