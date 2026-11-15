import React, { useState, useRef } from 'react';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  ReportSubmissionLog,
  calculateWeekMetrics,
  formatNumberId,
  normalizeAndRepairWeeklyTables,
} from '../data/initialData';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  normalizeCvInventoryOpsList,
} from '../data/cvDiv1Data';
import {
  Save,
  Download,
  Upload,
  CheckCircle2,
  BookmarkPlus,
  RotateCcw,
  Trash2,
  FileSpreadsheet,
  FileJson,
  Clock,
  X,
  HardDriveDownload,
  ShieldCheck,
} from 'lucide-react';

export interface SaveCheckpointItem {
  id: string;
  name: string;
  createdAt: string;
  totalWeeks: number;
  totalSales: number;
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  submissionLogs: ReportSubmissionLog[];
  cvDiv1Records?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
}

const STORAGE_KEY_CHECKPOINTS = 'bazar_kapuas_save_checkpoints_v1';

interface SaveToolsModalProps {
  isOpen: boolean;
  onClose: () => void;
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  submissionLogs: ReportSubmissionLog[];
  cvDiv1Records?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  lastSavedTimeStr: string;
  isSaving: boolean;
  onManualSaveNow: (customLabel?: string) => void;
  onRestoreState: (
    nextTables: WeeklySalesTable[],
    nextReviews: ReviewMonitoringDoc[],
    nextLogs: ReportSubmissionLog[],
    sourceLabel: string,
    nextCvRecords?: CvDailyRecord[],
    nextCvOps?: CvInventoryOpsEntry[]
  ) => void;
  onExportExcel: () => void;
}

export const SaveToolsModal: React.FC<SaveToolsModalProps> = ({
  isOpen,
  onClose,
  weeklyTables,
  reviewDocs,
  submissionLogs,
  cvDiv1Records,
  cvInventoryOps,
  lastSavedTimeStr,
  isSaving,
  onManualSaveNow,
  onRestoreState,
  onExportExcel,
}) => {
  const [checkpoints, setCheckpoints] = useState<SaveCheckpointItem[]>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY_CHECKPOINTS);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });

  const [checkpointName, setCheckpointName] = useState<string>('');
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  if (!isOpen) return null;

  const totalSalesAllWeeks = weeklyTables.reduce(
    (acc, t) => acc + calculateWeekMetrics(t).weeklySales,
    0
  );
  const filledWeeksCount = weeklyTables.filter(
    (t) => calculateWeekMetrics(t).weeklySales > 0
  ).length;

  const showTempStatus = (msg: string) => {
    setStatusMessage(msg);
    setImportError(null);
    setTimeout(() => {
      setStatusMessage((prev) => (prev === msg ? null : prev));
    }, 3200);
  };

  const persistCheckpoints = (next: SaveCheckpointItem[]) => {
    setCheckpoints(next);
    try {
      localStorage.setItem(STORAGE_KEY_CHECKPOINTS, JSON.stringify(next));
    } catch {}
  };

  const handleCreateCheckpoint = (e: React.FormEvent) => {
    e.preventDefault();
    const nowStr = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    const label =
      checkpointName.trim() || `Save Slot (${nowStr.slice(0, 16)})`;

    const newItem: SaveCheckpointItem = {
      id: `cp-${Date.now()}`,
      name: label,
      createdAt: nowStr,
      totalWeeks: weeklyTables.length,
      totalSales: totalSalesAllWeeks,
      weeklyTables: JSON.parse(JSON.stringify(weeklyTables)),
      reviewDocs: JSON.parse(JSON.stringify(reviewDocs)),
      submissionLogs: JSON.parse(JSON.stringify(submissionLogs)),
      ...(Array.isArray(cvDiv1Records)
        ? { cvDiv1Records: JSON.parse(JSON.stringify(cvDiv1Records)) }
        : {}),
      ...(Array.isArray(cvInventoryOps)
        ? { cvInventoryOps: JSON.parse(JSON.stringify(cvInventoryOps)) }
        : {}),
    };

    const next = [newItem, ...checkpoints].slice(0, 15);
    persistCheckpoints(next);
    setCheckpointName('');
    onManualSaveNow(`Save Slot "${label}" berhasil disimpan`);
    showTempStatus(`Save Slot "${label}" berhasil dibuat & disimpan!`);
  };

  const handleRestoreCheckpoint = (cp: SaveCheckpointItem) => {
    onRestoreState(
      normalizeAndRepairWeeklyTables(cp.weeklyTables),
      cp.reviewDocs,
      cp.submissionLogs,
      `Pulihkan Save Slot: ${cp.name}`,
      Array.isArray(cp.cvDiv1Records) ? cp.cvDiv1Records : undefined,
      Array.isArray(cp.cvInventoryOps)
        ? normalizeCvInventoryOpsList(cp.cvInventoryOps)
        : undefined
    );
    showTempStatus(`Data berhasil dipulihkan dari "${cp.name}"!`);
  };

  const handleDeleteCheckpoint = (id: string) => {
    const next = checkpoints.filter((c) => c.id !== id);
    persistCheckpoints(next);
    showTempStatus('Save Slot dihapus.');
  };

  const handleDownloadJsonBackup = (customData?: {
    name?: string;
    weeklyTables: WeeklySalesTable[];
    reviewDocs: ReviewMonitoringDoc[];
    submissionLogs: ReportSubmissionLog[];
    cvDiv1Records?: CvDailyRecord[];
    cvInventoryOps?: CvInventoryOpsEntry[];
  }) => {
    const targetCvRecords = customData ? customData.cvDiv1Records : cvDiv1Records;
    const targetCvOps = customData ? customData.cvInventoryOps : cvInventoryOps;
    const payload = {
      app: 'Bazar Matahari Tanjung Sales & Monitoring System',
      version: '2.1',
      savedAt: new Date().toISOString(),
      savedAtFormatted: new Date().toLocaleString('id-ID'),
      weeklyTables: customData ? customData.weeklyTables : weeklyTables,
      reviewDocs: customData ? customData.reviewDocs : reviewDocs,
      submissionLogs: customData ? customData.submissionLogs : submissionLogs,
      ...(Array.isArray(targetCvRecords) ? { cvDiv1Records: targetCvRecords } : {}),
      ...(Array.isArray(targetCvOps) ? { cvInventoryOps: targetCvOps } : {}),
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const dateStamp = new Date().toISOString().slice(0, 10);
    const slug = customData?.name
      ? `_${customData.name.replace(/[^a-zA-Z0-9_-]/g, '_')}`
      : '';
    a.href = url;
    a.download = `Bazar_Tanjung_Save_Backup_${dateStamp}${slug}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showTempStatus('File Save Backup (.json) berhasil diunduh ke perangkat!');
  };

  const handleImportJsonFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const text = String(ev.target?.result || '');
        const parsed = JSON.parse(text);
        if (!parsed || !Array.isArray(parsed.weeklyTables)) {
          setImportError(
            'Format file tidak valid. Pastikan memilih file Save Backup (.json) Bazar Tanjung.'
          );
          return;
        }
        const repairedTables = normalizeAndRepairWeeklyTables(parsed.weeklyTables);
        const nextReviews = Array.isArray(parsed.reviewDocs)
          ? parsed.reviewDocs
          : reviewDocs;
        const nextLogs = Array.isArray(parsed.submissionLogs)
          ? parsed.submissionLogs
          : submissionLogs;
        const nextCvRecords = Array.isArray(parsed.cvDiv1Records)
          ? parsed.cvDiv1Records
          : undefined;
        const nextCvOps = Array.isArray(parsed.cvInventoryOps)
          ? normalizeCvInventoryOpsList(parsed.cvInventoryOps)
          : undefined;

        onRestoreState(
          repairedTables,
          nextReviews,
          nextLogs,
          `Impor File Save (${file.name})`,
          nextCvRecords,
          nextCvOps
        );
        showTempStatus(
          `File "${file.name}" berhasil dimuat & disimpan ke server!`
        );
      } catch {
        setImportError(
          'Gagal membaca file Save (.json). Pastikan file tidak rusak.'
        );
      } finally {
        if (fileInputRef.current) {
          fileInputRef.current.value = '';
        }
      }
    };
    reader.readAsText(file);
  };

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center p-3 sm:p-4 bg-slate-950/65 backdrop-blur-xs">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-5 py-4 bg-slate-900 text-white flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-emerald-500 text-slate-950 flex items-center justify-center shrink-0 font-bold">
              <Save className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-bold text-white">
                Tools Save & Backup Data Bazar Tanjung
              </h2>
              <p className="text-[11px] text-slate-300">
                Simpan permanen ke server, buat Save Slot (Checkpoint), atau unduh/muat file Save (.json / .xlsx)
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 transition-colors"
            aria-label="Tutup Tools Save"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 overflow-y-auto space-y-5 text-xs">
          {statusMessage && (
            <div className="p-3 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-900 font-semibold flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{statusMessage}</span>
            </div>
          )}

          {importError && (
            <div className="p-3 rounded-xl bg-rose-50 border border-rose-200 text-rose-800 font-semibold">
              {importError}
            </div>
          )}

          {/* Section 1: Quick Save to Server & LocalStorage */}
          <div className="p-4 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <ShieldCheck className="w-4 h-4 text-emerald-700" />
                <span className="font-bold text-slate-900 text-xs sm:text-sm">
                  1. Simpan Langsung ke Server & Browser (Ctrl+S)
                </span>
              </div>
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-white border border-emerald-200 text-emerald-800">
                <Clock className="w-3 h-3" />
                Terakhir disimpan: {lastSavedTimeStr}
              </span>
            </div>

            <div className="grid grid-cols-3 gap-2.5 pt-1">
              <div className="p-2.5 rounded-lg bg-white border border-emerald-200/80">
                <div className="text-[11px] text-slate-500">Total Tabel Weekly</div>
                <div className="text-sm font-bold text-slate-900 font-mono-num mt-0.5">
                  {weeklyTables.length} Minggu ({filledWeeksCount} Terisi)
                </div>
              </div>
              <div className="p-2.5 rounded-lg bg-white border border-emerald-200/80">
                <div className="text-[11px] text-slate-500">Total Akumulasi Sales</div>
                <div className="text-sm font-bold text-emerald-700 font-mono-num mt-0.5">
                  Rp {formatNumberId(totalSalesAllWeeks) || '0'}
                </div>
              </div>
              <div className="p-2.5 rounded-lg bg-white border border-emerald-200/80">
                <div className="text-[11px] text-slate-500">Log Report Masuk</div>
                <div className="text-sm font-bold text-slate-900 font-mono-num mt-0.5">
                  {submissionLogs.length} Riwayat
                </div>
              </div>
            </div>

            <button
              type="button"
              disabled={isSaving}
              onClick={() => {
                onManualSaveNow();
                showTempStatus(
                  'Seluruh data Tabel Harian, Weekly, Bulanan & Review berhasil disimpan ke Server!'
                );
              }}
              className="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white font-bold flex items-center justify-center gap-2 shadow-xs transition-colors"
            >
              <Save className="w-4 h-4" />
              <span>
                {isSaving
                  ? 'Sedang Menyimpan Data...'
                  : 'Simpan Semua Perubahan Sekarang (Save Now)'}
              </span>
            </button>
          </div>

          {/* Section 2: Save Slots / Checkpoints */}
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <BookmarkPlus className="w-4 h-4 text-indigo-600" />
                <span className="font-bold text-slate-900 text-xs sm:text-sm">
                  2. Save Slot / Checkpoint (Simpan Versi Data)
                </span>
              </div>
              <span className="text-[11px] text-slate-500">
                {checkpoints.length} / 15 Slot Tersimpan
              </span>
            </div>

            <form onSubmit={handleCreateCheckpoint} className="flex flex-col sm:flex-row gap-2">
              <input
                type="text"
                value={checkpointName}
                onChange={(e) => setCheckpointName(e.target.value)}
                placeholder="Nama Save Slot (contoh: Rekap W41 Final / Sebelum Reset)..."
                className="flex-1 px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs focus:outline-none focus:border-indigo-600"
              />
              <button
                type="submit"
                className="inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg whitespace-nowrap transition-colors"
              >
                <BookmarkPlus className="w-3.5 h-3.5" />
                Buat Save Slot Baru
              </button>
            </form>

            {checkpoints.length > 0 ? (
              <div className="space-y-2 max-h-48 overflow-y-auto pr-1 pt-1">
                {checkpoints.map((cp) => (
                  <div
                    key={cp.id}
                    className="p-2.5 bg-white border border-slate-200 rounded-lg flex flex-wrap items-center justify-between gap-2"
                  >
                    <div className="min-w-0">
                      <div className="font-bold text-slate-900 truncate">{cp.name}</div>
                      <div className="text-[11px] text-slate-500 font-mono-num">
                        {cp.createdAt} · {cp.totalWeeks} Minggu · Rp{' '}
                        {formatNumberId(cp.totalSales) || '0'}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => handleRestoreCheckpoint(cp)}
                        className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-semibold transition-colors"
                        title="Pulihkan / Muat Save Slot ini"
                      >
                        <RotateCcw className="w-3 h-3" />
                        Muat
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          handleDownloadJsonBackup({
                            name: cp.name,
                            weeklyTables: cp.weeklyTables,
                            reviewDocs: cp.reviewDocs,
                            submissionLogs: cp.submissionLogs,
                            cvDiv1Records: cp.cvDiv1Records,
                            cvInventoryOps: cp.cvInventoryOps,
                          })
                        }
                        className="inline-flex items-center gap-1 px-2 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold transition-colors"
                        title="Unduh Save Slot ini sebagai file .json"
                      >
                        <Download className="w-3 h-3" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleDeleteCheckpoint(cp.id)}
                        className="inline-flex items-center gap-1 px-2 py-1.5 rounded-md bg-rose-50 hover:bg-rose-100 text-rose-700 font-semibold transition-colors"
                        title="Hapus Save Slot"
                      >
                        <Trash2 className="w-3 h-3" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="text-[11px] text-slate-500 italic py-1">
                Belum ada Save Slot. Buat Save Slot di atas untuk menyimpan titik cadangan data yang bisa dipulihkan kapan saja.
              </div>
            )}
          </div>

          {/* Section 3: Export & Import File Save (.json & .xlsx) */}
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
            <div className="flex items-center gap-2">
              <HardDriveDownload className="w-4 h-4 text-amber-600" />
              <span className="font-bold text-slate-900 text-xs sm:text-sm">
                3. Simpan ke File Perangkat & Muat File Save (Backup / Restore)
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
              <button
                type="button"
                onClick={() => handleDownloadJsonBackup()}
                className="p-3 rounded-xl bg-white hover:bg-slate-100 border border-slate-200 text-left flex flex-col justify-between gap-2 transition-colors"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900">Unduh File Save (.json)</span>
                  <FileJson className="w-4 h-4 text-indigo-600" />
                </div>
                <span className="text-[11px] text-slate-500">
                  Simpan cadangan penuh seluruh data ke HP/Komputer
                </span>
              </button>

              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="p-3 rounded-xl bg-white hover:bg-slate-100 border border-slate-200 text-left flex flex-col justify-between gap-2 transition-colors"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900">Muat File Save (.json)</span>
                  <Upload className="w-4 h-4 text-emerald-600" />
                </div>
                <span className="text-[11px] text-slate-500">
                  Pulihkan data dari file backup .json yang tersimpan
                </span>
              </button>

              <button
                type="button"
                onClick={() => {
                  onExportExcel();
                }}
                className="p-3 rounded-xl bg-white hover:bg-slate-100 border border-slate-200 text-left flex flex-col justify-between gap-2 transition-colors"
              >
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900">Export Excel (.xlsx)</span>
                  <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                </div>
                <span className="text-[11px] text-slate-500">
                  Buka menu ekspor lengkap format WPS Office / Microsoft Excel
                </span>
              </button>
            </div>

            <input
              ref={fileInputRef}
              type="file"
              accept=".json,application/json"
              onChange={handleImportJsonFile}
              className="hidden"
            />
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-5 py-3 bg-slate-100 border-t border-slate-200 flex items-center justify-between text-xs">
          <span className="text-slate-500">
            Pintasan Keyboard: <strong>Ctrl + S</strong> (Simpan Cepat) · <strong>Ctrl + Z</strong> (Undo)
          </span>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white font-semibold rounded-lg"
          >
            Selesai
          </button>
        </div>
      </div>
    </div>
  );
};
