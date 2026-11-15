import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  ReviewSection,
  ReviewDivisionRow,
  formatReviewBulletItem,
  getReviewDocWeekNumber,
  getReviewDocMonthKey,
  getReviewDocWeeklyLabel,
  createReviewDocForWeek,
  createEmptyReviewSections,
  countReviewDocItems,
  isReviewDocFilled,
  generateReviewDocFromWeeklySales,
  mergeTwoReviewDocsNonDestructive,
  ANNUAL_2026_WEEKS,
  getWeeklyPeriodeText,
  calculateWeekMetrics,
  formatNumberId,
  getIndoDayName,
} from '../data/initialData';
import { exportWeeklyTablesToWPSExcel } from '../utils/exportExcel';
import { copyTextToClipboard } from '../utils/clipboard';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  markAutoDraftClean,
  clearAutoDraftFromStorage,
} from '../utils/autoDraft';
import { AutoDraftStatusBadge } from './AutoDraftStatusBanner';
import {
  Plus,
  Trash2,
  Copy,
  Check,
  Edit2,
  LayoutGrid,
  Table2,
  Undo2,
  Redo2,
  Save,
  Filter,
  Calendar,
  CalendarRange,
  Download,
  FileSpreadsheet,
  X,
  Lock,
  Sparkles,
  Eraser,
  ChevronLeft,
  ChevronRight,
  ClipboardEdit,
  CheckCircle2,
  AlertCircle,
} from 'lucide-react';

interface ReviewMonitoringSheetProps {
  docs: ReviewMonitoringDoc[];
  activeDocId: string;
  onSelectDoc: (id: string) => void;
  onChangeDocs: (updated: ReviewMonitoringDoc[]) => void;
  weeklyTables?: WeeklySalesTable[];
  selectedWeeklyFilter?: string;
  onSelectWeeklyFilter?: (weekId: string) => void;
  showDivZero?: boolean;
  canUndo?: boolean;
  canRedo?: boolean;
  onUndo?: () => void;
  onRedo?: () => void;
  onManualSave?: () => void;
  isPublicReadOnly?: boolean;
}

const MONTH_FILTER_OPTIONS = [
  { key: 'ALL', label: 'Semua Bulan (Jan–Des 2026)' },
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

export const ReviewMonitoringSheet: React.FC<ReviewMonitoringSheetProps> = ({
  docs,
  activeDocId,
  onSelectDoc,
  onChangeDocs,
  weeklyTables = [],
  selectedWeeklyFilter = 'ALL',
  onSelectWeeklyFilter,
  showDivZero = true,
  canUndo = false,
  canRedo = false,
  onUndo,
  onRedo,
  onManualSave,
  isPublicReadOnly = false,
}) => {
  const [viewMode, setViewMode] = useState<'spreadsheet' | 'form' | 'structured'>(
    'spreadsheet'
  );
  const [editMode, setEditMode] = useState<boolean>(false);
  const [editingCellKey, setEditingCellKey] = useState<string | null>(null);
  const [copiedNotice, setCopiedNotice] = useState<boolean>(false);
  const [actionBannerMsg, setActionBannerMsg] = useState<string | null>(null);
  const [sectionFilter, setSectionFilter] = useState<string>('ALL');

  // Filter Weekly states inside Menu Review
  const [reviewMonthFilter, setReviewMonthFilter] = useState<string>('ALL');
  const [reviewStatusFilter, setReviewStatusFilter] = useState<
    'ALL' | 'FILLED' | 'EMPTY'
  >('ALL');
  const [reviewWeeklyScope, setReviewWeeklyScope] = useState<'SINGLE' | 'ALL_WEEKS'>(
    'SINGLE'
  );

  // Track the selected weekNumber (1..53) in SINGLE mode so every Weekly (even not yet saved in docs) can be viewed & filled independently
  const [selectedWeekNumber, setSelectedWeekNumber] = useState<number>(() => {
    if (selectedWeeklyFilter && selectedWeeklyFilter !== 'ALL') {
      const matchedTbl = weeklyTables.find((t) => t.id === selectedWeeklyFilter);
      if (matchedTbl) return matchedTbl.weekNumber;
      const numMatch = selectedWeeklyFilter.match(/(\d+)/);
      if (numMatch) {
        const n = Number(numMatch[1]);
        if (n >= 1 && n <= 53) return n;
      }
    }
    const foundDoc = docs.find((d) => d.id === activeDocId) || docs[0];
    return (foundDoc && getReviewDocWeekNumber(foundDoc)) || 40;
  });

  // Track per-table slot week overrides when viewing multiple Review tables so changing e.g. WEEKLY 41 ▼ to WEEKLY 1 ▼ updates that table directly
  const [tableSlotOverrides, setTableSlotOverrides] = useState<
    Record<number, number>
  >({});

  // Copy-from-other-weekly selector state
  const [copySourceWeekNum, setCopySourceWeekNum] = useState<number>(39);

  // Export Review Modal / Bar state with Weekly Filter
  const [isExportReviewOpen, setIsExportReviewOpen] = useState<boolean>(false);
  const [exportReviewFilterMode, setExportReviewFilterMode] = useState<
    'CURRENT_WEEK' | 'ALL_WEEKS' | 'BY_MONTH' | 'SPECIFIC_WEEK'
  >('CURRENT_WEEK');
  const [exportReviewMonthKey, setExportReviewMonthKey] = useState<string>('ALL');
  const [exportReviewWeekNum, setExportReviewWeekNum] = useState<number>(40);
  const [includeSalesInReviewExport, setIncludeSalesInReviewExport] =
    useState<boolean>(false);
  const [isExportingReview, setIsExportingReview] = useState<boolean>(false);
  const [isReviewDraftDirty, setIsReviewDraftDirty] = useState<boolean>(false);
  const [recoveredReviewDraft, setRecoveredReviewDraft] =
    useState<ReviewMonitoringDoc | null>(null);

  const showTransientMessage = (msg: string) => {
    setActionBannerMsg(msg);
    setTimeout(() => {
      setActionBannerMsg((prev) => (prev === msg ? null : prev));
    }, 3000);
  };

  // Sync with global selectedWeeklyFilter when user clicks a specific Weekly in the sidebar or top filter
  useEffect(() => {
    if (!selectedWeeklyFilter || selectedWeeklyFilter === 'ALL') return;
    const matchedTable = weeklyTables.find((t) => t.id === selectedWeeklyFilter);
    const wk =
      matchedTable?.weekNumber ||
      Number((selectedWeeklyFilter.match(/(\d+)/) || [])[1] || 0);
    if (wk >= 1 && wk <= 53) {
      setReviewWeeklyScope('SINGLE');
      setSelectedWeekNumber(wk);
      const existingDoc = docs.find((d) => getReviewDocWeekNumber(d) === wk);
      if (existingDoc && existingDoc.id !== activeDocId) {
        onSelectDoc(existingDoc.id);
      }
    }
  }, [selectedWeeklyFilter]);

  // Also sync if activeDocId is changed externally
  useEffect(() => {
    const matchedDoc = docs.find((d) => d.id === activeDocId);
    const wk = matchedDoc ? getReviewDocWeekNumber(matchedDoc) : null;
    if (wk && wk >= 1 && wk <= 53 && wk !== selectedWeekNumber) {
      setSelectedWeekNumber(wk);
    }
  }, [activeDocId]);

  const latestDocsRef = useRef<ReviewMonitoringDoc[]>(docs);
  latestDocsRef.current = docs;

  // Retrieve or build a clean (non-cloned) ReviewMonitoringDoc for any weekNumber 1..53
  const getOrBuildDocForWeek = (weekNum: number): ReviewMonitoringDoc => {
    const existing = docs.find((d) => getReviewDocWeekNumber(d) === weekNum);
    if (existing) return existing;
    return createReviewDocForWeek(weekNum, docs[0], false);
  };

  const activeDoc = useMemo(() => {
    return getOrBuildDocForWeek(selectedWeekNumber);
  }, [docs, selectedWeekNumber]);

  const activeDocWeekNum = selectedWeekNumber;

  useEffect(() => {
    if (isPublicReadOnly) return;
    const savedEnv = loadAutoDraftFromStorage<ReviewMonitoringDoc>(
      AUTO_DRAFT_KEYS.REVIEW_MONITORING
    );
    if (savedEnv && savedEnv.isDirty && savedEnv.data?.sections) {
      const savedWk = getReviewDocWeekNumber(savedEnv.data) || 40;
      const existing = docs.find((d) => getReviewDocWeekNumber(d) === savedWk);
      if (
        !existing ||
        JSON.stringify(existing.sections) !==
          JSON.stringify(savedEnv.data.sections)
      ) {
        setRecoveredReviewDraft(savedEnv.data);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const reviewDraftTimeStr = '';
  const isReviewOnline = true;

  const handleRestoreReviewDraft = () => {
    if (!recoveredReviewDraft) return;
    const wk = getReviewDocWeekNumber(recoveredReviewDraft) || selectedWeekNumber;
    updateDocByWeekNumber(wk, (currentDoc) =>
      mergeTwoReviewDocsNonDestructive(recoveredReviewDraft, currentDoc)
    );
    setSelectedWeekNumber(wk);
    setRecoveredReviewDraft(null);
    setIsReviewDraftDirty(false);
    markAutoDraftClean(AUTO_DRAFT_KEYS.REVIEW_MONITORING);
    showTransientMessage(`Draf Review WEEKLY ${wk} berhasil dipulihkan dari localStorage`);
  };

  useEffect(() => {
    if (activeDocWeekNum) {
      setExportReviewWeekNum(activeDocWeekNum);
    }
  }, [activeDocWeekNum]);

  // Helper to select a specific weekNumber (1..53)
  const handleSelectWeeklyNumber = (weekNum: number) => {
    if (!weekNum || weekNum < 1 || weekNum > 53) return;
    setReviewWeeklyScope('SINGLE');
    setSelectedWeekNumber(weekNum);
    setEditingCellKey(null);

    // Sync with global selectedWeeklyFilter if that week exists in weeklyTables
    const matchedSalesTable = weeklyTables.find((t) => t.weekNumber === weekNum);
    if (matchedSalesTable && onSelectWeeklyFilter) {
      onSelectWeeklyFilter(matchedSalesTable.id);
    }

    // Find existing review doc for this weekNumber
    const existingDoc = docs.find((d) => getReviewDocWeekNumber(d) === weekNum);
    if (existingDoc) {
      onSelectDoc(existingDoc.id);
      return;
    }

    // Create a clean empty ReviewMonitoringDoc for this Weekly period if Admin
    if (!isPublicReadOnly) {
      const newDoc = createReviewDocForWeek(weekNum, docs[0], false);
      const updatedList = [...docs, newDoc].sort(
        (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
      );
      onChangeDocs(updatedList);
      onSelectDoc(newDoc.id);
    }
  };

  // Update a specific week's ReviewMonitoringDoc (creates it in docs if not yet present)
  const updateDocByWeekNumber = (
    weekNum: number,
    updater: (doc: ReviewMonitoringDoc) => ReviewMonitoringDoc
  ) => {
    if (isPublicReadOnly) return;
    setIsReviewDraftDirty(true);
    const currentDocs = latestDocsRef.current;
    const existing = currentDocs.find((d) => getReviewDocWeekNumber(d) === weekNum);
    if (existing) {
      const updatedList = currentDocs.map((d) =>
        getReviewDocWeekNumber(d) === weekNum || d.id === existing.id
          ? updater(d)
          : d
      );
      latestDocsRef.current = updatedList;
      onChangeDocs(updatedList);
    } else {
      const baseNew = createReviewDocForWeek(weekNum, currentDocs[0], false);
      const updatedNew = updater(baseNew);
      const updatedList = [...currentDocs, updatedNew].sort(
        (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
      );
      latestDocsRef.current = updatedList;
      onChangeDocs(updatedList);
      onSelectDoc(updatedNew.id);
    }
  };

  const updateActiveDoc = (
    updater: (doc: ReviewMonitoringDoc) => ReviewMonitoringDoc
  ) => {
    updateDocByWeekNumber(activeDocWeekNum, updater);
  };

  // Build combined list of Weekly items (from docs + weeklyTables) so every active Weekly is accessible
  const quickWeeklyItems = useMemo(() => {
    const map = new Map<
      number,
      {
        weekNumber: number;
        label: string;
        periodeText: string;
        monthKey: string;
        hasReviewDoc: boolean;
        isReviewFilled: boolean;
        filledItemsCount: number;
        docId?: string;
        hasSalesTable: boolean;
        weeklySales: number;
      }
    >();

    for (const d of docs) {
      const wk = getReviewDocWeekNumber(d);
      if (wk) {
        const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
        const matchedTbl = weeklyTables.find((t) => t.weekNumber === wk);
        const sls = matchedTbl ? calculateWeekMetrics(matchedTbl).weeklySales : 0;
        const cnt = countReviewDocItems(d);
        map.set(wk, {
          weekNumber: wk,
          label: `Weekly ${wk}`,
          periodeText: sched
            ? sched.periodeText
            : d.periode.replace(/^Periode\s*:\s*/i, ''),
          monthKey: getReviewDocMonthKey(d),
          hasReviewDoc: true,
          isReviewFilled: cnt > 0,
          filledItemsCount: cnt,
          docId: d.id,
          hasSalesTable: Boolean(matchedTbl),
          weeklySales: sls,
        });
      }
    }

    for (const t of weeklyTables) {
      const existing = map.get(t.weekNumber);
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === t.weekNumber);
      const sls = calculateWeekMetrics(t).weeklySales;
      if (!existing) {
        map.set(t.weekNumber, {
          weekNumber: t.weekNumber,
          label: `Weekly ${t.weekNumber}`,
          periodeText: getWeeklyPeriodeText(t),
          monthKey: sched ? sched.monthKey : '10',
          hasReviewDoc: false,
          isReviewFilled: false,
          filledItemsCount: 0,
          hasSalesTable: true,
          weeklySales: sls,
        });
      } else {
        existing.hasSalesTable = true;
        existing.weeklySales = sls;
      }
    }

    // Ensure currently selected weekNumber is always in the map
    if (!map.has(selectedWeekNumber)) {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === selectedWeekNumber);
      if (sched) {
        map.set(selectedWeekNumber, {
          weekNumber: selectedWeekNumber,
          label: `Weekly ${selectedWeekNumber}`,
          periodeText: sched.periodeText,
          monthKey: sched.monthKey,
          hasReviewDoc: false,
          isReviewFilled: false,
          filledItemsCount: 0,
          hasSalesTable: false,
          weeklySales: 0,
        });
      }
    }

    return Array.from(map.values())
      .filter((item) => {
        if (reviewMonthFilter !== 'ALL' && item.monthKey !== reviewMonthFilter) {
          return false;
        }
        if (reviewStatusFilter === 'FILLED' && !item.isReviewFilled) {
          return false;
        }
        if (reviewStatusFilter === 'EMPTY' && item.isReviewFilled) {
          return false;
        }
        return true;
      })
      .sort((a, b) => a.weekNumber - b.weekNumber);
  }, [
    docs,
    weeklyTables,
    reviewMonthFilter,
    reviewStatusFilter,
    selectedWeekNumber,
  ]);

  // Docs to render in main area:
  // - If SINGLE: [activeDoc]
  // - If ALL_WEEKS: every Weekly in quickWeeklyItems (each with its own ReviewMonitoringDoc, supporting per-table filter selection 1..53!)
  const visibleReviewDocs = useMemo(() => {
    if (reviewWeeklyScope === 'SINGLE') {
      return [activeDoc];
    }
    const list = quickWeeklyItems.map((item, idx) => {
      const overriddenWk = tableSlotOverrides[idx] || item.weekNumber;
      return getOrBuildDocForWeek(overriddenWk);
    });
    return list.length > 0 ? list : [activeDoc];
  }, [reviewWeeklyScope, activeDoc, quickWeeklyItems, docs, tableSlotOverrides]);

  const handleTableHeaderWeekFilterChange = (
    docIndex: number,
    targetWeekNum: number
  ) => {
    if (!targetWeekNum || targetWeekNum < 1 || targetWeekNum > 53) return;
    if (reviewWeeklyScope === 'ALL_WEEKS') {
      setTableSlotOverrides((prev) => ({
        ...prev,
        [docIndex]: targetWeekNum,
      }));
      setEditingCellKey(null);
      const existingDoc = docs.find(
        (d) => getReviewDocWeekNumber(d) === targetWeekNum
      );
      if (!existingDoc && !isPublicReadOnly) {
        const newDoc = createReviewDocForWeek(targetWeekNum, docs[0], false);
        const updatedList = [...docs, newDoc].sort(
          (a, b) =>
            (getReviewDocWeekNumber(a) || 999) -
            (getReviewDocWeekNumber(b) || 999)
        );
        onChangeDocs(updatedList);
      }
      return;
    }
    handleSelectWeeklyNumber(targetWeekNum);
  };

  // Filled docs available to copy from
  const filledDocsList = useMemo(() => {
    return docs
      .filter((d) => isReviewDocFilled(d))
      .sort(
        (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
      );
  }, [docs]);

  const handleDuplicateOrAddNextWeekly = () => {
    const existingFilledWeekNums = new Set(
      docs
        .filter((d) => isReviewDocFilled(d))
        .map((d) => getReviewDocWeekNumber(d))
        .filter(Boolean) as number[]
    );
    let nextWeekNum = (activeDocWeekNum || 40) + 1;
    if (nextWeekNum > 53) nextWeekNum = 1;
    if (existingFilledWeekNums.has(nextWeekNum)) {
      const candidate = ANNUAL_2026_WEEKS.find(
        (w) => !existingFilledWeekNums.has(w.weekNumber) && w.weekNumber >= 40
      );
      if (candidate) nextWeekNum = candidate.weekNumber;
    }
    handleSelectWeeklyNumber(nextWeekNum);
    showTransientMessage(
      `Membuka lembar isian Review baru untuk WEEKLY ${nextWeekNum}`
    );
  };

  // Copy content from another Weekly into target weekNum
  const handleCopyFromAnotherWeek = (targetWeekNum: number, sourceWeekNum: number) => {
    if (isPublicReadOnly) return;
    const sourceDoc = docs.find((d) => getReviewDocWeekNumber(d) === sourceWeekNum);
    if (!sourceDoc) return;
    const copiedDoc = createReviewDocForWeek(targetWeekNum, sourceDoc, true);
    updateDocByWeekNumber(targetWeekNum, () => copiedDoc);
    showTransientMessage(
      `Isian WEEKLY ${sourceWeekNum} berhasil disalin ke WEEKLY ${targetWeekNum}`
    );
  };

  // Clear all items in target weekNum so it is empty/clean
  const handleClearWeekReview = (targetWeekNum: number) => {
    if (isPublicReadOnly) return;
    updateDocByWeekNumber(targetWeekNum, (doc) => ({
      ...doc,
      sections: createEmptyReviewSections(doc.sections),
    }));
    showTransientMessage(`Isian Review pada WEEKLY ${targetWeekNum} telah dikosongkan`);
  };

  // Auto-populate Review Monitoring from that week's Daily Sales & Shift Reasons
  const handleAutoPullFromSalesReasons = (targetWeekNum: number) => {
    if (isPublicReadOnly) return;
    const matchedTable = weeklyTables.find((t) => t.weekNumber === targetWeekNum);
    if (!matchedTable) {
      showTransientMessage(
        `Tabel Sales Weekly ${targetWeekNum} belum memiliki catatan Reason Harian`
      );
      return;
    }
    const currentDoc = getOrBuildDocForWeek(targetWeekNum);
    const generated = generateReviewDocFromWeeklySales(matchedTable, currentDoc);
    updateDocByWeekNumber(targetWeekNum, () => generated);
    showTransientMessage(
      `Ringkasan otomatis dari Reason Harian Sales Weekly ${targetWeekNum} berhasil dimasukkan`
    );
  };

  // Copy WhatsApp text (supports current Weekly filter or All Filtered Weeks)
  const handleCopyWhatsAppText = () => {
    const docsToCopy = visibleReviewDocs;

    const lines: string[] = [];
    lines.push('*REPORT BAZAR MATAHARI*');
    lines.push('*Tanjung*');
    lines.push('PT MDS Retailing Tbk');
    lines.push('');

    for (const docItem of docsToCopy) {
      lines.push(`==============================`);
      lines.push(`*${docItem.title}*`);
      lines.push(`*Periode : ${getReviewDocWeeklyLabel(docItem)}*`);
      lines.push('');

      const secsToCopy =
        sectionFilter === 'ALL'
          ? docItem.sections
          : docItem.sections.filter((s) => s.id === sectionFilter);

      for (const sec of secsToCopy) {
        lines.push(`*${sec.code}*`);
        if (sec.subRows && sec.subRows.length > 0) {
          for (const sub of sec.subRows) {
            lines.push(`_${sub.label}_:`);
            const validItems = sub.items
              .map((it) => formatReviewBulletItem(it))
              .filter(Boolean);
            if (validItems.length > 0) {
              for (const formatted of validItems) {
                lines.push(`  ${formatted}`);
              }
            } else {
              lines.push(`  - (Belum ada isian)`);
            }
          }
        } else if (sec.directItems) {
          const validDirect = sec.directItems
            .map((it) => formatReviewBulletItem(it))
            .filter(Boolean);
          if (validDirect.length > 0) {
            for (const formatted of validDirect) {
              lines.push(`  ${formatted}`);
            }
          } else {
            lines.push(`  - (Belum ada isian)`);
          }
        }
        lines.push('');
      }
    }

    void copyTextToClipboard(lines.join('\n')).then((ok) => {
      if (ok) {
        setCopiedNotice(true);
        setTimeout(() => setCopiedNotice(false), 2000);
      }
    });
  };

  // Export Review to Excel (.xlsx) with Filter Weekly
  const handleExecuteExportReviewExcel = async () => {
    if (isPublicReadOnly) return;
    setIsExportingReview(true);
    try {
      let reviewWeekFilterArg = 'ALL';
      let customMonthArg = 'ALL';
      let selectedTableIds: string[] = weeklyTables.map((t) => t.id);
      let fileSuffix = 'Semua_Weekly';

      if (exportReviewFilterMode === 'CURRENT_WEEK') {
        if (reviewWeeklyScope === 'ALL_WEEKS') {
          reviewWeekFilterArg = 'ALL';
          customMonthArg = reviewMonthFilter;
          fileSuffix =
            reviewMonthFilter === 'ALL'
              ? 'Semua_Weekly'
              : `Bulan_${reviewMonthFilter}`;
        } else if (activeDocWeekNum) {
          reviewWeekFilterArg = `w-${activeDocWeekNum}`;
          const matchedTbl = weeklyTables.find(
            (t) => t.weekNumber === activeDocWeekNum
          );
          if (matchedTbl) selectedTableIds = [matchedTbl.id];
          fileSuffix = `Weekly_${activeDocWeekNum}`;
        } else {
          reviewWeekFilterArg = activeDoc.id;
          fileSuffix = 'Weekly_Aktif';
        }
      } else if (exportReviewFilterMode === 'SPECIFIC_WEEK') {
        reviewWeekFilterArg = `w-${exportReviewWeekNum}`;
        const matchedTbl = weeklyTables.find(
          (t) => t.weekNumber === exportReviewWeekNum
        );
        if (matchedTbl) selectedTableIds = [matchedTbl.id];
        fileSuffix = `Weekly_${exportReviewWeekNum}`;
      } else if (exportReviewFilterMode === 'BY_MONTH') {
        reviewWeekFilterArg = 'ALL';
        customMonthArg = exportReviewMonthKey;
        fileSuffix =
          exportReviewMonthKey === 'ALL'
            ? 'Semua_Bulan'
            : `Bulan_${exportReviewMonthKey}`;
      } else {
        reviewWeekFilterArg = 'ALL';
        fileSuffix = 'Semua_Weekly';
      }

      // Ensure all visibleReviewDocs are included in the exported pool
      const exportPoolMap = new Map<number, ReviewMonitoringDoc>();
      for (const d of docs) {
        const w = getReviewDocWeekNumber(d);
        if (w) exportPoolMap.set(w, d);
      }
      for (const vd of visibleReviewDocs) {
        const w = getReviewDocWeekNumber(vd);
        if (w && !exportPoolMap.has(w)) {
          exportPoolMap.set(w, vd);
        }
      }

      const dateStr = new Date().toISOString().slice(0, 10);
      await exportWeeklyTablesToWPSExcel(weeklyTables, activeDoc, showDivZero, {
        includeWeeklyMatrix: includeSalesInReviewExport,
        includeDailyTable: false,
        includeMonthlyTable: false,
        includeReasonEvaluation: false,
        includeStaffAndReview: true,
        includeChartDiagrams: includeSalesInReviewExport,
        sheetLayout: includeSalesInReviewExport
          ? 'COMBINED_AND_SEPARATE'
          : 'SEPARATE_SHEETS',
        selectedWeekIds: selectedTableIds,
        selectedMonthKey: customMonthArg,
        allReviewDocs: Array.from(exportPoolMap.values()),
        selectedReviewWeekFilter: reviewWeekFilterArg,
        customFileName: `Review_Monitoring_Bazar_Tanjung_${fileSuffix}_${dateStr}`,
      });
      setIsExportReviewOpen(false);
    } finally {
      setIsExportingReview(false);
    }
  };

  const handleSubRowItemChangeForWeek = (
    weekNum: number,
    secId: string,
    subRowId: string,
    newText: string
  ) => {
    const lines = newText === '' ? [] : newText.split('\n');
    updateDocByWeekNumber(weekNum, (doc) => ({
      ...doc,
      sections: doc.sections.map((sec) => {
        if (sec.id !== secId || !sec.subRows) return sec;
        return {
          ...sec,
          subRows: sec.subRows.map((sr) =>
            sr.id === subRowId ? { ...sr, items: lines } : sr
          ),
        };
      }),
    }));
  };

  const handleDirectItemsChangeForWeek = (
    weekNum: number,
    secId: string,
    newText: string
  ) => {
    const lines = newText === '' ? [] : newText.split('\n');
    updateDocByWeekNumber(weekNum, (doc) => ({
      ...doc,
      sections: doc.sections.map((sec) =>
        sec.id === secId ? { ...sec, directItems: lines } : sec
      ),
    }));
  };

  const handleAddSubRowForWeek = (weekNum: number, secId: string) => {
    updateDocByWeekNumber(weekNum, (doc) => ({
      ...doc,
      sections: doc.sections.map((sec) => {
        if (sec.id !== secId) return sec;
        const nextRow: ReviewDivisionRow = {
          id: `sub-${Date.now()}`,
          label: 'Divisi Baru',
          isItalic: true,
          items: ['- Item baru'],
        };
        return {
          ...sec,
          subRows: [...(sec.subRows || []), nextRow],
        };
      }),
    }));
  };

  const handleDeleteSubRowForWeek = (
    weekNum: number,
    secId: string,
    subRowId: string
  ) => {
    updateDocByWeekNumber(weekNum, (doc) => ({
      ...doc,
      sections: doc.sections.map((sec) => {
        if (sec.id !== secId || !sec.subRows) return sec;
        return {
          ...sec,
          subRows: sec.subRows.filter((sr) => sr.id !== subRowId),
        };
      }),
    }));
  };

  const handleAddSectionForWeek = (weekNum: number) => {
    const nextSec: ReviewSection = {
      id: `sec-${Date.now()}`,
      code: 'E. CATATAN TAMBAHAN',
      isMainHeader: true,
      directItems: ['- Tambahan evaluasi bazar'],
    };
    updateDocByWeekNumber(weekNum, (doc) => ({
      ...doc,
      sections: [...doc.sections, nextSec],
    }));
  };

  // Render a single Weekly Review Spreadsheet Table (Authentic WPS Office style + interactive TOTAL WEEKLY X bar)
  const renderSpreadsheetTableForDoc = (
    docItem: ReviewMonitoringDoc,
    docIndex: number
  ) => {
    const wkNum = getReviewDocWeekNumber(docItem) || activeDocWeekNum || 40;
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum);
    const matchedSalesTable = weeklyTables.find((t) => t.weekNumber === wkNum);
    const salesMetrics = matchedSalesTable
      ? calculateWeekMetrics(matchedSalesTable)
      : null;
    const periodeClean = sched
      ? sched.periodeText
      : docItem.periode
          .replace(/^WEEKLY\s*\d+\s*\(/i, '')
          .replace(/\)$/, '')
          .replace(/^Periode\s*:\s*/i, '');
    const filledCount = countReviewDocItems(docItem);
    const isFilled = filledCount > 0;

    const secsToRender =
      sectionFilter === 'ALL'
        ? docItem.sections
        : docItem.sections.filter((s) => s.id === sectionFilter);

    let currentRowNumber = 2;

    return (
      <div
        key={docItem.id || `rev-sheet-${wkNum}-${docIndex}`}
        className="bg-white border border-slate-300 rounded-xl overflow-hidden shadow-xs space-y-0"
      >
        {/* Per-Weekly Quick Action & Status Strip */}
        <div className="no-print px-3.5 py-2 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-2 border-b border-slate-800">
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <span className="px-2 py-0.5 rounded bg-amber-400 text-slate-950 font-extrabold text-[11px]">
              WEEKLY {wkNum}
            </span>
            <span className="text-[11px] font-semibold text-amber-200">
              {periodeClean}
            </span>
            <div className="inline-flex items-center gap-1.5 bg-slate-800 border border-slate-700 rounded-lg px-2 py-0.5 text-xs">
              <Calendar className="w-3.5 h-3.5 text-amber-400 shrink-0" />
              <label htmlFor={`rev-cal-${wkNum}-${docIndex}`} className="text-[10px] font-bold text-slate-300">
                Kalender:
              </label>
              <input
                id={`rev-cal-${wkNum}-${docIndex}`}
                type="date"
                value={sched?.startDateIso || '2026-10-05'}
                onChange={(e) => {
                  const val = e.target.value;
                  if (val) {
                    const found = ANNUAL_2026_WEEKS.find((w) => {
                      const [y, m, d] = val.split('-').map(Number);
                      const t = new Date(y, m - 1, d).getTime();
                      const [sy, sm, sd] = w.startDateIso.split('-').map(Number);
                      const st = new Date(sy, sm - 1, sd).getTime();
                      const [my, mm, md] = w.endDateIso.split('-').map(Number);
                      const et = new Date(my, mm - 1, md).getTime();
                      return t >= st && t <= et;
                    });
                    if (found) {
                      handleTableHeaderWeekFilterChange(docIndex, found.weekNumber);
                    }
                  }
                }}
                className="text-[11px] font-bold bg-slate-900 text-white border border-slate-700 rounded px-1.5 py-0.5 cursor-pointer focus:outline-none focus:border-amber-400"
                title="Pilih tanggal kalender untuk otomatis membuka evaluasi weekly tersebut"
              />
            </div>
            <span
              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold ${
                isFilled
                  ? 'bg-emerald-500/20 text-emerald-200 border border-emerald-400/30'
                  : 'bg-amber-500/20 text-amber-200 border border-amber-400/30'
              }`}
            >
              {isFilled ? (
                <>
                  <CheckCircle2 className="w-3 h-3 text-emerald-300" />
                  <span>{filledCount} Poin Terisi</span>
                </>
              ) : (
                <>
                  <AlertCircle className="w-3 h-3 text-amber-300" />
                  <span>Belum Diisi</span>
                </>
              )}
            </span>
            {salesMetrics && salesMetrics.weeklySales > 0 && (
              <span className="px-2 py-0.5 rounded bg-indigo-500/25 border border-indigo-400/30 text-[10px] font-mono-num font-bold text-indigo-100">
                Sales: Rp {formatNumberId(salesMetrics.weeklySales)}
              </span>
            )}
          </div>

          {!isPublicReadOnly && (
            <div className="flex flex-wrap items-center gap-1.5">
              {matchedSalesTable && (
                <button
                  type="button"
                  onClick={() => handleAutoPullFromSalesReasons(wkNum)}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-bold bg-amber-400/20 hover:bg-amber-400/30 text-amber-200 border border-amber-400/30 transition-colors cursor-pointer"
                  title={`Tarik ringkasan dari Reason/Event Harian Sales Weekly ${wkNum}`}
                >
                  <Sparkles className="w-3 h-3" />
                  <span>Tarik Reason Harian</span>
                </button>
              )}

              {filledDocsList.length > 0 && (
                <div className="inline-flex items-center gap-1 bg-white/10 rounded-md px-1.5 py-0.5 border border-white/15">
                  <select
                    aria-label={`Pilih sumber salin Weekly untuk W${wkNum}`}
                    value={String(copySourceWeekNum)}
                    onChange={(e) => setCopySourceWeekNum(Number(e.target.value))}
                    className="bg-transparent text-[11px] font-bold text-amber-200 focus:outline-none cursor-pointer"
                  >
                    {filledDocsList.map((fd) => {
                      const fwk = getReviewDocWeekNumber(fd) || 39;
                      return (
                        <option
                          key={fd.id}
                          value={String(fwk)}
                          className="text-slate-900"
                        >
                          Salin W{fwk}
                        </option>
                      );
                    })}
                  </select>
                  <button
                    type="button"
                    onClick={() => handleCopyFromAnotherWeek(wkNum, copySourceWeekNum)}
                    className="px-1.5 py-0.5 rounded bg-white/15 hover:bg-white/25 text-[10px] font-bold text-white transition-colors cursor-pointer"
                    title={`Salin isi Review dari Weekly ${copySourceWeekNum} ke Weekly ${wkNum}`}
                  >
                    Tempel
                  </button>
                </div>
              )}

              {isFilled && (
                <button
                  type="button"
                  onClick={() => handleClearWeekReview(wkNum)}
                  className="inline-flex items-center gap-1 px-2 py-1 rounded-md text-[11px] font-semibold bg-rose-500/20 hover:bg-rose-500/30 text-rose-200 border border-rose-400/30 transition-colors cursor-pointer"
                  title={`Kosongkan seluruh isian pada Weekly ${wkNum}`}
                >
                  <Eraser className="w-3 h-3" />
                  <span>Kosongkan</span>
                </button>
              )}
            </div>
          )}
        </div>

        <div className="overflow-x-auto spreadsheet-scroll">
          <table className="w-full border-collapse text-sm bg-white min-w-[760px]">
            <tbody>
              {/* Row 2: REPORT BAZAR MATAHARI (Red Banner) */}
              <tr>
                <td className="w-10 min-w-[40px] bg-[#F8FAFC] border-r border-b border-slate-300 text-center text-xs text-slate-500 font-mono-num select-none py-2">
                  {currentRowNumber++}
                </td>
                <td
                  colSpan={2}
                  className="bg-[#C00000] text-white font-bold text-lg sm:text-xl text-center py-2.5 px-4 border border-slate-900 tracking-wide"
                >
                  REPORT BAZAR MATAHARI
                </td>
              </tr>

              {/* Row 3: Tanjung (Yellow/Amber Banner) */}
              <tr>
                <td className="bg-[#F8FAFC] border-r border-b border-slate-300 text-center text-xs text-slate-500 font-mono-num select-none py-1.5">
                  {currentRowNumber++}
                </td>
                <td
                  colSpan={2}
                  className="bg-[#FFC000] text-slate-950 font-bold text-base sm:text-lg text-center py-2 px-4 border border-slate-900"
                >
                  Tanjung
                </td>
              </tr>

              {/* Row 4: PT MDS Retailing Tbk (Dark Navy Banner) */}
              <tr>
                <td className="bg-[#F8FAFC] border-r border-b border-slate-300 text-center text-xs text-slate-500 font-mono-num select-none py-1.5">
                  {currentRowNumber++}
                </td>
                <td
                  colSpan={2}
                  className="bg-[#0F172A] text-[#FDE047] font-bold text-sm sm:text-base text-center py-2 px-4 border border-slate-900"
                >
                  PT MDS Retailing Tbk
                </td>
              </tr>

              {/* Row 5: Blank Separator Row */}
              <tr className="h-5">
                <td className="bg-[#F8FAFC] border-r border-b border-slate-200 text-center text-xs text-slate-500 font-mono-num select-none">
                  {currentRowNumber++}
                </td>
                <td colSpan={2} className="bg-white"></td>
              </tr>

              {/* Row 6: BAGIAN 5: REVIEW MONITORING BAZAR — REPORT BAZAR MATAHARI (Dark Green Banner) */}
              <tr>
                <td className="bg-[#F8FAFC] border-r border-b border-slate-300 text-center text-xs text-slate-500 font-mono-num select-none py-2">
                  {currentRowNumber++}
                </td>
                <td
                  colSpan={2}
                  className="bg-[#064E3B] text-[#FDE047] font-bold text-xs sm:text-sm text-left py-2 px-3 border border-slate-900 tracking-wide"
                >
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span>
                      BAGIAN 5: REVIEW MONITORING BAZAR — REPORT BAZAR MATAHARI
                    </span>
                    <span className="text-[11px] font-semibold text-emerald-200">
                      WEEKLY {wkNum} ({periodeClean})
                    </span>
                  </div>
                </td>
              </tr>

              {/* Row 7: WEEKLY X ▼ (Filter Pilih Weekly 1 s/d 53 & Tanda Filter) | REPORT BAZAR MATAHARI — WEEKLY X (Periode) */}
              <tr>
                <td className="bg-[#F8FAFC] border-r border-b border-slate-300 text-center text-xs text-slate-500 font-mono-num select-none py-1.5">
                  {currentRowNumber++}
                </td>
                <td className="w-[175px] min-w-[165px] bg-[#1E293B] text-[#FDE047] font-extrabold text-xs text-center py-2 px-2 border-2 border-[#FFC000] align-middle">
                  <div
                    className="flex items-center justify-center gap-1"
                    title="Klik tanda filter ▼ untuk pilih WEEKLY 1 s/d WEEKLY 53"
                  >
                    <Filter className="w-3.5 h-3.5 text-[#FFC000] shrink-0" />
                    <select
                      aria-label={`Filter Tabel Weekly ${wkNum}`}
                      value={String(wkNum)}
                      onChange={(e) =>
                        handleTableHeaderWeekFilterChange(
                          docIndex,
                          Number(e.target.value)
                        )
                      }
                      className="w-full bg-transparent text-[#FDE047] font-extrabold text-xs text-center focus:outline-none cursor-pointer"
                    >
                      {ANNUAL_2026_WEEKS.map((w) => (
                        <option
                          key={w.weekNumber}
                          value={String(w.weekNumber)}
                          className="text-slate-900 bg-white font-bold"
                        >
                          WEEKLY {w.weekNumber} ▼
                        </option>
                      ))}
                    </select>
                  </div>
                </td>
                <td className="bg-[#0F172A] text-[#FDE047] font-bold text-xs text-left py-2 px-3 border border-slate-900 align-middle">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span>
                      {docItem.title} — WEEKLY {wkNum} ({periodeClean})
                      {salesMetrics && salesMetrics.weeklySales > 0
                        ? ` — TOTAL SALES WEEKLY: Rp ${formatNumberId(salesMetrics.weeklySales)}`
                        : ''}
                    </span>
                    <span className="text-[11px] text-emerald-300 font-semibold">
                      {isFilled
                        ? `Isian Weekly ${wkNum}: ${filledCount} Poin Terisi`
                        : `Isian Weekly ${wkNum}: Klik baris di bawah untuk mengisi`}
                    </span>
                  </div>
                </td>
              </tr>

              {/* Rows 8..23: 2-Column Table matching photo format with direct per-Weekly inline editing */}
              {secsToRender.map((sec) => {
                const hasSubRows = sec.subRows && sec.subRows.length > 0;

                if (hasSubRows) {
                  const sectionHeaderRowNum = currentRowNumber++;
                  return (
                    <React.Fragment key={sec.id}>
                      {/* Section Header Row (e.g. A. MERCHANDISE, B.PROMO) */}
                      <tr>
                        <td className="bg-[#F8FAFC] border-r border-b border-slate-200 text-center text-xs text-slate-500 font-mono-num select-none py-1">
                          {sectionHeaderRowNum}
                        </td>
                        <td className="w-[175px] min-w-[165px] border border-slate-900 px-2.5 py-1.5 text-left text-[13px] font-bold text-slate-950 bg-slate-50/70 align-middle">
                          <input
                            type="text"
                            readOnly={isPublicReadOnly}
                            value={sec.code}
                            onChange={(e) =>
                              updateDocByWeekNumber(wkNum, (d) => ({
                                ...d,
                                sections: d.sections.map((s) =>
                                  s.id === sec.id
                                    ? { ...s, code: e.target.value }
                                    : s
                                ),
                              }))
                            }
                            className="w-full bg-transparent font-bold focus:bg-amber-50 focus:outline-none rounded"
                          />
                        </td>
                        <td className="border border-slate-900 px-3 py-1 bg-slate-50/40">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[11px] text-slate-500 font-medium">
                              Isian {sec.code} — WEEKLY {wkNum} ({periodeClean})
                            </span>
                            {!isPublicReadOnly && editMode && (
                              <button
                                type="button"
                                onClick={() => handleAddSubRowForWeek(wkNum, sec.id)}
                                className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded"
                              >
                                <Plus className="w-3 h-3" />
                                Tambah Baris Divisi
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>

                      {/* Sub-rows (Divisi 1..4 right-aligned italic; FAST MOVING & SLOW MOVING left-aligned middle tall) */}
                      {sec.subRows!.map((sub) => {
                        const rowNum = currentRowNumber++;
                        const cellKey = `${wkNum}-${sec.id}-${sub.id}`;
                        const isCellEditing =
                          !isPublicReadOnly &&
                          (editMode || editingCellKey === cellKey);
                        const joinedText = sub.items.join('\n');
                        const validLines = sub.items.filter(
                          (line) => (line || '').trim().length > 0
                        );
                        const isDivisi =
                          Boolean(sub.isItalic) ||
                          sub.label.trim().toLowerCase().startsWith('divisi');
                        const isFastOrSlow =
                          sub.label.toUpperCase().includes('FAST MOVING') ||
                          sub.label.toUpperCase().includes('SLOW MOVING');
                        const rowCount = isFastOrSlow
                          ? Math.max(5, sub.items.length)
                          : Math.max(2, sub.items.length);

                        return (
                          <tr
                            key={sub.id}
                            className={isFastOrSlow ? 'min-h-[110px]' : ''}
                          >
                            <td className="bg-[#F8FAFC] border-r border-b border-slate-200 text-center text-xs text-slate-500 font-mono-num select-none py-1.5 align-middle">
                              {rowNum}
                            </td>
                            <td
                              className={`w-[175px] min-w-[165px] border border-slate-900 px-2.5 bg-white align-middle ${
                                isFastOrSlow ? 'py-8' : 'py-1.5'
                              }`}
                            >
                              <div className="flex items-center justify-between gap-1">
                                <input
                                  type="text"
                                  readOnly={isPublicReadOnly}
                                  value={sub.label}
                                  onChange={(e) =>
                                    updateDocByWeekNumber(wkNum, (d) => ({
                                      ...d,
                                      sections: d.sections.map((s) =>
                                        s.id === sec.id
                                          ? {
                                              ...s,
                                              subRows: s.subRows?.map((sr) =>
                                                sr.id === sub.id
                                                  ? { ...sr, label: e.target.value }
                                                  : sr
                                              ),
                                            }
                                          : s
                                      ),
                                    }))
                                  }
                                  className={`w-full text-[13px] text-slate-950 bg-transparent focus:bg-amber-50 focus:outline-none rounded ${
                                    isDivisi
                                      ? 'italic text-right font-normal'
                                      : 'text-left font-normal'
                                  }`}
                                />
                                {!isPublicReadOnly && editMode && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      handleDeleteSubRowForWeek(
                                        wkNum,
                                        sec.id,
                                        sub.id
                                      )
                                    }
                                    className="p-1 text-red-600 hover:bg-red-50 rounded"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                )}
                              </div>
                            </td>
                            <td
                              className={`border border-slate-900 px-3 bg-white ${
                                isFastOrSlow
                                  ? 'py-2.5 align-top min-h-[110px]'
                                  : 'py-1.5 align-middle'
                              }`}
                            >
                              {isCellEditing ? (
                                <div className="space-y-1">
                                  <textarea
                                    autoFocus={editingCellKey === cellKey}
                                    value={joinedText}
                                    onChange={(e) =>
                                      handleSubRowItemChangeForWeek(
                                        wkNum,
                                        sec.id,
                                        sub.id,
                                        e.target.value
                                      )
                                    }
                                    placeholder={`Ketik isian ${sub.label} (${sec.code}) untuk WEEKLY ${wkNum} (satu poin per baris)...`}
                                    rows={rowCount}
                                    className="w-full text-[13px] leading-snug text-slate-950 bg-amber-50/40 border border-emerald-500 rounded p-2 focus:outline-none"
                                  />
                                  {!editMode && (
                                    <div className="flex justify-end">
                                      <button
                                        type="button"
                                        onClick={() => setEditingCellKey(null)}
                                        className="px-2.5 py-0.5 rounded bg-emerald-600 text-white text-[11px] font-bold hover:bg-emerald-700"
                                      >
                                        Selesai Baris Ini
                                      </button>
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <div
                                  onClick={() => {
                                    if (!isPublicReadOnly) {
                                      setEditingCellKey(cellKey);
                                    }
                                  }}
                                  className={`space-y-0.5 text-[13px] leading-snug text-slate-950 rounded px-1 py-0.5 transition-colors ${
                                    !isPublicReadOnly
                                      ? 'cursor-text hover:bg-amber-50/60'
                                      : ''
                                  } ${
                                    isFastOrSlow ? 'min-h-[88px]' : 'min-h-[22px]'
                                  }`}
                                  title={
                                    isPublicReadOnly
                                      ? undefined
                                      : `Klik untuk mengisi / mengedit ${sub.label} pada WEEKLY ${wkNum}`
                                  }
                                >
                                  {validLines.length > 0 ? (
                                    validLines.map((line, lIdx) => (
                                      <div key={lIdx}>
                                        {formatReviewBulletItem(line) || line}
                                      </div>
                                    ))
                                  ) : (
                                    <div className="text-xs italic text-slate-400 py-0.5">
                                      {isPublicReadOnly
                                        ? `- Belum ada isian ${sub.label} pada Weekly ${wkNum}`
                                        : `+ Klik untuk mengisi ${sub.label} (${sec.code}) pada Weekly ${wkNum}...`}
                                    </div>
                                  )}
                                </div>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  );
                }

                // Direct sections: C. PLATE/PEOPLE, D. ACTION PLAN, REASON/KENDALA, NEED SUPPORT
                const directRowNum = currentRowNumber++;
                const cellKey = `${wkNum}-${sec.id}-direct`;
                const isCellEditing =
                  !isPublicReadOnly && (editMode || editingCellKey === cellKey);
                const directJoined = (sec.directItems || []).join('\n');
                const validDirectLines = (sec.directItems || []).filter(
                  (line) => (line || '').trim().length > 0
                );
                const directLinesCount = Math.max(
                  4,
                  (sec.directItems || []).length
                );

                return (
                  <tr key={sec.id}>
                    <td className="bg-[#F8FAFC] border-r border-b border-slate-200 text-center text-xs text-slate-500 font-mono-num select-none py-2 align-middle">
                      {directRowNum}
                    </td>
                    <td className="w-[175px] min-w-[165px] border border-slate-900 px-2.5 py-7 align-middle text-left text-[13px] font-bold text-slate-950 bg-white">
                      <input
                        type="text"
                        readOnly={isPublicReadOnly}
                        value={sec.code}
                        onChange={(e) =>
                          updateDocByWeekNumber(wkNum, (d) => ({
                            ...d,
                            sections: d.sections.map((s) =>
                              s.id === sec.id
                                ? { ...s, code: e.target.value }
                                : s
                            ),
                          }))
                        }
                        className="w-full bg-transparent font-bold focus:bg-amber-50 focus:outline-none rounded"
                      />
                    </td>
                    <td className="border border-slate-900 px-3 py-2.5 align-top bg-white">
                      {isCellEditing ? (
                        <div className="space-y-1">
                          <textarea
                            autoFocus={editingCellKey === cellKey}
                            value={directJoined}
                            onChange={(e) =>
                              handleDirectItemsChangeForWeek(
                                wkNum,
                                sec.id,
                                e.target.value
                              )
                            }
                            placeholder={`Ketik isian ${sec.code} untuk WEEKLY ${wkNum} (satu poin per baris)...`}
                            rows={directLinesCount}
                            className="w-full text-[13px] leading-snug text-slate-950 bg-amber-50/40 border border-emerald-500 rounded p-2 focus:outline-none"
                          />
                          {!editMode && (
                            <div className="flex justify-end">
                              <button
                                type="button"
                                onClick={() => setEditingCellKey(null)}
                                className="px-2.5 py-0.5 rounded bg-emerald-600 text-white text-[11px] font-bold hover:bg-emerald-700"
                              >
                                Selesai Baris Ini
                              </button>
                            </div>
                          )}
                        </div>
                      ) : (
                        <div
                          onClick={() => {
                            if (!isPublicReadOnly) {
                              setEditingCellKey(cellKey);
                            }
                          }}
                          className={`space-y-0.5 text-[13px] leading-snug text-slate-950 rounded px-1 py-0.5 transition-colors min-h-[76px] ${
                            !isPublicReadOnly
                              ? 'cursor-text hover:bg-amber-50/60'
                              : ''
                          }`}
                          title={
                            isPublicReadOnly
                              ? undefined
                              : `Klik untuk mengisi / mengedit ${sec.code} pada WEEKLY ${wkNum}`
                          }
                        >
                          {validDirectLines.length > 0 ? (
                            validDirectLines.map((line, lIdx) => (
                              <div key={lIdx}>
                                {formatReviewBulletItem(line) || line}
                              </div>
                            ))
                          ) : (
                            <div className="text-xs italic text-slate-400 py-1">
                              {isPublicReadOnly
                                ? `- Belum ada isian ${sec.code} pada Weekly ${wkNum}`
                                : `+ Klik untuk mengisi ${sec.code} pada Weekly ${wkNum}...`}
                            </div>
                          )}
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
    );
  };

  // Render dedicated "Form Isian Setiap Weekly" view
  const renderFormIsianWeeklyView = (docItem: ReviewMonitoringDoc) => {
    const wkNum = getReviewDocWeekNumber(docItem) || activeDocWeekNum || 40;
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum);
    const matchedSalesTable = weeklyTables.find((t) => t.weekNumber === wkNum);
    const salesMetrics = matchedSalesTable
      ? calculateWeekMetrics(matchedSalesTable)
      : null;
    const periodeClean = sched
      ? sched.periodeText
      : docItem.periode
          .replace(/^WEEKLY\s*\d+\s*\(/i, '')
          .replace(/\)$/, '')
          .replace(/^Periode\s*:\s*/i, '');
    const filledCount = countReviewDocItems(docItem);

    return (
      <div
        key={`form-weekly-${wkNum}`}
        className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
      >
        {/* Top Banner of Form Isian Weekly */}
        <div className="bg-gradient-to-r from-emerald-950 via-teal-900 to-slate-900 text-white px-5 py-4 flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="flex flex-wrap items-center gap-2 text-xs font-bold text-amber-300">
              <span>FORM ISIAN REVIEW MONITORING BAZAR SETIAP WEEKLY</span>
              <span>·</span>
              <span className="px-2 py-0.5 rounded bg-amber-400 text-slate-950 font-extrabold">
                WEEKLY {wkNum}
              </span>
            </div>
            <h2 className="text-base sm:text-lg font-extrabold text-white mt-1">
              Isian Evaluasi Weekly {wkNum} — Periode: {periodeClean}
            </h2>
            <p className="text-xs text-emerald-200 mt-0.5">
              Setiap perubahan otomatis tersimpan ke lembar kerja{' '}
              <strong>WEEKLY {wkNum}</strong> tanpa mengubah isian Weekly lainnya.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex items-center gap-1 bg-white/10 p-1 rounded-xl border border-white/15">
              <button
                type="button"
                disabled={wkNum <= 1}
                onClick={() => handleSelectWeeklyNumber(Math.max(1, wkNum - 1))}
                className="p-1.5 rounded-lg hover:bg-white/20 text-white disabled:opacity-35 cursor-pointer"
                title="Weekly Sebelumnya"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <select
                aria-label="Pilih Weekly di Form Isian"
                value={String(wkNum)}
                onChange={(e) => handleSelectWeeklyNumber(Number(e.target.value))}
                className="px-2.5 py-1 rounded-lg bg-amber-400 text-slate-950 text-xs font-extrabold focus:outline-none cursor-pointer"
              >
                {ANNUAL_2026_WEEKS.map((w) => {
                  const exDoc = docs.find(
                    (d) => getReviewDocWeekNumber(d) === w.weekNumber
                  );
                  const hasIsian = isReviewDocFilled(exDoc);
                  return (
                    <option key={w.weekNumber} value={String(w.weekNumber)}>
                      WEEKLY {w.weekNumber} ({w.periodeText}){' '}
                      {hasIsian ? '✓ [Terisi]' : '[Kosong]'}
                    </option>
                  );
                })}
              </select>
              <button
                type="button"
                disabled={wkNum >= 53}
                onClick={() => handleSelectWeeklyNumber(Math.min(53, wkNum + 1))}
                className="p-1.5 rounded-lg hover:bg-white/20 text-white disabled:opacity-35 cursor-pointer"
                title="Weekly Berikutnya"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>

            <button
              type="button"
              onClick={() => setViewMode('spreadsheet')}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white text-slate-900 hover:bg-slate-100 text-xs font-extrabold shadow-2xs cursor-pointer"
            >
              <Table2 className="w-3.5 h-3.5 text-emerald-700" />
              <span>Lihat Tabel WPS W{wkNum}</span>
            </button>
          </div>
        </div>

        {/* Helper Bar: Sales Reference & 1-Click Auto-Fill / Copy / Clear for this Weekly */}
        <div className="px-5 py-3 bg-emerald-50/70 border-b border-emerald-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <span className="font-extrabold text-emerald-950">
              Status Isian W{wkNum}:
            </span>
            <span
              className={`px-2 py-0.5 rounded-full font-bold ${
                filledCount > 0
                  ? 'bg-emerald-600 text-white'
                  : 'bg-amber-200 text-amber-950'
              }`}
            >
              {filledCount > 0
                ? `${filledCount} Poin Evaluasi Terisi`
                : 'Masih Kosong (Silakan Isi di Bawah)'}
            </span>
            {salesMetrics && (
              <span className="text-slate-700 font-mono-num">
                · Sales W{wkNum}:{' '}
                <strong>Rp {formatNumberId(salesMetrics.weeklySales) || '0'}</strong>{' '}
                ({formatNumberId(salesMetrics.weeklyQty) || '0'} pcs /{' '}
                {formatNumberId(salesMetrics.weeklyTrx) || '0'} trx)
              </span>
            )}
          </div>

          {!isPublicReadOnly && (
            <div className="flex flex-wrap items-center gap-2">
              {matchedSalesTable && (
                <button
                  type="button"
                  onClick={() => handleAutoPullFromSalesReasons(wkNum)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold shadow-2xs cursor-pointer"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>Tarik Otomatis dari Reason Harian W{wkNum}</span>
                </button>
              )}

              {filledDocsList.length > 0 && (
                <div className="inline-flex items-center gap-1 bg-white border border-emerald-300 rounded-lg px-2 py-1">
                  <select
                    aria-label="Pilih Weekly sumber salin"
                    value={String(copySourceWeekNum)}
                    onChange={(e) => setCopySourceWeekNum(Number(e.target.value))}
                    className="text-xs font-bold text-slate-800 bg-transparent focus:outline-none cursor-pointer"
                  >
                    {filledDocsList.map((fd) => {
                      const fwk = getReviewDocWeekNumber(fd) || 39;
                      return (
                        <option key={fd.id} value={String(fwk)}>
                          Salin dari Weekly {fwk}
                        </option>
                      );
                    })}
                  </select>
                  <button
                    type="button"
                    onClick={() =>
                      handleCopyFromAnotherWeek(wkNum, copySourceWeekNum)
                    }
                    className="px-2 py-0.5 rounded bg-emerald-100 hover:bg-emerald-200 text-emerald-950 text-xs font-extrabold cursor-pointer"
                  >
                    Salin ke W{wkNum}
                  </button>
                </div>
              )}

              {filledCount > 0 && (
                <button
                  type="button"
                  onClick={() => handleClearWeekReview(wkNum)}
                  className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold cursor-pointer"
                >
                  <Eraser className="w-3.5 h-3.5" />
                  <span>Kosongkan W{wkNum}</span>
                </button>
              )}
            </div>
          )}
        </div>

        {/* Daily Shift Reasons Reference from Sales Weekly (if available) */}
        {matchedSalesTable &&
          matchedSalesTable.days.some(
            (d) =>
              (d.reasonPagi && d.reasonPagi.trim()) ||
              (d.reasonSiang && d.reasonSiang.trim())
          ) && (
            <div className="mx-5 mt-4 p-3.5 bg-amber-50/70 border border-amber-200 rounded-xl space-y-2">
              <div className="text-xs font-extrabold text-amber-950 flex items-center justify-between">
                <span>
                  📌 Referensi Catatan Reason/Event Harian pada Sales Weekly {wkNum}{' '}
                  ({periodeClean}):
                </span>
                <span className="text-[11px] font-semibold text-amber-800">
                  Dapat dijadikan acuan pengisian Review Weekly {wkNum}
                </span>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-40 overflow-y-auto pr-1 text-[11px] text-slate-700">
                {matchedSalesTable.days.map((d, idx) => {
                  const hasR =
                    (d.reasonPagi && d.reasonPagi.trim()) ||
                    (d.reasonSiang && d.reasonSiang.trim());
                  if (!hasR) return null;
                  return (
                    <div
                      key={d.id || idx}
                      className="p-2 bg-white border border-amber-200/80 rounded-lg space-y-0.5"
                    >
                      <div className="font-bold text-slate-900">
                        Hari {idx + 1} ({getIndoDayName(d.tanggal)}, {d.tanggal})
                      </div>
                      {d.reasonPagi && (
                        <div>
                          <strong className="text-amber-800">
                            Pagi ({d.pagi.nama || '-'}):
                          </strong>{' '}
                          {d.reasonPagi}
                        </div>
                      )}
                      {d.reasonSiang && (
                        <div>
                          <strong className="text-indigo-800">
                            Siang ({d.siang.nama || '-'}):
                          </strong>{' '}
                          {d.reasonSiang}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

        {/* All Sections Input Form for Weekly X */}
        <div className="p-5 space-y-5">
          {docItem.sections.map((sec) => {
            if (sec.subRows && sec.subRows.length > 0) {
              return (
                <div
                  key={sec.id}
                  className="bg-slate-50/80 border border-slate-200 rounded-xl p-4 space-y-3.5"
                >
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <span className="text-sm font-extrabold text-slate-900">
                      {sec.code} — WEEKLY {wkNum}
                    </span>
                    <span className="text-[11px] font-semibold text-slate-500">
                      Satu poin per baris (otomatis diformat bullet &quot;- &quot;)
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                    {sec.subRows.map((sub) => {
                      const isFastOrSlow =
                        sub.label.toUpperCase().includes('FAST MOVING') ||
                        sub.label.toUpperCase().includes('SLOW MOVING');
                      const joined = sub.items.join('\n');
                      return (
                        <div
                          key={sub.id}
                          className={isFastOrSlow ? 'md:col-span-1' : ''}
                        >
                          <label className="block text-xs font-bold text-slate-800 mb-1">
                            {sub.label} ({sec.code}) :
                          </label>
                          <textarea
                            readOnly={isPublicReadOnly}
                            rows={isFastOrSlow ? 5 : 3}
                            value={joined}
                            onChange={(e) =>
                              handleSubRowItemChangeForWeek(
                                wkNum,
                                sec.id,
                                sub.id,
                                e.target.value
                              )
                            }
                            placeholder={`Ketik isian ${sub.label} pada WEEKLY ${wkNum}...`}
                            className="w-full p-2.5 text-xs leading-relaxed bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-emerald-600"
                          />
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            }

            const directJoined = (sec.directItems || []).join('\n');
            return (
              <div
                key={sec.id}
                className="bg-slate-50/80 border border-slate-200 rounded-xl p-4 space-y-2"
              >
                <div className="flex items-center justify-between">
                  <label className="text-sm font-extrabold text-slate-900">
                    {sec.code} — WEEKLY {wkNum}
                  </label>
                  <span className="text-[11px] font-semibold text-slate-500">
                    Satu poin per baris
                  </span>
                </div>
                <textarea
                  readOnly={isPublicReadOnly}
                  rows={4}
                  value={directJoined}
                  onChange={(e) =>
                    handleDirectItemsChangeForWeek(wkNum, sec.id, e.target.value)
                  }
                  placeholder={`Ketik isian ${sec.code} pada WEEKLY ${wkNum}...`}
                  className="w-full p-3 text-xs leading-relaxed bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-emerald-600"
                />
              </div>
            );
          })}

          <div className="pt-2 flex flex-wrap items-center justify-between gap-3 border-t border-slate-200">
            <div className="text-xs text-slate-500">
              Seluruh isian di atas tersimpan khusus untuk{' '}
              <strong className="text-slate-800">
                WEEKLY {wkNum} ({periodeClean})
              </strong>
              .
            </div>
            <div className="flex flex-wrap items-center gap-2">
              {!isPublicReadOnly && (
                <AutoDraftStatusBadge
                  menuLabel={`Review W${wkNum}`}
                  lastDraftTimeStr={reviewDraftTimeStr}
                  isDirty={isReviewDraftDirty}
                  isOnline={isReviewOnline}
                />
              )}
              {onManualSave && !isPublicReadOnly && (
                <button
                  type="button"
                  onClick={() => {
                    onManualSave();
                    setIsReviewDraftDirty(false);
                    markAutoDraftClean(AUTO_DRAFT_KEYS.REVIEW_MONITORING);
                    showTransientMessage(
                      `Isian Review WEEKLY ${wkNum} berhasil disimpan permanen`
                    );
                  }}
                  className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold shadow-xs cursor-pointer"
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>Simpan Isian Weekly {wkNum}</span>
                </button>
              )}
              <button
                type="button"
                onClick={() => setViewMode('spreadsheet')}
                className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-amber-300 text-xs font-extrabold shadow-xs cursor-pointer"
              >
                <Table2 className="w-3.5 h-3.5" />
                <span>Selesai &amp; Lihat Tabel WPS Weekly {wkNum}</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-5">
      {!isPublicReadOnly && (
        <AutoDraftStatusBadge
          variant="banner"
          menuLabel={`Review Monitoring Bazar (WEEKLY ${activeDocWeekNum})`}
          lastDraftTimeStr={reviewDraftTimeStr}
          isDirty={isReviewDraftDirty}
          isOnline={isReviewOnline}
          hasRestorableDraft={Boolean(recoveredReviewDraft)}
          restorableDraftSummary={
            recoveredReviewDraft
              ? `WEEKLY ${getReviewDocWeekNumber(recoveredReviewDraft) || activeDocWeekNum}`
              : undefined
          }
          onRestoreDraft={handleRestoreReviewDraft}
          onClearDraft={() => {
            setRecoveredReviewDraft(null);
            clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REVIEW_MONITORING);
          }}
          onManualSave={() => {
            if (onManualSave) onManualSave();
            setIsReviewDraftDirty(false);
            markAutoDraftClean(AUTO_DRAFT_KEYS.REVIEW_MONITORING);
            showTransientMessage(
              `Isian Review WEEKLY ${activeDocWeekNum} berhasil disimpan permanen`
            );
          }}
        />
      )}

      {actionBannerMsg && (
        <div className="no-print bg-emerald-900 text-white px-4 py-2.5 rounded-xl flex items-center justify-between text-xs font-bold shadow-xs">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-amber-300 shrink-0" />
            <span>{actionBannerMsg}</span>
          </div>
          <button
            type="button"
            onClick={() => setActionBannerMsg(null)}
            className="text-emerald-200 hover:text-white"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* =====================================================================
          1. FILTER WEEKLY & PERIODE BAR DI BAGIAN MENU REVIEW
          ===================================================================== */}
      <div className="no-print bg-white border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
        {/* Top Banner of Review Monitoring */}
        <div className="px-4 py-3.5 bg-gradient-to-r from-slate-900 via-teal-950 to-emerald-950 text-white flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
          <div className="flex items-start sm:items-center gap-3 min-w-0">
            <div className="w-9 h-9 rounded-xl bg-emerald-500 text-slate-950 flex items-center justify-center shrink-0 font-bold shadow-xs mt-0.5 sm:mt-0">
              <ClipboardEdit className="w-4 h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <div className="flex flex-wrap items-center gap-1.5">
                <h2 className="text-sm sm:text-base font-extrabold tracking-tight text-white">
                  Review Monitoring Bazar (W1–W53)
                </h2>
                <span className="px-2 py-0.5 rounded-md text-[10px] font-extrabold bg-amber-400 text-slate-950">
                  {reviewWeeklyScope === 'ALL_WEEKS'
                    ? `Semua Weekly (${visibleReviewDocs.length} Periode)`
                    : `WEEKLY ${activeDocWeekNum}`}
                </span>
                <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-emerald-500/20 text-emerald-200 border border-emerald-400/30">
                  {countReviewDocItems(activeDoc)} Poin Terisi
                </span>
              </div>
              <p className="text-[11px] text-emerald-200/90 mt-0.5 leading-snug">
                {getReviewDocWeeklyLabel(activeDoc)} · Evaluasi Merchandise, Promo, People, Action Plan &amp; Support
              </p>
            </div>
          </div>

          {/* Right Quick Actions in Header */}
          <div className="flex items-center gap-2 shrink-0">
            {!isPublicReadOnly && (
              <button
                type="button"
                onClick={handleDuplicateOrAddNextWeekly}
                className="flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold text-emerald-100 bg-white/10 hover:bg-white/15 border border-white/15 transition-colors cursor-pointer whitespace-nowrap"
                title="Buka Periode Weekly Berikutnya untuk Mengisi Review"
              >
                <Plus className="w-3.5 h-3.5 text-amber-300" />
                <span>Weekly Baru</span>
              </button>
            )}
            <button
              type="button"
              disabled={isPublicReadOnly}
              onClick={() => {
                if (isPublicReadOnly) return;
                setIsExportReviewOpen(!isExportReviewOpen);
              }}
              title={
                isPublicReadOnly
                  ? 'Mode Publik (Hanya Lihat): Tidak bisa klik Export Review'
                  : 'Export Review ke Excel (.xlsx)'
              }
              className="flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-extrabold bg-amber-400 hover:bg-amber-300 text-slate-950 disabled:bg-slate-700 disabled:text-slate-400 disabled:opacity-60 disabled:cursor-not-allowed shadow-xs transition-colors cursor-pointer whitespace-nowrap"
            >
              {isPublicReadOnly ? (
                <Lock className="w-3.5 h-3.5" />
              ) : (
                <Download className="w-3.5 h-3.5" />
              )}
              <span>Export Review (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Compact Filter Bar for Review Monitoring */}
        <div className="p-3 bg-slate-50 border-b border-slate-200 space-y-2">
          {/* Primary Row: Weekly Selector with Prev/Next inside border */}
          <div className="flex items-center gap-1.5 bg-white px-2.5 py-1.5 rounded-xl border border-emerald-400 shadow-2xs w-full">
            <Filter className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
            <span className="text-[11px] font-extrabold text-slate-600 whitespace-nowrap">
              Weekly:
            </span>
            <button
              type="button"
              disabled={reviewWeeklyScope === 'SINGLE' && activeDocWeekNum <= 1}
              onClick={() => {
                const prevWk = Math.max(1, (activeDocWeekNum || 40) - 1);
                handleSelectWeeklyNumber(prevWk);
              }}
              className="p-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 disabled:opacity-35 transition-colors cursor-pointer shrink-0"
              title="Weekly Sebelumnya"
            >
              <ChevronLeft className="w-3.5 h-3.5" />
            </button>
            <select
              aria-label="Filter Pilih Periode Weekly Review Monitoring"
              value={
                reviewWeeklyScope === 'ALL_WEEKS'
                  ? 'ALL_WEEKS'
                  : String(activeDocWeekNum)
              }
              onChange={(e) => {
                const val = e.target.value;
                if (val === 'ALL_WEEKS') {
                  setTableSlotOverrides({});
                  setReviewWeeklyScope('ALL_WEEKS');
                  if (onSelectWeeklyFilter) onSelectWeeklyFilter('ALL');
                } else {
                  const wk = Number(val);
                  if (wk >= 1 && wk <= 53) {
                    handleSelectWeeklyNumber(wk);
                  }
                }
              }}
              className="flex-1 min-w-0 bg-transparent text-xs font-extrabold text-slate-950 focus:outline-none cursor-pointer truncate"
            >
              <option value="ALL_WEEKS">
                Semua Weekly ({quickWeeklyItems.length} Periode)
              </option>
              {quickWeeklyItems.map((item) => (
                <option key={item.weekNumber} value={String(item.weekNumber)}>
                  WEEKLY {item.weekNumber} ({item.periodeText}){' '}
                  {item.isReviewFilled
                    ? `✓ [${item.filledItemsCount} Poin]`
                    : '[Kosong]'}
                </option>
              ))}
            </select>
            <button
              type="button"
              disabled={reviewWeeklyScope === 'SINGLE' && activeDocWeekNum >= 53}
              onClick={() => {
                const nextWk = Math.min(53, (activeDocWeekNum || 40) + 1);
                handleSelectWeeklyNumber(nextWk);
              }}
              className="p-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 disabled:opacity-35 transition-colors cursor-pointer shrink-0"
              title="Weekly Berikutnya"
            >
              <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Secondary Row: Compact 3-Column Filter Grid (Bulan | Status | Bagian) */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
            {/* Filter Bulan */}
            <div className="flex items-center gap-1.5 bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 min-w-0">
              <CalendarRange className="w-3.5 h-3.5 text-teal-600 shrink-0" />
              <span className="text-[11px] font-bold text-slate-500 shrink-0">
                Bulan:
              </span>
              <select
                aria-label="Filter Bulan Review Monitoring"
                value={reviewMonthFilter}
                onChange={(e) => {
                  const mKey = e.target.value;
                  setReviewMonthFilter(mKey);
                  if (mKey !== 'ALL' && reviewWeeklyScope === 'SINGLE') {
                    const firstInMonth = ANNUAL_2026_WEEKS.find(
                      (w) => w.monthKey === mKey
                    );
                    if (firstInMonth) {
                      handleSelectWeeklyNumber(firstInMonth.weekNumber);
                    }
                  }
                }}
                className="flex-1 min-w-0 bg-transparent text-xs font-bold text-slate-800 focus:outline-none cursor-pointer truncate"
              >
                {MONTH_FILTER_OPTIONS.map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                  </option>
                ))}
              </select>
            </div>

            {/* Filter Status Isian */}
            <div className="flex items-center gap-1.5 bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 min-w-0">
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
              <span className="text-[11px] font-bold text-slate-500 shrink-0">
                Status:
              </span>
              <select
                aria-label="Filter Status Isian Review Monitoring"
                value={reviewStatusFilter}
                onChange={(e) => {
                  const nextStatus = e.target.value as 'ALL' | 'FILLED' | 'EMPTY';
                  setReviewStatusFilter(nextStatus);
                  if (nextStatus === 'FILLED' && reviewWeeklyScope === 'SINGLE') {
                    const firstFilledDoc = docs.find((d) => isReviewDocFilled(d));
                    const wk = firstFilledDoc
                      ? getReviewDocWeekNumber(firstFilledDoc)
                      : null;
                    if (wk) handleSelectWeeklyNumber(wk);
                  }
                }}
                className="flex-1 min-w-0 bg-transparent text-xs font-bold text-slate-800 focus:outline-none cursor-pointer truncate"
              >
                <option value="ALL">Semua Status (Terisi &amp; Kosong)</option>
                <option value="FILLED">
                  ✓ Sudah Terisi ({filledDocsList.length} Weekly)
                </option>
                <option value="EMPTY">Masih Kosong</option>
              </select>
            </div>

            {/* Filter Kategori / Bagian Evaluasi */}
            <div className="flex items-center gap-1.5 bg-white px-2.5 py-1.5 rounded-lg border border-slate-200 min-w-0">
              <Filter className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
              <span className="text-[11px] font-bold text-slate-500 shrink-0">
                Bagian:
              </span>
              <select
                aria-label="Filter Bagian Kategori Review"
                value={sectionFilter}
                onChange={(e) => setSectionFilter(e.target.value)}
                className="flex-1 min-w-0 bg-transparent text-xs font-bold text-slate-800 focus:outline-none cursor-pointer truncate"
              >
                <option value="ALL">Semua Bagian (A–D &amp; Kendala)</option>
                {activeDoc.sections.map((sec) => (
                  <option key={sec.id} value={sec.id}>
                    {sec.code}
                  </option>
                ))}
              </select>
              {(reviewMonthFilter !== 'ALL' ||
                reviewStatusFilter !== 'ALL' ||
                sectionFilter !== 'ALL' ||
                reviewWeeklyScope === 'ALL_WEEKS') && (
                <button
                  type="button"
                  onClick={() => {
                    setReviewMonthFilter('ALL');
                    setReviewStatusFilter('ALL');
                    setSectionFilter('ALL');
                    setReviewWeeklyScope('SINGLE');
                  }}
                  className="p-1 text-rose-600 hover:bg-rose-50 rounded shrink-0"
                  title="Reset Filter"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>
        </div>

        {/* ===================================================================
            PANEL EXPORT REVIEW DENGAN FILTER WEEKLY (INLINE COLLAPSIBLE)
            =================================================================== */}
        {isExportReviewOpen && !isPublicReadOnly && (
          <div className="p-4 bg-gradient-to-r from-amber-50/90 via-emerald-50/80 to-teal-50/80 border-b border-emerald-300 space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
                <span className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Pengaturan Export Review ke Excel (.xlsx) — Pilih Filter Weekly
                </span>
              </div>
              <button
                type="button"
                onClick={() => setIsExportReviewOpen(false)}
                className="p-1 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-white/60"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <button
                type="button"
                onClick={() => setExportReviewFilterMode('CURRENT_WEEK')}
                className={`p-2.5 rounded-xl border text-left text-xs transition-all ${
                  exportReviewFilterMode === 'CURRENT_WEEK'
                    ? 'bg-emerald-600 text-white border-emerald-600 font-bold shadow-2xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                <div className="font-extrabold">1. Sesuai Filter Weekly Aktif</div>
                <div
                  className={`text-[11px] mt-0.5 truncate ${
                    exportReviewFilterMode === 'CURRENT_WEEK'
                      ? 'text-emerald-100'
                      : 'text-slate-500'
                  }`}
                >
                  {reviewWeeklyScope === 'ALL_WEEKS'
                    ? `Semua Weekly (${visibleReviewDocs.length})`
                    : getReviewDocWeeklyLabel(activeDoc)}
                </div>
              </button>

              <button
                type="button"
                onClick={() => setExportReviewFilterMode('SPECIFIC_WEEK')}
                className={`p-2.5 rounded-xl border text-left text-xs transition-all ${
                  exportReviewFilterMode === 'SPECIFIC_WEEK'
                    ? 'bg-emerald-600 text-white border-emerald-600 font-bold shadow-2xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                <div className="font-extrabold">2. Pilih Weekly Tertentu</div>
                <div
                  className={`text-[11px] mt-0.5 ${
                    exportReviewFilterMode === 'SPECIFIC_WEEK'
                      ? 'text-emerald-100'
                      : 'text-slate-500'
                  }`}
                >
                  Pilih dari W1 s/d W53
                </div>
              </button>

              <button
                type="button"
                onClick={() => setExportReviewFilterMode('ALL_WEEKS')}
                className={`p-2.5 rounded-xl border text-left text-xs transition-all ${
                  exportReviewFilterMode === 'ALL_WEEKS'
                    ? 'bg-emerald-600 text-white border-emerald-600 font-bold shadow-2xs'
                    : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50'
                }`}
              >
                <div className="font-extrabold">3. Semua Weekly Review</div>
                <div
                  className={`text-[11px] mt-0.5 ${
                    exportReviewFilterMode === 'ALL_WEEKS'
                      ? 'text-emerald-100'
                      : 'text-slate-500'
                  }`}
                >
                  Gabungkan semua periode
                </div>
              </button>
            </div>

            {exportReviewFilterMode === 'SPECIFIC_WEEK' && (
              <div className="flex flex-wrap items-center gap-2 pt-1">
                <span className="text-xs font-bold text-slate-700">
                  Pilih Periode Weekly yang Di-export:
                </span>
                <select
                  aria-label="Pilih Weekly untuk Export Review"
                  value={String(exportReviewWeekNum)}
                  onChange={(e) => setExportReviewWeekNum(Number(e.target.value))}
                  className="px-3 py-1.5 text-xs font-bold bg-white border border-emerald-400 rounded-lg text-slate-900"
                >
                  {ANNUAL_2026_WEEKS.map((w) => {
                    const exDoc = docs.find(
                      (d) => getReviewDocWeekNumber(d) === w.weekNumber
                    );
                    const hasIsian = isReviewDocFilled(exDoc);
                    return (
                      <option key={w.weekNumber} value={String(w.weekNumber)}>
                        WEEKLY {w.weekNumber} ({w.periodeText})
                        {hasIsian ? ' ✓ [Ada Isian Review]' : ''}
                      </option>
                    );
                  })}
                </select>
              </div>
            )}

            <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-emerald-200/80">
              <label className="inline-flex items-center gap-2 text-xs font-semibold text-slate-800 cursor-pointer">
                <input
                  type="checkbox"
                  checked={includeSalesInReviewExport}
                  onChange={(e) => setIncludeSalesInReviewExport(e.target.checked)}
                  className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500"
                />
                <span>
                  Sertakan juga Sheet Sales Table Weekly (WPS) di file Excel yang sama
                </span>
              </label>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsExportReviewOpen(false)}
                  className="px-3 py-1.5 rounded-lg border border-slate-300 bg-white text-xs font-bold text-slate-700 hover:bg-slate-100"
                >
                  Batal
                </button>
                <button
                  type="button"
                  disabled={isExportingReview}
                  onClick={handleExecuteExportReviewExcel}
                  className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-extrabold shadow-xs cursor-pointer disabled:opacity-50"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>
                    {isExportingReview
                      ? 'Mengunduh Excel...'
                      : 'Unduh Hasil Export Review (.xlsx)'}
                  </span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Action Toolbar: Mode Tampilan (Tabel WPS / Form Isian / Terstruktur), Salin WA, Edit Semua */}
        <div className="p-3 bg-white flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2.5">
          {/* 3-Mode View Switcher: Full-width 3-column grid on mobile so nothing overflows */}
          <div className="grid grid-cols-3 sm:inline-flex items-center gap-1 p-1 bg-slate-100 rounded-xl border border-slate-200/80 w-full sm:w-auto">
            <button
              type="button"
              onClick={() => setViewMode('spreadsheet')}
              className={`inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 text-xs font-bold rounded-lg transition-colors truncate cursor-pointer ${
                viewMode === 'spreadsheet'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Table2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
              <span className="truncate">Tabel WPS</span>
            </button>
            {!isPublicReadOnly && (
              <button
                type="button"
                onClick={() => setViewMode('form')}
                className={`inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 text-xs font-bold rounded-lg transition-colors truncate cursor-pointer ${
                  viewMode === 'form'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'text-emerald-800 hover:bg-emerald-50'
                }`}
              >
                <ClipboardEdit className="w-3.5 h-3.5 shrink-0" />
                <span className="truncate">Form Isian</span>
              </button>
            )}
            <button
              type="button"
              onClick={() => setViewMode('structured')}
              className={`inline-flex items-center justify-center gap-1.5 px-2.5 py-1.5 text-xs font-bold rounded-lg transition-colors truncate cursor-pointer ${
                viewMode === 'structured'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <LayoutGrid className="w-3.5 h-3.5 shrink-0" />
              <span className="truncate">Terstruktur</span>
            </button>
          </div>

          {/* Right Quick Actions: Salin Teks WA & Edit Semua Baris */}
          <div className="flex items-center gap-2 w-full sm:w-auto">
            <button
              type="button"
              disabled={isPublicReadOnly}
              onClick={() => {
                if (isPublicReadOnly) return;
                handleCopyWhatsAppText();
              }}
              title={
                isPublicReadOnly
                  ? 'Mode Publik (Hanya Lihat): Tidak bisa menyalin/export laporan'
                  : 'Salin Teks WA'
              }
              className="flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3 py-1.5 text-xs font-bold text-slate-700 bg-slate-50 border border-slate-200 hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed rounded-lg transition-colors whitespace-nowrap cursor-pointer"
            >
              {copiedNotice ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                  <span>Tersalin (WA)</span>
                </>
              ) : (
                <>
                  {isPublicReadOnly ? (
                    <Lock className="w-3.5 h-3.5 shrink-0" />
                  ) : (
                    <Copy className="w-3.5 h-3.5 shrink-0" />
                  )}
                  <span>Salin Teks WA</span>
                </>
              )}
            </button>

            {!isPublicReadOnly && viewMode !== 'form' && (
              <button
                type="button"
                onClick={() => {
                  setEditMode(!editMode);
                  setEditingCellKey(null);
                }}
                className={`flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3.5 py-1.5 text-xs font-bold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
                  editMode
                    ? 'bg-emerald-600 text-white hover:bg-emerald-700'
                    : 'bg-indigo-600 text-white hover:bg-indigo-700'
                }`}
              >
                <Edit2 className="w-3.5 h-3.5 shrink-0" />
                <span>{editMode ? 'Selesai Edit' : 'Edit Semua Baris'}</span>
              </button>
            )}
          </div>
        </div>
      </div>

      {/* =====================================================================
          RINGKASAN FILTER CEPAT SAAT MODE SEMUA WEEKLY
          ===================================================================== */}
      {reviewWeeklyScope === 'ALL_WEEKS' && (
        <div className="no-print bg-emerald-50/70 border border-emerald-200 rounded-xl px-4 py-3 flex flex-wrap items-center justify-between gap-2.5">
          <div className="flex flex-wrap items-center gap-2 text-xs">
            <span className="font-extrabold text-emerald-950">
              Menampilkan {quickWeeklyItems.length} Periode Weekly:
            </span>
            <span className="px-2 py-0.5 rounded-full bg-emerald-600 text-white font-bold text-[11px]">
              ✓ {quickWeeklyItems.filter((i) => i.isReviewFilled).length} Terisi
            </span>
            <span className="px-2 py-0.5 rounded-full bg-amber-100 text-amber-900 border border-amber-300 font-bold text-[11px]">
              {quickWeeklyItems.filter((i) => !i.isReviewFilled).length} Kosong
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-lg border border-emerald-300">
              <Filter className="w-3.5 h-3.5 text-emerald-700" />
              <span className="text-[11px] font-bold text-slate-600">
                Lompat ke Weekly:
              </span>
              <select
                aria-label="Lompat ke Weekly Tertentu"
                value=""
                onChange={(e) => {
                  const wk = Number(e.target.value);
                  if (wk) handleSelectWeeklyNumber(wk);
                }}
                className="bg-transparent text-xs font-extrabold text-emerald-900 focus:outline-none cursor-pointer"
              >
                <option value="" disabled>
                  Pilih Weekly...
                </option>
                {quickWeeklyItems.map((item) => (
                  <option key={item.weekNumber} value={String(item.weekNumber)}>
                    WEEKLY {item.weekNumber} ({item.periodeText}){' '}
                    {item.isReviewFilled
                      ? `✓ [${item.filledItemsCount} Poin]`
                      : '[Kosong]'}
                  </option>
                ))}
              </select>
            </div>
          </div>
        </div>
      )}

      {/* =====================================================================
          MAIN CONTENT: RENDER SETIAP WEEKLY (FORM / STRUCTURED / SPREADSHEET)
          ===================================================================== */}
      {viewMode === 'form' ? (
        <div className="space-y-6">
          {visibleReviewDocs.map((docItem) => renderFormIsianWeeklyView(docItem))}
        </div>
      ) : viewMode === 'structured' ? (
        <div className="space-y-8">
          {visibleReviewDocs.map((docItem, dIdx) => {
            const wkNum =
              getReviewDocWeekNumber(docItem) || activeDocWeekNum || 40;
            const secsToRender =
              sectionFilter === 'ALL'
                ? docItem.sections
                : docItem.sections.filter((s) => s.id === sectionFilter);

            return (
              <div
                key={docItem.id || `struct-${wkNum}-${dIdx}`}
                className="space-y-4 bg-slate-100/70 p-4 rounded-2xl border border-slate-200"
              >
                {/* Document Header Card */}
                <div className="bg-white border border-slate-200 rounded-xl p-5 border-l-4 border-l-emerald-600 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div className="flex-1">
                    <div className="flex flex-wrap items-center gap-2 text-xs font-bold mb-1">
                      <span className="px-2 py-0.5 rounded bg-rose-600 text-white">
                        REPORT BAZAR MATAHARI
                      </span>
                      <span className="text-amber-800 bg-amber-100 border border-amber-300 px-2 py-0.5 rounded">
                        Tanjung
                      </span>
                      <span className="text-slate-600 font-semibold">
                        PT MDS Retailing Tbk
                      </span>
                      <span className="px-2 py-0.5 rounded bg-emerald-600 text-white font-extrabold">
                        WEEKLY {wkNum}
                      </span>
                    </div>
                    <input
                      type="text"
                      readOnly={isPublicReadOnly}
                      aria-label="Judul Laporan Review"
                      value={docItem.title}
                      onChange={(e) =>
                        updateDocByWeekNumber(wkNum, (d) => ({
                          ...d,
                          title: e.target.value,
                        }))
                      }
                      className="w-full text-lg sm:text-xl font-bold tracking-tight text-slate-950 bg-transparent focus:bg-slate-50 focus:outline-none rounded"
                    />
                    <div className="text-sm font-semibold text-emerald-800 mt-1">
                      {getReviewDocWeeklyLabel(docItem)}
                    </div>
                  </div>
                  {!isPublicReadOnly && (
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          handleSelectWeeklyNumber(wkNum);
                          setViewMode('form');
                        }}
                        className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg whitespace-nowrap cursor-pointer"
                      >
                        <ClipboardEdit className="w-3.5 h-3.5" />
                        <span>Buka Form Isian W{wkNum}</span>
                      </button>
                      {editMode && (
                        <button
                          type="button"
                          onClick={() => handleAddSectionForWeek(wkNum)}
                          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg whitespace-nowrap"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Tambah Kategori</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>

                {/* Structured Sections */}
                <div className="space-y-4">
                  {secsToRender.map((sec) => {
                    const hasSubRows = sec.subRows && sec.subRows.length > 0;

                    if (hasSubRows) {
                      return (
                        <section
                          key={sec.id}
                          className="bg-white border border-slate-200 rounded-xl overflow-hidden"
                        >
                          <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                            <span className="text-sm font-bold text-slate-950">
                              {sec.code}
                            </span>
                            {!isPublicReadOnly && editMode && (
                              <button
                                type="button"
                                onClick={() =>
                                  handleAddSubRowForWeek(wkNum, sec.id)
                                }
                                className="inline-flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-indigo-700 bg-indigo-50 hover:bg-indigo-100 rounded-md"
                              >
                                <Plus className="w-3 h-3" />
                                <span>Tambah Divisi / Sub-Baris</span>
                              </button>
                            )}
                          </div>

                          <div className="divide-y divide-slate-200">
                            {sec.subRows!.map((sub) => {
                              const joinedText = sub.items.join('\n');
                              const validLines = sub.items.filter(
                                (l) => (l || '').trim().length > 0
                              );
                              const rowCount = Math.max(2, sub.items.length);
                              const isFastMoving = sub.label
                                .toUpperCase()
                                .includes('FAST MOVING');
                              const isSlowMoving = sub.label
                                .toUpperCase()
                                .includes('SLOW MOVING');

                              return (
                                <div
                                  key={sub.id}
                                  className={`grid grid-cols-1 md:grid-cols-12 p-4 sm:px-5 gap-3 md:gap-6 items-start ${
                                    isFastMoving
                                      ? 'bg-emerald-50/20'
                                      : isSlowMoving
                                        ? 'bg-amber-50/20'
                                        : 'bg-white'
                                  }`}
                                >
                                  <div className="md:col-span-3 flex items-center justify-between">
                                    <span
                                      className={`text-sm font-bold ${
                                        sub.isItalic
                                          ? 'italic text-slate-800'
                                          : 'text-slate-950'
                                      }`}
                                    >
                                      {sub.label}
                                    </span>
                                  </div>

                                  <div className="md:col-span-9">
                                    {!isPublicReadOnly && editMode ? (
                                      <textarea
                                        value={joinedText}
                                        onChange={(e) =>
                                          handleSubRowItemChangeForWeek(
                                            wkNum,
                                            sec.id,
                                            sub.id,
                                            e.target.value
                                          )
                                        }
                                        placeholder={`Ketik isian ${sub.label} pada WEEKLY ${wkNum}...`}
                                        rows={rowCount}
                                        className="w-full text-xs sm:text-sm leading-relaxed text-slate-900 bg-slate-50 border border-slate-300 rounded-lg p-3 focus:outline-none focus:border-indigo-600"
                                      />
                                    ) : (
                                      <div
                                        onClick={() => {
                                          if (!isPublicReadOnly) setEditMode(true);
                                        }}
                                        className="space-y-1 text-xs sm:text-sm leading-relaxed text-slate-800 cursor-text hover:bg-slate-50 rounded-lg p-1.5 -m-1.5 transition-colors"
                                      >
                                        {validLines.length > 0 ? (
                                          validLines.map((line, lIdx) => (
                                            <div key={lIdx}>
                                              {formatReviewBulletItem(line) ||
                                                line}
                                            </div>
                                          ))
                                        ) : (
                                          <div className="text-xs italic text-slate-400">
                                            {isPublicReadOnly
                                              ? `- Belum ada isian ${sub.label} pada Weekly ${wkNum}`
                                              : `+ Klik untuk mengisi ${sub.label} pada Weekly ${wkNum}...`}
                                          </div>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </section>
                      );
                    }

                    const directJoined = (sec.directItems || []).join('\n');
                    const validDirect = (sec.directItems || []).filter(
                      (l) => (l || '').trim().length > 0
                    );
                    const directLinesCount = Math.max(
                      2,
                      (sec.directItems || []).length
                    );

                    return (
                      <section
                        key={sec.id}
                        className="bg-white border border-slate-200 rounded-xl overflow-hidden grid grid-cols-1 md:grid-cols-12"
                      >
                        <div className="md:col-span-3 px-5 py-4 bg-slate-50 border-b md:border-b-0 md:border-r border-slate-200 flex items-center">
                          <span className="text-sm font-bold text-slate-950">
                            {sec.code}
                          </span>
                        </div>
                        <div className="md:col-span-9 p-4 sm:px-5">
                          {!isPublicReadOnly && editMode ? (
                            <textarea
                              value={directJoined}
                              onChange={(e) =>
                                handleDirectItemsChangeForWeek(
                                  wkNum,
                                  sec.id,
                                  e.target.value
                                )
                              }
                              placeholder={`Ketik isian ${sec.code} pada WEEKLY ${wkNum}...`}
                              rows={directLinesCount}
                              className="w-full text-xs sm:text-sm leading-relaxed text-slate-900 bg-slate-50 border border-slate-300 rounded-lg p-3 focus:outline-none focus:border-indigo-600"
                            />
                          ) : (
                            <div
                              onClick={() => {
                                if (!isPublicReadOnly) setEditMode(true);
                              }}
                              className="space-y-1 text-xs sm:text-sm leading-relaxed text-slate-800 cursor-text hover:bg-slate-50 rounded-lg p-1.5 -m-1.5 transition-colors"
                            >
                              {validDirect.length > 0 ? (
                                validDirect.map((line, lIdx) => (
                                  <div key={lIdx}>
                                    {formatReviewBulletItem(line) || line}
                                  </div>
                                ))
                              ) : (
                                <div className="text-xs italic text-slate-400">
                                  {isPublicReadOnly
                                    ? `- Belum ada isian ${sec.code} pada Weekly ${wkNum}`
                                    : `+ Klik untuk mengisi ${sec.code} pada Weekly ${wkNum}...`}
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </section>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        /* VIEW MODE: AUTHENTIC WPS OFFICE SPREADSHEET SHEET (Renders active Weekly OR all Weeklies stacked!) */
        <div className="space-y-8">
          {visibleReviewDocs.map((docItem, dIdx) =>
            renderSpreadsheetTableForDoc(docItem, dIdx)
          )}
        </div>
      )}
    </div>
  );
};
