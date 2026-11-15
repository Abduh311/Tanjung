import React, { useState, useRef, useCallback, useEffect } from 'react';
import {
  WeeklySalesTable,
  DailySalesRecord,
  calculateDayMetrics,
  calculateWeekMetrics,
  formatNumberId,
  getWeeklyPeriodeText,
  generateSevenSlashDatesFromIso,
  getIndoDayName,
  isWeekendDate,
  getWeeklyDayCycleText,
  ANNUAL_2026_WEEKS,
  createWeeklyTableFromAnnualSchedule,
  getTableMonthKey,
} from '../data/initialData';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  saveAutoDraftToStorage,
  markAutoDraftClean,
  clearAutoDraftFromStorage,
} from '../utils/autoDraft';
import { AutoDraftStatusBadge } from './AutoDraftStatusBanner';
import {
  slashDateToIso,
  isoToSlashDate,
  shiftIsoDate,
  formatIndoDateLong,
} from '../utils/dateUtils';
import {
  Plus,
  Copy,
  Trash2,
  Edit3,
  Eraser,
  Check,
  X,
  Calendar,
  UserCheck,
  Users,
  Palette,
  Filter,
  Search,
  ChevronLeft,
  ChevronRight,
  CalendarRange,
  CalendarDays,
  Undo2,
  Redo2,
  Save,
} from 'lucide-react';
import {
  DailySalesSpreadsheetView,
  MonthlySalesSpreadsheetView,
  SalesTableHorizon,
} from './DailyAndMonthlySalesTables';

interface WeeklySalesSpreadsheetProps {
  tables: WeeklySalesTable[];
  onChangeTables: (updated: WeeklySalesTable[]) => void;
  showDivZero: boolean;
  onToggleDivZero: () => void;
  staffFilter: string;
  onChangeStaffFilter: (staff: string) => void;
  selectedWeeklyFilter: string; // 'ALL' or table.id
  onSelectWeeklyFilter: (filterId: string) => void;
  tableHorizon?: SalesTableHorizon;
  onChangeTableHorizon?: (mode: SalesTableHorizon) => void;
  canUndo?: boolean;
  canRedo?: boolean;
  undoCount?: number;
  redoCount?: number;
  onUndo?: () => void;
  onRedo?: () => void;
  onManualSave?: () => void;
  onOpenSaveTools?: () => void;
  onOpenReasonEvaluation?: (weekId: string) => void;
}

interface EditingDayModalState {
  tableId: string;
  dayIndex: number;
  dayData: DailySalesRecord;
}

interface PeriodeConfigModalState {
  tableId: string;
  weekNumber: number;
  title: string;
  startDateIso: string;
  dayCount: number;
  customPeriodeText: string;
}

const PRESET_PERIODS = [
  { label: 'Weekly 40 (Senin 28 Sep – Minggu 04 Okt)', weekNum: 40, startIso: '2026-09-28' },
  { label: 'Weekly 41 (Senin 05 Okt – Minggu 11 Okt)', weekNum: 41, startIso: '2026-10-05' },
  { label: 'Weekly 42 (Senin 12 Okt – Minggu 18 Okt)', weekNum: 42, startIso: '2026-10-12' },
  { label: 'Weekly 43 (Senin 19 Okt – Minggu 25 Okt)', weekNum: 43, startIso: '2026-10-19' },
  { label: 'Weekly 44 (Senin 26 Okt – Minggu 01 Nov)', weekNum: 44, startIso: '2026-10-26' },
];

const START_DAY_CYCLE_PRESETS = [
  { label: 'Siklus Senin – Minggu (Standar Tahunan)', startIso: '2026-10-05' },
  { label: 'Siklus Kamis – Rabu', startIso: '2026-10-01' },
  { label: 'Siklus Jumat – Kamis', startIso: '2026-10-02' },
  { label: 'Siklus Sabtu – Jumat', startIso: '2026-10-03' },
];

export const WeeklySalesSpreadsheet: React.FC<WeeklySalesSpreadsheetProps> = ({
  tables,
  onChangeTables,
  showDivZero,
  onToggleDivZero,
  staffFilter,
  onChangeStaffFilter,
  selectedWeeklyFilter,
  onSelectWeeklyFilter,
  tableHorizon,
  onChangeTableHorizon,
  canUndo = false,
  canRedo = false,
  undoCount = 0,
  redoCount = 0,
  onUndo,
  onRedo,
  onManualSave,
  onOpenSaveTools,
  onOpenReasonEvaluation,
}) => {
  const [internalHorizon, setInternalHorizon] = useState<SalesTableHorizon>('weekly');
  const activeHorizon: SalesTableHorizon = tableHorizon || internalHorizon;
  const handleSelectHorizon = (mode: SalesTableHorizon) => {
    setInternalHorizon(mode);
    if (onChangeTableHorizon) onChangeTableHorizon(mode);
  };
  const [editingModal, setEditingModal] = useState<EditingDayModalState | null>(null);
  const [periodeModal, setPeriodeModal] = useState<PeriodeConfigModalState | null>(null);
  const [isSalesDraftDirty, setIsSalesDraftDirty] = useState<boolean>(false);
  const [recoveredModalDraft, setRecoveredModalDraft] =
    useState<EditingDayModalState | null>(null);

  const handleTablesUpdateWithDraft = useCallback(
    (updated: WeeklySalesTable[]) => {
      setIsSalesDraftDirty(true);
      saveAutoDraftToStorage(
        AUTO_DRAFT_KEYS.SALES_MATRIX,
        `Tabel Sales (${activeHorizon.toUpperCase()})`,
        updated,
        true
      );
      onChangeTables(updated);
    },
    [onChangeTables, activeHorizon]
  );

  // Restore unsaved Quick Day Modal draft if connection dropped before "Simpan Perubahan" was pressed
  useEffect(() => {
    const savedModalEnv = loadAutoDraftFromStorage<EditingDayModalState>(
      AUTO_DRAFT_KEYS.QUICK_DAY_MODAL
    );
    if (savedModalEnv && savedModalEnv.isDirty && savedModalEnv.data?.dayData) {
      setRecoveredModalDraft(savedModalEnv.data);
    }
  }, []);

  const salesDraftTimeStr = '';
  const isSalesOnline = true;
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const [classicExcelColors, setClassicExcelColors] = useState<boolean>(false);
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'FILLED' | 'EMPTY'>('ALL');
  const [monthFilter, setMonthFilter] = useState<string>('ALL');
  const [rowViewFilter, setRowViewFilter] = useState<'ALL' | 'NO_REASON' | 'SUMMARY_ONLY'>('ALL');
  const [showDayMap, setShowDayMap] = useState<boolean>(false);
  const [tablePage, setTablePage] = useState<number>(1);
  const [dayOfWeekFilter, setDayOfWeekFilter] = useState<string>('ALL');
  const [calendarDateFilterIso, setCalendarDateFilterIso] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [modifiedCells, setModifiedCells] = useState<
    Record<string, { type: 'input' | 'calc'; seq: number }>
  >({});
  const [lastEditBanner, setLastEditBanner] = useState<{
    cellLabel: string;
    newValue: string;
    tableTitle: string;
  } | null>(null);
  const cellFlashClearTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const triggerCellFlash = useCallback(
    (
      inputKeys: string[],
      calcKeys: string[] = [],
      bannerInfo?: { cellLabel: string; newValue: string; tableTitle: string }
    ) => {
      if (bannerInfo) {
        setLastEditBanner(bannerInfo);
      }
      setModifiedCells((prev) => {
        const next = { ...prev };
        inputKeys.forEach((k) => {
          next[k] = { type: 'input', seq: (prev[k]?.seq || 0) + 1 };
        });
        calcKeys.forEach((k) => {
          next[k] = { type: 'calc', seq: (prev[k]?.seq || 0) + 1 };
        });
        return next;
      });

      // Single debounced timer to clear all cell highlights without spamming multiple renders
      if (cellFlashClearTimerRef.current) {
        clearTimeout(cellFlashClearTimerRef.current);
      }
      cellFlashClearTimerRef.current = setTimeout(() => {
        cellFlashClearTimerRef.current = null;
        setModifiedCells({});
        setLastEditBanner(null);
      }, 3200);
    },
    []
  );

  const getCellFlashClass = (cellKey: string): string => {
    const entry = modifiedCells[cellKey];
    if (!entry) return '';
    return entry.type === 'input' ? 'cell-just-modified' : 'cell-calc-updated';
  };

  // Reliably open any week 1..53 at position tablePositionIndex (0 = top, 1 = middle, 2 = bottom)
  // so that BOTH lower tables (+1, +2) AND upper tables (-1, -2) automatically follow the selected Total Weekly sequence!
  const handleOpenWeekNumber = (wkNum: number, tablePositionIndex: number = 0) => {
    if (!wkNum) return;
    setStatusFilter('ALL');
    setSearchQuery('');
    setTablePage(1);

    const safePos = Math.max(0, Math.min(2, tablePositionIndex));
    const topWeekNum = Math.max(1, Math.min(51, wkNum - safePos));

    let updatedTables = [...tables];
    let changed = false;
    let targetId = '';

    for (let offset = 0; offset <= 2; offset++) {
      const targetWk = topWeekNum + offset;
      if (targetWk >= 1 && targetWk <= 53) {
        const existing = updatedTables.find((t) => t.weekNumber === targetWk);
        if (existing) {
          if (offset === 0) targetId = existing.id;
        } else {
          const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === targetWk);
          if (sched) {
            const newTable = createWeeklyTableFromAnnualSchedule(sched);
            updatedTables.push(newTable);
            changed = true;
            if (offset === 0) targetId = newTable.id;
          }
        }
      }
    }

    if (changed) {
      updatedTables = updatedTables.sort((a, b) => a.weekNumber - b.weekNumber);
      handleTablesUpdateWithDraft(updatedTables);
    }

    if (targetId) {
      onSelectWeeklyFilter(targetId);
    }
  };

  const handleSelectWeeklyCalendarDate = (
    iso: string,
    tablePositionIndex: number = 0
  ) => {
    setCalendarDateFilterIso(iso);
    if (!iso) return;
    const slashTarget = isoToSlashDate(iso);
    const existing = tables.find((t) =>
      t.days.some((d) => d.tanggal.trim() === slashTarget)
    );
    if (existing) {
      handleOpenWeekNumber(existing.weekNumber, tablePositionIndex);
      return;
    }
    const sched =
      ANNUAL_2026_WEEKS.find((w) => w.slashDates.includes(slashTarget)) ||
      ANNUAL_2026_WEEKS.find(
        (w) => iso >= w.startDateIso && iso <= w.endDateIso
      );
    if (sched) {
      handleOpenWeekNumber(sched.weekNumber, tablePositionIndex);
    }
  };

  const handleDayFieldChange = (
    tableId: string,
    dayIndex: number,
    updater: (day: DailySalesRecord) => DailySalesRecord,
    flashMeta?: {
      fieldKey: string;
      cellLabel: string;
      newValue: string;
      affectsMetric?: 'sales' | 'qty' | 'trx' | 'none';
    }
  ) => {
    const targetTbl = tables.find((t) => t.id === tableId);
    const next = tables.map((t) => {
      if (t.id !== tableId) return t;
      const nextDays = t.days.map((d, idx) => (idx === dayIndex ? updater(d) : d));
      return { ...t, days: nextDays };
    });
    handleTablesUpdateWithDraft(next);

    if (flashMeta) {
      const inputKey = `${tableId}-d${dayIndex}-${flashMeta.fieldKey}`;
      const calcKeys: string[] = [];
      if (flashMeta.affectsMetric === 'sales') {
        calcKeys.push(
          `${tableId}-d${dayIndex}-totalSales`,
          `${tableId}-d${dayIndex}-pricePoint`,
          `${tableId}-d${dayIndex}-bigSize`,
          `${tableId}-weeklySales`,
          `${tableId}-weeklyPricePoint`,
          `${tableId}-weeklyBigSize`
        );
      } else if (flashMeta.affectsMetric === 'qty') {
        calcKeys.push(
          `${tableId}-d${dayIndex}-totalQty`,
          `${tableId}-d${dayIndex}-pricePoint`,
          `${tableId}-weeklyQty`,
          `${tableId}-weeklyPricePoint`
        );
      } else if (flashMeta.affectsMetric === 'trx') {
        calcKeys.push(
          `${tableId}-d${dayIndex}-totalTrx`,
          `${tableId}-d${dayIndex}-bigSize`,
          `${tableId}-weeklyTrx`,
          `${tableId}-weeklyBigSize`
        );
      }
      triggerCellFlash(
        [inputKey],
        calcKeys,
        {
          cellLabel: flashMeta.cellLabel,
          newValue: flashMeta.newValue,
          tableTitle: targetTbl?.title || `TOTAL WEEKLY`,
        }
      );
    }
  };

  const handleTableTitleChange = (tableId: string, newTitle: string) => {
    onChangeTables(
      tables.map((t) => (t.id === tableId ? { ...t, title: newTitle } : t))
    );
    triggerCellFlash([`${tableId}-title`], [], {
      cellLabel: 'Judul Tabel Weekly',
      newValue: newTitle,
      tableTitle: newTitle,
    });
  };

  const handleReasonSecondLabelChange = (tableId: string, newLabel: string) => {
    onChangeTables(
      tables.map((t) => (t.id === tableId ? { ...t, reasonSiangLabel: newLabel } : t))
    );
  };

  // Open Periode Weekly modal for a table
  const openPeriodeModal = (table: WeeklySalesTable) => {
    // Parse existing first day DD/MM/YYYY to YYYY-MM-DD
    const firstSlash = table.days[0]?.tanggal || '01/10/2026';
    const parts = firstSlash.split('/');
    const defaultIso =
      parts.length === 3
        ? `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`
        : '2026-10-01';

    setPeriodeModal({
      tableId: table.id,
      weekNumber: table.weekNumber,
      title: table.title,
      startDateIso: table.startDateIso || defaultIso,
      dayCount: table.days.length || 7,
      customPeriodeText: getWeeklyPeriodeText(table),
    });
  };

  // Apply Periode Weekly changes (auto-generate N days' dates from startDateIso)
  const applyPeriodeChanges = () => {
    if (!periodeModal) return;
    const { slashDates, periodeText } = generateSevenSlashDatesFromIso(
      periodeModal.startDateIso,
      periodeModal.dayCount
    );

    const next = tables.map((t) => {
      if (t.id !== periodeModal.tableId) return t;
      const updatedDays: DailySalesRecord[] = slashDates.map((tgl, idx) => {
        const existing = t.days[idx];
        if (existing) {
          return { ...existing, tanggal: tgl };
        }
        return {
          id: `${t.id}-d${idx + 1}-${Date.now()}`,
          tanggal: tgl,
          pagi: { nama: '', sales: null, qty: null, trx: null },
          siang: { nama: '', sales: null, qty: null, trx: null },
          reasonPagi: '',
          reasonSiang: '',
        };
      });

      return {
        ...t,
        weekNumber: periodeModal.weekNumber,
        title: periodeModal.title,
        startDateIso: periodeModal.startDateIso,
        periode: periodeText,
        days: updatedDays,
      };
    });

    onChangeTables(next);
    setPeriodeModal(null);
  };

  const handleAddWeek = () => {
    const existingNums = new Set(tables.map((t) => t.weekNumber));
    const currentNum =
      tables.find((t) => t.id === selectedWeeklyFilter)?.weekNumber || 41;
    let candidate = currentNum + 1;
    while (existingNums.has(candidate) && candidate <= 53) {
      candidate += 1;
    }
    if (candidate > 53) {
      candidate = 1;
      while (existingNums.has(candidate) && candidate <= 53) {
        candidate += 1;
      }
    }
    if (candidate <= 53) {
      handleOpenWeekNumber(candidate);
      return;
    }
    handleOpenWeekNumber(41);
  };

  const handleDuplicateWeek = (table: WeeklySalesTable) => {
    const maxWeek = tables.reduce((acc, t) => Math.max(acc, t.weekNumber), 40);
    const nextWeekNum = maxWeek + 1;
    const newId = `weekly-${Date.now()}`;
    const copy: WeeklySalesTable = {
      ...table,
      id: newId,
      weekNumber: nextWeekNum,
      title: `TOTAL WEEKLY ${nextWeekNum}`,
      days: table.days.map((d, i) => ({
        ...d,
        id: `w${nextWeekNum}-d${i + 1}-${Date.now()}`,
        pagi: { ...d.pagi },
        siang: { ...d.siang },
      })),
    };
    onChangeTables([...tables, copy]);
    onSelectWeeklyFilter(newId);
  };

  const handleClearWeekNumbers = (tableId: string) => {
    onChangeTables(
      tables.map((t) => {
        if (t.id !== tableId) return t;
        return {
          ...t,
          days: t.days.map((d) => ({
            ...d,
            pagi: { ...d.pagi, sales: null, qty: null, trx: null },
            siang: { ...d.siang, sales: null, qty: null, trx: null },
          })),
        };
      })
    );
  };

  const handleClearWeekNamesAndReasons = (tableId: string) => {
    onChangeTables(
      tables.map((t) => {
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
      })
    );
  };

  const handleDeleteWeek = (tableId: string) => {
    if (tables.length <= 1) {
      onChangeTables(
        tables.map((t) => ({
          ...t,
          days: t.days.map((d) => ({
            ...d,
            pagi: { nama: '', sales: null, qty: null, trx: null },
            siang: { nama: '', sales: null, qty: null, trx: null },
            reasonPagi: '',
            reasonSiang: '',
          })),
        }))
      );
      setConfirmDeleteId(null);
      return;
    }
    const remaining = tables.filter((t) => t.id !== tableId);
    onChangeTables(remaining);
    setConfirmDeleteId(null);
    onSelectWeeklyFilter(remaining[0]?.id || 'ALL');
  };

  const parseNumericInput = (raw: string): number | null => {
    const cleaned = raw.replace(/[^0-9]/g, '');
    if (cleaned === '') return null;
    const parsed = parseInt(cleaned, 10);
    return Number.isNaN(parsed) ? null : parsed;
  };

  const saveModalChanges = () => {
    if (!editingModal) return;
    const targetTbl = tables.find((t) => t.id === editingModal.tableId);
    const oldDay = targetTbl?.days[editingModal.dayIndex];
    const newDay = editingModal.dayData;
    const dIdx = editingModal.dayIndex;
    const tId = editingModal.tableId;

    const changedInputs: string[] = [];
    const changedCalcs = new Set<string>();

    if (oldDay) {
      if (oldDay.tanggal !== newDay.tanggal) changedInputs.push(`${tId}-d${dIdx}-tanggal`);
      if (oldDay.pagi.nama !== newDay.pagi.nama) changedInputs.push(`${tId}-d${dIdx}-pagi-nama`);
      if (oldDay.siang.nama !== newDay.siang.nama) changedInputs.push(`${tId}-d${dIdx}-siang-nama`);
      if (oldDay.pagi.sales !== newDay.pagi.sales) {
        changedInputs.push(`${tId}-d${dIdx}-pagi-sales`);
        changedCalcs.add(`${tId}-d${dIdx}-totalSales`);
        changedCalcs.add(`${tId}-d${dIdx}-pricePoint`);
        changedCalcs.add(`${tId}-d${dIdx}-bigSize`);
        changedCalcs.add(`${tId}-weeklySales`);
        changedCalcs.add(`${tId}-weeklyPricePoint`);
        changedCalcs.add(`${tId}-weeklyBigSize`);
      }
      if (oldDay.siang.sales !== newDay.siang.sales) {
        changedInputs.push(`${tId}-d${dIdx}-siang-sales`);
        changedCalcs.add(`${tId}-d${dIdx}-totalSales`);
        changedCalcs.add(`${tId}-d${dIdx}-pricePoint`);
        changedCalcs.add(`${tId}-d${dIdx}-bigSize`);
        changedCalcs.add(`${tId}-weeklySales`);
        changedCalcs.add(`${tId}-weeklyPricePoint`);
        changedCalcs.add(`${tId}-weeklyBigSize`);
      }
      if (oldDay.pagi.qty !== newDay.pagi.qty) {
        changedInputs.push(`${tId}-d${dIdx}-pagi-qty`);
        changedCalcs.add(`${tId}-d${dIdx}-totalQty`);
        changedCalcs.add(`${tId}-d${dIdx}-pricePoint`);
        changedCalcs.add(`${tId}-weeklyQty`);
        changedCalcs.add(`${tId}-weeklyPricePoint`);
      }
      if (oldDay.siang.qty !== newDay.siang.qty) {
        changedInputs.push(`${tId}-d${dIdx}-siang-qty`);
        changedCalcs.add(`${tId}-d${dIdx}-totalQty`);
        changedCalcs.add(`${tId}-d${dIdx}-pricePoint`);
        changedCalcs.add(`${tId}-weeklyQty`);
        changedCalcs.add(`${tId}-weeklyPricePoint`);
      }
      if (oldDay.pagi.trx !== newDay.pagi.trx) {
        changedInputs.push(`${tId}-d${dIdx}-pagi-trx`);
        changedCalcs.add(`${tId}-d${dIdx}-totalTrx`);
        changedCalcs.add(`${tId}-d${dIdx}-bigSize`);
        changedCalcs.add(`${tId}-weeklyTrx`);
        changedCalcs.add(`${tId}-weeklyBigSize`);
      }
      if (oldDay.siang.trx !== newDay.siang.trx) {
        changedInputs.push(`${tId}-d${dIdx}-siang-trx`);
        changedCalcs.add(`${tId}-d${dIdx}-totalTrx`);
        changedCalcs.add(`${tId}-d${dIdx}-bigSize`);
        changedCalcs.add(`${tId}-weeklyTrx`);
        changedCalcs.add(`${tId}-weeklyBigSize`);
      }
      if (oldDay.reasonPagi !== newDay.reasonPagi) changedInputs.push(`${tId}-d${dIdx}-reasonPagi`);
      if (oldDay.reasonSiang !== newDay.reasonSiang) changedInputs.push(`${tId}-d${dIdx}-reasonSiang`);
    }

    handleDayFieldChange(editingModal.tableId, editingModal.dayIndex, () => editingModal.dayData);
    if (changedInputs.length > 0) {
      triggerCellFlash(changedInputs, Array.from(changedCalcs), {
        cellLabel: `Input Harian (${newDay.tanggal})`,
        newValue: `${changedInputs.length} sel diperbarui`,
        tableTitle: targetTbl?.title || 'TOTAL WEEKLY',
      });
    }
    clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.QUICK_DAY_MODAL);
    setRecoveredModalDraft(null);
    setEditingModal(null);
  };

  // Apply Weekly Filter, Month Filter & Status Filter
  // When a specific starting Total Weekly is selected (e.g. TOTAL WEEKLY 1),
  // show that Total Weekly at the top and the next sequential Total Weekly tables (+1, +2) below it
  const sortedAllTables = [...tables].sort(
    (a, b) => a.weekNumber - b.weekNumber
  );

  const filteredTablesList = (() => {
    if (selectedWeeklyFilter !== 'ALL') {
      const startTable =
        sortedAllTables.find((t) => t.id === selectedWeeklyFilter) ||
        (() => {
          const m = selectedWeeklyFilter.match(/(\d+)/);
          if (!m) return undefined;
          const wk = Number(m[1]);
          const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
          return sched ? createWeeklyTableFromAnnualSchedule(sched) : undefined;
        })();
      if (startTable) {
        const startWk = startTable.weekNumber;
        const seq: WeeklySalesTable[] = [];
        for (let offset = 0; offset <= 2; offset++) {
          const wk = startWk + offset;
          if (wk <= 53) {
            const found = sortedAllTables.find((t) => t.weekNumber === wk);
            const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
            if (found) seq.push(found);
            else if (sched) seq.push(createWeeklyTableFromAnnualSchedule(sched));
          }
        }
        return seq;
      }
    }

    return sortedAllTables.filter((t) => {
      if (monthFilter !== 'ALL' && getTableMonthKey(t) !== monthFilter) {
        return false;
      }
      const wm = calculateWeekMetrics(t);
      if (statusFilter === 'FILLED' && wm.weeklySales <= 0) return false;
      if (statusFilter === 'EMPTY' && wm.weeklySales > 0) return false;
      return true;
    });
  })();

  // Show up to 12 sequential Weekly tables per page so all active Total Weekly tables (W40..W44) are visible without pagination cutoffs
  const tablesPerPage = 12;
  const totalTablePages =
    selectedWeeklyFilter === 'ALL'
      ? Math.max(1, Math.ceil(filteredTablesList.length / tablesPerPage))
      : 1;
  const safeTablePage = Math.min(tablePage, totalTablePages);
  const visibleTables =
    selectedWeeklyFilter === 'ALL'
      ? filteredTablesList.slice(
          (safeTablePage - 1) * tablesPerPage,
          safeTablePage * tablesPerPage
        )
      : filteredTablesList;

  const currentSelectedIdx = tables.findIndex((t) => t.id === selectedWeeklyFilter);
  const activeSingleTable =
    currentSelectedIdx !== -1 ? tables[currentSelectedIdx] : visibleTables[0] || tables[0];

  // Filter pill buttons by monthFilter and cap to 6 so the toolbar never wraps into a tall block
  const pillCandidateTables = tables.filter(
    (t) => monthFilter === 'ALL' || getTableMonthKey(t) === monthFilter
  );
  const compactPillTables = (() => {
    if (pillCandidateTables.length <= 6) return pillCandidateTables;
    const activeInCandidates = pillCandidateTables.findIndex(
      (t) => t.id === selectedWeeklyFilter
    );
    if (activeInCandidates === -1) return pillCandidateTables.slice(0, 6);
    const start = Math.max(
      0,
      Math.min(activeInCandidates - 2, pillCandidateTables.length - 6)
    );
    return pillCandidateTables.slice(start, start + 6);
  })();

  const annualDropdownWeeks = ANNUAL_2026_WEEKS.filter(
    (w) => monthFilter === 'ALL' || w.monthKey === monthFilter
  );

  // Styling classes based on Modern vs Classic WPS Excel theme
  const inputHeaderCellClass = classicExcelColors
    ? 'bg-[#FF0000] text-black font-semibold border border-slate-900'
    : 'bg-rose-700 text-white font-semibold border border-slate-300';

  const summaryHeaderCellClass = classicExcelColors
    ? 'bg-[#FFC000] text-black font-semibold border border-slate-900'
    : 'bg-amber-100 text-amber-950 font-semibold border border-slate-300';

  const totalWeeklyBannerClass = classicExcelColors
    ? 'bg-[#FFC000] text-slate-950 border border-slate-900'
    : 'bg-amber-400 text-slate-950 border border-slate-300';

  const totalWeeklyValueClass = classicExcelColors
    ? 'bg-[#FFC000] text-slate-950 border border-slate-900 font-bold'
    : 'bg-amber-100/90 text-slate-950 border border-slate-300 font-bold';

  const dataBorderClass = classicExcelColors
    ? 'border border-slate-900'
    : 'border border-slate-200';

  return (
    <div className="space-y-3.5 w-full max-w-full min-w-0">
      {/* Auto-Draft Banner for Sales Daily, Sales Weekly, and Sales Monthly */}
      <AutoDraftStatusBadge
        variant="banner"
        menuLabel={
          activeHorizon === 'harian'
            ? 'Sales Daily (Harian)'
            : activeHorizon === 'bulanan'
              ? 'Sales Monthly (Bulanan)'
              : 'Sales Weekly (Mingguan)'
        }
        lastDraftTimeStr={salesDraftTimeStr}
        isDirty={isSalesDraftDirty || Boolean(editingModal)}
        isOnline={isSalesOnline}
        hasRestorableDraft={Boolean(recoveredModalDraft)}
        restorableDraftSummary={
          recoveredModalDraft
            ? `Input Harian Cepat: ${recoveredModalDraft.dayData.tanggal}`
            : undefined
        }
        onRestoreDraft={() => {
          if (recoveredModalDraft) {
            setEditingModal(recoveredModalDraft);
            setRecoveredModalDraft(null);
          }
        }}
        onClearDraft={() => {
          setRecoveredModalDraft(null);
          clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.QUICK_DAY_MODAL);
        }}
        onManualSave={
          onManualSave
            ? () => {
                onManualSave();
                setIsSalesDraftDirty(false);
                markAutoDraftClean(AUTO_DRAFT_KEYS.SALES_MATRIX);
              }
            : undefined
        }
      />

      {/* MODE 1: SALES TABLE HARIAN */}
      {activeHorizon === 'harian' && (
        <DailySalesSpreadsheetView
          tables={tables}
          onChangeTables={handleTablesUpdateWithDraft}
          selectedWeeklyFilter={selectedWeeklyFilter}
          onSelectWeeklyFilter={onSelectWeeklyFilter}
          showDivZero={showDivZero}
          classicExcelColors={classicExcelColors}
          staffFilter={staffFilter}
          onChangeStaffFilter={onChangeStaffFilter}
          searchQuery={searchQuery}
          onChangeSearchQuery={setSearchQuery}
          onSelectHorizon={handleSelectHorizon}
          onToggleTheme={() => setClassicExcelColors(!classicExcelColors)}
          onToggleDivZero={onToggleDivZero}
          onManualSave={
            onManualSave
              ? () => {
                  onManualSave();
                  setIsSalesDraftDirty(false);
                  markAutoDraftClean(AUTO_DRAFT_KEYS.SALES_MATRIX);
                }
              : undefined
          }
        />
      )}

      {/* MODE 3: SALES TABLE BULANAN */}
      {activeHorizon === 'bulanan' && (
        <MonthlySalesSpreadsheetView
          tables={tables}
          onChangeTables={handleTablesUpdateWithDraft}
          onSelectWeeklyAndSwitchToWeeklyMode={(weekId) => {
            onSelectWeeklyFilter(weekId);
            handleSelectHorizon('weekly');
          }}
          showDivZero={showDivZero}
          classicExcelColors={classicExcelColors}
          onSelectHorizon={handleSelectHorizon}
          onToggleTheme={() => setClassicExcelColors(!classicExcelColors)}
          onToggleDivZero={onToggleDivZero}
          onManualSave={
            onManualSave
              ? () => {
                  onManualSave();
                  setIsSalesDraftDirty(false);
                  markAutoDraftClean(AUTO_DRAFT_KEYS.SALES_MATRIX);
                }
              : undefined
          }
        />
      )}

      {/* MODE 2: SALES TABLE WEEKLY */}
      {activeHorizon === 'weekly' && (
        <>

      {/* Pagination Banner when selectedWeeklyFilter === 'ALL' and multiple tables match */}
      {selectedWeeklyFilter === 'ALL' && totalTablePages > 1 && (
        <div className="no-print bg-indigo-50 border border-indigo-200 rounded-xl px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="font-semibold text-indigo-950">
            Mode Semua Periode: Menampilkan Tabel <strong>{safeTablePage}</strong> dari{' '}
            <strong>{totalTablePages}</strong> tabel mingguan (difilter agar tidak panjang ke bawah)
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              disabled={safeTablePage <= 1}
              onClick={() => setTablePage((p) => Math.max(1, p - 1))}
              className="px-2.5 py-1 rounded bg-white border border-indigo-200 font-bold text-indigo-700 disabled:opacity-40"
            >
              ◀ Sebelumnya
            </button>
            <span className="px-2.5 py-1 rounded bg-indigo-600 text-white font-bold font-mono-num">
              {safeTablePage} / {totalTablePages}
            </span>
            <button
              type="button"
              disabled={safeTablePage >= totalTablePages}
              onClick={() => setTablePage((p) => Math.min(totalTablePages, p + 1))}
              className="px-2.5 py-1 rounded bg-white border border-indigo-200 font-bold text-indigo-700 disabled:opacity-40"
            >
              Berikutnya ▶
            </button>
          </div>
        </div>
      )}

      {/* Empty Filter State */}
      {visibleTables.length === 0 && (
        <div className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-3">
          <div className="text-sm font-bold text-slate-900">
            Tidak ada Periode Weekly yang sesuai dengan filter saat ini
          </div>
          <p className="text-xs text-slate-500">
            Ubah pilihan Filter Periode Weekly atau Status Data untuk menampilkan seluruh tabel mingguan.
          </p>
          <button
            type="button"
            onClick={() => {
              onSelectWeeklyFilter('ALL');
              setStatusFilter('ALL');
              setSearchQuery('');
            }}
            className="px-4 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-lg"
          >
            Tampilkan Semua Periode Weekly
          </button>
        </div>
      )}

      {/* Render Each Filtered Weekly Table */}
      {visibleTables.map((table, tIdx) => {
        const weekMetrics = calculateWeekMetrics(table);
        const dayMetricsList = table.days.map((d) => calculateDayMetrics(d));
        const periodeLabel = getWeeklyPeriodeText(table);
        const dayCycleLabel = getWeeklyDayCycleText(table);

        // Calculate Weekday vs Weekend & filtered day metrics for this Weekly Periode
        let weekdaySales = 0;
        let weekdayDaysCount = 0;
        let weekendSales = 0;
        let weekendDaysCount = 0;
        let filteredDaySales = 0;
        let filteredDayQty = 0;
        let filteredDayTrx = 0;

        const dayMatchesFilter = (tanggal: string): boolean => {
          if (calendarDateFilterIso) {
            if (tanggal.trim() !== isoToSlashDate(calendarDateFilterIso)) {
              return false;
            }
          }
          if (dayOfWeekFilter === 'ALL') return true;
          const dName = getIndoDayName(tanggal);
          const wknd = isWeekendDate(tanggal);
          if (dayOfWeekFilter === 'WEEKDAY') return !wknd;
          if (dayOfWeekFilter === 'WEEKEND') return wknd;
          return dName.toLowerCase() === dayOfWeekFilter.toLowerCase();
        };

        table.days.forEach((d, idx) => {
          const dm = dayMetricsList[idx];
          if (isWeekendDate(d.tanggal)) {
            weekendSales += dm.totalSales;
            weekendDaysCount += 1;
          } else {
            weekdaySales += dm.totalSales;
            weekdayDaysCount += 1;
          }
          if (dayMatchesFilter(d.tanggal)) {
            filteredDaySales += dm.totalSales;
            filteredDayQty += dm.totalQty;
            filteredDayTrx += dm.totalTrx;
          }
        });

        return (
          <section
            key={table.id}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs print-break-after"
          >
            {/* Clean Integrated Dark Banner Header for Sales Table Weekly */}
            <div className="px-3.5 sm:px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white border-b border-slate-200 space-y-3 w-full max-w-full overflow-hidden">
              {/* Row 1: Title & Summary + Integrated Harian / Weekly / Bulanan Switcher */}
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
                      REPORT DP · SALES WEEKLY {table.weekNumber} — {periodeLabel}
                    </span>
                    <span className="px-2 py-0.5 text-[11px] font-semibold bg-indigo-500/30 text-indigo-200 border border-indigo-400/30 rounded-md">
                      {dayCycleLabel}
                    </span>
                  </div>
                  <div className="text-[11px] text-slate-300 font-mono-num mt-0.5">
                    Sales: Rp {formatNumberId(weekMetrics.weeklySales) || '0'} · Qty:{' '}
                    {formatNumberId(weekMetrics.weeklyQty) || '0'} pcs · Trx:{' '}
                    {formatNumberId(weekMetrics.weeklyTrx) || '0'} trx
                    {dayOfWeekFilter !== 'ALL' && (
                      <span className="ml-1.5 text-amber-300 font-bold">
                        | Filter ({dayOfWeekFilter}): Rp {formatNumberId(filteredDaySales) || '0'}
                      </span>
                    )}
                  </div>
                </div>

                <div className="no-print inline-flex items-center bg-white/10 p-1 rounded-xl border border-white/15 shrink-0">
                  <button
                    type="button"
                    onClick={() => handleSelectHorizon('harian')}
                    className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                  >
                    Sales Daily
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSelectHorizon('weekly')}
                    className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all bg-rose-600 text-white shadow-2xs"
                  >
                    Sales Weekly
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSelectHorizon('bulanan')}
                    className="px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all text-slate-200 hover:text-white hover:bg-white/10"
                  >
                    Sales Monthly
                  </button>
                </div>
              </div>

              {/* Row 2: Unified Weekly Switcher (◀ WEEKLY 41 ▼ ▶) + Filter Bulan + Action Controls */}
              <div className="no-print flex flex-wrap items-center justify-between gap-2 pt-2.5 border-t border-white/15 w-full max-w-full min-w-0">
                <div className="flex flex-wrap items-center gap-1.5 min-w-0 max-w-full">
                  <div className="inline-flex items-center gap-1 bg-white/10 p-0.5 rounded-lg border border-white/15 max-w-full min-w-0">
                    <button
                      type="button"
                      disabled={table.weekNumber <= 1}
                      onClick={() => handleOpenWeekNumber(Math.max(1, table.weekNumber - 1), tIdx)}
                      className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                      title="Weekly Sebelumnya"
                    >
                      ◀
                    </button>

                    <select
                      aria-label="Filter dan Pilih Weekly"
                      value={
                        selectedWeeklyFilter === 'ALL'
                          ? 'ALL'
                          : String(table.weekNumber)
                      }
                      onChange={(e) => {
                        const val = e.target.value;
                        if (val === 'ALL') {
                          onSelectWeeklyFilter('ALL');
                          setTablePage(1);
                          return;
                        }
                        handleOpenWeekNumber(Number(val), tIdx);
                      }}
                      className="px-2 py-1 rounded-md bg-amber-400 text-slate-950 text-xs font-extrabold focus:outline-none cursor-pointer min-w-0 max-w-[175px] sm:max-w-[230px] truncate shadow-2xs"
                    >
                      <option value="ALL">
                        Semua Total Weekly ({filteredTablesList.length})
                      </option>
                      {annualDropdownWeeks.map((w) => {
                        const ex = tables.find((t) => t.weekNumber === w.weekNumber);
                        const sls = ex ? calculateWeekMetrics(ex).weeklySales : 0;
                        return (
                          <option key={w.weekNumber} value={String(w.weekNumber)}>
                            TOTAL WEEKLY {w.weekNumber} ({w.seninSlash.slice(0, 5)}–{w.mingguSlash.slice(0, 5)}){' '}
                            {sls > 0 ? `[Rp ${formatNumberId(sls)}]` : ex ? '✓' : ''}
                          </option>
                        );
                      })}
                    </select>

                    <button
                      type="button"
                      disabled={table.weekNumber >= 53}
                      onClick={() => handleOpenWeekNumber(Math.min(53, table.weekNumber + 1), tIdx)}
                      className="px-2 py-1 rounded-md hover:bg-white/20 text-white text-xs font-bold disabled:opacity-35 transition-colors shrink-0"
                      title="Weekly Berikutnya"
                    >
                      ▶
                    </button>
                  </div>

                  {/* Tombol Cepat Pilihan Total Weekly Aktif */}
                  <div className="flex flex-wrap items-center gap-1">
                    {[40, 41, 42, 43, 44].map((wk) => {
                      const isCurrent =
                        table.weekNumber === wk && selectedWeeklyFilter !== 'ALL';
                      return (
                        <button
                          key={wk}
                          type="button"
                          onClick={() => handleOpenWeekNumber(wk)}
                          className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                            isCurrent
                              ? 'bg-amber-400 text-slate-950 font-extrabold shadow-sm ring-1 ring-amber-500'
                              : 'bg-white/10 text-white hover:bg-white/25 border border-white/20'
                          }`}
                          title={`Buka Total Weekly ${wk}`}
                        >
                          Weekly {wk}
                        </button>
                      );
                    })}
                    <button
                      type="button"
                      onClick={() => {
                        onSelectWeeklyFilter('ALL');
                        setTablePage(1);
                      }}
                      className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        selectedWeeklyFilter === 'ALL'
                          ? 'bg-amber-400 text-slate-950 font-extrabold shadow-sm ring-1 ring-amber-500'
                          : 'bg-white/10 text-white hover:bg-white/25 border border-white/20'
                      }`}
                      title="Tampilkan Semua Periode Weekly"
                    >
                      Semua
                    </button>
                  </div>

                  {/* Filter Bulan Dropdown */}
                  <select
                    aria-label="Filter Bulan pada Kartu Weekly"
                    value={monthFilter}
                    onChange={(e) => {
                      const m = e.target.value;
                      setMonthFilter(m);
                      setTablePage(1);
                      if (m !== 'ALL') {
                        const firstInMonth =
                          tables.find((t) => getTableMonthKey(t) === m) ||
                          ANNUAL_2026_WEEKS.find((w) => w.monthKey === m);
                        if (firstInMonth) {
                          handleOpenWeekNumber(firstInMonth.weekNumber);
                        }
                      }
                    }}
                    className="px-2.5 py-1.5 rounded-lg bg-white/10 hover:bg-white/15 border border-white/15 text-amber-200 text-xs font-bold focus:outline-none cursor-pointer min-w-0 max-w-[125px] truncate"
                  >
                    <option value="ALL" className="text-slate-900">Semua Bulan</option>
                    <option value="01" className="text-slate-900">Januari</option>
                    <option value="02" className="text-slate-900">Februari</option>
                    <option value="03" className="text-slate-900">Maret</option>
                    <option value="04" className="text-slate-900">April</option>
                    <option value="05" className="text-slate-900">Mei</option>
                    <option value="06" className="text-slate-900">Juni</option>
                    <option value="07" className="text-slate-900">Juli</option>
                    <option value="08" className="text-slate-900">Agustus</option>
                    <option value="09" className="text-slate-900">September</option>
                    <option value="10" className="text-slate-900">Oktober</option>
                    <option value="11" className="text-slate-900">November</option>
                    <option value="12" className="text-slate-900">Desember</option>
                  </select>

                  <div className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-white/10 border border-amber-400/40 text-xs">
                    <Calendar className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                    <input
                      type="date"
                      aria-label="Filter Kalender Tanggal Sales Weekly"
                      value={calendarDateFilterIso}
                      onChange={(e) =>
                        handleSelectWeeklyCalendarDate(e.target.value, tIdx)
                      }
                      title={
                        calendarDateFilterIso
                          ? `Menyorot tanggal ${formatIndoDateLong(calendarDateFilterIso)}`
                          : 'Pilih tanggal dari kalender untuk membuka & menyorot hari di Weekly'
                      }
                      className="bg-transparent text-white font-bold font-mono-num focus:outline-none cursor-pointer [color-scheme:dark]"
                    />
                    {calendarDateFilterIso && (
                      <button
                        type="button"
                        onClick={() => setCalendarDateFilterIso('')}
                        className="px-1.5 py-0.5 rounded bg-amber-400 text-slate-950 text-[10px] font-extrabold cursor-pointer shrink-0"
                        title="Hapus Filter Tanggal Kalender"
                      >
                        Semua
                      </button>
                    )}
                  </div>

                  <button
                    type="button"
                    onClick={handleAddWeek}
                    className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg transition-colors whitespace-nowrap"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    Weekly
                  </button>
                </div>
              </div>
            </div>

            {/* Clean Aligned Grid Filter Hari, Baris & Petugas (Tepat di Atas Tabel) */}
            <div className="no-print bg-slate-50 border-b border-slate-200 px-3.5 sm:px-4 py-2.5 space-y-2 w-full max-w-full overflow-hidden">
              {/* 2-Column on Mobile, 5-Column on Desktop so controls never stack awkwardly */}
              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 w-full">
                <select
                  aria-label="Filter Pilihan Total Weekly 1 s/d 53 pada Kartu"
                  value={
                    selectedWeeklyFilter === 'ALL'
                      ? `wk-${table.weekNumber}`
                      : `wk-${table.weekNumber}`
                  }
                  onChange={(e) => {
                    const val = e.target.value;
                    if (val === 'ALL') {
                      onSelectWeeklyFilter('ALL');
                      setTablePage(1);
                      return;
                    }
                    const wkNum = Number(val.replace('wk-', ''));
                    handleOpenWeekNumber(wkNum, tIdx);
                  }}
                  className="w-full px-2.5 py-1.5 rounded-lg bg-amber-50 border border-amber-400 text-xs font-extrabold text-amber-950 focus:outline-none focus:border-amber-600 cursor-pointer truncate"
                >
                  <option value="ALL">
                    📋 Semua Total Weekly ({tables.length})
                  </option>
                  {ANNUAL_2026_WEEKS.map((w) => {
                    const hasData = tables.some(
                      (t) => t.weekNumber === w.weekNumber
                    );
                    return (
                      <option key={w.weekNumber} value={`wk-${w.weekNumber}`}>
                        TOTAL WEEKLY {w.weekNumber} ({w.periodeText})
                        {hasData ? ' ★' : ''}
                      </option>
                    );
                  })}
                </select>

                <select
                  aria-label="Filter Hari pada Kartu Weekly"
                  value={dayOfWeekFilter}
                  onChange={(e) => setDayOfWeekFilter(e.target.value)}
                  className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
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

                <select
                  aria-label="Filter Tampilan Baris pada Kartu Weekly"
                  value={rowViewFilter}
                  onChange={(e) =>
                    setRowViewFilter(
                      e.target.value as 'ALL' | 'NO_REASON' | 'SUMMARY_ONLY'
                    )
                  }
                  className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-indigo-900 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
                >
                  <option value="ALL">Baris: Lengkap</option>
                  <option value="NO_REASON">Baris: Ringkas</option>
                  <option value="SUMMARY_ONLY">Baris: Hanya Total</option>
                </select>

                <select
                  aria-label="Filter Petugas pada Kartu Weekly"
                  value={staffFilter}
                  onChange={(e) => onChangeStaffFilter(e.target.value)}
                  className="w-full px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-800 focus:outline-none focus:border-indigo-600 cursor-pointer truncate"
                >
                  <option value="ALL">Petugas: Semua</option>
                  <option value="Gani">Petugas: Gani</option>
                  <option value="Irfan">Petugas: Irfan</option>
                  <option value="Feronika">Petugas: Feronika</option>
                  <option value="Abduh">Petugas: Abduh</option>
                  <option value="Fitri">Petugas: Fitri</option>
                </select>

                <button
                  type="button"
                  onClick={() => setShowDayMap((prev) => !prev)}
                  className={`w-full px-2.5 py-1.5 rounded-lg border text-xs font-bold transition-colors truncate ${
                    showDayMap
                      ? 'bg-indigo-600 text-white border-indigo-600'
                      : 'bg-white border-slate-300 text-indigo-700 hover:bg-indigo-50'
                  }`}
                >
                  {showDayMap ? 'Tutup 7 Hari ▲' : 'Kartu 7 Hari ▼'}
                </button>
              </div>

              {/* Side-by-side Weekday & Weekend Summary Pills */}
              <div className="grid grid-cols-2 gap-2 text-[11px] w-full">
                <button
                  type="button"
                  onClick={() =>
                    setDayOfWeekFilter(
                      dayOfWeekFilter === 'WEEKDAY' ? 'ALL' : 'WEEKDAY'
                    )
                  }
                  className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                    dayOfWeekFilter === 'WEEKDAY'
                      ? 'bg-indigo-600 text-white border-indigo-600'
                      : 'bg-indigo-50/80 text-indigo-900 border-indigo-200 hover:bg-indigo-100'
                  }`}
                >
                  Weekday ({weekdayDaysCount}H): Rp {formatNumberId(weekdaySales) || '0'} (
                  {weekMetrics.weeklySales > 0
                    ? Math.round((weekdaySales / weekMetrics.weeklySales) * 100)
                    : 0}
                  %)
                </button>
                <button
                  type="button"
                  onClick={() =>
                    setDayOfWeekFilter(
                      dayOfWeekFilter === 'WEEKEND' ? 'ALL' : 'WEEKEND'
                    )
                  }
                  className={`w-full px-2.5 py-1 rounded-lg border font-semibold font-mono-num text-center truncate transition-colors ${
                    dayOfWeekFilter === 'WEEKEND'
                      ? 'bg-rose-600 text-white border-rose-600'
                      : 'bg-rose-50/80 text-rose-900 border-rose-200 hover:bg-rose-100'
                  }`}
                >
                  Weekend ({weekendDaysCount}H): Rp {formatNumberId(weekendSales) || '0'} (
                  {weekMetrics.weeklySales > 0
                    ? Math.round((weekendSales / weekMetrics.weeklySales) * 100)
                    : 0}
                  %)
                </button>
              </div>

              {showDayMap && (
                <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2 pt-1">
                  {table.days.map((d, dIdx) => {
                    const dm = dayMetricsList[dIdx];
                    const dName = getIndoDayName(d.tanggal);
                    const wknd = isWeekendDate(d.tanggal);
                    const isFilteredMatch = dayMatchesFilter(d.tanggal);
                    const pct =
                      weekMetrics.weeklySales > 0
                        ? Math.round((dm.totalSales / weekMetrics.weeklySales) * 100)
                        : 0;

                    return (
                      <button
                        key={d.id}
                        type="button"
                        onClick={() =>
                          setDayOfWeekFilter(dayOfWeekFilter === dName ? 'ALL' : dName)
                        }
                        className={`p-2 rounded-lg border text-left transition-all ${
                          !isFilteredMatch
                            ? 'opacity-40 bg-slate-100 border-slate-200'
                            : dayOfWeekFilter === dName
                              ? 'ring-2 ring-indigo-600 bg-indigo-50/90 border-indigo-400'
                              : wknd
                                ? 'bg-rose-50/70 border-rose-200 hover:bg-rose-100/60'
                                : 'bg-white border-slate-200 hover:border-indigo-300'
                        }`}
                      >
                        <div className="flex items-center justify-between text-[10px] font-bold">
                          <span className={wknd ? 'text-rose-700' : 'text-indigo-700'}>
                            H{dIdx + 1} · {dName}
                          </span>
                          <span className="font-mono-num text-slate-500">{pct}%</span>
                        </div>
                        <div className="text-[11px] font-mono-num text-slate-600 mt-0.5">
                          {d.tanggal}
                        </div>
                        <div className="text-xs font-bold font-mono-num text-slate-900 mt-0.5 truncate">
                          {dm.totalSales > 0 ? `Rp ${formatNumberId(dm.totalSales)}` : 'Rp 0'}
                        </div>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Live Visual Indicator Banner when cells are recently modified */}
            {lastEditBanner && Object.keys(modifiedCells).length > 0 && (
              <div className="no-print bg-emerald-50/90 border-b border-emerald-300 px-3.5 sm:px-4 py-1.5 flex flex-wrap items-center justify-between gap-2 text-xs">
                <div className="flex items-center gap-2 text-emerald-950">
                  <span className="inline-block w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
                  <span className="font-bold">Sel Diubah:</span>
                  <span className="px-1.5 py-0.5 rounded bg-emerald-100 border border-emerald-300 font-bold text-emerald-900">
                    {lastEditBanner.cellLabel}
                  </span>
                  <span className="text-slate-600">→</span>
                  <span className="font-mono-num font-extrabold text-emerald-800">
                    {lastEditBanner.newValue || '(dikosongkan)'}
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

            {/* Mobile Swipe & Sticky Guide Hint */}
            <div className="flex sm:hidden items-center justify-between px-3 py-1.5 mb-1.5 bg-indigo-50 border border-indigo-200 rounded-lg text-[11px] text-indigo-900 font-medium">
              <span className="flex items-center gap-1">
                <span>👉</span> Geser ke samping untuk lihat seluruh hari &amp; total
              </span>
              <span className="font-bold text-indigo-700 bg-white px-2 py-0.5 rounded border border-indigo-200 shadow-2xs">
                Kolom Petugas Terkunci 📌
              </span>
            </div>

            {/* Spreadsheet Matrix Grid with Sticky Header & Column */}
            <div className="overflow-auto spreadsheet-scroll max-h-[82vh] relative border border-slate-300 rounded-xl shadow-xs bg-white overscroll-contain">
              <table className="w-full border-separate border-spacing-0 text-xs select-text min-w-[1420px]">
                <tbody>
                  {/* ROW HEADER WPS: TOTAL WEEKLY [N] ▼ (Filter Dropdown) | Periode : ... (Sticky Header) */}
                  <tr className="bg-slate-900 text-white sticky top-0 z-30">
                    <td className="bg-slate-900 border-b border-r-2 border-r-amber-400/90 border-slate-700 px-2 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky top-0 left-0 z-40 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.3)]">
                      <div className="flex items-center justify-between gap-1">
                        <select
                          aria-label={`Filter Total Weekly ${table.weekNumber}`}
                          value={String(table.weekNumber)}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val === 'ALL') {
                              onSelectWeeklyFilter('ALL');
                              setTablePage(1);
                            } else {
                              handleOpenWeekNumber(Number(val), tIdx);
                            }
                          }}
                          className="w-full bg-slate-900 text-amber-300 font-extrabold text-xs focus:outline-none cursor-pointer"
                          title="Klik tanda filter untuk pilih Total Weekly 1 s/d 53 (Tabel di atas & di bawah otomatis saling mengikuti urutan)"
                        >
                          {ANNUAL_2026_WEEKS.map((w) => (
                            <option
                              key={w.weekNumber}
                              value={String(w.weekNumber)}
                              className="bg-slate-900 text-amber-300 font-bold"
                            >
                              TOTAL WEEKLY {w.weekNumber} ▼
                            </option>
                          ))}
                        </select>
                      </div>
                    </td>
                    <td
                      colSpan={table.days.length * 2 + 1}
                      className="bg-slate-900 border-b border-slate-700 px-3 py-1.5 font-extrabold text-xs text-white text-left sticky top-0 z-30"
                    >
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <span>
                          Periode : {periodeLabel} ({dayCycleLabel})
                        </span>
                        <span className="text-[11px] font-semibold text-amber-300">
                          {tIdx === 0
                            ? `↕ Filter Urutan Dua Arah (Atas ⇄ Bawah): Mengikuti & diikuti TOTAL WEEKLY ${visibleTables[1]?.weekNumber ?? Math.min(53, table.weekNumber + 1)}`
                            : `↕ Mengikuti & Mengatur Urutan TOTAL WEEKLY ${visibleTables[tIdx - 1]?.weekNumber ?? Math.max(1, table.weekNumber - 1)} di atasnya`}
                        </span>
                      </div>
                    </td>
                  </tr>

                  {/* ROW 0: Hari (Urutan Hari dalam Weekly Periode) + TOTAL WEEKLY Merged Block (rowSpan=7) */}
                  <tr>
                    <td
                      className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}
                    >
                      Hari (Weekly)
                    </td>
                    {table.days.map((day, dIdx) => {
                      const dName = getIndoDayName(day.tanggal);
                      const wknd = isWeekendDate(day.tanggal);
                      const isDayActive = dayMatchesFilter(day.tanggal);
                      return (
                        <td
                          key={`${day.id}-hari`}
                          colSpan={2}
                          className={`${dataBorderClass} px-2 py-1.5 text-center font-bold min-w-[170px] transition-opacity ${
                            !isDayActive ? 'opacity-35' : ''
                          } ${
                            wknd
                              ? 'bg-rose-50 text-rose-800'
                              : 'bg-indigo-50/70 text-indigo-950'
                          }`}
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span className="px-1.5 py-0.2 rounded bg-slate-900/10 text-[10px] font-mono-num">
                              Hari {dIdx + 1}
                            </span>
                            <span>{dName}</span>
                            {wknd && (
                              <span className="px-1.5 py-0.2 rounded bg-rose-200/80 text-rose-900 text-[9px] uppercase font-bold">
                                Weekend
                              </span>
                            )}
                          </div>
                        </td>
                      );
                    })}
                    <td
                      rowSpan={rowViewFilter === 'SUMMARY_ONLY' ? 2 : 7}
                      className={`${totalWeeklyBannerClass} ${getCellFlashClass(`${table.id}-title`)} px-2 py-2 text-center align-middle w-[140px] min-w-[140px]`}
                    >
                      <select
                        aria-label="Filter Pilih Total Weekly 1 sampai 53"
                        value={String(table.weekNumber)}
                        onChange={(e) => {
                          const val = e.target.value;
                          if (val === 'ALL') {
                            onSelectWeeklyFilter('ALL');
                            setTablePage(1);
                          } else {
                            handleOpenWeekNumber(Number(val), tIdx);
                          }
                        }}
                        className="w-full mb-1 px-1.5 py-1 rounded bg-amber-500/30 hover:bg-amber-500/40 border border-amber-700/40 text-[11px] font-extrabold text-slate-950 text-center cursor-pointer focus:outline-none"
                        title="Klik untuk pilih Total Weekly 1 sampai 53"
                      >
                        <option value="ALL">Semua Weekly (1–53)</option>
                        {ANNUAL_2026_WEEKS.map((w) => (
                          <option key={w.weekNumber} value={String(w.weekNumber)}>
                            TOTAL WEEKLY {w.weekNumber}
                          </option>
                        ))}
                      </select>
                      <textarea
                        aria-label="Judul Total Weekly"
                        value={table.title}
                        onChange={(e) => handleTableTitleChange(table.id, e.target.value)}
                        rows={2}
                        className="w-full text-center font-bold text-sm leading-tight text-slate-950 bg-transparent resize-none focus:bg-amber-200/60 focus:outline-none rounded"
                      />
                      <div className="mt-1 text-[10px] font-semibold text-slate-800">
                        {dayCycleLabel}
                      </div>
                    </td>
                  </tr>

                  {/* ROW 1: Tanggal */}
                  <tr>
                    <td
                      className={`${inputHeaderCellClass} px-3 py-2 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}
                    >
                      Tanggal
                    </td>
                    {table.days.map((day, dIdx) => {
                      const q = searchQuery.trim().toLowerCase();
                      const isMatch =
                        q !== '' &&
                        (day.tanggal.toLowerCase().includes(q) ||
                          getIndoDayName(day.tanggal).toLowerCase().includes(q) ||
                          day.pagi.nama.toLowerCase().includes(q) ||
                          day.siang.nama.toLowerCase().includes(q) ||
                          day.reasonPagi.toLowerCase().includes(q) ||
                          day.reasonSiang.toLowerCase().includes(q));
                      const isDayActive = dayMatchesFilter(day.tanggal);

                      return (
                        <td
                          key={day.id}
                          colSpan={2}
                          className={`${
                            isMatch ? 'bg-amber-100' : 'bg-slate-50/70'
                          } ${
                            !isDayActive ? 'opacity-35' : ''
                          } ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-tanggal`)} px-2 py-1.5 text-center font-bold text-slate-900 min-w-[170px]`}
                        >
                          <div className="flex items-center justify-center gap-1">
                            <input
                              type="text"
                              aria-label={`Tanggal hari ke-${dIdx + 1}`}
                              value={day.tanggal}
                              onChange={(e) =>
                                handleDayFieldChange(
                                  table.id,
                                  dIdx,
                                  (d) => ({
                                    ...d,
                                    tanggal: e.target.value,
                                  }),
                                  {
                                    fieldKey: 'tanggal',
                                    cellLabel: `Tanggal Hari ke-${dIdx + 1}`,
                                    newValue: e.target.value,
                                    affectsMetric: 'none',
                                  }
                                )
                              }
                              className="w-28 text-center font-bold font-mono-num bg-transparent focus:bg-white focus:outline-1 focus:outline-indigo-600 rounded px-1 py-0.5"
                            />
                            <button
                              type="button"
                              onClick={() =>
                                setEditingModal({
                                  tableId: table.id,
                                  dayIndex: dIdx,
                                  dayData: {
                                    ...day,
                                    pagi: { ...day.pagi },
                                    siang: { ...day.siang },
                                  },
                                })
                              }
                              className="no-print opacity-60 hover:opacity-100 p-1 text-slate-600 hover:text-indigo-700 hover:bg-indigo-50 rounded transition-opacity"
                              title="Input Harian via Pop-up"
                            >
                              <Edit3 className="w-3 h-3" />
                            </button>
                          </div>
                        </td>
                      );
                    })}
                  </tr>

                  {/* ROW 2 to 6: Shift, Nama, Sales, Qty, Trx (Hidden in SUMMARY_ONLY mode) */}
                  {rowViewFilter !== 'SUMMARY_ONLY' && (
                    <>
                  {/* ROW 2: Shift */}
                  <tr>
                    <td className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Shift
                    </td>
                    {table.days.map((day) => (
                      <React.Fragment key={`${day.id}-shift`}>
                        <td
                          className={`bg-slate-50/40 ${dataBorderClass} px-2 py-1 text-center font-medium text-slate-600 w-[85px] min-w-[85px]`}
                        >
                          Pagi
                        </td>
                        <td
                          className={`bg-slate-50/40 ${dataBorderClass} px-2 py-1 text-center font-medium text-slate-600 w-[85px] min-w-[85px]`}
                        >
                          Siang
                        </td>
                      </React.Fragment>
                    ))}
                  </tr>

                  {/* ROW 3: Nama Petugas */}
                  <tr>
                    <td className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      <div className="flex items-center gap-1.5 font-bold">
                        <Users className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                        <span className="truncate">Nama Petugas</span>
                      </div>
                    </td>
                    {table.days.map((day, dIdx) => {
                      const highlightPagi =
                        staffFilter !== 'ALL' &&
                        day.pagi.nama.toLowerCase() === staffFilter.toLowerCase();
                      const highlightSiang =
                        staffFilter !== 'ALL' &&
                        day.siang.nama.toLowerCase() === staffFilter.toLowerCase();

                      return (
                        <React.Fragment key={`${day.id}-nama`}>
                          <td
                            className={`${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-pagi-nama`)} p-0 text-center ${
                              highlightPagi ? 'bg-amber-100 font-bold text-slate-950' : 'bg-white'
                            }`}
                          >
                            <input
                              type="text"
                              aria-label={`Nama Pagi ${day.tanggal}`}
                              value={day.pagi.nama}
                              onChange={(e) =>
                                handleDayFieldChange(
                                  table.id,
                                  dIdx,
                                  (d) => ({
                                    ...d,
                                    pagi: { ...d.pagi, nama: e.target.value },
                                  }),
                                  {
                                    fieldKey: 'pagi-nama',
                                    cellLabel: `Nama Pagi (${day.tanggal})`,
                                    newValue: e.target.value,
                                    affectsMetric: 'none',
                                  }
                                )
                              }
                              className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600"
                            />
                          </td>
                          <td
                            className={`${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-siang-nama`)} p-0 text-center ${
                              highlightSiang ? 'bg-amber-100 font-bold text-slate-950' : 'bg-white'
                            }`}
                          >
                            <input
                              type="text"
                              aria-label={`Nama Siang ${day.tanggal}`}
                              value={day.siang.nama}
                              onChange={(e) =>
                                handleDayFieldChange(
                                  table.id,
                                  dIdx,
                                  (d) => ({
                                    ...d,
                                    siang: { ...d.siang, nama: e.target.value },
                                  }),
                                  {
                                    fieldKey: 'siang-nama',
                                    cellLabel: `Nama Siang (${day.tanggal})`,
                                    newValue: e.target.value,
                                    affectsMetric: 'none',
                                  }
                                )
                              }
                              className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600"
                            />
                          </td>
                        </React.Fragment>
                      );
                    })}
                  </tr>

                  {/* ROW 4: Sales */}
                  <tr>
                    <td className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Sales (Rp)
                    </td>
                    {table.days.map((day, dIdx) => {
                      const highlightPagi =
                        staffFilter !== 'ALL' &&
                        day.pagi.nama.toLowerCase() === staffFilter.toLowerCase();
                      const highlightSiang =
                        staffFilter !== 'ALL' &&
                        day.siang.nama.toLowerCase() === staffFilter.toLowerCase();

                      return (
                        <React.Fragment key={`${day.id}-sales`}>
                          <td
                            className={`${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-pagi-sales`)} p-0 text-center font-mono-num ${
                              highlightPagi ? 'bg-amber-50 font-semibold' : 'bg-white'
                            }`}
                          >
                            <input
                              type="text"
                              inputMode="numeric"
                              aria-label={`Sales Pagi ${day.tanggal}`}
                              placeholder="-"
                              value={
                                day.pagi.sales !== null ? formatNumberId(day.pagi.sales) : ''
                              }
                              onChange={(e) => {
                                const parsed = parseNumericInput(e.target.value);
                                handleDayFieldChange(
                                  table.id,
                                  dIdx,
                                  (d) => ({
                                    ...d,
                                    pagi: {
                                      ...d.pagi,
                                      sales: parsed,
                                    },
                                  }),
                                  {
                                    fieldKey: 'pagi-sales',
                                    cellLabel: `Sales Pagi (${day.tanggal})`,
                                    newValue: parsed !== null ? `Rp ${formatNumberId(parsed)}` : '0',
                                    affectsMetric: 'sales',
                                  }
                                );
                              }}
                              className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                            />
                          </td>
                          <td
                            className={`${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-siang-sales`)} p-0 text-center font-mono-num ${
                              highlightSiang ? 'bg-amber-50 font-semibold' : 'bg-white'
                            }`}
                          >
                            <input
                              type="text"
                              inputMode="numeric"
                              aria-label={`Sales Siang ${day.tanggal}`}
                              placeholder="-"
                              value={
                                day.siang.sales !== null ? formatNumberId(day.siang.sales) : ''
                              }
                              onChange={(e) => {
                                const parsed = parseNumericInput(e.target.value);
                                handleDayFieldChange(
                                  table.id,
                                  dIdx,
                                  (d) => ({
                                    ...d,
                                    siang: {
                                      ...d.siang,
                                      sales: parsed,
                                    },
                                  }),
                                  {
                                    fieldKey: 'siang-sales',
                                    cellLabel: `Sales Siang (${day.tanggal})`,
                                    newValue: parsed !== null ? `Rp ${formatNumberId(parsed)}` : '0',
                                    affectsMetric: 'sales',
                                  }
                                );
                              }}
                              className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                            />
                          </td>
                        </React.Fragment>
                      );
                    })}
                  </tr>

                  {/* ROW 5: Qty */}
                  <tr>
                    <td className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Qty (pcs)
                    </td>
                    {table.days.map((day, dIdx) => (
                      <React.Fragment key={`${day.id}-qty`}>
                        <td
                          className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-pagi-qty`)} p-0 text-center font-mono-num`}
                        >
                          <input
                            type="text"
                            inputMode="numeric"
                            aria-label={`Qty Pagi ${day.tanggal}`}
                            placeholder="-"
                            value={day.pagi.qty !== null ? formatNumberId(day.pagi.qty) : ''}
                            onChange={(e) => {
                              const parsed = parseNumericInput(e.target.value);
                              handleDayFieldChange(
                                table.id,
                                dIdx,
                                (d) => ({
                                  ...d,
                                  pagi: {
                                    ...d.pagi,
                                    qty: parsed,
                                  },
                                }),
                                {
                                  fieldKey: 'pagi-qty',
                                  cellLabel: `Qty Pagi (${day.tanggal})`,
                                  newValue: parsed !== null ? `${formatNumberId(parsed)} pcs` : '0',
                                  affectsMetric: 'qty',
                                }
                              );
                            }}
                            className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                          />
                        </td>
                        <td
                          className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-siang-qty`)} p-0 text-center font-mono-num`}
                        >
                          <input
                            type="text"
                            inputMode="numeric"
                            aria-label={`Qty Siang ${day.tanggal}`}
                            placeholder="-"
                            value={day.siang.qty !== null ? formatNumberId(day.siang.qty) : ''}
                            onChange={(e) => {
                              const parsed = parseNumericInput(e.target.value);
                              handleDayFieldChange(
                                table.id,
                                dIdx,
                                (d) => ({
                                  ...d,
                                  siang: {
                                    ...d.siang,
                                    qty: parsed,
                                  },
                                }),
                                {
                                  fieldKey: 'siang-qty',
                                  cellLabel: `Qty Siang (${day.tanggal})`,
                                  newValue: parsed !== null ? `${formatNumberId(parsed)} pcs` : '0',
                                  affectsMetric: 'qty',
                                }
                              );
                            }}
                            className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                          />
                        </td>
                      </React.Fragment>
                    ))}
                  </tr>

                  {/* ROW 6: Struk */}
                  <tr>
                    <td className={`${inputHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Struk (Trx)
                    </td>
                    {table.days.map((day, dIdx) => (
                      <React.Fragment key={`${day.id}-trx`}>
                        <td
                          className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-pagi-trx`)} p-0 text-center font-mono-num`}
                        >
                          <input
                            type="text"
                            inputMode="numeric"
                            aria-label={`Trx Pagi ${day.tanggal}`}
                            placeholder="-"
                            value={day.pagi.trx !== null ? formatNumberId(day.pagi.trx) : ''}
                            onChange={(e) => {
                              const parsed = parseNumericInput(e.target.value);
                              handleDayFieldChange(
                                table.id,
                                dIdx,
                                (d) => ({
                                  ...d,
                                  pagi: {
                                    ...d.pagi,
                                    trx: parsed,
                                  },
                                }),
                                {
                                  fieldKey: 'pagi-trx',
                                  cellLabel: `Trx Pagi (${day.tanggal})`,
                                  newValue: parsed !== null ? `${formatNumberId(parsed)} trx` : '0',
                                  affectsMetric: 'trx',
                                }
                              );
                            }}
                            className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                          />
                        </td>
                        <td
                          className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-siang-trx`)} p-0 text-center font-mono-num`}
                        >
                          <input
                            type="text"
                            inputMode="numeric"
                            aria-label={`Trx Siang ${day.tanggal}`}
                            placeholder="-"
                            value={day.siang.trx !== null ? formatNumberId(day.siang.trx) : ''}
                            onChange={(e) => {
                              const parsed = parseNumericInput(e.target.value);
                              handleDayFieldChange(
                                table.id,
                                dIdx,
                                (d) => ({
                                  ...d,
                                  siang: {
                                    ...d.siang,
                                    trx: parsed,
                                  },
                                }),
                                {
                                  fieldKey: 'siang-trx',
                                  cellLabel: `Trx Siang (${day.tanggal})`,
                                  newValue: parsed !== null ? `${formatNumberId(parsed)} trx` : '0',
                                  affectsMetric: 'trx',
                                }
                              );
                            }}
                            className="w-full px-1.5 py-1.5 text-center bg-transparent focus:bg-indigo-50/50 focus:outline-1 focus:outline-indigo-600 placeholder:text-slate-300"
                          />
                        </td>
                      </React.Fragment>
                    ))}
                  </tr>
                    </>
                  )}

                  {/* ROW 7: Total Sales */}
                  <tr>
                    <td className={`${summaryHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-extrabold`}>
                      Total Sales
                    </td>
                    {dayMetricsList.map((m, idx) => (
                      <td
                        key={`${table.id}-tsales-${idx}`}
                        colSpan={2}
                        className={`bg-amber-50/25 ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${idx}-totalSales`)} px-2 py-1.5 text-center font-mono-num font-semibold text-slate-900`}
                      >
                        {formatNumberId(m.totalSales)}
                      </td>
                    ))}
                    <td
                      className={`${totalWeeklyValueClass} ${getCellFlashClass(`${table.id}-weeklySales`)} px-2 py-1.5 text-center font-mono-num`}
                    >
                      {formatNumberId(weekMetrics.weeklySales)}
                    </td>
                  </tr>

                  {/* ROW 8: Total Qty */}
                  <tr>
                    <td className={`${summaryHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Total Qty
                    </td>
                    {dayMetricsList.map((m, idx) => (
                      <td
                        key={`${table.id}-tqty-${idx}`}
                        colSpan={2}
                        className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${idx}-totalQty`)} px-2 py-1.5 text-center font-mono-num text-slate-800`}
                      >
                        {formatNumberId(m.totalQty)}
                      </td>
                    ))}
                    <td
                      className={`${totalWeeklyValueClass} ${getCellFlashClass(`${table.id}-weeklyQty`)} px-2 py-1.5 text-center font-mono-num`}
                    >
                      {formatNumberId(weekMetrics.weeklyQty)}
                    </td>
                  </tr>

                  {/* ROW 9: Total Trx */}
                  <tr>
                    <td className={`${summaryHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Total Trx
                    </td>
                    {dayMetricsList.map((m, idx) => (
                      <td
                        key={`${table.id}-ttrx-${idx}`}
                        colSpan={2}
                        className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${idx}-totalTrx`)} px-2 py-1.5 text-center font-mono-num text-slate-800`}
                      >
                        {formatNumberId(m.totalTrx)}
                      </td>
                    ))}
                    <td
                      className={`${totalWeeklyValueClass} ${getCellFlashClass(`${table.id}-weeklyTrx`)} px-2 py-1.5 text-center font-mono-num`}
                    >
                      {formatNumberId(weekMetrics.weeklyTrx)}
                    </td>
                  </tr>

                  {/* ROW 10: Price Point */}
                  <tr>
                    <td className={`${summaryHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Price Point
                    </td>
                    {dayMetricsList.map((m, idx) => (
                      <td
                        key={`${table.id}-pp-${idx}`}
                        colSpan={2}
                        className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${idx}-pricePoint`)} px-2 py-1.5 text-center font-mono-num text-slate-800`}
                      >
                        {m.pricePoint !== null
                          ? formatNumberId(m.pricePoint)
                          : showDivZero
                            ? '#DIV/0!'
                            : '0'}
                      </td>
                    ))}
                    <td
                      className={`${totalWeeklyValueClass} ${getCellFlashClass(`${table.id}-weeklyPricePoint`)} px-2 py-1.5 text-center font-mono-num`}
                    >
                      {weekMetrics.weeklyPricePoint !== null
                        ? formatNumberId(weekMetrics.weeklyPricePoint)
                        : showDivZero
                          ? '#DIV/0!'
                          : '0'}
                    </td>
                  </tr>

                  {/* ROW 11: Big Size */}
                  <tr>
                    <td className={`${summaryHeaderCellClass} px-3 py-1.5 w-[140px] min-w-[140px] max-w-[140px] sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90 font-bold`}>
                      Big Size
                    </td>
                    {dayMetricsList.map((m, idx) => (
                      <td
                        key={`${table.id}-bs-${idx}`}
                        colSpan={2}
                        className={`bg-white ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${idx}-bigSize`)} px-2 py-1.5 text-center font-mono-num text-slate-800`}
                      >
                        {m.bigSize !== null
                          ? formatNumberId(m.bigSize)
                          : showDivZero
                            ? '#DIV/0!'
                            : '0'}
                      </td>
                    ))}
                    <td
                      className={`${totalWeeklyValueClass} ${getCellFlashClass(`${table.id}-weeklyBigSize`)} px-2 py-1.5 text-center font-mono-num`}
                    >
                      {weekMetrics.weeklyBigSize !== null
                        ? formatNumberId(weekMetrics.weeklyBigSize)
                        : showDivZero
                          ? '#DIV/0!'
                          : '0'}
                    </td>
                  </tr>

                  {/* ROW 12 & 13: Reason / Event Shift Pagi & Siang (Shown only when rowViewFilter === 'ALL') */}
                  {rowViewFilter === 'ALL' && (
                    <>
                  {/* ROW 12: Reason / Event Shift Pagi */}
                  <tr>
                    <td
                      className={`${summaryHeaderCellClass} px-3 py-2 w-[140px] min-w-[140px] max-w-[140px] align-middle leading-snug sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90`}
                    >
                      <div className="font-bold text-slate-950">Reason / Event</div>
                      <div className="inline-flex items-center gap-1 mt-0.5 px-1.5 py-0.5 rounded bg-amber-300/80 text-amber-950 text-[10px] font-bold">
                        Shift Pagi
                      </div>
                    </td>
                    {table.days.map((day, dIdx) => (
                      <td
                        key={`${day.id}-rpagi`}
                        colSpan={2}
                        className={`bg-amber-50/15 ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-reasonPagi`)} p-0 align-top`}
                      >
                        <div className="px-2 pt-1.5 pb-0.5 flex items-center justify-between text-[10px] text-amber-900/80 border-b border-amber-100/80 bg-amber-50/40">
                          <span className="font-bold">Pagi: {day.pagi.nama || '-'}</span>
                          <span className="font-mono-num">
                            {day.pagi.sales !== null
                              ? `Rp ${formatNumberId(day.pagi.sales)}`
                              : '-'}
                          </span>
                        </div>
                        <textarea
                          aria-label={`Reason Shift Pagi ${day.tanggal}`}
                          value={day.reasonPagi}
                          onChange={(e) =>
                            handleDayFieldChange(
                              table.id,
                              dIdx,
                              (d) => ({
                                ...d,
                                reasonPagi: e.target.value,
                              }),
                              {
                                fieldKey: 'reasonPagi',
                                cellLabel: `Reason Pagi (${day.tanggal})`,
                                newValue: e.target.value,
                                affectsMetric: 'none',
                              }
                            )
                          }
                          rows={4}
                          placeholder="Isi Reason / Event Shift Pagi..."
                          className="w-full min-h-[76px] p-2 text-[11px] leading-relaxed text-slate-700 bg-transparent resize-y focus:bg-white focus:outline-1 focus:outline-amber-500"
                        />
                      </td>
                    ))}
                    <td className="bg-slate-50/40 border-none" />
                  </tr>

                  {/* ROW 13: Reason / Event Shift Siang */}
                  <tr>
                    <td
                      className={`${summaryHeaderCellClass} px-3 py-2 w-[140px] min-w-[140px] max-w-[140px] align-middle leading-snug sticky left-0 z-20 shadow-[3px_0_6px_-1px_rgba(0,0,0,0.22)] border-r-2 border-r-slate-400/90`}
                    >
                      <div className="font-bold text-slate-950">Reason / Event</div>
                      <div className="inline-flex items-center gap-1 mt-0.5 px-1.5 py-0.5 rounded bg-indigo-200 text-indigo-950 text-[10px] font-bold">
                        Shift Siang
                      </div>
                    </td>
                    {table.days.map((day, dIdx) => (
                      <td
                        key={`${day.id}-rsiang`}
                        colSpan={2}
                        className={`bg-indigo-50/15 ${dataBorderClass} ${getCellFlashClass(`${table.id}-d${dIdx}-reasonSiang`)} p-0 align-top`}
                      >
                        <div className="px-2 pt-1.5 pb-0.5 flex items-center justify-between text-[10px] text-indigo-900/80 border-b border-indigo-100/80 bg-indigo-50/40">
                          <span className="font-bold">Siang: {day.siang.nama || '-'}</span>
                          <span className="font-mono-num">
                            {day.siang.sales !== null
                              ? `Rp ${formatNumberId(day.siang.sales)}`
                              : '-'}
                          </span>
                        </div>
                        <textarea
                          aria-label={`Reason Shift Siang ${day.tanggal}`}
                          value={day.reasonSiang}
                          onChange={(e) =>
                            handleDayFieldChange(
                              table.id,
                              dIdx,
                              (d) => ({
                                ...d,
                                reasonSiang: e.target.value,
                              }),
                              {
                                fieldKey: 'reasonSiang',
                                cellLabel: `Reason Siang (${day.tanggal})`,
                                newValue: e.target.value,
                                affectsMetric: 'none',
                              }
                            )
                          }
                          rows={4}
                          placeholder="Isi Reason / Event Shift Siang..."
                          className="w-full min-h-[76px] p-2 text-[11px] leading-relaxed text-slate-700 bg-transparent resize-y focus:bg-white focus:outline-1 focus:outline-indigo-600"
                        />
                      </td>
                    ))}
                    <td className="bg-slate-50/40 border-none" />
                  </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>

            {/* ===================================================================
                GRAFIK DIAGRAM SALES TABLE WEEKLY (HANYA DIAGRAM SALES)
               =================================================================== */}
            {(() => {
              const totalPagiWeek = table.days.reduce((s, d) => s + (d.pagi.sales ?? 0), 0);
              const totalSiangWeek = table.days.reduce((s, d) => s + (d.siang.sales ?? 0), 0);
              const pagiPct =
                weekMetrics.weeklySales > 0
                  ? Math.round((totalPagiWeek / weekMetrics.weeklySales) * 100)
                  : 0;
              const siangPct = weekMetrics.weeklySales > 0 ? 100 - pagiPct : 0;
              const rawMaxDay = Math.max(
                0,
                ...dayMetricsList.map((m) => m.totalSales)
              );
              const maxDaySales = rawMaxDay > 0 ? rawMaxDay : 2000000;

              return (
                <div className="mt-4 mb-3 rounded-lg overflow-hidden border border-slate-300 bg-slate-100 shadow-xs">
                  {/* Dark Navy Header Banner matching attached photo */}
                  <div className="bg-slate-900 px-4 py-3 sm:px-5 space-y-1">
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <div className="text-xs sm:text-sm font-extrabold text-amber-400 tracking-tight">
                        DIAGRAM SALES WEEKLY — {table.title} ({periodeLabel})
                      </div>
                      <select
                        aria-label={`Filter Weekly Diagram Sales ${table.weekNumber}`}
                        value={String(table.weekNumber)}
                        onChange={(e) => {
                          handleOpenWeekNumber(Number(e.target.value), tIdx);
                        }}
                        className="no-print px-2 py-0.5 rounded bg-slate-800 border border-amber-400/70 text-[11px] font-extrabold text-amber-300 focus:outline-none focus:border-amber-400 cursor-pointer"
                        title="Pilih Total Weekly 1 s/d 53 — Grafik Diagram Sales otomatis mengikuti Filter Weekly yang dipilih"
                      >
                        {ANNUAL_2026_WEEKS.map((w) => (
                          <option key={w.weekNumber} value={String(w.weekNumber)}>
                            TOTAL WEEKLY {w.weekNumber} ▼ ({w.periodeText})
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="text-[11px] font-bold text-slate-300">
                      Diagram Batang Vertikal 7 Hari &amp; Garis Tren Penjualan Shift Pagi vs Shift Siang
                    </div>
                    <div className="flex flex-wrap items-center gap-4 pt-1 text-[11px] font-bold">
                      <span className="inline-flex items-center gap-1.5 text-amber-200">
                        <span className="w-2.5 h-2.5 bg-amber-500 inline-block" />
                        Shift Pagi: Rp {formatNumberId(totalPagiWeek) || '0'} ({pagiPct}%)
                      </span>
                      <span className="inline-flex items-center gap-1.5 text-indigo-200">
                        <span className="w-2.5 h-2.5 bg-indigo-600 inline-block" />
                        Shift Siang: Rp {formatNumberId(totalSiangWeek) || '0'} ({siangPct}%)
                      </span>
                      <span className="inline-flex items-center gap-1.5 text-emerald-300">
                        <span className="w-3.5 h-0.5 bg-emerald-500 inline-block" />
                        Garis Tren Total Sales
                      </span>
                    </div>
                  </div>

                  {/* Inner White Card matching attached photo */}
                  <div className="p-3 sm:p-4">
                    <div className="w-full bg-white border border-slate-300 rounded-lg p-3.5 sm:p-5 shadow-2xs space-y-3">
                      <div className="flex flex-wrap items-start justify-between gap-2">
                        <div className="space-y-1.5">
                          <span className="inline-block px-2.5 py-0.5 rounded bg-amber-600 text-white text-[10px] font-extrabold tracking-wide uppercase">
                            DIAGRAM SALES (OMZET RUPIAH)
                          </span>
                          <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                            Grafik Batang Shift Pagi vs Siang &amp; Garis Tren
                          </div>
                        </div>
                        <span className="px-3 py-1 rounded bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-extrabold font-mono-num">
                          Total: Rp {formatNumberId(weekMetrics.weeklySales) || '0'}
                        </span>
                      </div>

                      {/* Perfectly Aligned Plot Area (Y-Axis + Bars + SVG Trendline) */}
                      <div className="pt-5 pb-9 pl-12 pr-2 select-none">
                        <div className="relative h-40 w-full">
                          {/* Y-Axis Gridlines */}
                          {[1, 0.5, 0].map((ratio) => {
                            const val = Math.round(maxDaySales * ratio);
                            return (
                              <div
                                key={ratio}
                                style={{ bottom: `${ratio * 100}%` }}
                                className="absolute -left-12 right-0 flex items-center pointer-events-none translate-y-1/2"
                              >
                                <span className="w-11 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-500">
                                  {val >= 1000000
                                    ? `${(val / 1000000).toFixed(1)}jt`
                                    : val > 0
                                      ? `${Math.round(val / 1000)}rb`
                                      : '0'}
                                </span>
                                <div
                                  className={`flex-1 border-b ${
                                    ratio === 0
                                      ? 'border-solid border-slate-300'
                                      : 'border-dashed border-slate-200'
                                  }`}
                                />
                              </div>
                            );
                          })}

                          {/* SVG Trend Line Overlay (Exact 0..100% match with bar heights) */}
                          <svg
                            viewBox="0 0 700 200"
                            preserveAspectRatio="none"
                            className="absolute inset-0 w-full h-full pointer-events-none z-10 overflow-visible"
                          >
                            {(() => {
                              const pts = dayMetricsList.map((m, i) => {
                                const x = (i + 0.5) * (700 / 7);
                                const y =
                                  200 -
                                  Math.max(
                                    0,
                                    Math.min(200, (m.totalSales / maxDaySales) * 200)
                                  );
                                return { x, y, val: m.totalSales };
                              });
                              const poly = pts.map((p) => `${p.x},${p.y}`).join(' ');
                              return (
                                <>
                                  <polyline
                                    fill="none"
                                    stroke="#10b981"
                                    strokeWidth="2.5"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                    points={poly}
                                  />
                                  {pts.map((p, idx) => (
                                    <circle
                                      key={idx}
                                      cx={p.x}
                                      cy={p.y}
                                      r="4"
                                      fill="#ffffff"
                                      stroke="#059669"
                                      strokeWidth="2.2"
                                    />
                                  ))}
                                </>
                              );
                            })()}
                          </svg>

                          {/* 7 Vertical Stacked Columns */}
                          <div className="relative z-0 h-full grid grid-cols-7 gap-2 items-end">
                            {table.days.map((d, idx) => {
                              const m = dayMetricsList[idx];
                              const ps = d.pagi.sales ?? 0;
                              const ss = d.siang.sales ?? 0;
                              const hPct =
                                maxDaySales > 0
                                  ? Math.max(
                                      m.totalSales > 0 ? 6 : 2,
                                      Math.round((m.totalSales / maxDaySales) * 100)
                                    )
                                  : 2;
                              const pagiShare =
                                m.totalSales > 0 ? Math.round((ps / m.totalSales) * 100) : 0;
                              const siangShare = m.totalSales > 0 ? 100 - pagiShare : 0;
                              const wknd = isWeekendDate(d.tanggal);
                              const compactNominal =
                                m.totalSales >= 1000000
                                  ? `${(m.totalSales / 1000000).toFixed(1)}jt`
                                  : m.totalSales > 0
                                    ? `${Math.round(m.totalSales / 1000)}rb`
                                    : '0';

                              return (
                                <div
                                  key={d.id}
                                  className="h-full flex flex-col justify-end items-center group relative"
                                >
                                  <div
                                    style={{ height: `${hPct}%` }}
                                    className="w-full max-w-[44px] rounded-t-xs overflow-hidden flex flex-col transition-all relative"
                                    title={`${getIndoDayName(d.tanggal)} (${d.tanggal}) — Total: Rp ${formatNumberId(m.totalSales)} (Pagi: Rp ${formatNumberId(ps)} | Siang: Rp ${formatNumberId(ss)})`}
                                  >
                                    <div
                                      style={{ height: `${siangShare}%` }}
                                      className="w-full bg-indigo-600 transition-all"
                                    />
                                    <div
                                      style={{ height: `${pagiShare}%` }}
                                      className="w-full bg-amber-500 transition-all"
                                    />
                                  </div>

                                  {/* Top Value Label positioned cleanly above the bar top */}
                                  <div
                                    style={{ bottom: `calc(${hPct}% + 5px)` }}
                                    className="absolute inset-x-0 text-center pointer-events-none z-20"
                                  >
                                    <span className="text-[10px] font-mono-num font-extrabold text-slate-900 whitespace-nowrap">
                                      {compactNominal}
                                    </span>
                                  </div>

                                  {/* X-Axis Label: "Sen (H1)" + "28/09" */}
                                  <div className="absolute -bottom-8 inset-x-0 text-center">
                                    <div
                                      className={`text-[10px] font-extrabold truncate ${
                                        wknd ? 'text-rose-700' : 'text-slate-700'
                                      }`}
                                    >
                                      {getIndoDayName(d.tanggal).slice(0, 3)} (H{idx + 1})
                                    </div>
                                    <div className="text-[9px] font-mono-num font-semibold text-slate-500">
                                      {d.tanggal.slice(0, 5)}
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Compact Shortcut Bar to Dedicated Menu: Rekapitulasi & Evaluasi Harian Reason/Event Pagi vs Siang */}
            {onOpenReasonEvaluation && (
              <div className="no-print border-t border-slate-200 bg-slate-50/80 px-5 py-3 flex flex-wrap items-center justify-between gap-3">
                <div className="text-xs text-slate-600">
                  Ingin melihat perbandingan berdampingan & evaluasi mendalam{' '}
                  <strong className="text-amber-800">Reason / Event Pagi</strong> vs{' '}
                  <strong className="text-indigo-800">Reason / Event Siang</strong> untuk{' '}
                  <strong className="text-slate-900">{table.title}</strong>?
                </div>
                <button
                  type="button"
                  onClick={() => onOpenReasonEvaluation(table.id)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white transition-colors shadow-2xs"
                >
                  <span>Buka Menu Rekap Reason/Event Pagi vs Siang</span>
                </button>
              </div>
            )}
          </section>
        );
      })}
        </>
      )}

      {/* Modal Atur Periode Weekly & Generate 7 Hari Otomatis */}
      {periodeModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white border border-slate-200 rounded-xl shadow-xl max-w-lg w-full overflow-hidden">
            <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 bg-slate-900 text-white">
              <div className="flex items-center gap-2">
                <CalendarRange className="w-4 h-4 text-amber-400" />
                <h3 className="text-sm font-bold">
                  Atur Periode Weekly & Rentang 7 Tanggal Otomatis
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setPeriodeModal(null)}
                className="p-1 text-slate-400 hover:text-white rounded-md"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-4">
              {/* Full-Year 53-Week Monday-Sunday Selector */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">
                  1. Pilih dari Kalender 53 Weekly Setahun (Senin s/d Minggu — 2026):
                </label>
                <select
                  value={periodeModal.weekNumber}
                  onChange={(e) => {
                    const wkNum = Number(e.target.value);
                    const found = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wkNum);
                    if (found) {
                      setPeriodeModal({
                        ...periodeModal,
                        weekNumber: found.weekNumber,
                        title: `TOTAL WEEKLY ${found.weekNumber}`,
                        startDateIso: found.startDateIso,
                        dayCount: 7,
                        customPeriodeText: found.periodeText,
                      });
                    }
                  }}
                  className="w-full px-3 py-2 text-xs font-bold bg-indigo-50/70 border border-indigo-300 text-slate-900 rounded-lg focus:outline-none focus:border-indigo-600 mb-2"
                >
                  {ANNUAL_2026_WEEKS.map((w) => (
                    <option key={w.weekNumber} value={w.weekNumber}>
                      Weekly {w.weekNumber} ({w.monthName} · {w.quarter}) — Senin {w.seninSlash} s/d Minggu {w.mingguSlash}
                    </option>
                  ))}
                </select>
                <div className="flex flex-wrap gap-1.5">
                  {PRESET_PERIODS.map((pr) => (
                    <button
                      key={pr.weekNum}
                      type="button"
                      onClick={() => {
                        const { periodeText } = generateSevenSlashDatesFromIso(
                          pr.startIso,
                          periodeModal.dayCount
                        );
                        setPeriodeModal({
                          ...periodeModal,
                          weekNumber: pr.weekNum,
                          title: `TOTAL WEEKLY ${pr.weekNum}`,
                          startDateIso: pr.startIso,
                          customPeriodeText: periodeText,
                        });
                      }}
                      className={`px-2.5 py-1.5 text-xs font-medium rounded-lg border transition-colors ${
                        periodeModal.weekNumber === pr.weekNum &&
                        periodeModal.startDateIso === pr.startIso
                          ? 'bg-indigo-600 text-white border-indigo-600 font-semibold'
                          : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                      }`}
                    >
                      {pr.label}
                    </button>
                  ))}
                </div>
              </div>

              {/* Siklus Hari Mulai Weekly */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1.5">
                  2. Pilih Siklus Hari Mulai Weekly:
                </label>
                <div className="flex flex-wrap gap-1.5">
                  {START_DAY_CYCLE_PRESETS.map((cy) => (
                    <button
                      key={cy.label}
                      type="button"
                      onClick={() => {
                        const { periodeText } = generateSevenSlashDatesFromIso(
                          cy.startIso,
                          periodeModal.dayCount
                        );
                        setPeriodeModal({
                          ...periodeModal,
                          startDateIso: cy.startIso,
                          customPeriodeText: periodeText,
                        });
                      }}
                      className={`px-2.5 py-1.5 text-xs font-medium rounded-lg border transition-colors ${
                        periodeModal.startDateIso === cy.startIso
                          ? 'bg-amber-500 text-slate-950 border-amber-500 font-bold'
                          : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                      }`}
                    >
                      {cy.label}
                    </button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Minggu Ke- (Weekly)
                  </label>
                  <input
                    type="number"
                    value={periodeModal.weekNumber}
                    onChange={(e) => {
                      const wn = Number(e.target.value) || 41;
                      setPeriodeModal({
                        ...periodeModal,
                        weekNumber: wn,
                        title: `TOTAL WEEKLY ${wn}`,
                      });
                    }}
                    className="w-full px-3 py-2 text-xs font-bold font-mono-num border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Tanggal Mulai (Hari Ke-1)
                  </label>
                  <input
                    type="date"
                    value={periodeModal.startDateIso}
                    onChange={(e) => {
                      const iso = e.target.value;
                      if (!iso) return;
                      const { periodeText } = generateSevenSlashDatesFromIso(
                        iso,
                        periodeModal.dayCount
                      );
                      setPeriodeModal({
                        ...periodeModal,
                        startDateIso: iso,
                        customPeriodeText: periodeText,
                      });
                    }}
                    className="w-full px-3 py-2 text-xs font-bold font-mono-num border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Jumlah Hari (1–14 Hari)
                  </label>
                  <input
                    type="number"
                    min={1}
                    max={14}
                    value={periodeModal.dayCount}
                    onChange={(e) => {
                      const cnt = Math.max(1, Math.min(14, Number(e.target.value) || 7));
                      const { periodeText } = generateSevenSlashDatesFromIso(
                        periodeModal.startDateIso,
                        cnt
                      );
                      setPeriodeModal({
                        ...periodeModal,
                        dayCount: cnt,
                        customPeriodeText: periodeText,
                      });
                    }}
                    className="w-full px-3 py-2 text-xs font-bold font-mono-num border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              {/* Preview of Generated Days with Indonesian Day Names */}
              <div className="p-3.5 bg-indigo-50/60 border border-indigo-200 rounded-xl space-y-2">
                <div className="text-xs font-bold text-indigo-950">
                  Pratinjau Weekly Periode terhadap Hari:
                </div>
                <div className="text-xs font-semibold text-emerald-800">
                  {periodeModal.title} — Periode :{' '}
                  {
                    generateSevenSlashDatesFromIso(
                      periodeModal.startDateIso,
                      periodeModal.dayCount
                    ).periodeText
                  }{' '}
                  (
                  {
                    generateSevenSlashDatesFromIso(
                      periodeModal.startDateIso,
                      periodeModal.dayCount
                    ).dayCycleText
                  }
                  )
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-1.5 pt-1">
                  {generateSevenSlashDatesFromIso(
                    periodeModal.startDateIso,
                    periodeModal.dayCount
                  ).slashDates.map((dt, idx) => {
                    const dName = getIndoDayName(dt);
                    const wknd = isWeekendDate(dt);
                    return (
                      <span
                        key={idx}
                        className={`px-2 py-1 text-[11px] font-mono-num border rounded flex items-center justify-between ${
                          wknd
                            ? 'bg-rose-50 border-rose-200 text-rose-900 font-bold'
                            : 'bg-white border-indigo-200 text-slate-800'
                        }`}
                      >
                        <span className="font-sans font-semibold">
                          H{idx + 1} {dName}
                        </span>
                        <span>{dt.slice(0, 5)}</span>
                      </span>
                    );
                  })}
                </div>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2 px-5 py-3.5 border-t border-slate-200 bg-slate-50">
              <button
                type="button"
                onClick={() => setPeriodeModal(null)}
                className="px-4 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded-lg"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={applyPeriodeChanges}
                className="px-4 py-2 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg"
              >
                Terapkan Periode & 7 Tanggal
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Quick Daily Shift Editor Modal */}
      {editingModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="bg-white border border-slate-200 rounded-xl shadow-xl max-w-xl w-full overflow-hidden">
            <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 bg-slate-50">
              <div className="flex items-center gap-2">
                <Calendar className="w-4 h-4 text-indigo-600" />
                <h3 className="text-sm font-bold text-slate-900">
                  Input Harian Cepat — Tanggal {editingModal.dayData.tanggal}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setEditingModal(null)}
                className="p-1 text-slate-500 hover:text-slate-800 rounded-md"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 space-y-5 max-h-[80vh] overflow-y-auto">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Pilih Kalender Tanggal
                  </label>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => {
                        const prevIso = shiftIsoDate(
                          slashDateToIso(editingModal.dayData.tanggal),
                          -1
                        );
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            tanggal: isoToSlashDate(prevIso),
                          },
                        });
                      }}
                      title="Tanggal Sebelumnya"
                      className="p-2 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-700 shrink-0"
                    >
                      <ChevronLeft className="w-4 h-4" />
                    </button>
                    <input
                      type="date"
                      aria-label="Filter Kalender Tanggal Edit Harian"
                      value={slashDateToIso(editingModal.dayData.tanggal)}
                      onChange={(e) => {
                        if (e.target.value) {
                          setEditingModal({
                            ...editingModal,
                            dayData: {
                              ...editingModal.dayData,
                              tanggal: isoToSlashDate(e.target.value),
                            },
                          });
                        }
                      }}
                      className="w-full px-2.5 py-2 text-xs font-bold border border-indigo-300 bg-indigo-50/60 text-indigo-950 rounded-lg focus:outline-none focus:border-indigo-600 font-mono-num cursor-pointer"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        const nextIso = shiftIsoDate(
                          slashDateToIso(editingModal.dayData.tanggal),
                          1
                        );
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            tanggal: isoToSlashDate(nextIso),
                          },
                        });
                      }}
                      title="Tanggal Berikutnya"
                      className="p-2 rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-700 shrink-0"
                    >
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Tanggal (DD/MM/YYYY)
                  </label>
                  <input
                    type="text"
                    value={editingModal.dayData.tanggal}
                    onChange={(e) =>
                      setEditingModal({
                        ...editingModal,
                        dayData: { ...editingModal.dayData, tanggal: e.target.value },
                      })
                    }
                    className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600 font-mono-num"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg space-y-3">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <span className="text-xs font-bold text-slate-900">SHIFT PAGI</span>
                    <UserCheck className="w-3.5 h-3.5 text-slate-500" />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Nama Petugas</label>
                    <input
                      type="text"
                      value={editingModal.dayData.pagi.nama}
                      onChange={(e) =>
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            pagi: { ...editingModal.dayData.pagi, nama: e.target.value },
                          },
                        })
                      }
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md"
                    />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Sales (Rp)</label>
                    <input
                      type="number"
                      value={editingModal.dayData.pagi.sales ?? ''}
                      onChange={(e) =>
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            pagi: {
                              ...editingModal.dayData.pagi,
                              sales: e.target.value === '' ? null : Number(e.target.value),
                            },
                          },
                        })
                      }
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-xs text-slate-600 mb-1">Qty</label>
                      <input
                        type="number"
                        value={editingModal.dayData.pagi.qty ?? ''}
                        onChange={(e) =>
                          setEditingModal({
                            ...editingModal,
                            dayData: {
                              ...editingModal.dayData,
                              pagi: {
                                ...editingModal.dayData.pagi,
                                qty: e.target.value === '' ? null : Number(e.target.value),
                              },
                            },
                          })
                        }
                        className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                      />
                    </div>
                    <div>
                      <label className="block text-xs text-slate-600 mb-1">Trx</label>
                      <input
                        type="number"
                        value={editingModal.dayData.pagi.trx ?? ''}
                        onChange={(e) =>
                          setEditingModal({
                            ...editingModal,
                            dayData: {
                              ...editingModal.dayData,
                              pagi: {
                                ...editingModal.dayData.pagi,
                                trx: e.target.value === '' ? null : Number(e.target.value),
                              },
                            },
                          })
                        }
                        className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                      />
                    </div>
                  </div>
                </div>

                <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-lg space-y-3">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <span className="text-xs font-bold text-slate-900">SHIFT SIANG</span>
                    <UserCheck className="w-3.5 h-3.5 text-slate-500" />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Nama Petugas</label>
                    <input
                      type="text"
                      value={editingModal.dayData.siang.nama}
                      onChange={(e) =>
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            siang: { ...editingModal.dayData.siang, nama: e.target.value },
                          },
                        })
                      }
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md"
                    />
                  </div>
                  <div>
                    <label className="block text-xs text-slate-600 mb-1">Sales (Rp)</label>
                    <input
                      type="number"
                      value={editingModal.dayData.siang.sales ?? ''}
                      onChange={(e) =>
                        setEditingModal({
                          ...editingModal,
                          dayData: {
                            ...editingModal.dayData,
                            siang: {
                              ...editingModal.dayData.siang,
                              sales: e.target.value === '' ? null : Number(e.target.value),
                            },
                          },
                        })
                      }
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-xs text-slate-600 mb-1">Qty</label>
                      <input
                        type="number"
                        value={editingModal.dayData.siang.qty ?? ''}
                        onChange={(e) =>
                          setEditingModal({
                            ...editingModal,
                            dayData: {
                              ...editingModal.dayData,
                              siang: {
                                ...editingModal.dayData.siang,
                                qty: e.target.value === '' ? null : Number(e.target.value),
                              },
                            },
                          })
                        }
                        className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                      />
                    </div>
                    <div>
                      <label className="block text-xs text-slate-600 mb-1">Trx</label>
                      <input
                        type="number"
                        value={editingModal.dayData.siang.trx ?? ''}
                        onChange={(e) =>
                          setEditingModal({
                            ...editingModal,
                            dayData: {
                              ...editingModal.dayData,
                              siang: {
                                ...editingModal.dayData.siang,
                                trx: e.target.value === '' ? null : Number(e.target.value),
                              },
                            },
                          })
                        }
                        className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md font-mono-num"
                      />
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Reason / Event Shift Pagi
                  </label>
                  <textarea
                    rows={2}
                    value={editingModal.dayData.reasonPagi}
                    onChange={(e) =>
                      setEditingModal({
                        ...editingModal,
                        dayData: { ...editingModal.dayData, reasonPagi: e.target.value },
                      })
                    }
                    className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Reason / Event Shift Siang / Malam
                  </label>
                  <textarea
                    rows={2}
                    value={editingModal.dayData.reasonSiang}
                    onChange={(e) =>
                      setEditingModal({
                        ...editingModal,
                        dayData: { ...editingModal.dayData, reasonSiang: e.target.value },
                      })
                    }
                    className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            </div>

            <div className="flex items-center justify-between gap-2 px-5 py-3.5 border-t border-slate-200 bg-slate-50">
              <AutoDraftStatusBadge
                menuLabel="Modal Harian"
                lastDraftTimeStr={salesDraftTimeStr}
                isDirty={true}
                isOnline={isSalesOnline}
              />
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.QUICK_DAY_MODAL);
                    setEditingModal(null);
                  }}
                  className="px-4 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-100 rounded-lg"
                >
                  Batal
                </button>
                <button
                  type="button"
                  onClick={saveModalChanges}
                  className="px-4 py-2 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg"
                >
                  Simpan Perubahan
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
