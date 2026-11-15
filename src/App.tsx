/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import {
  INITIAL_WEEKLY_TABLES,
  INITIAL_REVIEW_DOCS,
  INITIAL_SUBMISSION_LOGS,
  PERSISTED_STORE_UPDATED_AT,
  WeeklySalesTable,
  DailySalesRecord,
  ReviewMonitoringDoc,
  ReportSubmissionLog,
  calculateWeekMetrics,
  formatNumberId,
  getWeeklyPeriodeText,
  generateSevenSlashDatesFromIso,
  ANNUAL_2026_WEEKS,
  normalizeAndRepairWeeklyTables,
  normalizeReviewDocs,
  createWeeklyTableFromAnnualSchedule,
  createEmptyInitialWeeklyTables,
  getTableMonthKey,
  getReviewDocWeekNumber,
  createReviewDocForWeek,
  mergeWeeklyTablesNonDestructive,
  mergeReviewDocsNonDestructive,
} from './data/initialData';
import { DashboardOverview, DashboardHorizon } from './components/DashboardOverview';
import { WeeklySalesSpreadsheet } from './components/WeeklySalesSpreadsheet';
import { SalesTableHorizon } from './components/DailyAndMonthlySalesTables';
import { DailyInputWorkspace } from './components/DailyInputWorkspace';
import { ReviewMonitoringSheet } from './components/ReviewMonitoringSheet';
import { AnalyticsSummaryView } from './components/AnalyticsSummaryView';
import { ReportLinkCenter } from './components/ReportLinkCenter';
import { AnnualWeeklyScheduleView } from './components/AnnualWeeklyScheduleView';
import { ReasonEventEvaluationView } from './components/ReasonEventEvaluationView';
import { SaveToolsModal } from './components/SaveToolsModal';
import { ExportExcelModal } from './components/ExportExcelModal';
import { ReportCvDiv1View, CvHorizon } from './components/ReportCvDiv1View';
import { TotalDpCvReportView, TotalHorizon } from './components/TotalDpCvReportView';
import { CvInventoryOpsTableSection } from './components/CvInventoryOpsTableSection';
import { exportCvDiv1ToExcel } from './utils/exportCvDiv1Excel';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  MONTH_NAMES_ID,
  loadCvDiv1Records,
  saveCvDiv1Records,
  loadCvInventoryOps,
  saveCvInventoryOps,
  normalizeCvInventoryOpsList,
  generateInitialCvDiv1Records,
  STORAGE_KEY_CV_DIV1,
  STORAGE_KEY_CV_INVENTORY_OPS,
} from './data/cvDiv1Data';
import {
  AppEntryAccessGate,
  AppShareLinkModal,
  AppLoginLogoutModal,
  AppAccessRole,
} from './components/AppAccessAndLinkModal';
import {
  setAutoDetectedPublicOrigin,
  getGooglePublicDataCvUrl,
} from './utils/shareUrl';
import { copyTextToClipboard } from './utils/clipboard';
import { isoToSlashDate } from './utils/dateUtils';
import {
  saveBazarWorkspaceToFirestore,
  saveReportSubmissionToFirestore,
  subscribeToBazarWorkspaceFirestore,
  fetchBazarWorkspaceFromFirestore,
  saveCvWorkspaceToFirestore,
  fetchCvWorkspaceFromFirestore,
  subscribeToCvWorkspaceFirestore,
} from './firebase';
import { exportWeeklyTablesToWPSExcel } from './utils/exportExcel';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  markAutoDraftClean,
} from './utils/autoDraft';
import { AutoDraftStatusBadge } from './components/AutoDraftStatusBanner';
import {
  FileSpreadsheet,
  PenSquare,
  ClipboardList,
  Users,
  Link2,
  Download,
  Printer,
  RotateCcw,
  Menu,
  X,
  Store,
  Filter,
  Plus,
  Calendar,
  CalendarRange,
  CalendarDays,
  Undo2,
  Redo2,
  MessageSquareText,
  Lock,
  Save,
  CheckCircle2,
  Eye,
  KeyRound,
  ShieldCheck,
  Share2,
  LogIn,
  LogOut,
  Tag,
  Sparkles,
  TrendingUp,
  Layers,
} from 'lucide-react';

interface HistorySnapshot {
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  cvDiv1Records?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  label: string;
}

export type NavMenuId =
  | 'dashboard'
  | 'sales_matrix'
  | 'report_cv_div1'
  | 'total_dp_cv'
  | 'data_cv_ops'
  | 'annual_schedule'
  | 'reason_evaluation'
  | 'daily_input'
  | 'report_link'
  | 'review_monitoring'
  | 'staff_analytics';

const STORAGE_KEY_TABLES = 'bazar_tanjung_weekly_tables_v6';
const STORAGE_KEY_REVIEWS = 'bazar_tanjung_review_docs_v6';
const STORAGE_KEY_LOGS = 'bazar_tanjung_submission_logs_v6';
const STORAGE_KEY_ADMIN_PASS = 'bazar_tanjung_admin_password_v3';
const STORAGE_KEY_UPDATED_AT = 'bazar_tanjung_updated_at_v6';
const STORAGE_KEY_PENDING_SYNC = 'bazar_tanjung_pending_sync_v2';
const STORAGE_KEY_UI_PREFS = 'bazar_tanjung_ui_prefs_v2';
const STORAGE_KEY_ACCESS_ROLE = 'bazar_tanjung_access_role_v2';
const SESSION_KEY_ACCESS_ROLE = 'bazar_tanjung_access_role_session_v2';

// Clear legacy stale cache versions so cleared tables immediately hydrate
try {
  if (typeof localStorage !== 'undefined') {
    const legacyKeys = [
      'bazar_kapuas_weekly_tables_v3',
      'bazar_kapuas_weekly_tables_v4',
      'bazar_kapuas_weekly_tables_v5',
      'bazar_kapuas_review_docs_v5',
      'bazar_kapuas_submission_logs_v5',
      'bazar_kapuas_updated_at_v3',
      'bazar_kapuas_updated_at_v4',
      'bazar_kapuas_updated_at_v5',
      'bazar_tanjung_cv_div1_records_v1',
    ];
    legacyKeys.forEach((k) => localStorage.removeItem(k));
  }
} catch {}

const IDB_NAME = 'bazar_tanjung_idb_v3';
const IDB_STORE = 'app_state';
const IDB_KEY = 'latest_snapshot';

function saveSnapshotToIndexedDB(payload: {
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  submissionLogs: ReportSubmissionLog[];
  cvDiv1Records?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  updatedAt: number;
}) {
  try {
    if (typeof indexedDB === 'undefined') return;
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => {
      try {
        const db = req.result;
        const tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).put(payload, IDB_KEY);
      } catch {}
    };
  } catch {}
}

function loadSnapshotFromIndexedDB(
  onLoaded: (payload: {
    weeklyTables?: WeeklySalesTable[];
    reviewDocs?: ReviewMonitoringDoc[];
    submissionLogs?: ReportSubmissionLog[];
    cvDiv1Records?: CvDailyRecord[];
    cvInventoryOps?: CvInventoryOpsEntry[];
    updatedAt?: number;
  }) => void
) {
  try {
    if (typeof indexedDB === 'undefined') return;
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => {
      try {
        const db = req.result;
        const tx = db.transaction(IDB_STORE, 'readonly');
        const getReq = tx.objectStore(IDB_STORE).get(IDB_KEY);
        getReq.onsuccess = () => {
          if (getReq.result && typeof getReq.result === 'object') {
            onLoaded(getReq.result);
          }
        };
      } catch {}
    };
  } catch {}
}

interface SavedUiPrefs {
  activeMenu?: NavMenuId;
  dashboardHorizon?: DashboardHorizon;
  salesTableHorizon?: SalesTableHorizon;
  selectedWeeklyFilter?: string;
  activeReviewId?: string;
  showDivZero?: boolean;
}

function loadSavedUiPrefs(): SavedUiPrefs {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_UI_PREFS);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') return parsed;
    }
  } catch {}
  return {};
}

export default function App() {
  // Parse URL query params for direct report form links (?form=shift or ?form=review) and access mode (?mode=public | gate | admin)
  const searchParams = (() => {
    if (typeof window === 'undefined') return new URLSearchParams();
    const primary = new URLSearchParams(window.location.search);
    if (window.location.hash && window.location.hash.includes('?')) {
      const hashQuery = window.location.hash.slice(
        window.location.hash.indexOf('?')
      );
      const fromHash = new URLSearchParams(hashQuery);
      fromHash.forEach((v, k) => {
        if (!primary.has(k)) primary.set(k, v);
      });
    }
    return primary;
  })();
  const rawPathname =
    typeof window !== 'undefined' && window.location
      ? window.location.pathname.toLowerCase().replace(/\/+$/, '')
      : '';
  const pathFormOverride =
    rawPathname.endsWith('/isi-report-cv') ||
    rawPathname.endsWith('/report-cv') ||
    rawPathname.endsWith('/form-cv')
      ? 'cv_report'
      : rawPathname.endsWith('/isi-data-cv') ||
          rawPathname.endsWith('/data-cv') ||
          rawPathname.endsWith('/form-data-cv')
        ? 'data_cv'
        : rawPathname.endsWith('/isi-report') ||
            rawPathname.endsWith('/report') ||
            rawPathname.endsWith('/form-shift')
          ? 'shift'
          : rawPathname.endsWith('/isi-review') ||
              rawPathname.endsWith('/review') ||
              rawPathname.endsWith('/form-review')
            ? 'review'
            : null;
  const pathModeOverride =
    rawPathname.endsWith('/publik') ||
    rawPathname.endsWith('/public') ||
    rawPathname.endsWith('/buka-aplikasi')
      ? 'public'
      : rawPathname.endsWith('/masuk') || rawPathname.endsWith('/gate')
        ? 'gate'
        : null;

  const urlFormParam = searchParams.get('form') || pathFormOverride; // 'shift' | 'review' | 'cv_report' | 'data_cv' | null
  const urlLockParam = searchParams.get('lock');
  const urlMenuParam = searchParams.get('menu');
  const urlModeParam = searchParams.get('mode') || pathModeOverride; // 'public' | 'gate' | 'admin' | null
  const isOpenedFromLockedLink = Boolean(
    urlFormParam === 'shift' ||
      urlFormParam === 'review' ||
      urlFormParam === 'cv_report' ||
      urlFormParam === 'data_cv' ||
      urlLockParam === '1' ||
      urlLockParam === '10' ||
      urlMenuParam === 'report_link'
  );
  const [directReportFormMode, setDirectReportFormMode] = useState<
    'shift' | 'review' | 'cv_report' | 'data_cv' | null
  >(
    urlFormParam === 'shift'
      ? 'shift'
      : urlFormParam === 'review'
        ? 'review'
        : urlFormParam === 'cv_report'
          ? 'cv_report'
          : urlFormParam === 'data_cv'
            ? 'data_cv'
            : null
  );
  const urlWeekParam = searchParams.get('week') || undefined;
  const urlDateParam = searchParams.get('date') || undefined;
  const urlPeriodTypeParam: 'weekly' | 'monthly' =
    searchParams.get('periodType') === 'monthly' ? 'monthly' : 'weekly';
  const urlMonthParam = searchParams.get('month') || undefined;
  const rawBrandParam = (searchParams.get('brand') || '').toUpperCase();
  const urlBrandParam: 'ALL' | 'GRAPHIS' | 'SIMPLICITY' =
    rawBrandParam === 'GRAPHIS'
      ? 'GRAPHIS'
      : rawBrandParam === 'SIMPLICITY'
        ? 'SIMPLICITY'
        : 'ALL';
  const urlDayParam = searchParams.get('day')
    ? Number(searchParams.get('day'))
    : undefined;
  const rawShift = searchParams.get('shift');
  const urlShiftParam =
    rawShift === 'Siang'
      ? 'Siang'
      : rawShift === 'Lembur' || rawShift === 'BOTH'
        ? 'Lembur'
        : 'Pagi';
  const urlStaffParam = searchParams.get('staff') || '';

  const initialUiPrefs = loadSavedUiPrefs();

  // Admin Password state (default: 'MDS311', synced with localStorage & backend)
  const [adminPassword, setAdminPassword] = useState<string>(() => {
    try {
      const savedPass = localStorage.getItem(STORAGE_KEY_ADMIN_PASS);
      if (savedPass && savedPass.trim().length >= 3 && savedPass.trim() !== 'admin123') {
        return savedPass.trim();
      }
    } catch {}
    return 'MDS311';
  });

  // Access Role state: 'public' (Read-Only / Hanya Lihat) vs 'admin' (Can Edit)
  const [accessRole, setAccessRole] = useState<AppAccessRole>(() => {
    if (urlModeParam === 'public') return 'public';
    try {
      const savedSessionRole = sessionStorage.getItem(SESSION_KEY_ACCESS_ROLE);
      if (savedSessionRole === 'admin' || savedSessionRole === 'public') {
        return savedSessionRole;
      }
      const savedPersistentRole = localStorage.getItem(STORAGE_KEY_ACCESS_ROLE);
      if (savedPersistentRole === 'admin' || savedPersistentRole === 'public') {
        return savedPersistentRole;
      }
    } catch {}
    return 'admin';
  });

  // Entry Gate screen shown on initial visit unless ?mode=public or opened from locked report form link
  const [showEntryGate, setShowEntryGate] = useState<boolean>(() => {
    if (isOpenedFromLockedLink) return false;
    if (urlModeParam === 'public') return false;
    return true;
  });

  const [isAppShareModalOpen, setIsAppShareModalOpen] = useState<boolean>(false);
  const [isLoginLogoutModalOpen, setIsLoginLogoutModalOpen] = useState<boolean>(false);

  const [isLockedReportOnly, setIsLockedReportOnly] = useState<boolean>(
    isOpenedFromLockedLink
  );

  const [activeMenu, setActiveMenu] = useState<NavMenuId>(() => {
    if (isOpenedFromLockedLink) {
      return 'report_link';
    }
    return initialUiPrefs.activeMenu || 'dashboard';
  });
  const [dashboardHorizon, setDashboardHorizon] = useState<DashboardHorizon>(
    () => initialUiPrefs.dashboardHorizon || 'weekly'
  );
  const [salesTableHorizon, setSalesTableHorizon] = useState<SalesTableHorizon>(
    () => initialUiPrefs.salesTableHorizon || 'weekly'
  );
  const [cvHorizon, setCvHorizon] = useState<CvHorizon>('harian');
  const [totalDpCvHorizon, setTotalDpCvHorizon] = useState<TotalHorizon>('harian');
  const [mobileMenuOpen, setMobileMenuOpen] = useState<boolean>(false);

  const [weeklyTables, setWeeklyTables] = useState<WeeklySalesTable[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_TABLES);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return normalizeAndRepairWeeklyTables(parsed);
        }
      }
    } catch {
      // ignore storage errors
    }
    return normalizeAndRepairWeeklyTables(INITIAL_WEEKLY_TABLES);
  });

  const [reviewDocs, setReviewDocs] = useState<ReviewMonitoringDoc[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_REVIEWS);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Prioritize user's saved localStorage edits (parsed as primary) while healing any missing sections from INITIAL_REVIEW_DOCS
          return mergeReviewDocsNonDestructive(parsed, INITIAL_REVIEW_DOCS);
        }
      }
    } catch {
      // ignore storage errors
    }
    return normalizeReviewDocs(INITIAL_REVIEW_DOCS);
  });

  const [submissionLogs, setSubmissionLogs] = useState<ReportSubmissionLog[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_LOGS);
      const savedTs = Number(localStorage.getItem(STORAGE_KEY_UPDATED_AT) || 0);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          if (savedTs < PERSISTED_STORE_UPDATED_AT) {
            const seen = new Set<string>();
            const combined: ReportSubmissionLog[] = [];
            for (const lg of [...INITIAL_SUBMISSION_LOGS, ...parsed]) {
              if (lg && lg.id && !seen.has(lg.id) && !/^log-[123]$/.test(lg.id)) {
                seen.add(lg.id);
                combined.push(lg);
              }
            }
            return combined;
          }
          return parsed.filter((lg: any) => lg && !/^log-[123]$/.test(lg.id));
        }
      }
    } catch {
      // ignore storage errors
    }
    return INITIAL_SUBMISSION_LOGS;
  });

  const [cvDiv1Records, setCvDiv1Records] = useState<CvDailyRecord[]>(() => {
    return loadCvDiv1Records();
  });

  const handleUpdateCvDiv1Records = (nextRecs: CvDailyRecord[]) => {
    if (!requireAdminForEdit()) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    lastServerUpdatedAtRef.current = Math.max(
      Date.now(),
      lastServerUpdatedAtRef.current + 1
    );
    const currentTables = latestTablesRef.current;
    const currentReviews = latestReviewsRef.current;
    const currentCvRecs = latestCvRecordsRef.current;
    const currentCvOps = latestCvOpsRef.current;
    pendingCvSyncRef.current = true;
    latestCvRecordsRef.current = nextRecs;
    setCvDiv1Records(nextRecs);
    saveCvDiv1Records(nextRecs);
    scheduleDebouncedEditSync(
      currentTables,
      currentReviews,
      'Perubahan Report CV Div. 1',
      currentCvRecs,
      currentCvOps
    );
  };

  const [cvInventoryOps, setCvInventoryOps] = useState<CvInventoryOpsEntry[]>(() => {
    return loadCvInventoryOps();
  });

  const handleUpdateCvInventoryOps = (nextOps: CvInventoryOpsEntry[]) => {
    if (!requireAdminForEdit()) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    lastServerUpdatedAtRef.current = Math.max(
      Date.now(),
      lastServerUpdatedAtRef.current + 1
    );
    const normalizedOps = normalizeCvInventoryOpsList(nextOps);
    const currentTables = latestTablesRef.current;
    const currentReviews = latestReviewsRef.current;
    const currentCvRecs = latestCvRecordsRef.current;
    const currentCvOps = latestCvOpsRef.current;
    pendingCvSyncRef.current = true;
    latestCvOpsRef.current = normalizedOps;
    setCvInventoryOps(normalizedOps);
    saveCvInventoryOps(normalizedOps);
    scheduleDebouncedEditSync(
      currentTables,
      currentReviews,
      'Perubahan Data CV Ops',
      currentCvRecs,
      currentCvOps
    );
  };

  const [dataCvMonthKey, setDataCvMonthKey] = useState<string>('11');
  const [dataCvWeekNumber, setDataCvWeekNumber] = useState<number>(45);
  const [dataCvToast, setDataCvToast] = useState<string | null>(null);
  const showDataCvToast = (msg: string) => {
    setDataCvToast(msg);
    setTimeout(() => setDataCvToast(null), 3000);
  };

  // Global Weekly Filter state ('ALL' or specific table.id like 'weekly-40')
  const [selectedWeeklyFilter, setSelectedWeeklyFilter] = useState<string>(() => {
    if (urlWeekParam) return urlWeekParam;
    if (urlDateParam) {
      const slash = isoToSlashDate(urlDateParam);
      const matchedTbl = weeklyTables.find((t) =>
        t.days.some((d) => d.tanggal.trim() === slash)
      );
      if (matchedTbl) return matchedTbl.id;
      const matchedSched = ANNUAL_2026_WEEKS.find(
        (w) =>
          w.slashDates.includes(slash) ||
          (urlDateParam >= w.startDateIso && urlDateParam <= w.endDateIso)
      );
      if (matchedSched) return `weekly-${matchedSched.weekNumber}`;
    }
    if (initialUiPrefs.selectedWeeklyFilter) {
      if (
        initialUiPrefs.selectedWeeklyFilter === 'ALL' ||
        weeklyTables.some((t) => t.id === initialUiPrefs.selectedWeeklyFilter)
      ) {
        return initialUiPrefs.selectedWeeklyFilter;
      }
    }
    const firstFilled = weeklyTables.find(
      (t) => calculateWeekMetrics(t).weeklySales > 0
    );
    return firstFilled?.id || weeklyTables[0]?.id || INITIAL_WEEKLY_TABLES[0].id;
  });
  const [activeReviewId, setActiveReviewId] = useState<string>(() => {
    if (urlDateParam) {
      const slash = isoToSlashDate(urlDateParam);
      const matchedSched = ANNUAL_2026_WEEKS.find(
        (w) =>
          w.slashDates.includes(slash) ||
          (urlDateParam >= w.startDateIso && urlDateParam <= w.endDateIso)
      );
      if (matchedSched) {
        const revMatch = reviewDocs.find(
          (d) =>
            getReviewDocWeekNumber(d) === matchedSched.weekNumber ||
            d.id === `review-w${matchedSched.weekNumber}`
        );
        if (revMatch) return revMatch.id;
      }
    }
    if (
      initialUiPrefs.activeReviewId &&
      reviewDocs.some((d) => d.id === initialUiPrefs.activeReviewId)
    ) {
      return initialUiPrefs.activeReviewId;
    }
    return reviewDocs[0]?.id || INITIAL_REVIEW_DOCS[0].id;
  });
  const [showDivZero, setShowDivZero] = useState<boolean>(() =>
    typeof initialUiPrefs.showDivZero === 'boolean' ? initialUiPrefs.showDivZero : true
  );
  const [staffFilter, setStaffFilter] = useState<string>('ALL');
  const [confirmReset, setConfirmReset] = useState<boolean>(false);

  // Undo & Redo History Stacks (up to 50 steps)
  const [undoStack, setUndoStack] = useState<HistorySnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([]);
  const [historyNotice, setHistoryNotice] = useState<string | null>(null);
  const [isSaveToolsOpen, setIsSaveToolsOpen] = useState<boolean>(false);
  const [isExportModalOpen, setIsExportModalOpen] = useState<boolean>(false);
  const [isSavingState, setIsSavingState] = useState<boolean>(false);
  const [hasUnsavedGlobalDraft, setHasUnsavedGlobalDraft] = useState<boolean>(false);
  const [lastSavedTimeStr, setLastSavedTimeStr] = useState<string>(() =>
    new Date().toLocaleTimeString('id-ID', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    })
  );

  // Initialize lastServerUpdatedAtRef from localStorage or PERSISTED_STORE_UPDATED_AT
  const lastServerUpdatedAtRef = useRef<number>((() => {
    try {
      const savedTs = Number(localStorage.getItem(STORAGE_KEY_UPDATED_AT) || 0);
      if (savedTs > 0) return Math.max(savedTs, PERSISTED_STORE_UPDATED_AT);
    } catch {}
    return PERSISTED_STORE_UPDATED_AT || 0;
  })());
  const pendingSaveCountRef = useRef<number>(0);
  const pendingCvSyncRef = useRef<boolean>(false);
  const lastCvSyncedAtRef = useRef<number>(0);
  const lastLocalEditAtRef = useRef<number>(0);
  const localEditSeqRef = useRef<number>(0);
  const debouncedSyncTimerRef = useRef<number | null>(null);
  const pendingUndoBaseRef = useRef<HistorySnapshot | null>(null);
  const lastServerSignatureRef = useRef<string>('');
  const pullFromServerRef = useRef<
    ((isInitial: boolean, forceApply?: boolean) => void) | null
  >(null);

  // Keep refs to latest state so beforeunload / pullFromServer always access latest data
  const latestTablesRef = useRef<WeeklySalesTable[]>(weeklyTables);
  latestTablesRef.current = weeklyTables;
  const latestReviewsRef = useRef<ReviewMonitoringDoc[]>(reviewDocs);
  latestReviewsRef.current = reviewDocs;
  const latestLogsRef = useRef<ReportSubmissionLog[]>(submissionLogs);
  latestLogsRef.current = submissionLogs;
  const latestCvRecordsRef = useRef<CvDailyRecord[]>(cvDiv1Records);
  latestCvRecordsRef.current = cvDiv1Records;
  const latestCvOpsRef = useRef<CvInventoryOpsEntry[]>(cvInventoryOps);
  latestCvOpsRef.current = cvInventoryOps;

  const {
    lastDraftTimeStr: globalDraftTimeStr,
    isOnline: isGlobalOnline,
    flushDraftNow: flushGlobalDraftNow,
  } = usePeriodicAutoDraft({
    storageKey: AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE,
    menuName: 'Seluruh Menu Aplikasi',
    data: {
      weeklyTables,
      reviewDocs,
      submissionLogs,
      cvDiv1Records,
      cvInventoryOps,
      activeMenu,
      selectedWeeklyFilter,
    },
    isDirty: hasUnsavedGlobalDraft,
    enabled: accessRole === 'admin' || isLockedReportOnly,
    intervalMs: 3000,
  });

  // Recover from Global Auto-Draft only if there was an offline unsynced edit
  useEffect(() => {
    let hasPendingSync = false;
    try {
      hasPendingSync = localStorage.getItem(STORAGE_KEY_PENDING_SYNC) === '1';
    } catch {}
    if (!hasPendingSync) return;

    const draftEnv = loadAutoDraftFromStorage<{
      weeklyTables?: WeeklySalesTable[];
      reviewDocs?: ReviewMonitoringDoc[];
      submissionLogs?: ReportSubmissionLog[];
      cvDiv1Records?: CvDailyRecord[];
      cvInventoryOps?: CvInventoryOpsEntry[];
    }>(AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE);
    if (
      draftEnv &&
      draftEnv.isDirty &&
      draftEnv.savedAt > lastServerUpdatedAtRef.current &&
      Array.isArray(draftEnv.data?.weeklyTables) &&
      draftEnv.data.weeklyTables.length > 0
    ) {
      const merged = mergeWeeklyTablesNonDestructive(
        draftEnv.data.weeklyTables,
        latestTablesRef.current
      );
      setWeeklyTables(merged);
      if (
        Array.isArray(draftEnv.data.reviewDocs) &&
        draftEnv.data.reviewDocs.length > 0
      ) {
        const mergedRevs = mergeReviewDocsNonDestructive(
          latestReviewsRef.current,
          draftEnv.data.reviewDocs
        );
        latestReviewsRef.current = mergedRevs;
        setReviewDocs(mergedRevs);
      }
      if (
        Array.isArray(draftEnv.data.cvDiv1Records) &&
        draftEnv.data.cvDiv1Records.length > 0
      ) {
        latestCvRecordsRef.current = draftEnv.data.cvDiv1Records;
        setCvDiv1Records(draftEnv.data.cvDiv1Records);
        saveCvDiv1Records(draftEnv.data.cvDiv1Records);
      }
      if (
        Array.isArray(draftEnv.data.cvInventoryOps) &&
        draftEnv.data.cvInventoryOps.length > 0
      ) {
        const normOps = normalizeCvInventoryOpsList(draftEnv.data.cvInventoryOps);
        latestCvOpsRef.current = normOps;
        setCvInventoryOps(normOps);
        saveCvInventoryOps(normOps);
      }
      setHasUnsavedGlobalDraft(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const cloneState = <T,>(obj: T): T => JSON.parse(JSON.stringify(obj));

  const persistDataToLocalStorage = (
    nextTables: WeeklySalesTable[],
    nextReviews: ReviewMonitoringDoc[],
    nextLogs: ReportSubmissionLog[],
    timestamp?: number,
    nextCvRecords?: CvDailyRecord[],
    nextCvOps?: CvInventoryOpsEntry[]
  ) => {
    const effectiveTs =
      timestamp && timestamp > 0
        ? timestamp
        : lastServerUpdatedAtRef.current || PERSISTED_STORE_UPDATED_AT;
    const cvRecsToPersist = nextCvRecords ?? latestCvRecordsRef.current;
    const cvOpsToPersist = nextCvOps ?? latestCvOpsRef.current;
    try {
      localStorage.setItem(STORAGE_KEY_TABLES, JSON.stringify(nextTables));
      localStorage.setItem(STORAGE_KEY_REVIEWS, JSON.stringify(nextReviews));
      localStorage.setItem(STORAGE_KEY_LOGS, JSON.stringify(nextLogs));
      if (effectiveTs > 0) {
        localStorage.setItem(STORAGE_KEY_UPDATED_AT, String(effectiveTs));
      }
    } catch {}
    saveCvDiv1Records(cvRecsToPersist);
    saveCvInventoryOps(cvOpsToPersist);
    saveSnapshotToIndexedDB({
      weeklyTables: nextTables,
      reviewDocs: nextReviews,
      submissionLogs: nextLogs,
      cvDiv1Records: cvRecsToPersist,
      cvInventoryOps: cvOpsToPersist,
      updatedAt: effectiveTs,
    });
  };

  // Automatically persist UI preferences (selected menu, weekly filter, active review doc)
  useEffect(() => {
    if (isLockedReportOnly) return;
    try {
      const prefs: SavedUiPrefs = {
        activeMenu,
        dashboardHorizon,
        salesTableHorizon,
        selectedWeeklyFilter,
        activeReviewId,
        showDivZero,
      };
      localStorage.setItem(STORAGE_KEY_UI_PREFS, JSON.stringify(prefs));
    } catch {}
  }, [
    activeMenu,
    dashboardHorizon,
    salesTableHorizon,
    selectedWeeklyFilter,
    activeReviewId,
    showDivZero,
    isLockedReportOnly,
  ]);

  // Automatically mirror data state to localStorage & IndexedDB (debounced so typing is never blocked)
  useEffect(() => {
    const timer = window.setTimeout(() => {
      persistDataToLocalStorage(
        weeklyTables,
        reviewDocs,
        submissionLogs,
        undefined,
        cvDiv1Records,
        cvInventoryOps
      );
    }, 550);
    return () => window.clearTimeout(timer);
  }, [weeklyTables, reviewDocs, submissionLogs, cvDiv1Records, cvInventoryOps]);

  // Recover from IndexedDB if localStorage was ever cleared by browser/iframe lifecycle
  useEffect(() => {
    loadSnapshotFromIndexedDB((idbSnap) => {
      if (!idbSnap || !Array.isArray(idbSnap.weeklyTables)) return;
      if (localEditSeqRef.current > 0 || Date.now() - lastLocalEditAtRef.current < 8000) {
        return;
      }
      const idbTs = Number(idbSnap.updatedAt || 0);
      if (idbTs > lastServerUpdatedAtRef.current) {
        const repairedTables = normalizeAndRepairWeeklyTables(idbSnap.weeklyTables);
        lastServerUpdatedAtRef.current = idbTs;
        latestTablesRef.current = repairedTables;
        setWeeklyTables(repairedTables);
        if (Array.isArray(idbSnap.reviewDocs) && idbSnap.reviewDocs.length > 0) {
          const repairedRevs = normalizeReviewDocs(idbSnap.reviewDocs);
          latestReviewsRef.current = repairedRevs;
          setReviewDocs(repairedRevs);
        }
        if (Array.isArray(idbSnap.submissionLogs) && idbSnap.submissionLogs.length > 0) {
          latestLogsRef.current = idbSnap.submissionLogs;
          setSubmissionLogs(idbSnap.submissionLogs);
        }
        if (Array.isArray(idbSnap.cvDiv1Records) && idbSnap.cvDiv1Records.length > 0) {
          latestCvRecordsRef.current = idbSnap.cvDiv1Records;
          setCvDiv1Records(idbSnap.cvDiv1Records);
          saveCvDiv1Records(idbSnap.cvDiv1Records);
        }
        if (Array.isArray(idbSnap.cvInventoryOps) && idbSnap.cvInventoryOps.length > 0) {
          const normOps = normalizeCvInventoryOpsList(idbSnap.cvInventoryOps);
          latestCvOpsRef.current = normOps;
          setCvInventoryOps(normOps);
          saveCvInventoryOps(normOps);
        }
      }
    });
  }, []);

  const showBriefHistoryNotice = (msg: string) => {
    setHistoryNotice(msg);
    setTimeout(() => {
      setHistoryNotice((prev) => (prev === msg ? null : prev));
    }, 2800);
  };

  const broadcastLocalSync = (
    nextTables: WeeklySalesTable[],
    nextReviews: ReviewMonitoringDoc[],
    nextLogs: ReportSubmissionLog[],
    updatedAt: number,
    nextCvRecords?: CvDailyRecord[],
    nextCvOps?: CvInventoryOpsEntry[]
  ) => {
    try {
      if (typeof BroadcastChannel !== 'undefined') {
        const ch = new BroadcastChannel('bazar_kapuas_realtime_sync');
        ch.postMessage({
          weeklyTables: nextTables,
          reviewDocs: nextReviews,
          submissionLogs: nextLogs,
          cvDiv1Records: nextCvRecords ?? latestCvRecordsRef.current,
          cvInventoryOps: nextCvOps ?? latestCvOpsRef.current,
          updatedAt,
        });
        ch.close();
      }
    } catch {}
  };

  // Persist to localStorage, IndexedDB & sync to backend server (which also writes bazar_store.json + persistedStore.ts)
  const syncStateToBackend = (
    nextTables: WeeklySalesTable[],
    nextReviews: ReviewMonitoringDoc[],
    nextLogs: ReportSubmissionLog[],
    isInitialSync: boolean = false,
    syncCv: boolean = false
  ) => {
    const nowTs = Math.max(Date.now(), lastServerUpdatedAtRef.current + 1);
    lastServerUpdatedAtRef.current = nowTs;
    if (!isInitialSync) {
      lastLocalEditAtRef.current = Date.now();
    }
    const cvRecsToSync = latestCvRecordsRef.current;
    const cvOpsToSync = latestCvOpsRef.current;
    const shouldSyncCvFirestore = syncCv || isInitialSync || pendingCvSyncRef.current;
    pendingCvSyncRef.current = false;

    latestTablesRef.current = nextTables;
    latestReviewsRef.current = nextReviews;
    latestLogsRef.current = nextLogs;
    try {
      if (!isInitialSync) {
        localStorage.setItem(STORAGE_KEY_PENDING_SYNC, '1');
      }
    } catch {}
    persistDataToLocalStorage(
      nextTables,
      nextReviews,
      nextLogs,
      nowTs,
      cvRecsToSync,
      cvOpsToSync
    );
    pendingSaveCountRef.current += 1;
    setIsSavingState(true);
    broadcastLocalSync(
      nextTables,
      nextReviews,
      nextLogs,
      nowTs,
      cvRecsToSync,
      cvOpsToSync
    );
    const firestorePromise = saveBazarWorkspaceToFirestore(
      nextTables,
      nextReviews,
      nextLogs,
      nowTs,
      accessRole === 'admin' ? 'admin' : 'public_sync'
    );
    const cvFirestorePromise = shouldSyncCvFirestore
      ? saveCvWorkspaceToFirestore(
          cvRecsToSync,
          cvOpsToSync,
          nowTs,
          accessRole === 'admin' ? 'admin' : 'public_sync'
        )
      : Promise.resolve(lastCvSyncedAtRef.current);
    const expressPromise = fetch('/api/bazar-data', {
      method: 'PUT',
      keepalive: true,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        weeklyTables: nextTables,
        reviewDocs: nextReviews,
        submissionLogs: nextLogs,
        cvDiv1Records: cvRecsToSync,
        cvInventoryOps: cvOpsToSync,
        updatedAt: nowTs,
        isInitialSync,
      }),
    }).then((r) => (r.ok ? r.json() : null));

    Promise.allSettled([firestorePromise, cvFirestorePromise, expressPromise])
      .then(([fsRes, cvFsRes, expRes]) => {
        const fsOk = fsRes.status === 'fulfilled';
        const fsTs = fsRes.status === 'fulfilled' ? Number(fsRes.value || 0) : 0;
        const cvFsTs = cvFsRes.status === 'fulfilled' ? Number(cvFsRes.value || 0) : 0;
        if (cvFsTs > lastCvSyncedAtRef.current) {
          lastCvSyncedAtRef.current = cvFsTs;
        }
        const resp = expRes.status === 'fulfilled' ? expRes.value : null;
        if (fsOk || resp) {
          try {
            localStorage.setItem(STORAGE_KEY_PENDING_SYNC, '0');
          } catch {}
          setHasUnsavedGlobalDraft(false);
          markAutoDraftClean(AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE);
          const confirmedTs = Math.max(
            lastServerUpdatedAtRef.current,
            fsTs,
            Number(resp?.state?.updatedAt || 0)
          );
          if (confirmedTs > 0) {
            lastServerUpdatedAtRef.current = confirmedTs;
            try {
              localStorage.setItem(STORAGE_KEY_UPDATED_AT, String(confirmedTs));
            } catch {}
          }
          setLastSavedTimeStr(
            new Date().toLocaleTimeString('id-ID', {
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit',
            })
          );
        } else {
          setHasUnsavedGlobalDraft(true);
        }
      })
      .finally(() => {
        pendingSaveCountRef.current = Math.max(0, pendingSaveCountRef.current - 1);
        setIsSavingState(false);
      });
  };

  // Fetch shared state from backend server on mount & poll automatically for new report submissions
  useEffect(() => {
    const computeServerSignature = (
      tablesData: any,
      reviewsData: any,
      logsData: any
    ): string => {
      try {
        const tSig = Array.isArray(tablesData)
          ? tablesData.map((t: any) => [t.weekNumber, t.days])
          : [];
        const lSig = Array.isArray(logsData)
          ? logsData
              .slice(0, 5)
              .map((l: any) => [
                l.id,
                l.timestamp,
                l.salesAmount,
                l.qtyAmount,
                l.trxAmount,
                l.summaryNote,
              ])
          : [];
        return JSON.stringify([tSig, reviewsData, lSig, logsData?.length || 0]);
      } catch {
        return '';
      }
    };

    const mirrorStateToLocalExpress = (
      tablesToMirror: WeeklySalesTable[],
      reviewsToMirror: ReviewMonitoringDoc[],
      logsToMirror: ReportSubmissionLog[],
      updatedAtMs: number,
      cvRecordsToMirror?: CvDailyRecord[],
      cvOpsToMirror?: CvInventoryOpsEntry[]
    ) => {
      fetch('/api/bazar-data', {
        method: 'PUT',
        keepalive: true,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          weeklyTables: tablesToMirror,
          reviewDocs: reviewsToMirror,
          submissionLogs: logsToMirror,
          ...(Array.isArray(cvRecordsToMirror) ? { cvDiv1Records: cvRecordsToMirror } : {}),
          ...(Array.isArray(cvOpsToMirror) ? { cvInventoryOps: cvOpsToMirror } : {}),
          updatedAt: updatedAtMs,
          isInitialSync: true,
          isMirrorFromFirestore: true,
        }),
      }).catch(() => {});
    };

    const applyRemoteCvState = (
      remoteCvRecsRaw: CvDailyRecord[] | undefined,
      remoteCvOpsRaw: CvInventoryOpsEntry[] | undefined,
      remoteCvTs: number,
      source: 'firestore' | 'express'
    ) => {
      if (pendingSaveCountRef.current > 0 || debouncedSyncTimerRef.current !== null) {
        return;
      }
      if (Date.now() - lastLocalEditAtRef.current < 8000) {
        return;
      }
      const nextRecs =
        Array.isArray(remoteCvRecsRaw) && remoteCvRecsRaw.length > 0
          ? remoteCvRecsRaw
          : latestCvRecordsRef.current;
      const nextOps =
        Array.isArray(remoteCvOpsRaw) && remoteCvOpsRaw.length > 0
          ? normalizeCvInventoryOpsList(remoteCvOpsRaw)
          : latestCvOpsRef.current;

      latestCvRecordsRef.current = nextRecs;
      latestCvOpsRef.current = nextOps;
      setCvDiv1Records(nextRecs);
      setCvInventoryOps(nextOps);
      saveCvDiv1Records(nextRecs);
      saveCvInventoryOps(nextOps);
      if (remoteCvTs > lastCvSyncedAtRef.current) {
        lastCvSyncedAtRef.current = remoteCvTs;
      }
      saveSnapshotToIndexedDB({
        weeklyTables: latestTablesRef.current,
        reviewDocs: latestReviewsRef.current,
        submissionLogs: latestLogsRef.current,
        cvDiv1Records: nextRecs,
        cvInventoryOps: nextOps,
        updatedAt: Math.max(lastServerUpdatedAtRef.current, remoteCvTs),
      });
      if (source === 'firestore') {
        mirrorStateToLocalExpress(
          latestTablesRef.current,
          latestReviewsRef.current,
          latestLogsRef.current,
          Math.max(lastServerUpdatedAtRef.current, remoteCvTs),
          nextRecs,
          nextOps
        );
      }
    };

    const applyAuthoritativeRemoteState = (
      remoteTablesRaw: WeeklySalesTable[],
      remoteReviewsRaw: ReviewMonitoringDoc[],
      remoteLogsRaw: ReportSubmissionLog[],
      remoteTs: number,
      isInitial: boolean,
      forceApply: boolean,
      source: 'firestore' | 'express'
    ) => {
      if (pendingSaveCountRef.current > 0 || debouncedSyncTimerRef.current !== null) {
        return;
      }
      if (!forceApply && Date.now() - lastLocalEditAtRef.current < 8000) {
        return;
      }
      // Never overwrite while admin is actively focused in an input/textarea editing data
      if (!isInitial && !forceApply) {
        const activeEl = document.activeElement;
        if (
          activeEl &&
          (activeEl.tagName === 'INPUT' || activeEl.tagName === 'TEXTAREA')
        ) {
          return;
        }
      }
      const repairedTables =
        Array.isArray(remoteTablesRaw) && remoteTablesRaw.length > 0
          ? normalizeAndRepairWeeklyTables(remoteTablesRaw)
          : latestTablesRef.current;
      const repairedReviews =
        Array.isArray(remoteReviewsRaw) && remoteReviewsRaw.length > 0
          ? normalizeReviewDocs(remoteReviewsRaw)
          : latestReviewsRef.current;
      const cleanLogs = Array.isArray(remoteLogsRaw)
        ? remoteLogsRaw.filter((lg: any) => lg && !/^log-[123]$/.test(lg.id))
        : latestLogsRef.current;

      const newSig = computeServerSignature(
        repairedTables,
        repairedReviews,
        cleanLogs
      );
      const sigChanged =
        Boolean(newSig) && newSig !== lastServerSignatureRef.current;

      if (
        !forceApply &&
        !isInitial &&
        remoteTs < lastServerUpdatedAtRef.current
      ) {
        return;
      }

      if (
        !forceApply &&
        !isInitial &&
        !sigChanged &&
        remoteTs <= lastServerUpdatedAtRef.current
      ) {
        return;
      }

      const prevTop = latestLogsRef.current[0];
      const newTop = cleanLogs[0];
      const isNewOrUpdatedReport =
        newTop &&
        (!prevTop ||
          newTop.id !== prevTop.id ||
          newTop.salesAmount !== prevTop.salesAmount ||
          newTop.qtyAmount !== prevTop.qtyAmount ||
          newTop.trxAmount !== prevTop.trxAmount ||
          newTop.summaryNote !== prevTop.summaryNote);

      const effectiveTs = Math.max(
        lastServerUpdatedAtRef.current,
        remoteTs || Date.now()
      );
      lastServerUpdatedAtRef.current = effectiveTs;
      lastServerSignatureRef.current = newSig;

      latestTablesRef.current = repairedTables;
      latestReviewsRef.current = repairedReviews;
      latestLogsRef.current = cleanLogs;
      setWeeklyTables(repairedTables);
      setReviewDocs(repairedReviews);
      setSubmissionLogs(cleanLogs);
      setHasUnsavedGlobalDraft(false);
      markAutoDraftClean(AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE);
      persistDataToLocalStorage(
        repairedTables,
        repairedReviews,
        cleanLogs,
        effectiveTs
      );
      setLastSavedTimeStr(
        new Date().toLocaleTimeString('id-ID', {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        })
      );

      if (!isInitial && isNewOrUpdatedReport) {
        showBriefHistoryNotice(
          `✓ Data tersinkron otomatis: ${newTop.submittedBy} (${newTop.tanggal || newTop.weekTitle || ''})`
        );
      }

      if (source === 'firestore') {
        mirrorStateToLocalExpress(
          repairedTables,
          repairedReviews,
          cleanLogs,
          effectiveTs
        );
      }
    };

    const pullFromServer = async (
      isInitial: boolean,
      forceApply: boolean = false
    ) => {
      if (
        pendingSaveCountRef.current > 0 ||
        debouncedSyncTimerRef.current !== null ||
        (!forceApply && Date.now() - lastLocalEditAtRef.current < 8000)
      ) {
        return;
      }
      const seqBeforePull = localEditSeqRef.current;
      const sinceParam =
        !isInitial && lastServerUpdatedAtRef.current > 0
          ? `?since=${lastServerUpdatedAtRef.current}`
          : '';

      // Routine background polling only needs a fast ~40-byte check from Express;
      // Firestore already pushes real-time changes instantly via subscriptions
      const shouldCheckFirestore = isInitial || forceApply;
      const [fsData, cvFsData, expressData] = await Promise.all([
        shouldCheckFirestore
          ? fetchBazarWorkspaceFromFirestore().catch(() => null)
          : Promise.resolve(null),
        shouldCheckFirestore
          ? fetchCvWorkspaceFromFirestore().catch(() => null)
          : Promise.resolve(null),
        fetch(`/api/bazar-data${sinceParam}`, { cache: 'no-store' })
          .then((r) => (r.ok ? r.json() : null))
          .catch(() => null),
      ]);

      if (
        pendingSaveCountRef.current > 0 ||
        debouncedSyncTimerRef.current !== null ||
        localEditSeqRef.current !== seqBeforePull ||
        (!forceApply && Date.now() - lastLocalEditAtRef.current < 8000)
      ) {
        return;
      }

      if (expressData?.notModified && !fsData && !cvFsData) {
        if (
          expressData.publicTunnelUrl &&
          typeof expressData.publicTunnelUrl === 'string'
        ) {
          setAutoDetectedPublicOrigin(expressData.publicTunnelUrl);
        }
        return;
      }

      if (
        expressData &&
        typeof expressData.publicTunnelUrl === 'string' &&
        expressData.publicTunnelUrl
      ) {
        setAutoDetectedPublicOrigin(expressData.publicTunnelUrl);
      }

      if (
        expressData &&
        typeof expressData.adminPassword === 'string' &&
        expressData.adminPassword.trim().length >= 3 &&
        expressData.adminPassword.trim() !== 'admin123'
      ) {
        const cleanPass = expressData.adminPassword.trim();
        setAdminPassword(cleanPass);
        try {
          localStorage.setItem(STORAGE_KEY_ADMIN_PASS, cleanPass);
        } catch {}
      }

      const fsTime = Number(fsData?.updatedAtMs || 0);
      const cvFsTime = Number(cvFsData?.updatedAtMs || 0);
      const serverTime = Number(expressData?.updatedAt || 0);
      const serverIsPersisted = Boolean(
        expressData?.isPersisted && serverTime > 0
      );
      const highestKnown = Math.max(
        lastServerUpdatedAtRef.current,
        fsTime
      );

      // Reconcile CV workspace between Firestore (/bazarWorkspaces/cv) and Express
      if (cvFsData && cvFsTime > 0 && (!serverIsPersisted || cvFsTime >= serverTime)) {
        if (cvFsTime > lastCvSyncedAtRef.current || isInitial || forceApply) {
          applyRemoteCvState(
            cvFsData.cvDiv1Records,
            cvFsData.cvInventoryOps,
            cvFsTime,
            'firestore'
          );
        }
      } else {
        const hasExpCvRecs =
          Array.isArray(expressData?.cvDiv1Records) && expressData.cvDiv1Records.length > 0;
        const hasExpCvOps =
          Array.isArray(expressData?.cvInventoryOps) && expressData.cvInventoryOps.length > 0;
        if (hasExpCvRecs || hasExpCvOps) {
          applyRemoteCvState(
            hasExpCvRecs ? expressData.cvDiv1Records : undefined,
            hasExpCvOps ? expressData.cvInventoryOps : undefined,
            serverTime || Date.now(),
            'express'
          );
          if (shouldCheckFirestore && (!cvFsData || serverTime > cvFsTime)) {
            saveCvWorkspaceToFirestore(
              hasExpCvRecs ? expressData.cvDiv1Records : latestCvRecordsRef.current,
              hasExpCvOps
                ? normalizeCvInventoryOpsList(expressData.cvInventoryOps)
                : latestCvOpsRef.current,
              serverTime || Date.now(),
              'admin'
            ).catch(() => {});
          }
        } else if (isInitial && !cvFsData) {
          saveCvWorkspaceToFirestore(
            latestCvRecordsRef.current,
            latestCvOpsRef.current,
            Math.max(Date.now(), lastServerUpdatedAtRef.current),
            'admin'
          ).catch(() => {});
        }
      }

      // 1. If Firestore has data and is newer than or equal to Express
      if (fsData && fsTime > 0 && (!serverIsPersisted || fsTime >= serverTime)) {
        if (fsTime > lastServerUpdatedAtRef.current || isInitial || forceApply) {
          applyAuthoritativeRemoteState(
            fsData.weeklyTables,
            fsData.reviewDocs,
            fsData.submissionLogs,
            fsTime,
            isInitial,
            forceApply,
            'firestore'
          );
        }
        return;
      }

      // 2. If Express server (Google AI Studio) has newer persisted data than Firestore and newer than confirmed client state
      if (
        expressData &&
        serverIsPersisted &&
        serverTime > highestKnown &&
        (!shouldCheckFirestore || serverTime > fsTime)
      ) {
        applyAuthoritativeRemoteState(
          expressData.weeklyTables,
          expressData.reviewDocs,
          expressData.submissionLogs,
          serverTime,
          isInitial,
          forceApply,
          'express'
        );
        // Push AI Studio's newer state to Firestore so Vercel App & Report Link immediately match AI Studio
        const tablesToPush =
          Array.isArray(expressData.weeklyTables) &&
          expressData.weeklyTables.length > 0
            ? normalizeAndRepairWeeklyTables(expressData.weeklyTables)
            : latestTablesRef.current;
        const reviewsToPush =
          Array.isArray(expressData.reviewDocs) &&
          expressData.reviewDocs.length > 0
            ? normalizeReviewDocs(expressData.reviewDocs)
            : latestReviewsRef.current;
        const logsToPush = Array.isArray(expressData.submissionLogs)
          ? expressData.submissionLogs.filter(
              (lg: any) => lg && !/^log-[123]$/.test(lg.id)
            )
          : latestLogsRef.current;
        saveBazarWorkspaceToFirestore(
          tablesToPush,
          reviewsToPush,
          logsToPush,
          serverTime,
          'admin'
        ).catch(() => {});
        return;
      }

      // 3. If neither Firestore nor Express is persisted yet, seed both from current state on initial load
      if (!fsData && !serverIsPersisted && isInitial) {
        syncStateToBackend(
          latestTablesRef.current,
          latestReviewsRef.current,
          latestLogsRef.current,
          true,
          true
        );
      }
    };

    pullFromServerRef.current = pullFromServer;
    pullFromServer(true, true);

    // Poll every 8s for external report submissions (never interrupts active editing, pauses in background)
    const intervalId = window.setInterval(() => {
      if (document.hidden) return;
      pullFromServer(false, false);
    }, 8000);

    const onFocus = () => pullFromServer(false, false);
    const onVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        pullFromServer(false, false);
      }
    };
    window.addEventListener('focus', onFocus);
    document.addEventListener('visibilitychange', onVisibilityChange);

    // Flush to localStorage on page unload / hide so closing or refreshing the app never loses state
    const onBeforeUnload = () => {
      persistDataToLocalStorage(
        latestTablesRef.current,
        latestReviewsRef.current,
        latestLogsRef.current,
        lastServerUpdatedAtRef.current,
        latestCvRecordsRef.current,
        latestCvOpsRef.current
      );
    };
    window.addEventListener('beforeunload', onBeforeUnload);

    // Instant cross-tab sync via BroadcastChannel & storage event
    let channel: BroadcastChannel | null = null;
    try {
      if (typeof BroadcastChannel !== 'undefined') {
        channel = new BroadcastChannel('bazar_kapuas_realtime_sync');
        channel.onmessage = (ev) => {
          const msg = ev.data;
          if (!msg) return;
          if (
            pendingSaveCountRef.current > 0 ||
            debouncedSyncTimerRef.current !== null ||
            Date.now() - lastLocalEditAtRef.current < 8000
          ) {
            return;
          }
          const incomingTs = Number(msg.updatedAt || 0);
          if (incomingTs > 0 && incomingTs <= lastServerUpdatedAtRef.current) {
            return;
          }
          if (incomingTs > 0) {
            lastServerUpdatedAtRef.current = incomingTs;
            try {
              localStorage.setItem(
                STORAGE_KEY_UPDATED_AT,
                String(lastServerUpdatedAtRef.current)
              );
            } catch {}
          }
          if (Array.isArray(msg.weeklyTables)) {
            const nextT = normalizeAndRepairWeeklyTables(msg.weeklyTables);
            latestTablesRef.current = nextT;
            setWeeklyTables(nextT);
          }
          if (Array.isArray(msg.reviewDocs)) {
            const nextR = normalizeReviewDocs(msg.reviewDocs);
            latestReviewsRef.current = nextR;
            setReviewDocs(nextR);
          }
          if (Array.isArray(msg.cvDiv1Records) && msg.cvDiv1Records.length > 0) {
            latestCvRecordsRef.current = msg.cvDiv1Records;
            setCvDiv1Records(msg.cvDiv1Records);
            saveCvDiv1Records(msg.cvDiv1Records);
          }
          if (Array.isArray(msg.cvInventoryOps) && msg.cvInventoryOps.length > 0) {
            const normOps = normalizeCvInventoryOpsList(msg.cvInventoryOps);
            latestCvOpsRef.current = normOps;
            setCvInventoryOps(normOps);
            saveCvInventoryOps(normOps);
          }
          if (Array.isArray(msg.submissionLogs)) {
            latestLogsRef.current = msg.submissionLogs;
            setSubmissionLogs(msg.submissionLogs);
            const newest = msg.submissionLogs[0];
            if (newest) {
              showBriefHistoryNotice(
                `Report baru otomatis masuk: ${newest.submittedBy} (${newest.tanggal || newest.weekTitle || ''})`
              );
            }
          }
        };
      }
    } catch {}

    const onStorage = (e: StorageEvent) => {
      if (
        pendingSaveCountRef.current > 0 ||
        debouncedSyncTimerRef.current !== null ||
        Date.now() - lastLocalEditAtRef.current < 8000
      ) {
        return;
      }
      try {
        if (e.key === STORAGE_KEY_UPDATED_AT && e.newValue) {
          const ts = Number(e.newValue);
          if (ts > 0) {
            lastServerUpdatedAtRef.current = Math.max(
              lastServerUpdatedAtRef.current,
              ts
            );
          }
        } else if (e.key === STORAGE_KEY_TABLES && e.newValue) {
          const parsed = JSON.parse(e.newValue);
          if (Array.isArray(parsed)) {
            const nextT = normalizeAndRepairWeeklyTables(parsed);
            latestTablesRef.current = nextT;
            setWeeklyTables(nextT);
          }
        } else if (e.key === STORAGE_KEY_REVIEWS && e.newValue) {
          const parsed = JSON.parse(e.newValue);
          if (Array.isArray(parsed)) {
            const nextR = normalizeReviewDocs(parsed);
            latestReviewsRef.current = nextR;
            setReviewDocs(nextR);
          }
        } else if (e.key === STORAGE_KEY_LOGS && e.newValue) {
          const parsed = JSON.parse(e.newValue);
          if (Array.isArray(parsed)) {
            latestLogsRef.current = parsed;
            setSubmissionLogs(parsed);
          }
        } else if (e.key === STORAGE_KEY_CV_DIV1 && e.newValue) {
          const parsed = JSON.parse(e.newValue);
          if (Array.isArray(parsed) && parsed.length > 0) {
            latestCvRecordsRef.current = parsed;
            setCvDiv1Records(parsed);
          }
        } else if (e.key === STORAGE_KEY_CV_INVENTORY_OPS && e.newValue) {
          const parsed = JSON.parse(e.newValue);
          if (Array.isArray(parsed)) {
            const normOps = normalizeCvInventoryOpsList(parsed);
            latestCvOpsRef.current = normOps;
            setCvInventoryOps(normOps);
          }
        }
      } catch {}
    };
    window.addEventListener('storage', onStorage);

    // Real-time Firebase Firestore subscriptions (syncs AI Studio <-> Vercel instantaneously)
    let unsubFirestore: (() => void) | null = null;
    let unsubCvFirestore: (() => void) | null = null;
    try {
      unsubFirestore = subscribeToBazarWorkspaceFirestore((fsData) => {
        if (
          !fsData ||
          pendingSaveCountRef.current > 0 ||
          debouncedSyncTimerRef.current !== null ||
          Date.now() - lastLocalEditAtRef.current < 8000
        ) {
          return;
        }
        const fsTime = Number(fsData.updatedAtMs || 0);
        if (fsTime <= 0 || fsTime <= lastServerUpdatedAtRef.current) return;
        applyAuthoritativeRemoteState(
          fsData.weeklyTables,
          fsData.reviewDocs,
          fsData.submissionLogs,
          fsTime,
          false,
          false,
          'firestore'
        );
      });
    } catch {}
    try {
      unsubCvFirestore = subscribeToCvWorkspaceFirestore((cvFsData) => {
        if (
          !cvFsData ||
          pendingSaveCountRef.current > 0 ||
          debouncedSyncTimerRef.current !== null ||
          Date.now() - lastLocalEditAtRef.current < 8000
        ) {
          return;
        }
        const cvFsTime = Number(cvFsData.updatedAtMs || 0);
        if (cvFsTime <= 0 || cvFsTime <= lastCvSyncedAtRef.current) return;
        applyRemoteCvState(
          cvFsData.cvDiv1Records,
          cvFsData.cvInventoryOps,
          cvFsTime,
          'firestore'
        );
      });
    } catch {}

    return () => {
      window.clearInterval(intervalId);
      window.removeEventListener('focus', onFocus);
      document.removeEventListener('visibilitychange', onVisibilityChange);
      window.removeEventListener('beforeunload', onBeforeUnload);
      window.removeEventListener('storage', onStorage);
      if (channel) channel.close();
      if (unsubFirestore) unsubFirestore();
      if (unsubCvFirestore) unsubCvFirestore();
    };
  }, []);

  // Access Control Helpers (Public View-Only vs Admin Edit)
  const handleEnterPublicMode = () => {
    setAccessRole('public');
    setShowEntryGate(false);
    try {
      sessionStorage.setItem(SESSION_KEY_ACCESS_ROLE, 'public');
      localStorage.setItem(STORAGE_KEY_ACCESS_ROLE, 'public');
    } catch {}
    // Immediately pull and auto-save any report submitted via link when Pelihat enters the app
    pullFromServerRef.current?.(false, true);
    showBriefHistoryNotice('Mode Publik aktif (Hanya Melihat / Read-Only) — Data Report Tersinkron Otomatis');
  };

  const handleLogoutToEntryGate = () => {
    setAccessRole('public');
    setIsLockedReportOnly(false);
    setShowEntryGate(true);
    try {
      sessionStorage.removeItem(SESSION_KEY_ACCESS_ROLE);
      localStorage.removeItem(STORAGE_KEY_ACCESS_ROLE);
    } catch {}
    showBriefHistoryNotice('Berhasil Logout — Kembali ke halaman Login utama');
  };

  const handleLoginAdminWithPassword = (passwordInput: string): boolean => {
    const cleanInput = passwordInput.trim();
    if (cleanInput === adminPassword || cleanInput === 'MDS311') {
      setAccessRole('admin');
      setShowEntryGate(false);
      try {
        sessionStorage.setItem(SESSION_KEY_ACCESS_ROLE, 'admin');
        localStorage.setItem(STORAGE_KEY_ACCESS_ROLE, 'admin');
      } catch {}
      // Immediately pull and auto-save any report submitted via link when Admin enters the app
      pullFromServerRef.current?.(false, true);
      showBriefHistoryNotice('Mode Admin aktif — Akses Edit dibuka & Data Report Tersinkron Otomatis');
      return true;
    }
    return false;
  };

  const handleChangeAdminPassword = (newPassword: string) => {
    const clean = newPassword.trim();
    if (clean.length < 3) return;
    setAdminPassword(clean);
    try {
      localStorage.setItem(STORAGE_KEY_ADMIN_PASS, clean);
    } catch {}
    fetch('/api/bazar-data', {
      method: 'PUT',
      keepalive: true,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        weeklyTables: latestTablesRef.current,
        reviewDocs: latestReviewsRef.current,
        submissionLogs: latestLogsRef.current,
        cvDiv1Records: latestCvRecordsRef.current,
        cvInventoryOps: latestCvOpsRef.current,
        adminPassword: clean,
        updatedAt: Date.now(),
      }),
    }).catch(() => {});
    showBriefHistoryNotice('Sandi Admin berhasil diperbarui');
  };

  const requireAdminForEdit = (): boolean => {
    if (accessRole === 'admin') return true;
    showBriefHistoryNotice(
      '🔒 Mode Publik (Hanya Lihat): Masukkan Sandi Admin untuk mengedit data'
    );
    return false;
  };

  const flushDebouncedEditSyncNow = () => {
    if (debouncedSyncTimerRef.current !== null) {
      window.clearTimeout(debouncedSyncTimerRef.current);
      debouncedSyncTimerRef.current = null;
    }
    if (pendingUndoBaseRef.current) {
      const snap = pendingUndoBaseRef.current;
      pendingUndoBaseRef.current = null;
      setUndoStack((prev) => [
        ...prev.slice(-49),
        {
          weeklyTables: cloneState(snap.weeklyTables),
          reviewDocs: cloneState(snap.reviewDocs),
          ...(snap.cvDiv1Records ? { cvDiv1Records: cloneState(snap.cvDiv1Records) } : {}),
          ...(snap.cvInventoryOps ? { cvInventoryOps: cloneState(snap.cvInventoryOps) } : {}),
          label: snap.label,
        },
      ]);
      setRedoStack([]);
    }
  };

  const scheduleDebouncedEditSync = (
    baseTables: WeeklySalesTable[],
    baseReviews: ReviewMonitoringDoc[],
    label: string,
    baseCvRecords?: CvDailyRecord[],
    baseCvOps?: CvInventoryOpsEntry[]
  ) => {
    if (!pendingUndoBaseRef.current) {
      pendingUndoBaseRef.current = {
        weeklyTables: baseTables,
        reviewDocs: baseReviews,
        cvDiv1Records: baseCvRecords ?? latestCvRecordsRef.current,
        cvInventoryOps: baseCvOps ?? latestCvOpsRef.current,
        label,
      };
    }
    if (debouncedSyncTimerRef.current !== null) {
      window.clearTimeout(debouncedSyncTimerRef.current);
    }
    debouncedSyncTimerRef.current = window.setTimeout(() => {
      debouncedSyncTimerRef.current = null;
      if (pendingUndoBaseRef.current) {
        const snap = pendingUndoBaseRef.current;
        pendingUndoBaseRef.current = null;
        setUndoStack((prev) => [
          ...prev.slice(-49),
          {
            weeklyTables: cloneState(snap.weeklyTables),
            reviewDocs: cloneState(snap.reviewDocs),
            ...(snap.cvDiv1Records ? { cvDiv1Records: cloneState(snap.cvDiv1Records) } : {}),
            ...(snap.cvInventoryOps ? { cvInventoryOps: cloneState(snap.cvInventoryOps) } : {}),
            label: snap.label,
          },
        ]);
        setRedoStack([]);
      }
      syncStateToBackend(
        latestTablesRef.current,
        latestReviewsRef.current,
        latestLogsRef.current
      );
    }, 900);
  };

  const handleManualSaveNow = (customLabel?: string) => {
    flushDebouncedEditSyncNow();
    const nowTs = Math.max(Date.now(), lastServerUpdatedAtRef.current + 1);
    lastServerUpdatedAtRef.current = nowTs;
    setHasUnsavedGlobalDraft(false);
    markAutoDraftClean(AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE);
    markAutoDraftClean(AUTO_DRAFT_KEYS.DAILY_INPUT);
    markAutoDraftClean(AUTO_DRAFT_KEYS.SALES_MATRIX);
    markAutoDraftClean(AUTO_DRAFT_KEYS.REASON_EVALUATION);
    markAutoDraftClean(AUTO_DRAFT_KEYS.REVIEW_MONITORING);
    persistDataToLocalStorage(
      latestTablesRef.current,
      latestReviewsRef.current,
      latestLogsRef.current,
      nowTs,
      latestCvRecordsRef.current,
      latestCvOpsRef.current
    );
    syncStateToBackend(
      latestTablesRef.current,
      latestReviewsRef.current,
      latestLogsRef.current,
      false,
      true
    );
    const timeStr = new Date().toLocaleTimeString('id-ID', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    setLastSavedTimeStr(timeStr);
    showBriefHistoryNotice(
      customLabel || `Data berhasil disimpan (${timeStr})`
    );
  };

  const handleRestoreSavedState = (
    nextTables: WeeklySalesTable[],
    nextReviews: ReviewMonitoringDoc[],
    nextLogs: ReportSubmissionLog[],
    sourceLabel: string,
    nextCvRecords?: CvDailyRecord[],
    nextCvOps?: CvInventoryOpsEntry[]
  ) => {
    if (!requireAdminForEdit()) return;
    flushDebouncedEditSyncNow();
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    setUndoStack((prev) => [
      ...prev.slice(-49),
      {
        weeklyTables: cloneState(latestTablesRef.current),
        reviewDocs: cloneState(latestReviewsRef.current),
        cvDiv1Records: cloneState(latestCvRecordsRef.current),
        cvInventoryOps: cloneState(latestCvOpsRef.current),
        label: sourceLabel,
      },
    ]);
    setRedoStack([]);
    latestTablesRef.current = nextTables;
    latestReviewsRef.current = nextReviews;
    latestLogsRef.current = nextLogs;
    setWeeklyTables(nextTables);
    setReviewDocs(nextReviews);
    setSubmissionLogs(nextLogs);
    if (Array.isArray(nextCvRecords) && nextCvRecords.length > 0) {
      latestCvRecordsRef.current = nextCvRecords;
      setCvDiv1Records(nextCvRecords);
      saveCvDiv1Records(nextCvRecords);
    }
    if (Array.isArray(nextCvOps) && nextCvOps.length > 0) {
      const normalizedOps = normalizeCvInventoryOpsList(nextCvOps);
      latestCvOpsRef.current = normalizedOps;
      setCvInventoryOps(normalizedOps);
      saveCvInventoryOps(normalizedOps);
    }
    if (nextTables[0]) {
      setSelectedWeeklyFilter(nextTables[0].id);
    }
    syncStateToBackend(nextTables, nextReviews, nextLogs, false, true);
    showBriefHistoryNotice(`${sourceLabel} berhasil diterapkan & disimpan`);
  };

  const handleUpdateTables = (nextTables: WeeklySalesTable[]) => {
    if (!requireAdminForEdit()) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    lastServerUpdatedAtRef.current = Math.max(
      Date.now(),
      lastServerUpdatedAtRef.current + 1
    );
    const currentTables = latestTablesRef.current;
    const currentReviews = latestReviewsRef.current;
    latestTablesRef.current = nextTables;
    setWeeklyTables(nextTables);
    scheduleDebouncedEditSync(
      currentTables,
      currentReviews,
      'Perubahan Tabel Sales'
    );
  };

  const handleUpdateReviewDocs = (nextReviews: ReviewMonitoringDoc[]) => {
    if (!requireAdminForEdit()) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    lastServerUpdatedAtRef.current = Math.max(
      Date.now(),
      lastServerUpdatedAtRef.current + 1
    );
    const currentTables = latestTablesRef.current;
    const currentReviews = latestReviewsRef.current;
    latestReviewsRef.current = nextReviews;
    setReviewDocs(nextReviews);
    scheduleDebouncedEditSync(
      currentTables,
      currentReviews,
      'Perubahan Review Bazar'
    );
  };

  const handleUndo = () => {
    if (!requireAdminForEdit()) return;
    flushDebouncedEditSyncNow();
    if (undoStack.length === 0) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    const lastSnapshot = undoStack[undoStack.length - 1];
    const nextUndo = undoStack.slice(0, -1);

    setRedoStack((prev) => [
      ...prev.slice(-49),
      {
        weeklyTables: cloneState(latestTablesRef.current),
        reviewDocs: cloneState(latestReviewsRef.current),
        cvDiv1Records: cloneState(latestCvRecordsRef.current),
        cvInventoryOps: cloneState(latestCvOpsRef.current),
        label: lastSnapshot.label,
      },
    ]);
    setUndoStack(nextUndo);

    latestTablesRef.current = lastSnapshot.weeklyTables;
    latestReviewsRef.current = lastSnapshot.reviewDocs;
    setWeeklyTables(lastSnapshot.weeklyTables);
    setReviewDocs(lastSnapshot.reviewDocs);
    let hasCvChange = false;
    if (Array.isArray(lastSnapshot.cvDiv1Records)) {
      latestCvRecordsRef.current = lastSnapshot.cvDiv1Records;
      setCvDiv1Records(lastSnapshot.cvDiv1Records);
      saveCvDiv1Records(lastSnapshot.cvDiv1Records);
      hasCvChange = true;
    }
    if (Array.isArray(lastSnapshot.cvInventoryOps)) {
      const normOps = normalizeCvInventoryOpsList(lastSnapshot.cvInventoryOps);
      latestCvOpsRef.current = normOps;
      setCvInventoryOps(normOps);
      saveCvInventoryOps(normOps);
      hasCvChange = true;
    }
    syncStateToBackend(
      lastSnapshot.weeklyTables,
      lastSnapshot.reviewDocs,
      latestLogsRef.current,
      false,
      hasCvChange
    );
    showBriefHistoryNotice('Undo berhasil & disimpan');
  };

  const handleRedo = () => {
    if (!requireAdminForEdit()) return;
    flushDebouncedEditSyncNow();
    if (redoStack.length === 0) return;
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    const nextSnapshot = redoStack[redoStack.length - 1];
    const nextRedo = redoStack.slice(0, -1);

    setUndoStack((prev) => [
      ...prev.slice(-49),
      {
        weeklyTables: cloneState(latestTablesRef.current),
        reviewDocs: cloneState(latestReviewsRef.current),
        cvDiv1Records: cloneState(latestCvRecordsRef.current),
        cvInventoryOps: cloneState(latestCvOpsRef.current),
        label: nextSnapshot.label,
      },
    ]);
    setRedoStack(nextRedo);

    latestTablesRef.current = nextSnapshot.weeklyTables;
    latestReviewsRef.current = nextSnapshot.reviewDocs;
    setWeeklyTables(nextSnapshot.weeklyTables);
    setReviewDocs(nextSnapshot.reviewDocs);
    let hasCvChange = false;
    if (Array.isArray(nextSnapshot.cvDiv1Records)) {
      latestCvRecordsRef.current = nextSnapshot.cvDiv1Records;
      setCvDiv1Records(nextSnapshot.cvDiv1Records);
      saveCvDiv1Records(nextSnapshot.cvDiv1Records);
      hasCvChange = true;
    }
    if (Array.isArray(nextSnapshot.cvInventoryOps)) {
      const normOps = normalizeCvInventoryOpsList(nextSnapshot.cvInventoryOps);
      latestCvOpsRef.current = normOps;
      setCvInventoryOps(normOps);
      saveCvInventoryOps(normOps);
      hasCvChange = true;
    }
    syncStateToBackend(
      nextSnapshot.weeklyTables,
      nextSnapshot.reviewDocs,
      latestLogsRef.current,
      false,
      hasCvChange
    );
    showBriefHistoryNotice('Redo berhasil & disimpan');
  };

  // Global keyboard shortcuts: Ctrl+S (Save), Ctrl+Z / Cmd+Z (Undo), Ctrl+Y / Cmd+Shift+Z (Redo)
  const shortcutHandlersRef = useRef({
    onSave: () => handleManualSaveNow(),
    onUndo: () => handleUndo(),
    onRedo: () => handleRedo(),
    canUndo: false,
    canRedo: false,
  });
  shortcutHandlersRef.current = {
    onSave: () => handleManualSaveNow(),
    onUndo: () => handleUndo(),
    onRedo: () => handleRedo(),
    canUndo: undoStack.length > 0,
    canRedo: redoStack.length > 0,
  };

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const isCtrlOrMeta = e.ctrlKey || e.metaKey;
      if (!isCtrlOrMeta) return;

      if (e.key.toLowerCase() === 's') {
        e.preventDefault();
        shortcutHandlersRef.current.onSave();
      } else if (e.key.toLowerCase() === 'z' && !e.shiftKey) {
        if (shortcutHandlersRef.current.canUndo) {
          e.preventDefault();
          shortcutHandlersRef.current.onUndo();
        }
      } else if (
        e.key.toLowerCase() === 'y' ||
        (e.key.toLowerCase() === 'z' && e.shiftKey)
      ) {
        if (shortcutHandlersRef.current.canRedo) {
          e.preventDefault();
          shortcutHandlersRef.current.onRedo();
        }
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  const handleAddSubmissionLog = (newLog: ReportSubmissionLog) => {
    const nextLogs = [newLog, ...latestLogsRef.current];
    latestLogsRef.current = nextLogs;
    setSubmissionLogs(nextLogs);
    try {
      localStorage.setItem(STORAGE_KEY_LOGS, JSON.stringify(nextLogs));
    } catch {}
    syncStateToBackend(latestTablesRef.current, latestReviewsRef.current, nextLogs);
  };

  const handleAtomicReportSubmit = (payload: {
    nextTables?: WeeklySalesTable[];
    nextReviewDocs?: ReviewMonitoringDoc[];
    newLog: ReportSubmissionLog;
    targetWeekId?: string;
    targetTable?: WeeklySalesTable;
    dayIdx?: number;
    updatedDay?: DailySalesRecord;
    isAutoSave?: boolean;
    shiftPatch?: {
      tanggal?: string;
      shouldUpdatePagi: boolean;
      shouldUpdateSiang: boolean;
      pagi?: {
        nama: string;
        sales: number | null;
        qty: number | null;
        trx: number | null;
        traffic?: number | null;
      };
      siang?: {
        nama: string;
        sales: number | null;
        qty: number | null;
        trx: number | null;
        traffic?: number | null;
      };
      reasonPagi?: string;
      reasonSiang?: string;
    };
  }) => {
    flushDebouncedEditSyncNow();
    localEditSeqRef.current += 1;
    const seqAtSubmit = localEditSeqRef.current;
    const finalTables = payload.nextTables
      ? normalizeAndRepairWeeklyTables(payload.nextTables)
      : latestTablesRef.current;
    const finalReviews = payload.nextReviewDocs
      ? mergeReviewDocsNonDestructive(
          payload.nextReviewDocs,
          latestReviewsRef.current
        )
      : latestReviewsRef.current;
    const finalLogs = [
      payload.newLog,
      ...latestLogsRef.current.filter((lg) => lg.id !== payload.newLog.id),
    ];

    if (!payload.isAutoSave) {
      setUndoStack((prev) => [
        ...prev.slice(-49),
        {
          weeklyTables: cloneState(latestTablesRef.current),
          reviewDocs: cloneState(latestReviewsRef.current),
          label: 'Pengisian Report via Link',
        },
      ]);
      setRedoStack([]);
    }

    lastLocalEditAtRef.current = Date.now();
    latestTablesRef.current = finalTables;
    latestReviewsRef.current = finalReviews;
    latestLogsRef.current = finalLogs;
    setWeeklyTables(finalTables);
    setReviewDocs(finalReviews);
    setSubmissionLogs(finalLogs);
    setHasUnsavedGlobalDraft(false);
    markAutoDraftClean(AUTO_DRAFT_KEYS.GLOBAL_WORKSPACE);

    if (payload.targetWeekId && selectedWeeklyFilter !== 'ALL') {
      setSelectedWeeklyFilter(payload.targetWeekId);
    }

    const nowTs = Math.max(Date.now(), lastServerUpdatedAtRef.current + 1);
    lastServerUpdatedAtRef.current = nowTs;
    persistDataToLocalStorage(finalTables, finalReviews, finalLogs, nowTs);
    pendingSaveCountRef.current += 1;
    broadcastLocalSync(finalTables, finalReviews, finalLogs, nowTs);
    const fsSavePromise = saveBazarWorkspaceToFirestore(
      finalTables,
      finalReviews,
      finalLogs,
      nowTs,
      'report_link'
    );
    const fsLogPromise = saveReportSubmissionToFirestore(payload.newLog);
    setLastSavedTimeStr(
      new Date().toLocaleTimeString('id-ID', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      })
    );

    const expressSubmitPromise = fetch('/api/bazar-submit-report', {
      method: 'POST',
      keepalive: true,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        targetTable: payload.targetTable,
        dayIdx: payload.dayIdx,
        updatedDay: payload.updatedDay,
        shiftPatch: payload.shiftPatch,
        updatedReviewDocs: finalReviews,
        newLog: payload.newLog,
      }),
    }).then((r) => (r.ok ? r.json() : null));

    Promise.allSettled([fsSavePromise, fsLogPromise, expressSubmitPromise])
      .then(([fsRes, , expRes]) => {
        if (localEditSeqRef.current !== seqAtSubmit) return;
        const fsOk = fsRes.status === 'fulfilled';
        const fsTs = fsRes.status === 'fulfilled' ? Number(fsRes.value || 0) : 0;
        const resp = expRes.status === 'fulfilled' ? expRes.value : null;
        const confirmedTs = Math.max(
          lastServerUpdatedAtRef.current,
          fsTs,
          Number(resp?.state?.updatedAt || nowTs)
        );
        lastServerUpdatedAtRef.current = confirmedTs;
        if (fsOk) {
          persistDataToLocalStorage(
            finalTables,
            finalReviews,
            finalLogs,
            confirmedTs
          );
          broadcastLocalSync(
            finalTables,
            finalReviews,
            finalLogs,
            confirmedTs
          );
        } else if (resp?.state) {
          const serverTables = Array.isArray(resp.state.weeklyTables)
            ? normalizeAndRepairWeeklyTables(resp.state.weeklyTables)
            : finalTables;
          const serverReviews = Array.isArray(resp.state.reviewDocs)
            ? normalizeReviewDocs(resp.state.reviewDocs)
            : finalReviews;
          const serverLogs = Array.isArray(resp.state.submissionLogs)
            ? resp.state.submissionLogs
            : finalLogs;
          latestTablesRef.current = serverTables;
          latestReviewsRef.current = serverReviews;
          latestLogsRef.current = serverLogs;
          setWeeklyTables(serverTables);
          setReviewDocs(serverReviews);
          setSubmissionLogs(serverLogs);
          persistDataToLocalStorage(
            serverTables,
            serverReviews,
            serverLogs,
            confirmedTs
          );
          broadcastLocalSync(
            serverTables,
            serverReviews,
            serverLogs,
            confirmedTs
          );
        }
      })
      .finally(() => {
        pendingSaveCountRef.current = Math.max(0, pendingSaveCountRef.current - 1);
      });

    if (!payload.isAutoSave) {
      showBriefHistoryNotice(
        `✓ Report berhasil masuk & tersimpan di aplikasi (${payload.newLog.tanggal || payload.newLog.weekTitle || ''})`
      );
    }
  };

  const handleAtomicCvReportSubmit = (payload: {
    updatedCvRecord?: CvDailyRecord;
    updatedCvOpsEntries?: CvInventoryOpsEntry[];
    newLog: ReportSubmissionLog;
  }) => {
    flushDebouncedEditSyncNow();
    localEditSeqRef.current += 1;
    const seqAtSubmit = localEditSeqRef.current;

    let finalCvRecords = [...latestCvRecordsRef.current];
    if (payload.updatedCvRecord) {
      const targetTgl = payload.updatedCvRecord.tanggal.trim();
      const idx = finalCvRecords.findIndex(
        (r) => r.tanggal.trim() === targetTgl
      );
      if (idx >= 0) {
        finalCvRecords[idx] = payload.updatedCvRecord;
      } else {
        finalCvRecords.push(payload.updatedCvRecord);
      }
    }

    let finalCvOps = [...latestCvOpsRef.current];
    if (
      Array.isArray(payload.updatedCvOpsEntries) &&
      payload.updatedCvOpsEntries.length > 0
    ) {
      const incomingMap = new Map<string, CvInventoryOpsEntry>();
      payload.updatedCvOpsEntries.forEach((entry) => {
        incomingMap.set(entry.id, entry);
      });
      finalCvOps = finalCvOps.map((existing) =>
        incomingMap.has(existing.id)
          ? { ...existing, ...incomingMap.get(existing.id)! }
          : existing
      );
      incomingMap.forEach((entry, id) => {
        if (!finalCvOps.some((e) => e.id === id)) {
          finalCvOps.push(entry);
        }
      });
      finalCvOps = normalizeCvInventoryOpsList(finalCvOps);
    }

    const finalLogs = [
      payload.newLog,
      ...latestLogsRef.current.filter((lg) => lg.id !== payload.newLog.id),
    ];

    lastLocalEditAtRef.current = Date.now();
    latestCvRecordsRef.current = finalCvRecords;
    latestCvOpsRef.current = finalCvOps;
    latestLogsRef.current = finalLogs;
    pendingCvSyncRef.current = true;
    setCvDiv1Records(finalCvRecords);
    saveCvDiv1Records(finalCvRecords);
    setCvInventoryOps(finalCvOps);
    saveCvInventoryOps(finalCvOps);
    setSubmissionLogs(finalLogs);

    const nowTs = Math.max(Date.now(), lastServerUpdatedAtRef.current + 1);
    lastServerUpdatedAtRef.current = nowTs;
    persistDataToLocalStorage(
      latestTablesRef.current,
      latestReviewsRef.current,
      finalLogs,
      nowTs,
      finalCvRecords,
      finalCvOps
    );
    pendingSaveCountRef.current += 1;
    broadcastLocalSync(
      latestTablesRef.current,
      latestReviewsRef.current,
      finalLogs,
      nowTs,
      finalCvRecords,
      finalCvOps
    );

    const fsCvPromise = saveCvWorkspaceToFirestore(
      finalCvRecords,
      finalCvOps,
      nowTs,
      'report_link'
    );
    const fsMainPromise = saveBazarWorkspaceToFirestore(
      latestTablesRef.current,
      latestReviewsRef.current,
      finalLogs,
      nowTs,
      'report_link'
    );
    const fsLogPromise = saveReportSubmissionToFirestore(payload.newLog);

    const expressSubmitPromise = fetch('/api/bazar-submit-cv-report', {
      method: 'POST',
      keepalive: true,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        updatedCvRecord: payload.updatedCvRecord,
        updatedCvOpsEntries: payload.updatedCvOpsEntries,
        newLog: payload.newLog,
      }),
    }).then((r) => (r.ok ? r.json() : null));

    setLastSavedTimeStr(
      new Date().toLocaleTimeString('id-ID', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
      })
    );

    Promise.allSettled([
      fsCvPromise,
      fsMainPromise,
      fsLogPromise,
      expressSubmitPromise,
    ])
      .then(([fsCvRes, , , expRes]) => {
        if (localEditSeqRef.current !== seqAtSubmit) return;
        const fsTs =
          fsCvRes.status === 'fulfilled' ? Number(fsCvRes.value || 0) : 0;
        const resp = expRes.status === 'fulfilled' ? expRes.value : null;
        const confirmedTs = Math.max(
          lastServerUpdatedAtRef.current,
          fsTs,
          Number(resp?.state?.updatedAt || nowTs)
        );
        lastServerUpdatedAtRef.current = confirmedTs;
        persistDataToLocalStorage(
          latestTablesRef.current,
          latestReviewsRef.current,
          finalLogs,
          confirmedTs,
          finalCvRecords,
          finalCvOps
        );
      })
      .finally(() => {
        pendingSaveCountRef.current = Math.max(
          0,
          pendingSaveCountRef.current - 1
        );
      });

    showBriefHistoryNotice(
      `✓ ${
        payload.newLog.reportType === 'data_cv' ? 'Data CV' : 'Report CV'
      } berhasil tersimpan & disinkronkan (${
        payload.newLog.tanggal || payload.newLog.weekTitle || ''
      })`
    );
  };

  const handleResetData = (mode: 'empty' | 'initial' = 'initial') => {
    if (!requireAdminForEdit()) return;
    flushDebouncedEditSyncNow();
    localEditSeqRef.current += 1;
    lastLocalEditAtRef.current = Date.now();
    setUndoStack((prev) => [
      ...prev.slice(-49),
      {
        weeklyTables: cloneState(latestTablesRef.current),
        reviewDocs: cloneState(latestReviewsRef.current),
        cvDiv1Records: cloneState(latestCvRecordsRef.current),
        cvInventoryOps: cloneState(latestCvOpsRef.current),
        label: mode === 'empty' ? 'Reset Kosongkan Data (0)' : 'Reset Data Awal',
      },
    ]);
    setRedoStack([]);

    const nextTables =
      mode === 'empty' ? createEmptyInitialWeeklyTables() : INITIAL_WEEKLY_TABLES;
    const nextReviews = INITIAL_REVIEW_DOCS;
    const nextLogs = mode === 'empty' ? [] : INITIAL_SUBMISSION_LOGS;

    const nextCv = generateInitialCvDiv1Records();
    const nextCvOps = normalizeCvInventoryOpsList([]);
    latestTablesRef.current = nextTables;
    latestReviewsRef.current = nextReviews;
    latestLogsRef.current = nextLogs;
    latestCvRecordsRef.current = nextCv;
    latestCvOpsRef.current = nextCvOps;
    setWeeklyTables(nextTables);
    setReviewDocs(nextReviews);
    setSubmissionLogs(nextLogs);
    setCvDiv1Records(nextCv);
    saveCvDiv1Records(nextCv);
    setCvInventoryOps(nextCvOps);
    saveCvInventoryOps(nextCvOps);
    setSelectedWeeklyFilter(nextTables[0].id);
    setActiveReviewId(nextReviews[0].id);
    setConfirmReset(false);

    const nowTs = Math.max(Date.now(), lastServerUpdatedAtRef.current + 1);
    lastServerUpdatedAtRef.current = nowTs;
    persistDataToLocalStorage(nextTables, nextReviews, nextLogs, nowTs, nextCv, nextCvOps);
    pendingSaveCountRef.current += 1;
    broadcastLocalSync(nextTables, nextReviews, nextLogs, nowTs, nextCv, nextCvOps);
    const fsResetPromise = saveBazarWorkspaceToFirestore(
      nextTables,
      nextReviews,
      nextLogs,
      nowTs,
      'admin'
    );
    const cvFsResetPromise = saveCvWorkspaceToFirestore(
      nextCv,
      nextCvOps,
      nowTs,
      'admin'
    );

    const expressResetPromise = fetch('/api/bazar-reset', {
      method: 'POST',
      keepalive: true,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mode }),
    }).then((r) => (r.ok ? r.json() : null));

    Promise.allSettled([fsResetPromise, cvFsResetPromise, expressResetPromise])
      .then(([, , expRes]) => {
        const resp = expRes.status === 'fulfilled' ? expRes.value : null;
        const updatedAt = resp?.state?.updatedAt || resp?.updatedAt;
        if (updatedAt) {
          const confirmedTs = Math.max(
            lastServerUpdatedAtRef.current,
            Number(updatedAt)
          );
          lastServerUpdatedAtRef.current = confirmedTs;
          try {
            localStorage.setItem(STORAGE_KEY_UPDATED_AT, String(confirmedTs));
          } catch {}
        }
      })
      .finally(() => {
        pendingSaveCountRef.current = Math.max(0, pendingSaveCountRef.current - 1);
      });

    showBriefHistoryNotice(
      mode === 'empty'
        ? 'Seluruh data sales berhasil dikosongkan (Reset 0)'
        : 'Data berhasil dikembalikan ke data awal'
    );
  };

  const handleOpenWeekly = (
    weekNumber: number,
    targetMenu?: 'sales_matrix' | 'dashboard' | 'daily_input' | 'reason_evaluation'
  ) => {
    const existing = weeklyTables.find((t) => t.weekNumber === weekNumber);
    let targetId = existing?.id;

    if (!existing) {
      if (accessRole !== 'admin') {
        showBriefHistoryNotice(
          '🔒 Mode Publik: Hanya Admin yang dapat membuka/menambah periode Weekly baru'
        );
        return;
      }
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNumber);
      if (sched) {
        const created = createWeeklyTableFromAnnualSchedule(sched);
        const next = [...weeklyTables, created].sort(
          (a, b) => a.weekNumber - b.weekNumber
        );
        handleUpdateTables(next);
        targetId = created.id;
      }
    }

    if (targetId) {
      setSelectedWeeklyFilter(targetId);
    }

    // If currently on Review Monitoring menu, also switch/prepare the Review doc for that Weekly
    const existingRev = reviewDocs.find(
      (d) => getReviewDocWeekNumber(d) === weekNumber
    );
    if (existingRev) {
      setActiveReviewId(existingRev.id);
    } else if (activeMenu === 'review_monitoring' && accessRole === 'admin') {
      const createdRev = createReviewDocForWeek(weekNumber, reviewDocs[0]);
      const nextRevs = [...reviewDocs, createdRev].sort(
        (a, b) =>
          (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
      );
      handleUpdateReviewDocs(nextRevs);
      setActiveReviewId(createdRev.id);
    }

    if (targetMenu === 'dashboard') {
      setDashboardHorizon('weekly');
      handleNavigate('dashboard');
    } else if (targetMenu) {
      handleNavigate(targetMenu);
    } else if (activeMenu === 'annual_schedule') {
      handleNavigate('sales_matrix');
    }
  };

  const handleQuickAddWeek = () => {
    const existingNums = new Set(weeklyTables.map((t) => t.weekNumber));
    const activeT =
      weeklyTables.find((t) => t.id === selectedWeeklyFilter) || weeklyTables[0];
    let nextWeekNum = (activeT?.weekNumber || 41) + 1;
    if (nextWeekNum > 53 || existingNums.has(nextWeekNum)) {
      const firstUnopened = ANNUAL_2026_WEEKS.find(
        (w) => !existingNums.has(w.weekNumber)
      );
      nextWeekNum = firstUnopened ? firstUnopened.weekNumber : 43;
    }
    handleOpenWeekly(nextWeekNum, 'sales_matrix');
  };

  const handleExportExcel = () => {
    if (accessRole !== 'admin') {
      showBriefHistoryNotice(
        '🔒 Mode Publik (Hanya Lihat): Tidak dapat klik Export Excel'
      );
      return;
    }
    setIsExportModalOpen(true);
  };

  const handlePrint = () => {
    if (accessRole !== 'admin') {
      showBriefHistoryNotice(
        '🔒 Mode Publik (Hanya Lihat): Tidak dapat klik Cetak / Export'
      );
      return;
    }
    window.print();
  };

  const handleNavigate = (menu: NavMenuId) => {
    // Jika sedang dalam mode Link Pengisian Report Terkunci, blokir semua navigasi ke menu lain
    if (isLockedReportOnly && menu !== 'report_link') {
      setActiveMenu('report_link');
      setMobileMenuOpen(false);
      window.scrollTo({ top: 0, behavior: 'instant' });
      return;
    }
    setActiveMenu(menu);
    setMobileMenuOpen(false);
    window.scrollTo({ top: 0, behavior: 'instant' });
  };

  const handleSelectDashboardHorizon = (h: DashboardHorizon) => {
    if (isLockedReportOnly) return;
    setDashboardHorizon(h);
    setActiveMenu('dashboard');
    setMobileMenuOpen(false);
    window.scrollTo({ top: 0, behavior: 'instant' });
  };

  const handleSelectSalesTableHorizon = (h: SalesTableHorizon) => {
    if (isLockedReportOnly) return;
    setSalesTableHorizon(h);
    setActiveMenu('sales_matrix');
    setMobileMenuOpen(false);
    window.scrollTo({ top: 0, behavior: 'instant' });
  };

  const handleSelectCvHorizon = (h: CvHorizon) => {
    if (isLockedReportOnly) return;
    setCvHorizon(h);
    setActiveMenu('report_cv_div1');
    setMobileMenuOpen(false);
    window.scrollTo({ top: 0, behavior: 'instant' });
  };

  const handleSelectTotalDpCvHorizon = (h: TotalHorizon) => {
    if (isLockedReportOnly) return;
    setTotalDpCvHorizon(h);
    setActiveMenu('total_dp_cv');
    setMobileMenuOpen(false);
    window.scrollTo({ top: 0, behavior: 'instant' });
  };

  // Sync activeWeekIndex for DailyInputWorkspace with selectedWeeklyFilter
  const activeWeekIndex = Math.max(
    0,
    weeklyTables.findIndex((t) => t.id === selectedWeeklyFilter)
  );

  const handleSelectWeekIndex = (idx: number) => {
    const target = weeklyTables[idx];
    if (target) {
      setSelectedWeeklyFilter(target.id);
    }
  };

  const currentReviewDoc =
    reviewDocs.find((d) => d.id === activeReviewId) || reviewDocs[0];

  const getHeaderBreadcrumbTitle = () => {
    if (activeMenu === 'dashboard') {
      if (dashboardHorizon === 'harian') return 'Dashboard Harian';
      if (dashboardHorizon === 'weekly') return 'Dashboard Weekly (Mingguan)';
      return 'Dashboard Bulanan';
    }
    if (activeMenu === 'sales_matrix') {
      if (salesTableHorizon === 'harian') return 'Report DP · Sales Daily (Pagi vs Siang & Rekap Harian)';
      if (salesTableHorizon === 'bulanan') return 'Report DP · Sales Monthly (Sales Tanggal 1–30/31 & Rekap 12 Bulan)';
      return 'Report DP · Sales Weekly (Mingguan)';
    }
    if (activeMenu === 'annual_schedule')
      return 'Data Kalender 52 Weekly Setahun (Senin s/d Minggu)';
    if (activeMenu === 'report_cv_div1') {
      if (cvHorizon === 'harian')
        return 'Report CV · Sales Daily (Dept 6201 Graphis & 6244 Simplicity)';
      if (cvHorizon === 'weekly')
        return 'Report CV · Sales Weekly (Dept 6201 Graphis & 6244 Simplicity)';
      if (cvHorizon === 'bulanan')
        return 'Report CV · Sales Monthly (Dept 6201 Graphis & 6244 Simplicity)';
      return 'Report CV · Sales Daily (Dept 6201 Graphis & 6244 Simplicity)';
    }
    if (activeMenu === 'total_dp_cv') {
      if (totalDpCvHorizon === 'harian')
        return 'Total Report (DP + CV) · Konsolidasi Sales Daily';
      if (totalDpCvHorizon === 'weekly')
        return 'Total Report (DP + CV) · Konsolidasi Sales Weekly';
      if (totalDpCvHorizon === 'bulanan')
        return 'Total Report (DP + CV) · Konsolidasi Sales Monthly (1 Tahun)';
      return 'Total Report (DP + CV) · Konsolidasi';
    }
    if (activeMenu === 'data_cv_ops')
      return 'Data CV — Inventory, Promo, Reason/Kendala & Action Plan';
    if (activeMenu === 'reason_evaluation')
      return 'Rekapitulasi & Evaluasi Harian: Reason/Event Pagi vs Siang';
    if (activeMenu === 'review_monitoring') return 'Review Monitoring — REPORT BAZAR MATAHARI';
    if (activeMenu === 'report_link') return 'Link Pengisian Report Online';
    if (activeMenu === 'daily_input') return 'Input Shift Harian';
    return 'Performa Petugas';
  };

  const renderSidebarNavigationContent = (isMobile: boolean = false) => (
    <div className="p-3.5 space-y-5">
      {/* GROUP 1: DASHBOARD HARIAN, WEEKLY & BULANAN */}
      <div>
        <div className="px-3 mb-2 text-[11px] font-bold tracking-wider text-slate-400">
          1. DASHBOARD ANALITIK
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => handleSelectDashboardHorizon('harian')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'dashboard' && dashboardHorizon === 'harian'
                ? 'bg-gradient-to-r from-amber-500 to-orange-500 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'dashboard' && dashboardHorizon === 'harian'
                  ? 'bg-white/20 text-white'
                  : 'bg-amber-500/15 text-amber-400'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Dashboard Harian</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectDashboardHorizon('weekly')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'dashboard' && dashboardHorizon === 'weekly'
                ? 'bg-gradient-to-r from-emerald-500 to-teal-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'dashboard' && dashboardHorizon === 'weekly'
                  ? 'bg-white/20 text-white'
                  : 'bg-emerald-500/15 text-emerald-400'
              }`}
            >
              <CalendarRange className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Dashboard Weekly</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectDashboardHorizon('bulanan')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'dashboard' && dashboardHorizon === 'bulanan'
                ? 'bg-gradient-to-r from-sky-500 to-blue-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'dashboard' && dashboardHorizon === 'bulanan'
                  ? 'bg-white/20 text-white'
                  : 'bg-sky-500/15 text-sky-400'
              }`}
            >
              <CalendarDays className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Dashboard Bulanan</span>
          </button>
        </nav>
      </div>

      {/* GROUP 2: REPORT DP (DIRECT PURCHASE) */}
      <div>
        <div className="px-3 mb-2 flex items-center justify-between">
          <span className="text-[11px] font-bold tracking-wider text-slate-400">
            2. REPORT DP
          </span>
          <span className="text-[9px] px-1.5 py-0.5 rounded bg-blue-500/15 text-blue-400 font-semibold tracking-wide">
            DIRECT PURCHASE
          </span>
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => handleSelectSalesTableHorizon('harian')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'sales_matrix' && salesTableHorizon === 'harian'
                ? 'bg-gradient-to-r from-amber-500 to-orange-500 text-slate-950 font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'sales_matrix' && salesTableHorizon === 'harian'
                  ? 'bg-slate-950/15 text-slate-950'
                  : 'bg-amber-500/15 text-amber-400'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Daily</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectSalesTableHorizon('weekly')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'sales_matrix' && salesTableHorizon === 'weekly'
                ? 'bg-gradient-to-r from-rose-600 to-red-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'sales_matrix' && salesTableHorizon === 'weekly'
                  ? 'bg-white/20 text-white'
                  : 'bg-rose-500/15 text-rose-400'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Weekly</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectSalesTableHorizon('bulanan')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'sales_matrix' && salesTableHorizon === 'bulanan'
                ? 'bg-gradient-to-r from-sky-500 to-blue-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'sales_matrix' && salesTableHorizon === 'bulanan'
                  ? 'bg-white/20 text-white'
                  : 'bg-sky-500/15 text-sky-400'
              }`}
            >
              <CalendarDays className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Monthly</span>
          </button>
        </nav>
      </div>

      {/* GROUP 3: REPORT CV (KONSINYASI) */}
      <div>
        <div className="px-3 mb-2 flex items-center justify-between">
          <span className="text-[11px] font-bold tracking-wider text-slate-400">
            3. REPORT CV
          </span>
          <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-500/15 text-purple-400 font-semibold tracking-wide">
            KONSINYASI
          </span>
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => handleSelectCvHorizon('harian')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'report_cv_div1' && cvHorizon === 'harian'
                ? 'bg-gradient-to-r from-purple-700 via-indigo-700 to-purple-800 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'report_cv_div1' && cvHorizon === 'harian'
                  ? 'bg-white/20 text-white'
                  : 'bg-purple-500/15 text-purple-400'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Daily</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectCvHorizon('weekly')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'report_cv_div1' && cvHorizon === 'weekly'
                ? 'bg-gradient-to-r from-purple-700 via-indigo-700 to-purple-800 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'report_cv_div1' && cvHorizon === 'weekly'
                  ? 'bg-white/20 text-white'
                  : 'bg-purple-500/15 text-purple-400'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Weekly</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectCvHorizon('bulanan')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'report_cv_div1' && cvHorizon === 'bulanan'
                ? 'bg-gradient-to-r from-purple-700 via-indigo-700 to-purple-800 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'report_cv_div1' && cvHorizon === 'bulanan'
                  ? 'bg-white/20 text-white'
                  : 'bg-purple-500/15 text-purple-400'
              }`}
            >
              <CalendarDays className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Sales Monthly</span>
          </button>
        </nav>
      </div>

      {/* GROUP 4: TOTAL REPORT (DP + CV) */}
      <div>
        <div className="px-3 mb-2 flex items-center justify-between">
          <span className="text-[11px] font-bold tracking-wider text-slate-400">
            4. TOTAL REPORT (DP + CV)
          </span>
          <span className="text-[9px] px-1.5 py-0.5 rounded bg-emerald-500/15 text-emerald-400 font-semibold tracking-wide border border-emerald-500/30">
            TOTAL GABUNGAN
          </span>
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => handleSelectTotalDpCvHorizon('harian')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'harian'
                ? 'bg-gradient-to-r from-emerald-600 via-teal-600 to-cyan-700 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'harian'
                  ? 'bg-white/20 text-white'
                  : 'bg-emerald-500/15 text-emerald-400'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Total Daily</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectTotalDpCvHorizon('weekly')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'weekly'
                ? 'bg-gradient-to-r from-emerald-600 via-teal-600 to-cyan-700 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'weekly'
                  ? 'bg-white/20 text-white'
                  : 'bg-emerald-500/15 text-emerald-400'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Total Weekly</span>
          </button>

          <button
            type="button"
            onClick={() => handleSelectTotalDpCvHorizon('bulanan')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'bulanan'
                ? 'bg-gradient-to-r from-emerald-600 via-teal-600 to-cyan-700 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'total_dp_cv' && totalDpCvHorizon === 'bulanan'
                  ? 'bg-white/20 text-white'
                  : 'bg-emerald-500/15 text-emerald-400'
              }`}
            >
              <CalendarDays className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Total Monthly (1 Tahun)</span>
          </button>
        </nav>
      </div>

      {/* GROUP 5: MONITORING & EVALUASI BAZAR */}
      <div>
        <div className="px-3 mb-2 text-[11px] font-bold tracking-wider text-slate-400">
          5. MONITORING &amp; EVALUASI BAZAR
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => handleNavigate('review_monitoring')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'review_monitoring'
                ? 'bg-gradient-to-r from-teal-500 to-emerald-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'review_monitoring'
                  ? 'bg-white/20 text-white'
                  : 'bg-teal-500/15 text-teal-400'
              }`}
            >
              <ClipboardList className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Review Monitoring Bazar</span>
          </button>

          <button
            type="button"
            onClick={() => handleNavigate('reason_evaluation')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'reason_evaluation'
                ? 'bg-gradient-to-r from-orange-500 to-amber-500 text-slate-950 font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'reason_evaluation'
                  ? 'bg-slate-950/15 text-slate-950'
                  : 'bg-orange-500/15 text-orange-400'
              }`}
            >
              <MessageSquareText className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Reason/Event Pagi vs Siang</span>
          </button>

          <button
            type="button"
            onClick={() => handleNavigate('annual_schedule')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'annual_schedule'
                ? 'bg-gradient-to-r from-amber-500 to-yellow-500 text-slate-950 font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'annual_schedule'
                  ? 'bg-slate-950/15 text-slate-950'
                  : 'bg-amber-500/15 text-amber-400'
              }`}
            >
              <CalendarRange className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Data 52 Weekly Setahun</span>
          </button>

          <button
            type="button"
            onClick={() => handleNavigate('data_cv_ops')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'data_cv_ops'
                ? 'bg-gradient-to-r from-indigo-600 to-sky-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'data_cv_ops'
                  ? 'bg-white/20 text-white'
                  : 'bg-sky-500/15 text-sky-400'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Data CV</span>
          </button>
        </nav>
      </div>

      {/* GROUP 6: LINK APLIKASI & PELAPORAN */}
      <div>
        <div className="px-3 mb-2 text-[11px] font-bold tracking-wider text-slate-400">
          6. LINK APLIKASI &amp; PELAPORAN
        </div>
        <nav className="space-y-1">
          <button
            type="button"
            onClick={() => {
              setIsAppShareModalOpen(true);
              if (isMobile) setMobileMenuOpen(false);
            }}
            className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-bold transition-all text-left bg-gradient-to-r from-amber-400/20 to-indigo-500/20 border border-amber-400/40 text-amber-300 hover:bg-amber-400/30"
          >
            <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-amber-400 text-slate-950">
              <Share2 className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Link Aplikasi (Publik / Admin)</span>
          </button>

          <button
            type="button"
            onClick={() => handleNavigate('report_link')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
              activeMenu === 'report_link'
                ? 'bg-gradient-to-r from-indigo-500 to-blue-600 text-white font-bold shadow-xs'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <div
              className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                activeMenu === 'report_link'
                  ? 'bg-white/20 text-white'
                  : 'bg-indigo-500/15 text-indigo-400'
              }`}
            >
              <Link2 className="w-3.5 h-3.5" />
            </div>
            <span className="truncate">Link Pengisian Report</span>
          </button>

          {accessRole === 'admin' && (
            <button
              type="button"
              onClick={() => handleNavigate('daily_input')}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-medium transition-all text-left ${
                activeMenu === 'daily_input'
                  ? 'bg-gradient-to-r from-cyan-500 to-sky-600 text-white font-bold shadow-xs'
                  : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
              }`}
            >
              <div
                className={`w-6 h-6 rounded-lg flex items-center justify-center shrink-0 ${
                  activeMenu === 'daily_input'
                    ? 'bg-white/20 text-white'
                    : 'bg-cyan-500/15 text-cyan-400'
                }`}
              >
                <PenSquare className="w-3.5 h-3.5" />
              </div>
              <span className="truncate">Input Shift Harian</span>
            </button>
          )}
        </nav>
      </div>

      {/* GROUP 7: MENU LOGIN & LOGOUT */}
      <div>
        <div className="px-3 mb-2 text-[11px] font-bold tracking-wider text-slate-400">
          7. MENU LOGIN &amp; LOGOUT
        </div>
        <nav className="space-y-1">
          {accessRole === 'public' ? (
            <>
              <button
                type="button"
                onClick={() => {
                  setIsLoginLogoutModalOpen(true);
                  if (isMobile) setMobileMenuOpen(false);
                }}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-xs font-bold transition-all text-left bg-gradient-to-r from-indigo-600 to-blue-600 text-white hover:from-indigo-500 hover:to-blue-500 shadow-xs cursor-pointer"
              >
                <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-white/20 text-amber-300">
                  <LogIn className="w-3.5 h-3.5" />
                </div>
                <span className="truncate">Login Admin (Buka Edit)</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  if (isMobile) setMobileMenuOpen(false);
                  handleLogoutToEntryGate();
                }}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-xl text-xs font-semibold transition-all text-left text-rose-300 hover:bg-rose-950/60 hover:text-rose-200 border border-rose-800/40 cursor-pointer"
              >
                <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-rose-500/15 text-rose-400">
                  <LogOut className="w-3.5 h-3.5" />
                </div>
                <span className="truncate">Logout (Ke Layar Masuk)</span>
              </button>
            </>
          ) : (
            <>
              <button
                type="button"
                onClick={() => {
                  setIsLoginLogoutModalOpen(true);
                  if (isMobile) setMobileMenuOpen(false);
                }}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-xl text-xs font-bold transition-all text-left bg-emerald-950/70 border border-emerald-700/70 text-emerald-300 hover:bg-emerald-900/70 cursor-pointer"
              >
                <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-emerald-500/20 text-emerald-400">
                  <ShieldCheck className="w-3.5 h-3.5" />
                </div>
                <span className="truncate">Status: Login Admin Aktif</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  if (isMobile) setMobileMenuOpen(false);
                  handleEnterPublicMode();
                }}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-xl text-xs font-bold transition-all text-left bg-amber-500/15 border border-amber-500/40 text-amber-300 hover:bg-amber-500/25 cursor-pointer"
              >
                <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-amber-400 text-slate-950">
                  <Lock className="w-3.5 h-3.5" />
                </div>
                <span className="truncate">Logout ke Mode Publik</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  if (isMobile) setMobileMenuOpen(false);
                  handleLogoutToEntryGate();
                }}
                className="w-full flex items-center gap-3 px-3 py-2 rounded-xl text-xs font-bold transition-all text-left bg-rose-600/20 border border-rose-500/40 text-rose-300 hover:bg-rose-600/30 cursor-pointer"
              >
                <div className="w-6 h-6 rounded-lg flex items-center justify-center shrink-0 bg-rose-600 text-white">
                  <LogOut className="w-3.5 h-3.5" />
                </div>
                <span className="truncate">Logout (Keluar Aplikasi)</span>
              </button>
            </>
          )}
        </nav>
      </div>
    </div>
  );

  // Show Entry Access Gate when visitor first opens the application link (unless ?mode=public or locked report form)
  if (showEntryGate && !isLockedReportOnly) {
    return (
      <AppEntryAccessGate
        onEnterPublic={handleEnterPublicMode}
        onEnterAdmin={handleLoginAdminWithPassword}
        onOpenReportForm={(formType) => {
          setDirectReportFormMode(formType);
          setIsLockedReportOnly(true);
          setActiveMenu('report_link');
          setShowEntryGate(false);
          if (pullFromServerRef.current) {
            pullFromServerRef.current(false, true);
          }
        }}
        initialFocusAdmin={urlModeParam !== 'public'}
      />
    );
  }

  return (
    <div className="h-[100dvh] w-full max-w-full overflow-hidden flex bg-[#F4F7FB] text-slate-900">
      {/* Desktop Colorful Dark Slate Sidebar Navigation (Hidden completely when in Locked Report Link Mode) */}
      {!isLockedReportOnly && (
      <aside className="no-print hidden lg:flex lg:flex-col lg:w-[272px] lg:shrink-0 bg-slate-900 text-white border-r border-slate-800 h-full justify-between overflow-y-auto">
        <div>
          {/* Brand Identity Header */}
          <div className="px-4 py-3.5 border-b border-slate-800 bg-gradient-to-b from-slate-900 to-slate-950 flex items-start gap-3">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-rose-600 to-amber-500 text-white flex items-center justify-center shrink-0 shadow-xs mt-0.5">
              <Store className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <div className="text-[13px] font-extrabold tracking-tight text-white leading-tight">
                REPORT BAZAR MATAHARI
              </div>
              <div className="text-[11px] text-amber-400 font-bold truncate mt-0.5">
                Tanjung
              </div>
              <div className="text-[10px] text-slate-400 font-semibold tracking-wide uppercase truncate mt-0.5">
                PT MDS Retailing Tbk
              </div>
            </div>
          </div>

          {renderSidebarNavigationContent(false)}
        </div>

        {/* Bottom Sidebar Save, Login/Logout, Access Mode & Reset */}
        <div className="p-4 border-t border-slate-800 bg-slate-950/50 space-y-2">
          <div className="grid grid-cols-2 gap-1.5">
            {accessRole === 'public' ? (
              <button
                type="button"
                onClick={() => setIsLoginLogoutModalOpen(true)}
                className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 rounded-lg text-xs font-extrabold bg-amber-400 hover:bg-amber-300 text-slate-950 transition-colors cursor-pointer"
              >
                <LogIn className="w-3.5 h-3.5" />
                <span>Login Admin</span>
              </button>
            ) : (
              <button
                type="button"
                onClick={() => setIsLoginLogoutModalOpen(true)}
                className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 rounded-lg text-xs font-extrabold bg-emerald-950/90 border border-emerald-700 text-emerald-300 hover:bg-emerald-900/80 transition-colors cursor-pointer"
              >
                <ShieldCheck className="w-3.5 h-3.5" />
                <span>Akun Admin</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => {
                if (accessRole === 'admin') {
                  handleEnterPublicMode();
                } else {
                  handleLogoutToEntryGate();
                }
              }}
              className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 rounded-lg text-xs font-extrabold bg-rose-600 hover:bg-rose-500 text-white transition-colors cursor-pointer"
              title={
                accessRole === 'admin'
                  ? 'Logout dari Mode Admin ke Mode Publik'
                  : 'Logout & Kembali ke Layar Masuk'
              }
            >
              <LogOut className="w-3.5 h-3.5" />
              <span>Logout</span>
            </button>
          </div>

          <button
            type="button"
            onClick={() => setIsAppShareModalOpen(true)}
            className="w-full inline-flex items-center justify-between px-3 py-2 rounded-lg text-xs font-bold border border-indigo-500/40 bg-indigo-950/60 text-indigo-200 hover:bg-indigo-900/60 transition-colors cursor-pointer"
          >
            <span className="inline-flex items-center gap-1.5">
              <Share2 className="w-3.5 h-3.5 text-amber-400" />
              <span>Link Aplikasi &amp; Report</span>
            </span>
            <span className="text-[10px] underline">Salin Link</span>
          </button>

          {accessRole === 'admin' && (
            <>
              <div className="grid grid-cols-2 gap-1.5">
                <button
                  type="button"
                  onClick={() => handleManualSaveNow()}
                  className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 rounded-lg transition-colors shadow-2xs"
                  title="Simpan Cepat ke Server & Browser (Ctrl+S)"
                >
                  <Save className="w-3.5 h-3.5" />
                  Simpan
                </button>
                <button
                  type="button"
                  onClick={() => setIsSaveToolsOpen(true)}
                  className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-semibold text-amber-300 bg-slate-800 hover:bg-slate-700 border border-amber-500/30 rounded-lg transition-colors"
                  title="Buka Tools Save, Save Slot & Backup (.json)"
                >
                  <Save className="w-3.5 h-3.5" />
                  Tools Save
                </button>
              </div>
              {confirmReset ? (
                <div className="p-2.5 bg-amber-950/90 border border-amber-700 rounded-lg space-y-2 text-xs">
                  <div className="font-semibold text-amber-200">
                    Pilih Mode Reset Data:
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <button
                      type="button"
                      onClick={() => handleResetData('empty')}
                      className="w-full py-1.5 px-2 bg-rose-600 hover:bg-rose-500 text-white font-bold rounded text-left flex items-center justify-between"
                    >
                      <span>Kosongkan Semua Sales (0)</span>
                      <span className="text-[10px] opacity-90">Nolkan</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => handleResetData('initial')}
                      className="w-full py-1.5 px-2 bg-amber-400 hover:bg-amber-500 text-slate-950 font-bold rounded text-left flex items-center justify-between"
                    >
                      <span>Kembalikan Data Contoh Awal</span>
                      <span className="text-[10px] opacity-80">W41</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setConfirmReset(false)}
                      className="w-full py-1 bg-slate-800 border border-slate-700 text-slate-200 rounded text-center"
                    >
                      Batal
                    </button>
                  </div>
                </div>
              ) : (
                <button
                  type="button"
                  onClick={() => setConfirmReset(true)}
                  className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-300 hover:text-white bg-slate-800/70 hover:bg-slate-800 border border-slate-700/80 rounded-lg transition-colors"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  Reset Data (Nolkan / Awal)
                </button>
              )}
            </>
          )}
        </div>
      </aside>
      )}

      {/* Mobile Sidebar Drawer (Hidden completely when in Locked Report Link Mode) */}
      {!isLockedReportOnly && mobileMenuOpen && (
        <div className="no-print fixed inset-0 z-[100] lg:hidden flex">
          <div
            className="fixed inset-0 bg-slate-950/60 backdrop-blur-xs"
            onClick={() => setMobileMenuOpen(false)}
          />
          <div className="relative w-72 max-w-[85vw] bg-slate-900 text-white h-full flex flex-col justify-between z-10 shadow-2xl overflow-y-auto">
            <div>
              <div className="px-4 py-3 border-b border-slate-800 flex items-center justify-between sticky top-0 bg-slate-900 z-20">
                <div className="flex items-center gap-2.5 min-w-0">
                  <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-rose-600 to-amber-500 text-white flex items-center justify-center shrink-0">
                    <Store className="w-4 h-4" />
                  </div>
                  <div className="min-w-0">
                    <div className="text-xs font-extrabold text-white truncate">
                      REPORT BAZAR MATAHARI
                    </div>
                    <div className="text-[11px] font-bold text-amber-400 truncate">
                      Tanjung
                    </div>
                    <div className="text-[10px] text-slate-400 truncate">
                      PT MDS Retailing Tbk
                    </div>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => setMobileMenuOpen(false)}
                  className="p-2 text-slate-300 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
                  aria-label="Tutup Menu"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {renderSidebarNavigationContent(true)}
            </div>

            <div className="p-4 border-t border-slate-800 bg-slate-900 space-y-2">
              <div className="grid grid-cols-2 gap-1.5">
                <button
                  type="button"
                  onClick={() => {
                    setMobileMenuOpen(false);
                    setIsLoginLogoutModalOpen(true);
                  }}
                  className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-extrabold text-slate-950 bg-amber-400 hover:bg-amber-300 rounded-lg"
                >
                  <LogIn className="w-3.5 h-3.5" />
                  <span>{accessRole === 'admin' ? 'Akun Admin' : 'Login Admin'}</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setMobileMenuOpen(false);
                    if (accessRole === 'admin') {
                      handleEnterPublicMode();
                    } else {
                      handleLogoutToEntryGate();
                    }
                  }}
                  className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-extrabold text-white bg-rose-600 hover:bg-rose-500 rounded-lg"
                >
                  <LogOut className="w-3.5 h-3.5" />
                  <span>Logout</span>
                </button>
              </div>
              <button
                type="button"
                onClick={() => {
                  setMobileMenuOpen(false);
                  setIsAppShareModalOpen(true);
                }}
                className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-extrabold text-white bg-indigo-600 hover:bg-indigo-500 rounded-lg"
              >
                <Share2 className="w-3.5 h-3.5" />
                <span>Link Aplikasi &amp; Report</span>
              </button>
              {accessRole === 'admin' && (
                <>
                  <div className="grid grid-cols-2 gap-1.5">
                    <button
                      type="button"
                      onClick={() => {
                        handleManualSaveNow();
                        setMobileMenuOpen(false);
                      }}
                      className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-500 rounded-lg"
                    >
                      <Save className="w-3.5 h-3.5" />
                      Simpan
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setMobileMenuOpen(false);
                        setIsSaveToolsOpen(true);
                      }}
                      className="inline-flex items-center justify-center gap-1.5 px-2.5 py-2 text-xs font-bold text-amber-300 bg-slate-800 hover:bg-slate-700 border border-amber-500/30 rounded-lg"
                    >
                      <Save className="w-3.5 h-3.5" />
                      Tools Save
                    </button>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      handleResetData('empty');
                      setMobileMenuOpen(false);
                    }}
                    className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-500 rounded-lg"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    Reset Kosongkan Sales (0)
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      handleResetData('initial');
                      setMobileMenuOpen(false);
                    }}
                    className="w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-300 bg-slate-800 hover:bg-slate-700 rounded-lg"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    Kembalikan Data Contoh Awal
                  </button>
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Right Main Workspace Column — Frozen Header + Scrollable Content */}
      <div className="flex-1 flex flex-col min-w-0 max-w-full h-[100dvh] overflow-hidden">
        {/* Top Header Bar — Frozen (Dibekukan) & Compact Layout */}
        <header className="no-print shrink-0 sticky top-0 z-40 bg-white border-b border-slate-200 shadow-2xs w-full max-w-full">
          {isLockedReportOnly ? (
            <div className="px-3 sm:px-6 py-2.5 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2.5 sm:gap-4 bg-slate-900 text-white border-b border-slate-800">
              {/* Brand Title & Lock Icon */}
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="w-8 h-8 rounded-lg bg-amber-400 text-slate-950 flex items-center justify-center shrink-0 shadow-xs">
                  <Lock className="w-4 h-4" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex flex-wrap items-center gap-1.5 text-xs font-extrabold leading-tight">
                    <span className="text-white">REPORT BAZAR MATAHARI</span>
                    <span className="text-amber-400 font-bold">· Tanjung</span>
                    <span className="text-slate-400 font-medium hidden md:inline">· PT MDS Retailing Tbk</span>
                  </div>
                  <div className="text-[11px] font-semibold text-amber-200/90 truncate mt-0.5">
                    Menu Pengisian Report Online
                  </div>
                </div>
              </div>

              {/* Action Buttons: Clean & Responsive */}
              <div className="flex items-center justify-end gap-1.5 shrink-0 pt-1.5 sm:pt-0 border-t border-slate-800/80 sm:border-t-0">
                <button
                  type="button"
                  onClick={() => {
                    setIsLockedReportOnly(false);
                    handleEnterPublicMode();
                    setActiveMenu('dashboard');
                  }}
                  className="px-2.5 sm:px-3 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-500 text-[11px] sm:text-xs font-extrabold text-white whitespace-nowrap cursor-pointer transition-colors shadow-2xs"
                >
                  Buka Aplikasi<span className="hidden sm:inline"> (Publik)</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setIsLockedReportOnly(false);
                    setIsLoginLogoutModalOpen(true);
                  }}
                  className="px-2.5 sm:px-3 py-1.5 rounded-lg bg-amber-400 hover:bg-amber-300 text-[11px] sm:text-xs font-extrabold text-slate-950 inline-flex items-center gap-1 whitespace-nowrap cursor-pointer transition-colors shadow-2xs"
                >
                  <LogIn className="w-3.5 h-3.5" />
                  <span>Login Admin</span>
                </button>
                {!isOpenedFromLockedLink && (
                  <button
                    type="button"
                    onClick={() => setIsLockedReportOnly(false)}
                    className="px-2 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-[10px] sm:text-xs font-semibold text-amber-300 border border-amber-400/30 whitespace-nowrap transition-colors"
                  >
                    Buka Kunci
                  </button>
                )}
              </div>
            </div>
          ) : (
            <>
          <div className="px-3 sm:px-5 py-1.5 sm:py-2 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-1.5 w-full max-w-full min-w-0">
            {/* Top Row on Mobile / Left Side on Desktop: Menu Trigger + Official Title, Subtitle & Company */}
            <div className="flex items-center gap-2 min-w-0 flex-1 overflow-hidden">
              <button
                type="button"
                onClick={() => setMobileMenuOpen(true)}
                className="lg:hidden inline-flex items-center justify-center w-8 h-8 text-white bg-slate-900 hover:bg-slate-800 rounded-lg shrink-0 shadow-2xs"
                aria-label="Buka Menu Navigasi"
              >
                <Menu className="w-3.5 h-3.5" />
              </button>

              <div className="min-w-0 flex-1 overflow-hidden">
                <div className="font-extrabold text-slate-950 tracking-tight text-xs sm:text-sm leading-tight truncate">
                  REPORT BAZAR MATAHARI
                </div>
                <div className="flex flex-wrap items-center gap-1 mt-0.5 min-w-0">
                  <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-950 border border-amber-300 leading-none">
                    Tanjung
                  </span>
                  <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-rose-50 text-rose-800 border border-rose-200 leading-none">
                    PT MDS Retailing Tbk
                  </span>
                  {accessRole === 'admin' && (
                    <AutoDraftStatusBadge
                      menuLabel={getHeaderBreadcrumbTitle()}
                      lastDraftTimeStr={globalDraftTimeStr}
                      isDirty={hasUnsavedGlobalDraft}
                      isOnline={isGlobalOnline}
                      onManualSave={() => handleManualSaveNow()}
                    />
                  )}
                  <span className="hidden xl:inline-block text-[11px] font-semibold text-slate-500 truncate ml-1">
                    · {getHeaderBreadcrumbTitle()}
                  </span>
                </div>
              </div>
            </div>

            {/* Action Toolbar: Compact 2-Tier Symmetrical Layout (Top: Simpan & Export, Bottom: Link & Undo/Redo) */}
            <div className="flex flex-col gap-1 pt-1.5 lg:pt-0 border-t border-slate-100 lg:border-t-0 shrink-0 w-full lg:w-auto lg:min-w-[300px]">
              {/* SUSUN 1 (ATAS): SIMPAN & EXPORT */}
              <div className="grid grid-cols-2 gap-1 w-full">
                {/* Kiri Atas: Simpan | Slot */}
                <div
                  className={`flex items-center h-[26px] bg-emerald-50 p-0.5 rounded-md border border-emerald-200 w-full ${
                    accessRole !== 'admin' ? 'opacity-50' : ''
                  }`}
                >
                  <button
                    type="button"
                    disabled={accessRole !== 'admin'}
                    onClick={() => handleManualSaveNow()}
                    title={
                      accessRole !== 'admin'
                        ? 'Terkunci di Mode Publik (Hanya Lihat)'
                        : `Simpan Data Sekarang (Ctrl+S) — Terakhir: ${lastSavedTimeStr}`
                    }
                    className="flex-1 inline-flex items-center justify-center gap-1 h-full px-1.5 text-[11px] font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:hover:bg-emerald-600 disabled:cursor-not-allowed rounded transition-all whitespace-nowrap cursor-pointer leading-none"
                  >
                    {isSavingState ? (
                      <CheckCircle2 className="w-3 h-3 shrink-0 animate-pulse" />
                    ) : (
                      <Save className="w-3 h-3 shrink-0" />
                    )}
                    <span>{isSavingState ? 'Simpan...' : 'Simpan'}</span>
                  </button>
                  <button
                    type="button"
                    disabled={accessRole !== 'admin'}
                    onClick={() => {
                      if (!requireAdminForEdit()) return;
                      setIsSaveToolsOpen(true);
                    }}
                    title={
                      accessRole !== 'admin'
                        ? 'Terkunci di Mode Publik (Hanya Lihat)'
                        : 'Buka Tools Save, Save Slot (Checkpoint), & Backup/Restore File (.json)'
                    }
                    className="inline-flex items-center justify-center h-full px-2 text-[11px] font-bold text-emerald-900 hover:bg-white disabled:hover:bg-transparent disabled:cursor-not-allowed rounded transition-all whitespace-nowrap shrink-0 cursor-pointer leading-none"
                  >
                    <span>Slot</span>
                  </button>
                </div>

                {/* Kanan Atas: Export (& Cetak on Desktop) */}
                <div
                  className={`flex items-center gap-1 w-full ${
                    accessRole !== 'admin' ? 'opacity-50' : ''
                  }`}
                >
                  <button
                    type="button"
                    disabled={accessRole !== 'admin'}
                    onClick={handlePrint}
                    title={
                      accessRole !== 'admin'
                        ? 'Mode Publik (Hanya Lihat): Tidak bisa klik Cetak'
                        : 'Cetak Halaman'
                    }
                    className="hidden xl:inline-flex items-center justify-center gap-1 h-[26px] px-2 text-[11px] font-semibold text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 disabled:hover:bg-white disabled:cursor-not-allowed rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer leading-none"
                  >
                    <Printer className="w-3 h-3" />
                    <span>Cetak</span>
                  </button>
                  <button
                    type="button"
                    disabled={accessRole !== 'admin'}
                    onClick={handleExportExcel}
                    title={
                      accessRole !== 'admin'
                        ? 'Mode Publik (Hanya Lihat): Tidak bisa klik Export'
                        : 'Export File Excel (.xlsx)'
                    }
                    className="flex-1 inline-flex items-center justify-center gap-1 h-[26px] px-2 text-[11px] font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-400 disabled:hover:bg-slate-400 disabled:cursor-not-allowed rounded-md transition-colors whitespace-nowrap shadow-2xs w-full cursor-pointer leading-none"
                  >
                    {accessRole !== 'admin' ? (
                      <Lock className="w-3 h-3 shrink-0" />
                    ) : (
                      <Download className="w-3 h-3 shrink-0" />
                    )}
                    <span>Export</span>
                  </button>
                </div>
              </div>

              {/* SUSUN 2 (BAWAH): LINK & UNDO REDO */}
              <div className="grid grid-cols-2 gap-1 w-full">
                {/* Kiri Bawah: Link & Admin */}
                <div className="flex items-center h-[26px] bg-slate-100 p-0.5 rounded-md border border-slate-200/90 w-full">
                  <button
                    type="button"
                    onClick={() => setIsAppShareModalOpen(true)}
                    className="flex-1 inline-flex items-center justify-center gap-1 h-full px-1.5 text-[11px] font-bold text-white bg-indigo-600 hover:bg-indigo-700 rounded transition-colors cursor-pointer whitespace-nowrap leading-none"
                    title="Buat & Salin Link Aplikasi & Link Report"
                  >
                    <Share2 className="w-3 h-3 shrink-0" />
                    <span>Link</span>
                  </button>

                  {accessRole === 'public' ? (
                    <button
                      type="button"
                      onClick={() => setIsLoginLogoutModalOpen(true)}
                      className="inline-flex items-center justify-center gap-1 h-full px-1.5 ml-0.5 text-[11px] font-extrabold text-slate-950 bg-amber-400 hover:bg-amber-300 rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 leading-none"
                      title="Login Admin (Masukkan Sandi Admin untuk Membuka Akses Edit Data)"
                    >
                      <KeyRound className="w-2.5 h-2.5 shrink-0" />
                      <span>Admin</span>
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => setIsLoginLogoutModalOpen(true)}
                      className="inline-flex items-center justify-center gap-1 h-full px-1.5 ml-0.5 text-[11px] font-extrabold text-emerald-900 bg-emerald-100 hover:bg-emerald-200 rounded transition-colors cursor-pointer whitespace-nowrap shrink-0 leading-none"
                      title="Status Login Admin Aktif (Bisa Edit)"
                    >
                      <ShieldCheck className="w-3 h-3 shrink-0 text-emerald-700" />
                      <span>Admin</span>
                    </button>
                  )}
                </div>

                {/* Kanan Bawah: Undo & Redo */}
                <div
                  className={`grid grid-cols-2 items-center h-[26px] bg-slate-100 p-0.5 rounded-md border border-slate-200/90 w-full ${
                    accessRole !== 'admin' ? 'opacity-60' : ''
                  }`}
                >
                  <button
                    type="button"
                    onClick={handleUndo}
                    disabled={undoStack.length === 0 || accessRole !== 'admin'}
                    title={`Undo / Batalkan Perubahan (Ctrl+Z) — ${undoStack.length} langkah tersedia`}
                    className="inline-flex items-center justify-center gap-1 h-full px-1.5 text-[11px] font-semibold text-slate-700 hover:bg-white hover:text-slate-900 disabled:opacity-40 disabled:hover:bg-transparent rounded transition-all whitespace-nowrap cursor-pointer leading-none"
                  >
                    <Undo2 className="w-3 h-3 shrink-0" />
                    <span>Undo</span>
                    {undoStack.length > 0 && (
                      <span className="px-1 py-0.2 bg-indigo-600 text-white rounded text-[9px] font-mono-num leading-none">
                        {undoStack.length}
                      </span>
                    )}
                  </button>
                  <button
                    type="button"
                    onClick={handleRedo}
                    disabled={redoStack.length === 0 || accessRole !== 'admin'}
                    title={`Redo / Ulangi Perubahan (Ctrl+Y) — ${redoStack.length} langkah tersedia`}
                    className="inline-flex items-center justify-center gap-1 h-full px-1.5 text-[11px] font-semibold text-slate-700 hover:bg-white hover:text-slate-900 disabled:opacity-40 disabled:hover:bg-transparent rounded transition-all whitespace-nowrap border-l border-slate-200/80 cursor-pointer leading-none"
                  >
                    <Redo2 className="w-3 h-3 shrink-0" />
                    <span>Redo</span>
                    {redoStack.length > 0 && (
                      <span className="px-1 py-0.2 bg-emerald-600 text-white rounded text-[9px] font-mono-num leading-none">
                        {redoStack.length}
                      </span>
                    )}
                  </button>
                </div>
              </div>
            </div>
          </div>

          {/* Read-Only Public Mode Banner */}
          {accessRole === 'public' && (
            <div className="bg-gradient-to-r from-sky-950 via-indigo-950 to-slate-900 text-white px-3 sm:px-5 py-1 border-t border-sky-800/60 flex items-center justify-between gap-2 text-xs">
              <div className="flex items-center gap-1.5 min-w-0">
                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-sky-500/20 border border-sky-400/40 text-sky-200 font-extrabold text-[10px] shrink-0">
                  <Eye className="w-3 h-3 text-sky-300" />
                  <span>PUBLIK (HANYA LIHAT)</span>
                </span>
                <span className="hidden sm:inline text-slate-200 text-[11px] truncate">
                  Mode Hanya Lihat — Edit data dinonaktifkan (Screenshot &amp; Export tetap dapat digunakan).
                </span>
              </div>
              <div className="flex items-center gap-1.5 shrink-0">
                <button
                  type="button"
                  onClick={() => setIsLoginLogoutModalOpen(true)}
                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-[10px] transition-colors cursor-pointer whitespace-nowrap"
                >
                  <KeyRound className="w-2.5 h-2.5" />
                  <span>Sandi Admin</span>
                </button>
              </div>
            </div>
          )}
          </>
          )}
        </header>

        {/* Scrollable Viewport Area Underneath Frozen Header */}
        <div className="flex-1 overflow-y-auto overflow-x-hidden flex flex-col min-w-0">
          {/* Main Content Viewport — Protected in Public Read-Only Mode */}
          <main
            className="flex-1 w-full max-w-[1540px] mx-auto px-3 sm:px-6 py-4 min-w-0"
          onKeyDownCapture={(e) => {
            if (accessRole === 'public' && !isLockedReportOnly && activeMenu !== 'report_link') {
              const target = e.target as HTMLElement;
              if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) {
                if (
                  e.key.length === 1 ||
                  e.key === 'Backspace' ||
                  e.key === 'Delete' ||
                  e.key === 'Enter'
                ) {
                  e.preventDefault();
                  e.stopPropagation();
                  showBriefHistoryNotice(
                    '🔒 Mode Publik (Hanya Lihat): Masukkan Sandi Admin untuk mengedit data'
                  );
                }
              }
            }
          }}
          onPasteCapture={(e) => {
            if (accessRole === 'public' && !isLockedReportOnly && activeMenu !== 'report_link') {
              const target = e.target as HTMLElement;
              if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA')) {
                e.preventDefault();
                e.stopPropagation();
                showBriefHistoryNotice(
                  '🔒 Mode Publik (Hanya Lihat): Masukkan Sandi Admin untuk mengedit data'
                );
              }
            }
          }}
        >
          {!isLockedReportOnly &&
            accessRole === 'admin' &&
            (activeMenu === 'dashboard' ||
              activeMenu === 'annual_schedule' ||
              activeMenu === 'staff_analytics') && (
              <div className="mb-3.5">
                <AutoDraftStatusBadge
                  variant="banner"
                  menuLabel={getHeaderBreadcrumbTitle()}
                  lastDraftTimeStr={globalDraftTimeStr}
                  isDirty={hasUnsavedGlobalDraft}
                  isOnline={isGlobalOnline}
                  onManualSave={() => handleManualSaveNow()}
                />
              </div>
            )}

          <div
            key={
              isLockedReportOnly
                ? 'locked_report_link'
                : `${activeMenu}_${activeMenu === 'dashboard' ? dashboardHorizon : ''}_${activeMenu === 'sales_matrix' ? salesTableHorizon : ''}_${activeMenu === 'report_cv_div1' ? cvHorizon : ''}_${activeMenu === 'total_dp_cv' ? totalDpCvHorizon : ''}`
            }
            className="view-transition-fade w-full"
          >
            {!isLockedReportOnly && activeMenu === 'dashboard' && (
              <DashboardOverview
                tables={weeklyTables}
                selectedWeeklyFilter={selectedWeeklyFilter}
                onSelectWeeklyFilter={setSelectedWeeklyFilter}
                reviewDoc={currentReviewDoc}
                reviewDocs={reviewDocs}
                horizon={dashboardHorizon}
                onChangeHorizon={setDashboardHorizon}
                onNavigate={handleNavigate}
                onChangeTables={handleUpdateTables}
                isPublicReadOnly={accessRole !== 'admin'}
                cvRecords={cvDiv1Records}
                cvInventoryOps={cvInventoryOps}
                onNavigateToCv={handleSelectCvHorizon}
                onNavigateToTotalDpCv={handleSelectTotalDpCvHorizon}
                onOpenExportModal={
                  accessRole === 'admin' ? () => setIsExportModalOpen(true) : undefined
                }
              />
            )}

            {activeMenu === 'sales_matrix' && (
              <WeeklySalesSpreadsheet
                tables={weeklyTables}
                onChangeTables={handleUpdateTables}
                showDivZero={showDivZero}
                onToggleDivZero={() => setShowDivZero(!showDivZero)}
                staffFilter={staffFilter}
                onChangeStaffFilter={setStaffFilter}
                selectedWeeklyFilter={selectedWeeklyFilter}
                onSelectWeeklyFilter={setSelectedWeeklyFilter}
                tableHorizon={salesTableHorizon}
                onChangeTableHorizon={setSalesTableHorizon}
                canUndo={undoStack.length > 0 && accessRole === 'admin'}
                canRedo={redoStack.length > 0 && accessRole === 'admin'}
                undoCount={undoStack.length}
                redoCount={redoStack.length}
                onUndo={accessRole === 'admin' ? handleUndo : undefined}
                onRedo={accessRole === 'admin' ? handleRedo : undefined}
                onManualSave={accessRole === 'admin' ? () => handleManualSaveNow() : undefined}
                onOpenSaveTools={accessRole === 'admin' ? () => setIsSaveToolsOpen(true) : undefined}
                onOpenReasonEvaluation={(weekId) => {
                  setSelectedWeeklyFilter(weekId);
                  handleNavigate('reason_evaluation');
                }}
              />
            )}

            {activeMenu === 'reason_evaluation' && (
              <ReasonEventEvaluationView
                tables={weeklyTables}
                selectedWeeklyFilter={selectedWeeklyFilter}
                onSelectWeeklyFilter={setSelectedWeeklyFilter}
                onChangeTables={handleUpdateTables}
                onOpenSpreadsheet={(weekId) => {
                  setSelectedWeeklyFilter(weekId);
                  handleNavigate('sales_matrix');
                }}
                canUndo={undoStack.length > 0 && accessRole === 'admin'}
                canRedo={redoStack.length > 0 && accessRole === 'admin'}
                undoCount={undoStack.length}
                redoCount={redoStack.length}
                onUndo={accessRole === 'admin' ? handleUndo : undefined}
                onRedo={accessRole === 'admin' ? handleRedo : undefined}
                onManualSave={accessRole === 'admin' ? () => handleManualSaveNow() : undefined}
                isPublicReadOnly={accessRole !== 'admin'}
              />
            )}

            {activeMenu === 'annual_schedule' && (
              <AnnualWeeklyScheduleView
                tables={weeklyTables}
                onChangeTables={handleUpdateTables}
                onOpenWeekly={handleOpenWeekly}
                onSelectWeeklyAndOpenSpreadsheet={(weekId) => {
                  setSelectedWeeklyFilter(weekId);
                  handleNavigate('sales_matrix');
                }}
                onSelectWeeklyAndOpenDailyInput={(weekIdx) => {
                  handleSelectWeekIndex(weekIdx);
                  handleNavigate('daily_input');
                }}
                isPublicReadOnly={accessRole !== 'admin'}
              />
            )}

            {(isLockedReportOnly || activeMenu === 'report_link') && (
              <ReportLinkCenter
                tables={weeklyTables}
                onChangeTables={(nextTables) => {
                  setWeeklyTables(nextTables);
                  syncStateToBackend(nextTables, reviewDocs, submissionLogs);
                }}
                reviewDocs={reviewDocs}
                onChangeReviewDocs={(nextReviews) => {
                  setReviewDocs(nextReviews);
                  syncStateToBackend(weeklyTables, nextReviews, submissionLogs);
                }}
                submissionLogs={submissionLogs}
                onAddSubmissionLog={handleAddSubmissionLog}
                onAtomicSubmitReport={handleAtomicReportSubmit}
                cvRecords={cvDiv1Records}
                cvInventoryOps={cvInventoryOps}
                onAtomicSubmitCvReport={handleAtomicCvReportSubmit}
                initialFormMode={
                  directReportFormMode === 'shift' || urlFormParam === 'shift'
                    ? 'shift'
                    : directReportFormMode === 'review' || urlFormParam === 'review'
                      ? 'review'
                      : directReportFormMode === 'cv_report' || urlFormParam === 'cv_report'
                        ? 'cv_report'
                        : directReportFormMode === 'data_cv' || urlFormParam === 'data_cv'
                          ? 'data_cv'
                          : 'none'
                }
                initialWeekId={urlWeekParam}
                initialDayIdx={urlDayParam}
                initialDateIso={urlDateParam}
                initialShift={urlShiftParam}
                initialStaff={urlStaffParam}
                initialPeriodType={urlPeriodTypeParam}
                initialMonthKey={urlMonthParam}
                initialBrand={urlBrandParam}
                isLockedReportOnly={isLockedReportOnly}
                onToggleLockReportOnly={(locked) => {
                  setIsLockedReportOnly(locked);
                  if (locked) setActiveMenu('report_link');
                }}
              />
            )}

            {activeMenu === 'daily_input' && (
              <DailyInputWorkspace
                tables={weeklyTables}
                activeWeekIndex={activeWeekIndex}
                onSelectWeekIndex={handleSelectWeekIndex}
                onChangeTables={handleUpdateTables}
                canUndo={undoStack.length > 0}
                canRedo={redoStack.length > 0}
                onUndo={handleUndo}
                onRedo={handleRedo}
                onManualSave={() => handleManualSaveNow()}
              />
            )}

            {activeMenu === 'review_monitoring' && (
              <ReviewMonitoringSheet
                docs={reviewDocs}
                activeDocId={activeReviewId}
                onSelectDoc={setActiveReviewId}
                onChangeDocs={handleUpdateReviewDocs}
                weeklyTables={weeklyTables}
                selectedWeeklyFilter={selectedWeeklyFilter}
                onSelectWeeklyFilter={setSelectedWeeklyFilter}
                showDivZero={showDivZero}
                canUndo={undoStack.length > 0 && accessRole === 'admin'}
                canRedo={redoStack.length > 0 && accessRole === 'admin'}
                onUndo={accessRole === 'admin' ? handleUndo : undefined}
                onRedo={accessRole === 'admin' ? handleRedo : undefined}
                onManualSave={accessRole === 'admin' ? () => handleManualSaveNow() : undefined}
                isPublicReadOnly={accessRole !== 'admin'}
              />
            )}

            {activeMenu === 'staff_analytics' && (
              <AnalyticsSummaryView
                tables={weeklyTables}
                selectedWeeklyFilter={selectedWeeklyFilter}
                onSelectWeeklyFilter={setSelectedWeeklyFilter}
              />
            )}

            {activeMenu === 'report_cv_div1' && (
              <ReportCvDiv1View
                records={cvDiv1Records}
                onUpdateRecords={handleUpdateCvDiv1Records}
                inventoryOps={cvInventoryOps}
                onUpdateInventoryOps={handleUpdateCvInventoryOps}
                isAdmin={accessRole === 'admin'}
                onTriggerSave={() => handleManualSaveNow()}
                activeHorizon={cvHorizon}
                onChangeHorizon={setCvHorizon}
              />
            )}

            {activeMenu === 'total_dp_cv' && (
              <TotalDpCvReportView
                weeklyTables={weeklyTables}
                cvRecords={cvDiv1Records}
                activeHorizon={totalDpCvHorizon}
                onChangeHorizon={setTotalDpCvHorizon}
                onNavigateToDp={(h, weekId) => {
                  if (weekId) setSelectedWeeklyFilter(weekId);
                  handleSelectSalesTableHorizon(h);
                }}
                onNavigateToCv={(h) => {
                  handleSelectCvHorizon(h);
                }}
                isAdmin={accessRole === 'admin'}
                onUpdateTraffic={
                  accessRole === 'admin'
                    ? (tanggalSlash, shift, trafficVal) => {
                        const baseTables = latestTablesRef.current;
                        const foundWeekIdx = baseTables.findIndex((t) =>
                          t.days.some((d) => d.tanggal === tanggalSlash)
                        );
                        if (foundWeekIdx >= 0) {
                          const nextTables = baseTables.map((t, wIdx) =>
                            wIdx !== foundWeekIdx
                              ? t
                              : {
                                  ...t,
                                  days: t.days.map((d) =>
                                    d.tanggal === tanggalSlash
                                      ? {
                                          ...d,
                                          [shift]: {
                                            ...d[shift],
                                            traffic: trafficVal,
                                          },
                                        }
                                      : d
                                  ),
                                }
                          );
                          handleUpdateTables(nextTables);
                        } else {
                          const sched = ANNUAL_2026_WEEKS.find((w) =>
                            w.slashDates.includes(tanggalSlash)
                          );
                          if (!sched) return;
                          const created = createWeeklyTableFromAnnualSchedule(sched);
                          const patchedTable = {
                            ...created,
                            days: created.days.map((d) =>
                              d.tanggal === tanggalSlash
                                ? {
                                    ...d,
                                    [shift]: {
                                      ...d[shift],
                                      traffic: trafficVal,
                                    },
                                  }
                                : d
                            ),
                          };
                          const nextTables = [...baseTables, patchedTable].sort(
                            (a, b) => a.weekNumber - b.weekNumber
                          );
                          handleUpdateTables(nextTables);
                        }
                      }
                    : undefined
                }
              />
            )}

            {activeMenu === 'data_cv_ops' && (
              <div className="space-y-4 pb-12">
                {dataCvToast && (
                  <div className="fixed bottom-6 right-6 z-50 flex items-center gap-2 px-4 py-3 bg-slate-900 text-white rounded-xl shadow-2xl border border-amber-400 text-xs font-bold">
                    <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                    <span>{dataCvToast}</span>
                  </div>
                )}

                <div className="bg-white border border-slate-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <div className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
                      Monitoring &amp; Evaluasi Bazar · Divisi 1 (Young Fashion, Ladies &amp; Intimate)
                    </div>
                    <h2 className="text-base sm:text-lg font-extrabold text-slate-950 tracking-tight mt-0.5">
                      Data CV — Inventory, Promo, Reason/Kendala &amp; Action Plan
                    </h2>
                  </div>

                  <div className="flex flex-wrap items-center gap-2">
                    <div className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-50 border border-slate-300 text-xs">
                      <CalendarDays className="w-3.5 h-3.5 text-slate-700" />
                      <span className="font-semibold text-slate-700">Bulan:</span>
                      <select
                        value={dataCvMonthKey}
                        onChange={(e) => setDataCvMonthKey(e.target.value)}
                        className="font-bold text-slate-900 bg-transparent focus:outline-none cursor-pointer"
                      >
                        {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                          <option key={mKey} value={mKey}>
                            {mName} 2026
                          </option>
                        ))}
                      </select>
                    </div>

                    <button
                      type="button"
                      onClick={() => {
                        const url = getGooglePublicDataCvUrl({
                          periodType: 'weekly',
                          week: `weekly-${dataCvWeekNumber}`,
                        });
                        void copyTextToClipboard(url).then((ok) => {
                          showDataCvToast(
                            ok
                              ? 'Link pengisian Data CV berhasil disalin ke clipboard!'
                              : 'Gagal menyalin link.'
                          );
                        });
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-700 hover:bg-sky-800 text-white cursor-pointer whitespace-nowrap shrink-0 transition-colors"
                      title="Salin link form pengisian Data CV untuk dibagikan ke petugas"
                    >
                      <Link2 className="w-3.5 h-3.5" />
                      <span>Salin Link Isi Data CV</span>
                    </button>

                    <button
                      type="button"
                      onClick={async () => {
                        try {
                          await exportCvDiv1ToExcel(
                            cvDiv1Records,
                            dataCvWeekNumber,
                            dataCvMonthKey,
                            2026,
                            cvInventoryOps
                          );
                          showDataCvToast('Excel Data CV (.xlsx) berhasil diunduh!');
                        } catch {
                          showDataCvToast('Gagal mengunduh file Excel.');
                        }
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white cursor-pointer whitespace-nowrap shrink-0 transition-colors"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Ekspor Excel CV</span>
                    </button>
                  </div>
                </div>

                <CvInventoryOpsTableSection
                  defaultPeriodType="weekly"
                  records={cvDiv1Records}
                  entries={cvInventoryOps}
                  onUpdateEntries={handleUpdateCvInventoryOps}
                  selectedMonthKey={dataCvMonthKey}
                  onChangeMonthKey={(mk) => setDataCvMonthKey(mk)}
                  selectedWeekNumber={dataCvWeekNumber}
                  onSelectWeekNumber={(wn) => setDataCvWeekNumber(wn)}
                  isAdmin={accessRole === 'admin'}
                  onToast={showDataCvToast}
                />
              </div>
            )}
          </div>
        </main>

        <SaveToolsModal
          isOpen={isSaveToolsOpen}
          onClose={() => setIsSaveToolsOpen(false)}
          weeklyTables={weeklyTables}
          reviewDocs={reviewDocs}
          submissionLogs={submissionLogs}
          cvDiv1Records={cvDiv1Records}
          cvInventoryOps={cvInventoryOps}
          lastSavedTimeStr={lastSavedTimeStr}
          isSaving={isSavingState}
          onManualSaveNow={handleManualSaveNow}
          onRestoreState={handleRestoreSavedState}
          onExportExcel={() => {
            setIsSaveToolsOpen(false);
            setIsExportModalOpen(true);
          }}
        />

        <ExportExcelModal
          isOpen={isExportModalOpen && accessRole === 'admin'}
          onClose={() => setIsExportModalOpen(false)}
          weeklyTables={weeklyTables}
          reviewDoc={currentReviewDoc}
          reviewDocs={reviewDocs}
          selectedWeeklyFilter={selectedWeeklyFilter}
          showDivZero={showDivZero}
          cvRecords={cvDiv1Records}
          cvInventoryOps={cvInventoryOps}
          onExportSuccess={(msg) => showBriefHistoryNotice(msg)}
        />

        <AppShareLinkModal
          isOpen={isAppShareModalOpen}
          onClose={() => setIsAppShareModalOpen(false)}
          accessRole={accessRole}
          adminPassword={adminPassword}
          onLoginAdmin={handleLoginAdminWithPassword}
          onLogoutToPublic={handleEnterPublicMode}
          onLogoutToEntryGate={handleLogoutToEntryGate}
          onChangeAdminPassword={handleChangeAdminPassword}
          onOpenReportForm={(formType) => {
            setDirectReportFormMode(formType);
            setIsLockedReportOnly(true);
            setActiveMenu('report_link');
            setShowEntryGate(false);
          }}
        />

        <AppLoginLogoutModal
          isOpen={isLoginLogoutModalOpen}
          onClose={() => setIsLoginLogoutModalOpen(false)}
          accessRole={accessRole}
          onLoginAdmin={handleLoginAdminWithPassword}
          onEnterPublic={handleEnterPublicMode}
          onLogoutToPublic={handleEnterPublicMode}
          onLogoutToEntryGate={handleLogoutToEntryGate}
          onChangeAdminPassword={handleChangeAdminPassword}
          onOpenReportForm={(formType) => {
            setDirectReportFormMode(formType);
            setIsLockedReportOnly(true);
            setActiveMenu('report_link');
            setShowEntryGate(false);
          }}
        />

          {/* Quiet Workspace Footer */}
          <footer className="no-print border-t border-slate-200 bg-white px-6 py-3.5 text-xs text-slate-500">
            <div className="max-w-[1540px] mx-auto flex flex-wrap items-center justify-between gap-2">
              <div>
                <strong className="text-slate-800">REPORT BAZAR MATAHARI</strong> ·{' '}
                <span className="font-semibold text-slate-700">Tanjung</span> ·{' '}
                <span className="font-semibold text-rose-700">PT MDS Retailing Tbk</span>
              </div>
              <div>
                Filter Aktif:{' '}
                <strong className="text-slate-700">
                  {selectedWeeklyFilter === 'ALL'
                    ? `Semua Weekly (${weeklyTables.length} Minggu)`
                    : weeklyTables.find((t) => t.id === selectedWeeklyFilter)?.title}
                </strong>{' '}
                · Mendukung Link Pengisian Online &amp; Ekspor WPS/Excel (.xlsx)
              </div>
            </div>
          </footer>
        </div>
      </div>
    </div>
  );
}
