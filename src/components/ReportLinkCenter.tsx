import React, { useState, useEffect, useRef } from 'react';
import {
  WeeklySalesTable,
  DailySalesRecord,
  ReviewMonitoringDoc,
  ReportSubmissionLog,
  formatNumberId,
  ANNUAL_2026_WEEKS,
  generateAnnualWeeksMondayToSunday,
  createWeeklyTableFromAnnualSchedule,
  getIndoDayName,
  getWeeklyPeriodeText,
  normalizeReviewItemsList,
  getReviewDocWeekNumber,
  createReviewDocForWeek,
} from '../data/initialData';
import { copyTextToClipboard } from '../utils/clipboard';
import {
  Link2,
  Copy,
  Check,
  ExternalLink,
  Share2,
  FileSpreadsheet,
  ClipboardList,
  Clock,
  CheckCircle2,
  Sun,
  Sunset,
  Send,
  ArrowLeft,
  UserCheck,
  Calculator,
  Lock,
  ShieldCheck,
  Calendar,
  Sparkles,
  ChevronLeft,
  ChevronRight,
  Tag,
  Layers,
  PackageCheck,
  TrendingUp,
  AlertTriangle,
} from 'lucide-react';
import {
  CvDailyRecord,
  CvInventoryBrandKey,
  CvInventoryOpsEntry,
  CvInventoryPeriodType,
  CV_BRAND_CATALOG,
  CV_DIVISI_DEFAULT,
  MONTH_NAMES_ID,
  calcCvBrandShiftTotals,
  calcCvLiveSalesQtyMap,
  createDefaultCvOpsEntry,
  getCvOpsEntryId,
} from '../data/cvDiv1Data';
import {
  AUTO_DRAFT_KEYS,
  usePeriodicAutoDraft,
  loadAutoDraftFromStorage,
  clearAutoDraftFromStorage,
} from '../utils/autoDraft';
import {
  getPublicShareableBaseUrl,
  getGooglePublicAppUrl,
  getGooglePublicShiftReportUrl,
  getGooglePublicReviewReportUrl,
  getGooglePublicCvReportUrl,
  getGooglePublicDataCvUrl,
  getSavedCustomPublicOrigin,
  setSavedCustomPublicOrigin,
  setAutoDetectedPublicOrigin,
} from '../utils/shareUrl';
import { AutoDraftStatusBadge } from './AutoDraftStatusBanner';

interface ShiftFormDraftPayload {
  formCalendarIso: string;
  formWeekId: string;
  formDayIdx: number;
  formShiftType: 'Pagi' | 'Siang' | 'Lembur' | 'BOTH';
  formStaffPagi: string;
  formStaffLembur?: string;
  formSalesPagi: string;
  formQtyPagi: string;
  formTrxPagi: string;
  formTrafficPagi?: string;
  formReasonPagi: string;
  formStaffSiang: string;
  formSalesSiang: string;
  formQtySiang: string;
  formTrxSiang: string;
  formTrafficSiang?: string;
  formReasonSiang: string;
}

interface ReviewFormDraftPayload {
  formReviewCalendarIso: string;
  formMerchDiv1: string;
  formMerchDiv2: string;
  formMerchDiv3: string;
  formMerchDiv4: string;
  formFastMoving: string;
  formSlowMoving: string;
  formPromoDiv1: string;
  formPromoDiv2: string;
  formPromoDiv3: string;
  formPromoDiv4: string;
  formCustomerNote: string;
  formActionPlan: string;
  formKendala: string;
  formNeedSupport: string;
}

function slashDateToIso(slash: string): string {
  const parts = (slash || '').trim().split('/');
  if (parts.length !== 3) return '2026-10-05';
  const dd = parts[0].padStart(2, '0');
  const mm = parts[1].padStart(2, '0');
  const yyyy = parts[2];
  return `${yyyy}-${mm}-${dd}`;
}

function isoToSlashDate(iso: string): string {
  const parts = (iso || '').trim().split('-');
  if (parts.length !== 3) return '05/10/2026';
  const [yyyy, mm, dd] = parts;
  return `${dd.padStart(2, '0')}/${mm.padStart(2, '0')}/${yyyy}`;
}

function shiftIsoDate(iso: string, deltaDays: number): string {
  const [y, m, d] = (iso || '2026-10-05').split('-').map(Number);
  const dt = new Date(y || 2026, (m || 10) - 1, d || 5);
  dt.setDate(dt.getDate() + deltaDays);
  const yyyy = String(dt.getFullYear());
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function parseCleanNumber(raw: string): number | null {
  const trimmed = (raw || '').trim();
  if (!trimmed) return null;
  const digitsOnly = trimmed.replace(/[^0-9]/g, '');
  if (!digitsOnly) return null;
  const num = Number(digitsOnly);
  return Number.isNaN(num) ? null : num;
}

function formatAutoNominal(raw: string): string {
  const digitsOnly = (raw || '').replace(/[^0-9]/g, '');
  if (!digitsOnly) return '';
  const num = Number(digitsOnly);
  if (Number.isNaN(num)) return '';
  return num.toLocaleString('id-ID');
}

interface ReportLinkCenterProps {
  tables: WeeklySalesTable[];
  onChangeTables: (updated: WeeklySalesTable[]) => void;
  reviewDocs: ReviewMonitoringDoc[];
  onChangeReviewDocs: (updated: ReviewMonitoringDoc[]) => void;
  submissionLogs: ReportSubmissionLog[];
  onAddSubmissionLog: (log: ReportSubmissionLog) => void;
  onAtomicSubmitReport?: (payload: {
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
  }) => void;
  initialFormMode?: 'none' | 'shift' | 'review' | 'cv_report' | 'data_cv';
  initialWeekId?: string;
  initialDayIdx?: number;
  initialDateIso?: string;
  initialShift?: 'Pagi' | 'Siang' | 'Lembur' | 'BOTH';
  initialStaff?: string;
  initialPeriodType?: CvInventoryPeriodType;
  initialMonthKey?: string;
  initialBrand?: 'ALL' | CvInventoryBrandKey;
  cvRecords?: CvDailyRecord[];
  cvInventoryOps?: CvInventoryOpsEntry[];
  onAtomicSubmitCvReport?: (payload: {
    updatedCvRecord?: CvDailyRecord;
    updatedCvOpsEntries?: CvInventoryOpsEntry[];
    newLog: ReportSubmissionLog;
  }) => void;
  isLockedReportOnly?: boolean;
  onToggleLockReportOnly?: (locked: boolean) => void;
}

export const ReportLinkCenter: React.FC<ReportLinkCenterProps> = ({
  tables,
  onChangeTables,
  reviewDocs,
  onChangeReviewDocs,
  submissionLogs,
  onAddSubmissionLog,
  onAtomicSubmitReport,
  initialFormMode = 'none',
  initialWeekId,
  initialDayIdx = 0,
  initialDateIso,
  initialShift = 'Pagi',
  initialStaff = '',
  initialPeriodType = 'weekly',
  initialMonthKey = '11',
  initialBrand = 'ALL',
  cvRecords = [],
  cvInventoryOps = [],
  onAtomicSubmitCvReport,
  isLockedReportOnly = false,
  onToggleLockReportOnly,
}) => {
  // Mode: 'hub' | 'shift_form' | 'review_form' | 'cv_report_form' | 'data_cv_form'
  const [activeView, setActiveView] = useState<
    'hub' | 'shift_form' | 'review_form' | 'cv_report_form' | 'data_cv_form'
  >(
    initialFormMode === 'shift'
      ? 'shift_form'
      : initialFormMode === 'review'
        ? 'review_form'
        : initialFormMode === 'cv_report'
          ? 'cv_report_form'
          : initialFormMode === 'data_cv'
            ? 'data_cv_form'
            : 'hub'
  );

  useEffect(() => {
    if (initialFormMode === 'shift') setActiveView('shift_form');
    else if (initialFormMode === 'review') setActiveView('review_form');
    else if (initialFormMode === 'cv_report') setActiveView('cv_report_form');
    else if (initialFormMode === 'data_cv') setActiveView('data_cv_form');
  }, [initialFormMode]);

  // Determine default initial table & day (prefer weekly-41 if initialWeekId not specified)
  const preferredDefaultTable =
    tables.find((t) => t.id === initialWeekId) ||
    tables.find((t) => t.id === 'weekly-41') ||
    tables[0];
  const defaultInitialDay =
    preferredDefaultTable?.days[initialDayIdx] || preferredDefaultTable?.days[0];
  const defaultIso =
    initialDateIso ||
    (defaultInitialDay ? slashDateToIso(defaultInitialDay.tanggal) : '2026-10-05');

  // Link Customizer States (for Link Generator in Hub)
  const [linkDateFilterMode, setLinkDateFilterMode] = useState<'CALENDAR' | 'ANY'>('CALENDAR');
  const [linkCalendarIso, setLinkCalendarIso] = useState<string>(defaultIso);
  const [linkWeekId, setLinkWeekId] = useState<string>(
    preferredDefaultTable?.id || 'weekly-41'
  );
  const [linkDayIdx, setLinkDayIdx] = useState<string>(
    initialDayIdx !== undefined ? String(initialDayIdx) : '0'
  ); // 'ANY' or '0'..'6'
  const [linkShift, setLinkShift] = useState<'ANY' | 'Pagi' | 'Siang' | 'Lembur'>('ANY');
  const [linkStaff, setLinkStaff] = useState<string>('');
  const [linkReviewDateMode, setLinkReviewDateMode] = useState<'CALENDAR' | 'ANY'>('CALENDAR');
  const [linkReviewCalendarIso, setLinkReviewCalendarIso] = useState<string>(defaultIso);

  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [customPublicOrigin, setCustomPublicOrigin] = useState<string>(() =>
    getSavedCustomPublicOrigin()
  );
  const [, setReportTunnelTick] = useState(0);

  useEffect(() => {
    let active = true;
    const checkTunnel = () => {
      fetch('/api/public-url', { cache: 'no-store' })
        .then((r) => (r.ok ? r.json() : null))
        .then((d) => {
          if (active && d && d.publicTunnelUrl) {
            setAutoDetectedPublicOrigin(d.publicTunnelUrl);
            setReportTunnelTick((n) => n + 1);
          }
        })
        .catch(() => {});
    };
    checkTunnel();
    const timer = window.setInterval(checkTunnel, 2500);
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, []);

  // Form States for Shift Report Submission
  const [formCalendarIso, setFormCalendarIso] = useState<string>(defaultIso);
  const [formWeekId, setFormWeekId] = useState<string>(
    preferredDefaultTable?.id || 'weekly-41'
  );
  const [formDayIdx, setFormDayIdx] = useState<number>(initialDayIdx);
  const [formShiftType, setFormShiftType] = useState<'Pagi' | 'Siang' | 'Lembur'>(
    initialShift === 'Siang'
      ? 'Siang'
      : initialShift === 'Lembur' || initialShift === 'BOTH'
        ? 'Lembur'
        : 'Pagi'
  );
  const [formStaffPagi, setFormStaffPagi] = useState<string>(initialStaff || '');
  const [formSalesPagi, setFormSalesPagi] = useState<string>('');
  const [formQtyPagi, setFormQtyPagi] = useState<string>('');
  const [formTrxPagi, setFormTrxPagi] = useState<string>('');
  const [formTrafficPagi, setFormTrafficPagi] = useState<string>('');
  const [formReasonPagi, setFormReasonPagi] = useState<string>('');

  const [formStaffSiang, setFormStaffSiang] = useState<string>(initialStaff || '');
  const [formSalesSiang, setFormSalesSiang] = useState<string>('');
  const [formQtySiang, setFormQtySiang] = useState<string>('');
  const [formTrxSiang, setFormTrxSiang] = useState<string>('');
  const [formTrafficSiang, setFormTrafficSiang] = useState<string>('');
  const [formReasonSiang, setFormReasonSiang] = useState<string>('');

  const [formStaffLembur, setFormStaffLembur] = useState<string>(initialStaff || '');

  const [submitSuccessData, setSubmitSuccessData] =
    useState<ReportSubmissionLog | null>(null);
  const [sessionSeed, setSessionSeed] = useState<string>(() =>
    String(Date.now())
  );
  const [lastLiveAppSaveTimeStr, setLastLiveAppSaveTimeStr] = useState<
    string | null
  >(null);
  const shiftAutoSaveTimerRef = useRef<number | null>(null);
  const reviewAutoSaveTimerRef = useRef<number | null>(null);

  // Form States for Review Monitoring Submission
  const [formReviewCalendarIso, setFormReviewCalendarIso] = useState<string>(defaultIso);
  const [formMerchDiv1, setFormMerchDiv1] = useState<string>('');
  const [formMerchDiv2, setFormMerchDiv2] = useState<string>('');
  const [formMerchDiv3, setFormMerchDiv3] = useState<string>('');
  const [formMerchDiv4, setFormMerchDiv4] = useState<string>('');
  const [formFastMoving, setFormFastMoving] = useState<string>('');
  const [formSlowMoving, setFormSlowMoving] = useState<string>('');
  const [formPromoDiv1, setFormPromoDiv1] = useState<string>('');
  const [formPromoDiv2, setFormPromoDiv2] = useState<string>('');
  const [formPromoDiv3, setFormPromoDiv3] = useState<string>('');
  const [formPromoDiv4, setFormPromoDiv4] = useState<string>('');
  const [formCustomerNote, setFormCustomerNote] = useState<string>('');
  const [formActionPlan, setFormActionPlan] = useState<string>('');
  const [formKendala, setFormKendala] = useState<string>('');
  const [formNeedSupport, setFormNeedSupport] = useState<string>('');

  // Link Generator States for Report CV & Data CV
  const defaultCvIso =
    initialDateIso ||
    (cvRecords[0]?.tanggal ? slashDateToIso(cvRecords[0].tanggal) : '2026-11-01');
  const [linkCvDateMode, setLinkCvDateMode] = useState<'CALENDAR' | 'ANY'>('CALENDAR');
  const [linkCvCalendarIso, setLinkCvCalendarIso] = useState<string>(defaultCvIso);
  const [linkCvBrand, setLinkCvBrand] = useState<CvInventoryBrandKey>(
    initialBrand === 'SIMPLICITY' ? 'SIMPLICITY' : 'GRAPHIS'
  );
  const [linkCvStaff, setLinkCvStaff] = useState<string>('');

  const [linkDataCvPeriodType, setLinkDataCvPeriodType] =
    useState<CvInventoryPeriodType>(initialPeriodType);
  const [linkDataCvWeekNum, setLinkDataCvWeekNum] = useState<number>(() => {
    if (initialWeekId && initialWeekId.startsWith('weekly-')) {
      const n = Number(initialWeekId.replace('weekly-', ''));
      if (n >= 1 && n <= 53) return n;
    }
    return 45;
  });
  const [linkDataCvMonthKey, setLinkDataCvMonthKey] = useState<string>(initialMonthKey || '11');
  const [linkDataCvBrand, setLinkDataCvBrand] = useState<CvInventoryBrandKey>(
    initialBrand === 'SIMPLICITY' ? 'SIMPLICITY' : 'GRAPHIS'
  );
  const [linkDataCvStaff, setLinkDataCvStaff] = useState<string>('');

  // Form States for Standalone Report CV Form (cv_report_form)
  const [cvFormCalendarIso, setCvFormCalendarIso] = useState<string>(defaultCvIso);
  const [cvFormBrandScope, setCvFormBrandScope] = useState<CvInventoryBrandKey>(
    initialBrand === 'SIMPLICITY' ? 'SIMPLICITY' : 'GRAPHIS'
  );
  const [cvFormStaff, setCvFormStaff] = useState<string>(initialStaff || '');
  const [cvFormGraphisSales, setCvFormGraphisSales] = useState<string>('');
  const [cvFormGraphisQty, setCvFormGraphisQty] = useState<string>('');
  const [cvFormGraphisNota, setCvFormGraphisNota] = useState<string>('');
  const [cvFormSimplicitySales, setCvFormSimplicitySales] = useState<string>('');
  const [cvFormSimplicityQty, setCvFormSimplicityQty] = useState<string>('');
  const [cvFormSimplicityNota, setCvFormSimplicityNota] = useState<string>('');
  const [cvFormReasonEvent, setCvFormReasonEvent] = useState<string>('');

  // Form States for Standalone Data CV Form (data_cv_form)
  interface DataCvBrandDraft {
    divisi: string;
    stockAwal: string;
    barangDatang: string;
    salesQtyMode: 'live' | 'manual';
    salesQtyManual: string;
    barangKeluar: string;
    promo: string;
    reasonKendala: string;
    actionPlan: string;
  }
  const [dataCvFormPeriodType, setDataCvFormPeriodType] =
    useState<CvInventoryPeriodType>(initialPeriodType);
  const [dataCvFormWeekNum, setDataCvFormWeekNum] = useState<number>(() => {
    if (initialWeekId && initialWeekId.startsWith('weekly-')) {
      const n = Number(initialWeekId.replace('weekly-', ''));
      if (n >= 1 && n <= 53) return n;
    }
    if (initialDateIso) {
      const slash = isoToSlashDate(initialDateIso);
      const matched = ANNUAL_2026_WEEKS.find(
        (w) =>
          w.slashDates.includes(slash) ||
          (initialDateIso >= w.startDateIso && initialDateIso <= w.endDateIso)
      );
      if (matched) return matched.weekNumber;
    }
    return 45;
  });
  const [dataCvFormMonthKey, setDataCvFormMonthKey] = useState<string>(() => {
    if (initialMonthKey && MONTH_NAMES_ID[initialMonthKey]) return initialMonthKey;
    if (initialDateIso && initialDateIso.split('-').length === 3) {
      const mm = initialDateIso.split('-')[1];
      if (MONTH_NAMES_ID[mm]) return mm;
    }
    return '11';
  });
  const [dataCvFormBrandScope, setDataCvFormBrandScope] =
    useState<CvInventoryBrandKey>(
      initialBrand === 'SIMPLICITY' ? 'SIMPLICITY' : 'GRAPHIS'
    );
  const [dataCvFormStaff, setDataCvFormStaff] = useState<string>(initialStaff || '');
  const [dataCvGraphisDraft, setDataCvGraphisDraft] = useState<DataCvBrandDraft>({
    divisi: CV_DIVISI_DEFAULT,
    stockAwal: '0',
    barangDatang: '0',
    salesQtyMode: 'live',
    salesQtyManual: '0',
    barangKeluar: '0',
    promo: '',
    reasonKendala: '',
    actionPlan: '',
  });
  const [dataCvSimplicityDraft, setDataCvSimplicityDraft] = useState<DataCvBrandDraft>({
    divisi: CV_DIVISI_DEFAULT,
    stockAwal: '0',
    barangDatang: '0',
    salesQtyMode: 'live',
    salesQtyManual: '0',
    barangKeluar: '0',
    promo: '',
    reasonKendala: '',
    actionPlan: '',
  });

  const liveCvSalesMap = React.useMemo(
    () => calcCvLiveSalesQtyMap(cvRecords),
    [cvRecords]
  );

  const loadCvDayIntoForm = (isoDate: string) => {
    const slash = isoToSlashDate(isoDate);
    const rec = cvRecords.find((r) => r.tanggal.trim() === slash);
    if (!rec) {
      setCvFormGraphisSales('');
      setCvFormGraphisQty('');
      setCvFormGraphisNota('');
      setCvFormSimplicitySales('');
      setCvFormSimplicityQty('');
      setCvFormSimplicityNota('');
      setCvFormReasonEvent('');
      return;
    }
    const gTot = calcCvBrandShiftTotals(rec.graphis);
    const sTot = calcCvBrandShiftTotals(rec.simplicity);
    setCvFormStaff(rec.petugasPagi || rec.petugasSiang || initialStaff || '');
    setCvFormGraphisSales(
      gTot.sales > 0 ? formatAutoNominal(String(gTot.sales)) : ''
    );
    setCvFormGraphisQty(gTot.qty > 0 ? formatAutoNominal(String(gTot.qty)) : '');
    setCvFormGraphisNota(gTot.trx > 0 ? formatAutoNominal(String(gTot.trx)) : '');
    setCvFormSimplicitySales(
      sTot.sales > 0 ? formatAutoNominal(String(sTot.sales)) : ''
    );
    setCvFormSimplicityQty(sTot.qty > 0 ? formatAutoNominal(String(sTot.qty)) : '');
    setCvFormSimplicityNota(sTot.trx > 0 ? formatAutoNominal(String(sTot.trx)) : '');
    setCvFormReasonEvent(rec.reasonEvent || '');
  };

  const loadDataCvIntoForm = (
    pType: CvInventoryPeriodType,
    weekNum: number,
    monthKey: string
  ) => {
    const pKey = pType === 'weekly' ? String(weekNum) : monthKey;
    const makeBrandDraft = (brand: CvInventoryBrandKey): DataCvBrandDraft => {
      const id = getCvOpsEntryId(pType, pKey, brand);
      const entry =
        cvInventoryOps.find((e) => e.id === id) ||
        createDefaultCvOpsEntry(pType, pKey, brand);
      const liveQty =
        pType === 'weekly'
          ? liveCvSalesMap.weekly.get(weekNum)?.[brand] || 0
          : liveCvSalesMap.monthly.get(monthKey)?.[brand] || 0;
      const isManual =
        entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined;
      return {
        divisi: entry.divisi || CV_DIVISI_DEFAULT,
        stockAwal: String(entry.stockAwal ?? 0),
        barangDatang: String(entry.barangDatang ?? 0),
        salesQtyMode: isManual ? 'manual' : 'live',
        salesQtyManual: String(isManual ? entry.salesQtyOverride : liveQty),
        barangKeluar: String(entry.barangKeluar ?? 0),
        promo: entry.promo || '',
        reasonKendala: entry.reasonKendala || '',
        actionPlan: entry.actionPlan || '',
      };
    };
    setDataCvGraphisDraft(makeBrandDraft('GRAPHIS'));
    setDataCvSimplicityDraft(makeBrandDraft('SIMPLICITY'));
  };

  useEffect(() => {
    loadCvDayIntoForm(cvFormCalendarIso);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cvFormCalendarIso]);

  useEffect(() => {
    loadDataCvIntoForm(
      dataCvFormPeriodType,
      dataCvFormWeekNum,
      dataCvFormMonthKey
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dataCvFormPeriodType, dataCvFormWeekNum, dataCvFormMonthKey]);

  // Reset Shift Form input fields so filling the link again starts fresh/clean
  const resetShiftFormInputs = (alsoClearStorage: boolean = true) => {
    setFormStaffPagi('');
    setFormStaffSiang('');
    setFormStaffLembur('');
    setFormSalesPagi('');
    setFormQtyPagi('');
    setFormTrxPagi('');
    setFormTrafficPagi('');
    setFormReasonPagi('');
    setFormSalesSiang('');
    setFormQtySiang('');
    setFormTrxSiang('');
    setFormTrafficSiang('');
    setFormReasonSiang('');
    if (alsoClearStorage) {
      clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REPORT_FORM_SHIFT);
    }
  };

  // Reset Review Form input fields so filling the link again starts fresh/clean
  const resetReviewFormInputs = (alsoClearStorage: boolean = true) => {
    setFormMerchDiv1('');
    setFormMerchDiv2('');
    setFormMerchDiv3('');
    setFormMerchDiv4('');
    setFormFastMoving('');
    setFormSlowMoving('');
    setFormPromoDiv1('');
    setFormPromoDiv2('');
    setFormPromoDiv3('');
    setFormPromoDiv4('');
    setFormCustomerNote('');
    setFormActionPlan('');
    setFormKendala('');
    setFormNeedSupport('');
    if (alsoClearStorage) {
      clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REPORT_FORM_REVIEW);
    }
  };

  const hasShiftDraftContent = Boolean(
    formSalesPagi.trim() ||
      formQtyPagi.trim() ||
      formTrxPagi.trim() ||
      formTrafficPagi.trim() ||
      formReasonPagi.trim() ||
      formSalesSiang.trim() ||
      formQtySiang.trim() ||
      formTrxSiang.trim() ||
      formTrafficSiang.trim() ||
      formReasonSiang.trim()
  );

  const hasReviewDraftContent = Boolean(
    formMerchDiv1.trim() ||
      formMerchDiv2.trim() ||
      formMerchDiv3.trim() ||
      formMerchDiv4.trim() ||
      formFastMoving.trim() ||
      formSlowMoving.trim() ||
      formPromoDiv1.trim() ||
      formPromoDiv2.trim() ||
      formPromoDiv3.trim() ||
      formPromoDiv4.trim() ||
      formCustomerNote.trim() ||
      formActionPlan.trim() ||
      formKendala.trim() ||
      formNeedSupport.trim()
  );

  const {
    lastDraftTimeStr: shiftDraftTimeStr,
    isOnline: isShiftOnline,
  } = usePeriodicAutoDraft<ShiftFormDraftPayload>({
    storageKey: AUTO_DRAFT_KEYS.REPORT_FORM_SHIFT,
    menuName: 'Form Report DP',
    data: {
      formCalendarIso,
      formWeekId,
      formDayIdx,
      formShiftType,
      formStaffPagi,
      formSalesPagi,
      formQtyPagi,
      formTrxPagi,
      formTrafficPagi,
      formReasonPagi,
      formStaffSiang,
      formSalesSiang,
      formQtySiang,
      formTrxSiang,
      formTrafficSiang,
      formReasonSiang,
    },
    isDirty: hasShiftDraftContent,
    intervalMs: 2000,
  });

  const {
    lastDraftTimeStr: reviewDraftTimeStr,
    isOnline: isReviewOnline,
  } = usePeriodicAutoDraft<ReviewFormDraftPayload>({
    storageKey: AUTO_DRAFT_KEYS.REPORT_FORM_REVIEW,
    menuName: 'Form Data DP (Review Monitoring)',
    data: {
      formReviewCalendarIso,
      formMerchDiv1,
      formMerchDiv2,
      formMerchDiv3,
      formMerchDiv4,
      formFastMoving,
      formSlowMoving,
      formPromoDiv1,
      formPromoDiv2,
      formPromoDiv3,
      formPromoDiv4,
      formCustomerNote,
      formActionPlan,
      formKendala,
      formNeedSupport,
    },
    isDirty: hasReviewDraftContent,
    intervalMs: 2000,
  });

  // Load existing day values
  const loadDayIntoFormFromTable = (tbl: WeeklySalesTable | undefined, dIdx: number) => {
    const day = tbl?.days[dIdx];
    if (!day) return;
    setFormStaffPagi(day.pagi.nama || initialStaff || '');
    setFormSalesPagi(
      day.pagi.sales !== null ? formatAutoNominal(String(day.pagi.sales)) : ''
    );
    setFormQtyPagi(
      day.pagi.qty !== null ? formatAutoNominal(String(day.pagi.qty)) : ''
    );
    setFormTrxPagi(
      day.pagi.trx !== null ? formatAutoNominal(String(day.pagi.trx)) : ''
    );
    setFormTrafficPagi(
      day.pagi.traffic !== undefined && day.pagi.traffic !== null
        ? formatAutoNominal(String(day.pagi.traffic))
        : ''
    );
    setFormReasonPagi(day.reasonPagi || '');

    setFormStaffSiang(day.siang.nama || initialStaff || '');
    setFormStaffLembur(day.pagi.nama || day.siang.nama || initialStaff || '');
    setFormSalesSiang(
      day.siang.sales !== null ? formatAutoNominal(String(day.siang.sales)) : ''
    );
    setFormQtySiang(
      day.siang.qty !== null ? formatAutoNominal(String(day.siang.qty)) : ''
    );
    setFormTrxSiang(
      day.siang.trx !== null ? formatAutoNominal(String(day.siang.trx)) : ''
    );
    setFormTrafficSiang(
      day.siang.traffic !== undefined && day.siang.traffic !== null
        ? formatAutoNominal(String(day.siang.traffic))
        : ''
    );
    setFormReasonSiang(day.reasonSiang || '');
  };

  const loadDayIntoForm = (weekId: string, dIdx: number) => {
    const tbl = tables.find((t) => t.id === weekId) || tables[0];
    loadDayIntoFormFromTable(tbl, dIdx);
  };

  // Resolve any ISO calendar date (YYYY-MM-DD) -> automatically find or create its Weekly table & day index
  const resolveCalendarDateToWeekly = (
    isoDate: string,
    persistIfNew: boolean = true
  ): {
    table: WeeklySalesTable;
    dayIdx: number;
    slashDate: string;
  } => {
    const slashTarget = isoToSlashDate(isoDate);

    // 1. Check existing loaded tables first
    for (const tbl of tables) {
      const foundDayIdx = tbl.days.findIndex((d) => d.tanggal.trim() === slashTarget);
      if (foundDayIdx >= 0) {
        return { table: tbl, dayIdx: foundDayIdx, slashDate: slashTarget };
      }
    }

    // 2. Check Annual Schedule (2026 or the year from isoDate)
    const yearNum = Number((isoDate || '2026').split('-')[0]) || 2026;
    const annualList =
      yearNum === 2026
        ? ANNUAL_2026_WEEKS
        : generateAnnualWeeksMondayToSunday(yearNum);

    const matchedSched =
      annualList.find((w) => w.slashDates.includes(slashTarget)) ||
      annualList.find(
        (w) => isoDate >= w.startDateIso && isoDate <= w.endDateIso
      );

    if (matchedSched) {
      const dIdx = Math.max(0, matchedSched.slashDates.indexOf(slashTarget));
      // Check if table with this weekNumber exists in tables
      const existingByNum = tables.find(
        (t) => t.weekNumber === matchedSched.weekNumber
      );
      if (existingByNum) {
        return {
          table: existingByNum,
          dayIdx: Math.min(dIdx, existingByNum.days.length - 1),
          slashDate: slashTarget,
        };
      }

      // Create the Weekly table automatically so it is available immediately
      const createdTable = createWeeklyTableFromAnnualSchedule(matchedSched);
      if (persistIfNew) {
        const nextTables = [...tables, createdTable].sort(
          (a, b) => a.weekNumber - b.weekNumber
        );
        onChangeTables(nextTables);
      }
      return {
        table: createdTable,
        dayIdx: dIdx,
        slashDate: slashTarget,
      };
    }

    // Fallback to first table
    const fallbackTable = preferredDefaultTable || tables[0];
    return {
      table: fallbackTable,
      dayIdx: 0,
      slashDate: fallbackTable?.days[0]?.tanggal || slashTarget,
    };
  };

  // Handler when user picks a date on the Calendar in Hub Link Generator
  const handleSelectLinkCalendarDate = (isoDate: string) => {
    if (!isoDate) return;
    setLinkCalendarIso(isoDate);
    setLinkDateFilterMode('CALENDAR');
    const resolved = resolveCalendarDateToWeekly(isoDate, false);
    setLinkWeekId(resolved.table.id);
    setLinkDayIdx(String(resolved.dayIdx));
  };

  // Handler when user picks a date on the Calendar in Shift Report Form
  const handleSelectFormCalendarDate = (isoDate: string, shouldResetInputs: boolean = false) => {
    if (!isoDate) return;
    const isChangingDate = isoDate !== formCalendarIso;
    setFormCalendarIso(isoDate);
    const resolved = resolveCalendarDateToWeekly(isoDate, false);
    setFormWeekId(resolved.table.id);
    setFormDayIdx(resolved.dayIdx);
    if (shouldResetInputs || isChangingDate) {
      resetShiftFormInputs();
      setSessionSeed(String(Date.now()));
    }
  };

  // Load existing Review Monitoring doc for a specific weekNumber into Review Form
  const loadReviewIntoFormByWeekNumber = (
    weekNum: number,
    docsSource: ReviewMonitoringDoc[] = reviewDocs
  ) => {
    const doc = docsSource.find(
      (d) =>
        getReviewDocWeekNumber(d) === weekNum ||
        d.id === `review-w${weekNum}` ||
        d.periode.toUpperCase().includes(`WEEKLY ${weekNum}`)
    );
    if (!doc) {
      resetReviewFormInputs(false);
      return;
    }
    const merch = doc.sections.find((s) => s.code.includes('MERCHANDISE'));
    const mDiv1 = merch?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 1');
    const mDiv2 = merch?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 2');
    const mDiv3 = merch?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 3');
    const mDiv4 = merch?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 4');
    const fast = merch?.subRows?.find((r) => r.label.includes('FAST MOVING'));
    const slow = merch?.subRows?.find((r) => r.label.includes('SLOW MOVING'));

    const promo = doc.sections.find((s) => s.code.includes('PROMO'));
    const pDiv1 = promo?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 1');
    const pDiv2 = promo?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 2');
    const pDiv3 = promo?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 3');
    const pDiv4 = promo?.subRows?.find((r) => r.label.trim().toLowerCase() === 'divisi 4');

    const place = doc.sections.find((s) => s.code.includes('PLATE') || s.code.includes('PLACE'));
    const action = doc.sections.find((s) => s.code.includes('ACTION PLAN'));
    const reason = doc.sections.find((s) => s.code.includes('REASON'));
    const support = doc.sections.find((s) => s.code.includes('NEED SUPPORT'));

    setFormMerchDiv1((mDiv1?.items || []).join('\n'));
    setFormMerchDiv2((mDiv2?.items || []).join('\n'));
    setFormMerchDiv3((mDiv3?.items || []).join('\n'));
    setFormMerchDiv4((mDiv4?.items || []).join('\n'));
    setFormFastMoving((fast?.items || []).join('\n'));
    setFormSlowMoving((slow?.items || []).join('\n'));
    setFormPromoDiv1((pDiv1?.items || []).join('\n'));
    setFormPromoDiv2((pDiv2?.items || []).join('\n'));
    setFormPromoDiv3((pDiv3?.items || []).join('\n'));
    setFormPromoDiv4((pDiv4?.items || []).join('\n'));
    setFormCustomerNote((place?.directItems || []).join('\n'));
    setFormActionPlan((action?.directItems || []).join('\n'));
    setFormKendala((reason?.directItems || []).join('\n'));
    setFormNeedSupport((support?.directItems || []).join('\n'));
  };

  // Handler when user picks a date on the Calendar in Review Monitoring Form
  const handleSelectReviewCalendarDate = (isoDate: string) => {
    if (!isoDate) return;
    const prevWeekNum = resolveCalendarDateToWeekly(formReviewCalendarIso, false).table.weekNumber;
    const nextWeekNum = resolveCalendarDateToWeekly(isoDate, false).table.weekNumber;
    setFormReviewCalendarIso(isoDate);
    if (nextWeekNum !== prevWeekNum) {
      loadReviewIntoFormByWeekNumber(nextWeekNum);
    }
  };

  // Initialize form values on mount or restore unsaved Auto-Draft from localStorage if present
  useEffect(() => {
    const savedShiftDraft = loadAutoDraftFromStorage<ShiftFormDraftPayload>(
      AUTO_DRAFT_KEYS.REPORT_FORM_SHIFT
    );
    const savedReviewDraft = loadAutoDraftFromStorage<ReviewFormDraftPayload>(
      AUTO_DRAFT_KEYS.REPORT_FORM_REVIEW
    );

    let effectiveReviewIso = formReviewCalendarIso;
    if (initialDateIso) {
      handleSelectFormCalendarDate(initialDateIso, false);
      handleSelectLinkCalendarDate(initialDateIso);
      setFormReviewCalendarIso(initialDateIso);
      effectiveReviewIso = initialDateIso;
    } else {
      const initTbl =
        tables.find((t) => t.id === (initialWeekId || formWeekId)) ||
        preferredDefaultTable ||
        tables[0];
      const initDay = initTbl?.days[initialDayIdx] || initTbl?.days[0];
      if (initDay) {
        const iso = slashDateToIso(initDay.tanggal);
        setFormCalendarIso(iso);
        setLinkCalendarIso(iso);
        setFormReviewCalendarIso(iso);
        effectiveReviewIso = iso;
      }
    }

    // Pre-populate Review Form with existing saved ReviewMonitoringDoc for that week so all sections stay intact
    const initialRevWeekNum = resolveCalendarDateToWeekly(effectiveReviewIso, false).table.weekNumber;
    loadReviewIntoFormByWeekNumber(initialRevWeekNum, reviewDocs);

    // Restore Shift Form Auto-Draft if user had unsaved input before connection drop / reload
    if (savedShiftDraft && savedShiftDraft.isDirty && savedShiftDraft.data) {
      const sd = savedShiftDraft.data;
      if (sd.formCalendarIso && !initialDateIso) {
        handleSelectFormCalendarDate(sd.formCalendarIso, false);
      }
      if (sd.formShiftType) {
        setFormShiftType(
          sd.formShiftType === 'BOTH' ? 'Lembur' : sd.formShiftType
        );
      }
      if (sd.formStaffLembur) setFormStaffLembur(sd.formStaffLembur);
      if (sd.formStaffPagi) setFormStaffPagi(sd.formStaffPagi);
      if (sd.formSalesPagi) setFormSalesPagi(sd.formSalesPagi);
      if (sd.formQtyPagi) setFormQtyPagi(sd.formQtyPagi);
      if (sd.formTrxPagi) setFormTrxPagi(sd.formTrxPagi);
      if (sd.formTrafficPagi) setFormTrafficPagi(sd.formTrafficPagi);
      if (sd.formReasonPagi) setFormReasonPagi(sd.formReasonPagi);
      if (sd.formStaffSiang) setFormStaffSiang(sd.formStaffSiang);
      if (sd.formSalesSiang) setFormSalesSiang(sd.formSalesSiang);
      if (sd.formQtySiang) setFormQtySiang(sd.formQtySiang);
      if (sd.formTrxSiang) setFormTrxSiang(sd.formTrxSiang);
      if (sd.formTrafficSiang) setFormTrafficSiang(sd.formTrafficSiang);
      if (sd.formReasonSiang) setFormReasonSiang(sd.formReasonSiang);
    }

    // Restore Review Form Auto-Draft if user had unsaved input before connection drop / reload
    if (savedReviewDraft && savedReviewDraft.isDirty && savedReviewDraft.data) {
      const rd = savedReviewDraft.data;
      if (rd.formReviewCalendarIso && !initialDateIso) {
        setFormReviewCalendarIso(rd.formReviewCalendarIso);
        const draftWk = resolveCalendarDateToWeekly(rd.formReviewCalendarIso, false).table.weekNumber;
        loadReviewIntoFormByWeekNumber(draftWk, reviewDocs);
      }
      if (rd.formMerchDiv1) setFormMerchDiv1(rd.formMerchDiv1);
      if (rd.formMerchDiv2) setFormMerchDiv2(rd.formMerchDiv2);
      if (rd.formMerchDiv3) setFormMerchDiv3(rd.formMerchDiv3);
      if (rd.formMerchDiv4) setFormMerchDiv4(rd.formMerchDiv4);
      if (rd.formFastMoving) setFormFastMoving(rd.formFastMoving);
      if (rd.formSlowMoving) setFormSlowMoving(rd.formSlowMoving);
      if (rd.formPromoDiv1) setFormPromoDiv1(rd.formPromoDiv1);
      if (rd.formPromoDiv2) setFormPromoDiv2(rd.formPromoDiv2);
      if (rd.formPromoDiv3) setFormPromoDiv3(rd.formPromoDiv3);
      if (rd.formPromoDiv4) setFormPromoDiv4(rd.formPromoDiv4);
      if (rd.formCustomerNote) setFormCustomerNote(rd.formCustomerNote);
      if (rd.formActionPlan) setFormActionPlan(rd.formActionPlan);
      if (rd.formKendala) setFormKendala(rd.formKendala);
      if (rd.formNeedSupport) setFormNeedSupport(rd.formNeedSupport);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Construct shareable URLs (automatically converts private ais-dev-* to public Shared App URL ais-pre-* and uses clean paths /isi-report, /isi-review, /publik)
  const publicAppUrl = getGooglePublicAppUrl(
    customPublicOrigin,
    linkDateFilterMode === 'CALENDAR' && linkCalendarIso
      ? { date: linkCalendarIso }
      : undefined
  );

  const buildShiftReportLink = () => {
    const extraParams: Record<string, string> = {};
    if (linkDateFilterMode === 'CALENDAR' && linkCalendarIso) {
      extraParams.date = linkCalendarIso;
    }
    if (linkWeekId) {
      extraParams.week = linkWeekId;
    }
    if (linkDateFilterMode === 'CALENDAR' && linkDayIdx !== 'ANY') {
      extraParams.day = linkDayIdx;
    }
    if (linkShift !== 'ANY') extraParams.shift = linkShift;
    if (linkStaff.trim() && linkStaff.trim() !== 'ANY') extraParams.staff = linkStaff.trim();
    return getGooglePublicShiftReportUrl(extraParams, customPublicOrigin);
  };

  const buildReviewReportLink = () => {
    const resolvedRev =
      linkReviewDateMode === 'CALENDAR' && linkReviewCalendarIso
        ? resolveCalendarDateToWeekly(linkReviewCalendarIso, false)
        : null;
    const matchedDoc = resolvedRev
      ? reviewDocs.find(
          (d) =>
            getReviewDocWeekNumber(d) === resolvedRev.table.weekNumber ||
            d.id === `review-w${resolvedRev.table.weekNumber}`
        )
      : undefined;
    const extraParams: Record<string, string> = {};
    if (linkReviewDateMode === 'CALENDAR' && linkReviewCalendarIso) {
      extraParams.date = linkReviewCalendarIso;
    }
    return getGooglePublicReviewReportUrl(
      matchedDoc?.id || reviewDocs[0]?.id || 'review-sep-w4',
      customPublicOrigin,
      extraParams
    );
  };

  const buildCvReportLink = () => {
    const extraParams: Record<string, string> = {
      brand: linkCvBrand,
    };
    if (linkCvDateMode === 'CALENDAR' && linkCvCalendarIso) {
      extraParams.date = linkCvCalendarIso;
    }
    if (linkCvStaff.trim()) {
      extraParams.staff = linkCvStaff.trim();
    }
    return getGooglePublicCvReportUrl(extraParams, customPublicOrigin);
  };

  const buildDataCvLink = () => {
    const extraParams: Record<string, string> = {
      periodType: linkDataCvPeriodType,
      brand: linkDataCvBrand,
    };
    if (linkDataCvPeriodType === 'weekly') {
      extraParams.week = `weekly-${linkDataCvWeekNum}`;
    } else {
      extraParams.month = linkDataCvMonthKey;
    }
    if (linkDataCvStaff.trim()) {
      extraParams.staff = linkDataCvStaff.trim();
    }
    return getGooglePublicDataCvUrl(extraParams, customPublicOrigin);
  };

  const shiftReportUrl = buildShiftReportLink();
  const reviewReportUrl = buildReviewReportLink();
  const cvReportUrl = buildCvReportLink();
  const dataCvUrl = buildDataCvLink();

  const handleCopyText = (key: string, text: string) => {
    void copyTextToClipboard(text).then((ok) => {
      if (ok) {
        setCopiedKey(key);
        setTimeout(() => setCopiedKey(null), 2200);
      }
    });
  };

  const selectedLinkWeekObj =
    tables.find((t) => t.id === linkWeekId) ||
    resolveCalendarDateToWeekly(linkCalendarIso, false).table ||
    tables[0];

  const buildWhatsAppInviteText = () => {
    const dayText =
      linkDayIdx !== 'ANY'
        ? selectedLinkWeekObj?.days[Number(linkDayIdx)]?.tanggal || 'Hari Ini'
        : 'Harian';
    const shiftText = linkShift !== 'ANY' ? `Shift ${linkShift}` : 'Shift Pagi / Siang / Lembur';
    const staffText =
      linkStaff.trim() && linkStaff.trim() !== 'ANY'
        ? `(${linkStaff.trim()})`
        : 'Tim Bazar Tanjung';

    return [
      `*LINK PENGISIAN REPORT DP HARIAN BAZAR TANJUNG*`,
      `Halo ${staffText}, silakan isi laporan penjualan *${selectedLinkWeekObj?.title}* (${dayText} - ${shiftText}) melalui link berikut:`,
      ``,
      shiftReportUrl,
      ``,
      `Data *Sales, Qty, Trx, Traffic Pengunjung, Price Point, Big Size* & *Reason/Event* akan otomatis terupdate di Tabel Rekapitulasi & Total Conversion Rate (DP + CV). Terima kasih!`,
    ].join('\n');
  };

  // Build atomic Shift Report payload (used by both Live Auto-Save and Final Submit)
  const buildShiftSubmissionPayload = (customLogId?: string) => {
    const resolved = resolveCalendarDateToWeekly(formCalendarIso, false);
    const targetTable =
      tables.find((t) => t.weekNumber === resolved.table.weekNumber) ||
      tables.find((t) => t.id === resolved.table.id) ||
      resolved.table ||
      tables.find((t) => t.id === formWeekId) ||
      tables[0];

    const matchedDayBySlashIdx = targetTable
      ? targetTable.days.findIndex(
          (d) => d.tanggal.trim() === resolved.slashDate
        )
      : -1;
    const effectiveDayIdx =
      matchedDayBySlashIdx >= 0
        ? matchedDayBySlashIdx
        : targetTable?.days[resolved.dayIdx] !== undefined
          ? resolved.dayIdx
          : targetTable?.days[formDayIdx] !== undefined
            ? formDayIdx
            : 0;
    const targetDay =
      targetTable?.days[effectiveDayIdx] || targetTable?.days[0];
    if (!targetTable || !targetDay) return null;

    const pSalesNum = parseCleanNumber(formSalesPagi);
    const pQtyNum = parseCleanNumber(formQtyPagi);
    const pTrxNum = parseCleanNumber(formTrxPagi);
    const pTrafficNum = parseCleanNumber(formTrafficPagi);

    const sSalesNum = parseCleanNumber(formSalesSiang);
    const sQtyNum = parseCleanNumber(formQtySiang);
    const sTrxNum = parseCleanNumber(formTrxSiang);
    const sTrafficNum = parseCleanNumber(formTrafficSiang);

    const baseList = tables.some((t) => t.weekNumber === targetTable.weekNumber)
      ? tables
      : [...tables, targetTable].sort((a, b) => a.weekNumber - b.weekNumber);

    const hasPagiInput =
      pSalesNum !== null ||
      pQtyNum !== null ||
      pTrxNum !== null ||
      formReasonPagi.trim() !== '';
    const hasSiangInput =
      sSalesNum !== null ||
      sQtyNum !== null ||
      sTrxNum !== null ||
      formReasonSiang.trim() !== '';

    const shouldUpdatePagi =
      formShiftType === 'Pagi' || formShiftType === 'Lembur';
    const shouldUpdateSiang =
      formShiftType === 'Siang' || formShiftType === 'Lembur';

    let updatedDayRecord: DailySalesRecord = targetDay;

    const staffForPagi =
      formShiftType === 'Lembur'
        ? (formStaffLembur || formStaffPagi)
        : formStaffPagi;
    const staffForSiang =
      formShiftType === 'Lembur'
        ? (formStaffLembur || formStaffSiang)
        : formStaffSiang;

    const nextTables = baseList.map((t) => {
      if (t.weekNumber !== targetTable.weekNumber && t.id !== targetTable.id) {
        return t;
      }
      return {
        ...t,
        days: t.days.map((d, idx) => {
          if (idx !== effectiveDayIdx) return d;
          const nextDay: DailySalesRecord = {
            ...d,
            pagi: shouldUpdatePagi
              ? {
                  nama: staffForPagi || d.pagi.nama,
                  sales: pSalesNum !== null ? pSalesNum : d.pagi.sales,
                  qty: pQtyNum !== null ? pQtyNum : d.pagi.qty,
                  trx: pTrxNum !== null ? pTrxNum : d.pagi.trx,
                  traffic:
                    pTrafficNum !== null ? pTrafficNum : (d.pagi.traffic ?? null),
                }
              : d.pagi,
            siang: shouldUpdateSiang
              ? {
                  nama: staffForSiang || d.siang.nama,
                  sales: sSalesNum !== null ? sSalesNum : d.siang.sales,
                  qty: sQtyNum !== null ? sQtyNum : d.siang.qty,
                  trx: sTrxNum !== null ? sTrxNum : d.siang.trx,
                  traffic:
                    sTrafficNum !== null ? sTrafficNum : (d.siang.traffic ?? null),
                }
              : d.siang,
            reasonPagi: shouldUpdatePagi
              ? formReasonPagi.trim() !== ''
                ? formReasonPagi
                : d.reasonPagi
              : d.reasonPagi,
            reasonSiang: shouldUpdateSiang
              ? formReasonSiang.trim() !== ''
                ? formReasonSiang
                : d.reasonSiang
              : d.reasonSiang,
          };
          updatedDayRecord = nextDay;
          return nextDay;
        }),
      };
    });

    const nowStr = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });

    const submitterName =
      formShiftType === 'Pagi'
        ? (formStaffPagi || 'Petugas Pagi')
        : formShiftType === 'Siang'
          ? (formStaffSiang || 'Petugas Siang')
          : (formStaffLembur || formStaffPagi || formStaffSiang || 'Petugas Lembur');

    const newLog: ReportSubmissionLog = {
      id:
        customLogId ||
        `log-shift-${sessionSeed}-${formCalendarIso}-${formShiftType}`,
      timestamp: nowStr,
      reportType: 'shift_sales',
      submittedBy: submitterName,
      weekTitle: targetTable.title,
      tanggal: targetDay.tanggal,
      shift: formShiftType === 'Lembur' ? 'Lembur' : formShiftType,
      salesAmount:
        formShiftType === 'Pagi'
          ? pSalesNum ?? 0
          : formShiftType === 'Siang'
            ? sSalesNum ?? 0
            : (pSalesNum ?? 0) + (sSalesNum ?? 0),
      qtyAmount:
        formShiftType === 'Pagi'
          ? pQtyNum ?? 0
          : formShiftType === 'Siang'
            ? sQtyNum ?? 0
            : (pQtyNum ?? 0) + (sQtyNum ?? 0),
      trxAmount:
        formShiftType === 'Pagi'
          ? pTrxNum ?? 0
          : formShiftType === 'Siang'
            ? sTrxNum ?? 0
            : (pTrxNum ?? 0) + (sTrxNum ?? 0),
      trafficAmount:
        formShiftType === 'Pagi'
          ? pTrafficNum ?? 0
          : formShiftType === 'Siang'
            ? sTrafficNum ?? 0
            : (pTrafficNum ?? 0) + (sTrafficNum ?? 0),
      summaryNote:
        formShiftType === 'Pagi'
          ? formReasonPagi || 'Report Shift Pagi berhasil diinput'
          : formShiftType === 'Siang'
            ? formReasonSiang || 'Report Shift Siang berhasil diinput'
            : formReasonPagi && formReasonSiang
              ? `${formReasonPagi} | ${formReasonSiang}`
              : formReasonPagi || formReasonSiang || 'Report Shift Lembur berhasil diinput',
    };

    return {
      nextTables,
      newLog,
      targetWeekId: targetTable.id,
      targetTable,
      dayIdx: effectiveDayIdx,
      updatedDay: updatedDayRecord,
      shiftPatch: {
        tanggal: targetDay.tanggal,
        shouldUpdatePagi,
        shouldUpdateSiang,
        pagi: {
          nama: staffForPagi,
          sales: pSalesNum,
          qty: pQtyNum,
          trx: pTrxNum,
          traffic: pTrafficNum,
        },
        siang: {
          nama: staffForSiang,
          sales: sSalesNum,
          qty: sQtyNum,
          trx: sTrxNum,
          traffic: sTrafficNum,
        },
        reasonPagi: formReasonPagi,
        reasonSiang: formReasonSiang,
      },
    };
  };

  // Build atomic Review Report payload (used by both Live Auto-Save and Final Submit)
  const buildReviewSubmissionPayload = (customLogId?: string) => {
    const resolved = resolveCalendarDateToWeekly(formReviewCalendarIso, false);
    const targetPeriode = `WEEKLY ${resolved.table.weekNumber} (${getWeeklyPeriodeText(resolved.table)})`;
    const targetDocId = `review-w${resolved.table.weekNumber}`;

    const applySectionsUpdate = (doc: ReviewMonitoringDoc): ReviewMonitoringDoc => ({
      ...doc,
      periode: targetPeriode,
      sections: doc.sections.map((sec) => {
        if (sec.code.includes('MERCHANDISE') && sec.subRows) {
          return {
            ...sec,
            subRows: sec.subRows.map((sr) => {
              const lbl = sr.label.trim().toLowerCase();
              if (lbl === 'divisi 1' && formMerchDiv1.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formMerchDiv1.split('\n')) };
              }
              if (lbl === 'divisi 2' && formMerchDiv2.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formMerchDiv2.split('\n')) };
              }
              if (lbl === 'divisi 3' && formMerchDiv3.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formMerchDiv3.split('\n')) };
              }
              if (lbl === 'divisi 4' && formMerchDiv4.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formMerchDiv4.split('\n')) };
              }
              if (sr.label.includes('FAST MOVING') && formFastMoving.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formFastMoving.split('\n')) };
              }
              if (sr.label.includes('SLOW MOVING') && formSlowMoving.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formSlowMoving.split('\n')) };
              }
              return sr;
            }),
          };
        }
        if (sec.code.includes('PROMO') && sec.subRows) {
          return {
            ...sec,
            subRows: sec.subRows.map((sr) => {
              const lbl = sr.label.trim().toLowerCase();
              if (lbl === 'divisi 1' && formPromoDiv1.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formPromoDiv1.split('\n')) };
              }
              if (lbl === 'divisi 2' && formPromoDiv2.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formPromoDiv2.split('\n')) };
              }
              if (lbl === 'divisi 3' && formPromoDiv3.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formPromoDiv3.split('\n')) };
              }
              if (lbl === 'divisi 4' && formPromoDiv4.trim()) {
                return { ...sr, items: normalizeReviewItemsList(formPromoDiv4.split('\n')) };
              }
              return sr;
            }),
          };
        }
        if (
          (sec.code.includes('PLATE') || sec.code.includes('PLACE')) &&
          formCustomerNote.trim()
        ) {
          return { ...sec, directItems: normalizeReviewItemsList(formCustomerNote.split('\n')) };
        }
        if (sec.code.includes('ACTION PLAN') && formActionPlan.trim()) {
          return { ...sec, directItems: normalizeReviewItemsList(formActionPlan.split('\n')) };
        }
        if (sec.code.includes('REASON') && formKendala.trim()) {
          return { ...sec, directItems: normalizeReviewItemsList(formKendala.split('\n')) };
        }
        if (sec.code.includes('NEED SUPPORT') && formNeedSupport.trim()) {
          return { ...sec, directItems: normalizeReviewItemsList(formNeedSupport.split('\n')) };
        }
        return sec;
      }),
    });

    const targetWeekNum = resolved.table.weekNumber;
    const existingDoc = reviewDocs.find(
      (d) =>
        getReviewDocWeekNumber(d) === targetWeekNum ||
        d.id === targetDocId ||
        d.periode.toUpperCase().includes(`WEEKLY ${targetWeekNum}`)
    );

    let updatedDocs: ReviewMonitoringDoc[];
    if (existingDoc) {
      updatedDocs = reviewDocs.map((doc) =>
        getReviewDocWeekNumber(doc) === targetWeekNum || doc.id === existingDoc.id
          ? applySectionsUpdate(doc)
          : doc
      );
    } else {
      const cleanBase = createReviewDocForWeek(
        targetWeekNum,
        reviewDocs[0],
        false
      );
      const createdDoc = applySectionsUpdate({
        ...cleanBase,
        id: targetDocId,
        weekNumber: targetWeekNum,
        periode: targetPeriode,
      });
      updatedDocs = [...reviewDocs, createdDoc].sort(
        (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
      );
    }

    const nowStr = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });

    const newLog: ReportSubmissionLog = {
      id:
        customLogId ||
        `log-review-${sessionSeed}-w${resolved.table.weekNumber}`,
      timestamp: nowStr,
      reportType: 'review_monitoring',
      submittedBy: 'Review Monitoring',
      weekTitle: targetPeriode,
      tanggal: resolved.slashDate,
      summaryNote:
        'Memperbarui Review Monitoring Bazar (Fast/Slow Moving, Action Plan & Need Support)',
    };

    return {
      nextReviewDocs: updatedDocs,
      newLog,
    };
  };

  // Automatic live-save while typing is disabled so input stays smooth and never jumps back;
  // data is saved when the user clicks "Kirim & Simpan Report Shift" or "Simpan Review Monitoring".

  // Handle Shift Form Submission
  const handleSubmitShiftReport = (e: React.FormEvent) => {
    e.preventDefault();
    if (shiftAutoSaveTimerRef.current) {
      window.clearTimeout(shiftAutoSaveTimerRef.current);
    }
    const payload = buildShiftSubmissionPayload();
    if (!payload) return;

    if (onAtomicSubmitReport) {
      onAtomicSubmitReport({
        ...payload,
        isAutoSave: false,
      });
    } else {
      onChangeTables(payload.nextTables);
      onAddSubmissionLog(payload.newLog);
    }
    // Reset form inputs immediately so when the user fills the link again, all fields are clean
    resetShiftFormInputs();
    setSessionSeed(String(Date.now()));
    setSubmitSuccessData(payload.newLog);
  };

  // Handle Review Form Submission
  const handleSubmitReviewReport = (e: React.FormEvent) => {
    e.preventDefault();
    if (reviewAutoSaveTimerRef.current) {
      window.clearTimeout(reviewAutoSaveTimerRef.current);
    }
    const payload = buildReviewSubmissionPayload();
    if (!payload) return;

    if (onAtomicSubmitReport) {
      onAtomicSubmitReport({
        ...payload,
        isAutoSave: false,
      });
    } else {
      onChangeReviewDocs(payload.nextReviewDocs);
      onAddSubmissionLog(payload.newLog);
    }
    clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REPORT_FORM_REVIEW);
    setSessionSeed(String(Date.now()));
    setSubmitSuccessData(payload.newLog);
  };

  // Handle Report CV Submission (Divisi 1: 6201 Graphis & 6244 Simplicity)
  const handleSubmitCvReport = (e: React.FormEvent) => {
    e.preventDefault();
    const slashTgl = isoToSlashDate(cvFormCalendarIso);
    const parts = cvFormCalendarIso.split('-');
    const monthKey = parts[1] || '11';
    const matchedWeek = ANNUAL_2026_WEEKS.find(
      (w) =>
        w.slashDates.includes(slashTgl) ||
        (cvFormCalendarIso >= w.startDateIso && cvFormCalendarIso <= w.endDateIso)
    );
    const weekNum = matchedWeek?.weekNumber || 45;
    const existingRec = cvRecords.find((r) => r.tanggal.trim() === slashTgl);

    const gSales = parseCleanNumber(cvFormGraphisSales);
    const gQty = parseCleanNumber(cvFormGraphisQty);
    const gNota = parseCleanNumber(cvFormGraphisNota);

    const sSales = parseCleanNumber(cvFormSimplicitySales);
    const sQty = parseCleanNumber(cvFormSimplicityQty);
    const sNota = parseCleanNumber(cvFormSimplicityNota);

    const includeGraphis = cvFormBrandScope === 'GRAPHIS';
    const includeSimplicity = cvFormBrandScope === 'SIMPLICITY';

    const emptyBrandShift = {
      pagi: { sales: null, qty: null, trx: null },
      siang: { sales: null, qty: null, trx: null },
    };

    const nextRec: CvDailyRecord = {
      ...(existingRec || {
        tanggal: slashTgl,
        isoDate: cvFormCalendarIso,
        dayName: getIndoDayName(slashTgl),
        weekNumber: weekNum,
        monthKey,
        graphis: emptyBrandShift,
        simplicity: emptyBrandShift,
      }),
      tanggal: slashTgl,
      isoDate: cvFormCalendarIso,
      dayName: getIndoDayName(slashTgl),
      weekNumber: existingRec?.weekNumber || weekNum,
      monthKey: existingRec?.monthKey || monthKey,
      petugasPagi: cvFormStaff.trim() || existingRec?.petugasPagi || '',
      targetDaily: existingRec?.targetDaily,
      reasonEvent:
        cvFormReasonEvent.trim() !== ''
          ? cvFormReasonEvent.trim()
          : existingRec?.reasonEvent || '',
      graphis: includeGraphis
        ? {
            pagi: { sales: gSales, qty: gQty, trx: gNota },
            siang: { sales: null, qty: null, trx: null },
          }
        : existingRec?.graphis || emptyBrandShift,
      simplicity: includeSimplicity
        ? {
            pagi: { sales: sSales, qty: sQty, trx: sNota },
            siang: { sales: null, qty: null, trx: null },
          }
        : existingRec?.simplicity || emptyBrandShift,
    };

    const finalGTot = calcCvBrandShiftTotals(nextRec.graphis);
    const finalSTot = calcCvBrandShiftTotals(nextRec.simplicity);
    const submittedTot = includeGraphis ? finalGTot : finalSTot;

    const nowStr = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });

    const brandMeta = CV_BRAND_CATALOG[cvFormBrandScope];
    const brandLabelText = `Dept ${brandMeta.deptCode} · ${brandMeta.label}`;

    const newLog: ReportSubmissionLog = {
      id: `log-cv-${Date.now()}-${cvFormCalendarIso}-${cvFormBrandScope}`,
      timestamp: nowStr,
      reportType: 'cv_report',
      submittedBy: cvFormStaff.trim() || `Petugas ${brandMeta.label}`,
      weekTitle: `WEEK ${weekNum} (${brandLabelText})`,
      tanggal: slashTgl,
      shift: `CV ${brandMeta.label}`,
      salesAmount: submittedTot.sales,
      qtyAmount: submittedTot.qty,
      trxAmount: submittedTot.trx,
      summaryNote: `Report CV ${brandLabelText}: Rp ${formatNumberId(submittedTot.sales)} (${submittedTot.qty} pcs / ${submittedTot.trx} nota)${
        nextRec.reasonEvent ? ` — ${nextRec.reasonEvent}` : ''
      }`,
    };

    if (onAtomicSubmitCvReport) {
      onAtomicSubmitCvReport({
        updatedCvRecord: nextRec,
        newLog,
      });
    } else {
      onAddSubmissionLog(newLog);
    }

    clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REPORT_FORM_CV);
    setSessionSeed(String(Date.now()));
    setSubmitSuccessData(newLog);
  };

  // Handle Data CV Submission (Inventory, Promo, Reason/Kendala & Action Plan)
  const handleSubmitDataCv = (e: React.FormEvent) => {
    e.preventDefault();
    const pType = dataCvFormPeriodType;
    const pKey =
      pType === 'weekly' ? String(dataCvFormWeekNum) : dataCvFormMonthKey;
    const matchedWeekObj = ANNUAL_2026_WEEKS.find(
      (w) => w.weekNumber === dataCvFormWeekNum
    );

    const buildOpsEntry = (
      brand: CvInventoryBrandKey,
      draft: DataCvBrandDraft
    ): CvInventoryOpsEntry => {
      const id = getCvOpsEntryId(pType, pKey, brand);
      const existing =
        cvInventoryOps.find((x) => x.id === id) ||
        createDefaultCvOpsEntry(pType, pKey, brand);
      return {
        ...existing,
        divisi: draft.divisi.trim() || CV_DIVISI_DEFAULT,
        stockAwal: Math.max(0, Number(draft.stockAwal) || 0),
        barangDatang: Math.max(0, Number(draft.barangDatang) || 0),
        barangKeluar: Math.max(0, Number(draft.barangKeluar) || 0),
        salesQtyOverride:
          draft.salesQtyMode === 'manual'
            ? Math.max(0, Number(draft.salesQtyManual) || 0)
            : null,
        promo: draft.promo.trim(),
        reasonKendala: draft.reasonKendala.trim(),
        actionPlan: draft.actionPlan.trim(),
        updatedAt: new Date().toISOString(),
      };
    };

    const entriesToSave: CvInventoryOpsEntry[] = [];
    const summaryParts: string[] = [];
    let totalStockAkhir = 0;
    let totalSalesQty = 0;

    const processBrand = (
      brand: CvInventoryBrandKey,
      draft: DataCvBrandDraft
    ) => {
      const entry = buildOpsEntry(brand, draft);
      entriesToSave.push(entry);
      const liveQty =
        pType === 'weekly'
          ? liveCvSalesMap.weekly.get(dataCvFormWeekNum)?.[brand] || 0
          : liveCvSalesMap.monthly.get(dataCvFormMonthKey)?.[brand] || 0;
      const effSalesQty =
        entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined
          ? entry.salesQtyOverride
          : liveQty;
      const stockAkhir =
        entry.stockAwal + entry.barangDatang - effSalesQty - entry.barangKeluar;
      totalStockAkhir += stockAkhir;
      totalSalesQty += effSalesQty;
      summaryParts.push(
        `${CV_BRAND_CATALOG[brand].label}: Awal ${entry.stockAwal}, Dtg +${entry.barangDatang}, Sls -${effSalesQty}, Klr -${entry.barangKeluar} => Akhir ${stockAkhir} pcs`
      );
    };

    if (dataCvFormBrandScope === 'GRAPHIS') {
      processBrand('GRAPHIS', dataCvGraphisDraft);
    } else {
      processBrand('SIMPLICITY', dataCvSimplicityDraft);
    }

    const nowStr = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });

    const dataCvBrandMeta = CV_BRAND_CATALOG[dataCvFormBrandScope];
    const periodTitle =
      pType === 'weekly'
        ? `Week ${dataCvFormWeekNum} (${matchedWeekObj?.periodeText || ''}) · ${dataCvBrandMeta.label}`
        : `${MONTH_NAMES_ID[dataCvFormMonthKey] || dataCvFormMonthKey} 2026 · ${dataCvBrandMeta.label}`;

    const newLog: ReportSubmissionLog = {
      id: `log-datacv-${Date.now()}-${pType}-${pKey}-${dataCvFormBrandScope}`,
      timestamp: nowStr,
      reportType: 'data_cv',
      submittedBy: dataCvFormStaff.trim() || `Petugas ${dataCvBrandMeta.label}`,
      weekTitle: periodTitle,
      tanggal:
        pType === 'weekly'
          ? matchedWeekObj?.slashDates[0] || '-'
          : `01/${dataCvFormMonthKey}/2026`,
      shift:
        pType === 'weekly'
          ? `Data CV Weekly (${dataCvBrandMeta.label})`
          : `Data CV Monthly (${dataCvBrandMeta.label})`,
      salesAmount: totalStockAkhir,
      qtyAmount: totalSalesQty,
      trxAmount: entriesToSave.length,
      summaryNote: summaryParts.join(' | '),
    };

    if (onAtomicSubmitCvReport) {
      onAtomicSubmitCvReport({
        updatedCvOpsEntries: entriesToSave,
        newLog,
      });
    } else {
      onAddSubmissionLog(newLog);
    }

    clearAutoDraftFromStorage(AUTO_DRAFT_KEYS.REPORT_FORM_DATA_CV);
    setSessionSeed(String(Date.now()));
    setSubmitSuccessData(newLog);
  };

  // RENDER 1: SHIFT REPORT FILLING FORM (Standalone / Petugas View)
  if (activeView === 'shift_form') {
    const resolvedFormWeek = resolveCalendarDateToWeekly(formCalendarIso, false);
    const activeFormTable =
      tables.find((t) => t.weekNumber === resolvedFormWeek.table.weekNumber) ||
      tables.find((t) => t.id === resolvedFormWeek.table.id) ||
      resolvedFormWeek.table ||
      tables.find((t) => t.id === formWeekId) ||
      tables[0];
    const matchedActiveDayIdx = activeFormTable
      ? activeFormTable.days.findIndex(
          (d) => d.tanggal.trim() === resolvedFormWeek.slashDate
        )
      : -1;
    const effectiveActiveDayIdx =
      matchedActiveDayIdx >= 0
        ? matchedActiveDayIdx
        : activeFormTable?.days[resolvedFormWeek.dayIdx] !== undefined
          ? resolvedFormWeek.dayIdx
          : formDayIdx;
    const activeFormDay =
      activeFormTable?.days[effectiveActiveDayIdx] || activeFormTable?.days[0];

    const previewSales =
      (formShiftType !== 'Siang' ? (parseCleanNumber(formSalesPagi) ?? 0) : 0) +
      (formShiftType !== 'Pagi' ? (parseCleanNumber(formSalesSiang) ?? 0) : 0);
    const previewQty =
      (formShiftType !== 'Siang' ? (parseCleanNumber(formQtyPagi) ?? 0) : 0) +
      (formShiftType !== 'Pagi' ? (parseCleanNumber(formQtySiang) ?? 0) : 0);
    const previewTrx =
      (formShiftType !== 'Siang' ? (parseCleanNumber(formTrxPagi) ?? 0) : 0) +
      (formShiftType !== 'Pagi' ? (parseCleanNumber(formTrxSiang) ?? 0) : 0);
    const previewTraffic =
      (formShiftType !== 'Siang' ? (parseCleanNumber(formTrafficPagi) ?? 0) : 0) +
      (formShiftType !== 'Pagi' ? (parseCleanNumber(formTrafficSiang) ?? 0) : 0);
    const previewPP = previewQty > 0 ? Math.round(previewSales / previewQty) : 0;
    const previewBS = previewTrx > 0 ? Math.round(previewSales / previewTrx) : 0;
    const previewConvRate = previewTraffic > 0 ? (previewTrx / previewTraffic) * 100 : 0;

    return (
      <div className="max-w-2xl mx-auto space-y-6">
        {/* Top Navigation Inside Report Menu Only (Never opens other menus) */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 bg-white border border-slate-200 rounded-xl p-3 shadow-2xs">
          <div className="flex flex-wrap items-center gap-2">
            {initialFormMode === 'none' && (
              <button
                type="button"
                onClick={() => {
                  setSubmitSuccessData(null);
                  setActiveView('hub');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-700"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                Daftar Link Report
              </button>
            )}
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-amber-50 border border-amber-200 text-amber-900 text-xs font-bold">
              <Lock className="w-3.5 h-3.5 text-amber-600" />
              Menu Pengisian Report
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-1.5 self-end sm:self-auto">
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('shift_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 text-white shadow-xs cursor-pointer"
            >
              Report DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('review_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Data DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('cv_report_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-purple-50 hover:bg-purple-100 text-purple-800 border border-purple-200 cursor-pointer"
            >
              Report CV
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('data_cv_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 cursor-pointer"
            >
              Data CV
            </button>
          </div>
        </div>

        {submitSuccessData ? (
          <div className="bg-white border border-emerald-200 rounded-xl p-6 sm:p-8 text-center space-y-5 shadow-xs">
            <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto">
              <CheckCircle2 className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">
                Report DP Berhasil Dikirim &amp; Tersimpan!
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Data penjualan telah langsung masuk ke tabel <strong>{submitSuccessData.weekTitle}</strong> pada tanggal <strong>{submitSuccessData.tanggal}</strong>.
              </p>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 text-left text-xs space-y-2 font-mono-num">
              <div className="flex justify-between">
                <span className="text-slate-500">Petugas / Shift:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.submittedBy} ({submitSuccessData.shift})
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">Tanggal / Minggu:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.tanggal} · {submitSuccessData.weekTitle}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">Sales / Qty / Trx:</span>
                <span className="font-bold text-emerald-700">
                  Rp {formatNumberId(submitSuccessData.salesAmount || 0)} ·{' '}
                  {submitSuccessData.qtyAmount} pcs · {submitSuccessData.trxAmount} trx
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">Traffic / Est. Conv. Rate:</span>
                <span className="font-bold text-indigo-700">
                  {formatNumberId(submitSuccessData.trafficAmount || 0) || '0'} pengunjung ·{' '}
                  {(submitSuccessData.trafficAmount || 0) > 0
                    ? `${(((submitSuccessData.trxAmount || 0) / (submitSuccessData.trafficAmount || 1)) * 100).toFixed(1)}%`
                    : '—'}
                </span>
              </div>
              <div className="pt-2 border-t border-slate-200 text-slate-700 font-sans">
                <strong>Reason / Event:</strong> {submitSuccessData.summaryNote}
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => {
                  resetShiftFormInputs();
                  setSubmitSuccessData(null);
                }}
                className="px-5 py-2.5 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl shadow-xs"
              >
                Isi Report Lagi (Form Otomatis Direset)
              </button>
              {initialFormMode === 'none' && (
                <button
                  type="button"
                  onClick={() => {
                    setSubmitSuccessData(null);
                    setActiveView('hub');
                  }}
                  className="px-4 py-2.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl"
                >
                  Kembali ke Menu Link Pengisian Report
                </button>
              )}
            </div>
          </div>
        ) : (
          <>
          <AutoDraftStatusBadge
            variant="banner"
            menuLabel="Form Report DP"
            lastDraftTimeStr={shiftDraftTimeStr}
            isDirty={hasShiftDraftContent}
            isOnline={isShiftOnline}
            hasRestorableDraft={hasShiftDraftContent}
            onClearDraft={() => resetShiftFormInputs(true)}
          />
          <form
            onSubmit={handleSubmitShiftReport}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
          >
            {/* Form Header */}
            <div className="bg-slate-900 text-white px-4 sm:px-6 py-4 sm:py-5 border-b border-slate-800">
              <div className="text-[11px] sm:text-xs text-amber-400 font-extrabold tracking-wide uppercase">
                REPORT BAZAR MATAHARI · Tanjung
              </div>
              <h2 className="text-base sm:text-lg font-bold mt-0.5 text-white">
                Pengisian Report DP (Penjualan Shift Harian)
              </h2>
              <p className="text-xs text-slate-300 mt-0.5">
                Isi data penjualan Shift Pagi, Shift Siang, atau Shift Lembur beserta catatan evaluasi harian
              </p>
            </div>

            <div className="p-6 space-y-5">
              {/* Step 1: Pilih Tanggal (Filter Kalender) -> Weekly Otomatis */}
              <div className="p-4 bg-indigo-50/70 border border-indigo-200 rounded-xl space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2">
                    <Calendar className="w-4 h-4 text-indigo-600" />
                    <span className="text-xs font-bold text-slate-900">
                      Filter Kalender Tanggal — Weekly Terpilih Otomatis
                    </span>
                  </div>
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-indigo-600 text-white">
                    <Sparkles className="w-3 h-3" />
                    Otomatis: {activeFormTable?.title} (W{activeFormTable?.weekNumber})
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  {/* 1. Filter Kalender Tanggal */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      1. Pilih Tanggal (Kalender)
                    </label>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          handleSelectFormCalendarDate(shiftIsoDate(formCalendarIso, -1))
                        }
                        title="Hari Sebelumnya"
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                      </button>
                      <input
                        type="date"
                        value={formCalendarIso}
                        onChange={(e) => handleSelectFormCalendarDate(e.target.value)}
                        className="w-full px-2.5 py-1.5 text-xs font-bold font-mono-num bg-white border border-indigo-300 rounded-lg text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer"
                      />
                      <button
                        type="button"
                        onClick={() =>
                          handleSelectFormCalendarDate(shiftIsoDate(formCalendarIso, 1))
                        }
                        title="Hari Berikutnya"
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                      >
                        <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                    </div>
                    <div className="text-[11px] text-indigo-900 font-medium mt-1 font-mono-num">
                      {getIndoDayName(activeFormDay?.tanggal || '')},{' '}
                      <strong>{activeFormDay?.tanggal}</strong> (Hari ke-{formDayIdx + 1})
                    </div>
                  </div>

                  {/* 2. Weekly Otomatis */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      2. Minggu / Weekly (Otomatis)
                    </label>
                    <div className="px-3 py-2 bg-white border border-indigo-200 rounded-lg">
                      <div className="text-xs font-extrabold text-indigo-950 flex items-center justify-between">
                        <span>{activeFormTable?.title}</span>
                        <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold">
                          Auto
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-500 truncate mt-0.5">
                        {getWeeklyPeriodeText(activeFormTable)}
                      </div>
                    </div>
                  </div>

                  {/* 3. Pilih Shift */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      3. Pilih Shift
                    </label>
                    <div className="grid grid-cols-3 gap-1 p-1 bg-white border border-slate-200 rounded-lg">
                      {(['Pagi', 'Siang', 'Lembur'] as const).map((sh) => (
                        <button
                          key={sh}
                          type="button"
                          onClick={() => setFormShiftType(sh)}
                          className={`py-1.5 text-xs font-bold rounded-md transition-colors ${
                            formShiftType === sh
                              ? 'bg-indigo-600 text-white shadow-xs'
                              : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                          }`}
                        >
                          {sh === 'Pagi' ? 'Shift Pagi' : sh === 'Siang' ? 'Shift Siang' : 'Shift Lembur'}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </div>

              {/* Step 2: Input Shift Pagi */}
              {formShiftType === 'Pagi' && (
                <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2.5">
                    <div className="flex items-center gap-2">
                      <Sun className="w-4 h-4 text-amber-500" />
                      <span className="text-xs font-bold text-slate-900">
                        DATA SHIFT PAGI ({activeFormDay?.tanggal})
                      </span>
                    </div>
                    <UserCheck className="w-4 h-4 text-slate-400" />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Nama Petugas Shift Pagi
                    </label>
                    <input
                      type="text"
                      required
                      value={formStaffPagi}
                      onChange={(e) => setFormStaffPagi(e.target.value)}
                      placeholder="Ketik nama petugas..."
                      className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Sales Shift Pagi (Rp)
                      </label>
                      <div className="relative flex items-center">
                        <span className="absolute left-3 text-xs font-bold text-slate-400 pointer-events-none select-none">
                          Rp
                        </span>
                        <input
                          type="text"
                          inputMode="numeric"
                          required
                          value={formSalesPagi}
                          onChange={(e) =>
                            setFormSalesPagi(formatAutoNominal(e.target.value))
                          }
                          placeholder="Contoh: 705.900"
                          className="w-full pl-8 pr-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Qty Terjual (Pcs)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        required
                        value={formQtyPagi}
                        onChange={(e) =>
                          setFormQtyPagi(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 20"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Jumlah Transaksi (Trx)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        required
                        value={formTrxPagi}
                        onChange={(e) =>
                          setFormTrxPagi(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 7"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-indigo-900 mb-1">
                        Traffic Pengunjung (Org)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={formTrafficPagi}
                        onChange={(e) =>
                          setFormTrafficPagi(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 35"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-indigo-50/40 border border-indigo-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Reason / Event Shift Pagi
                    </label>
                    <textarea
                      rows={3}
                      value={formReasonPagi}
                      onChange={(e) => setFormReasonPagi(e.target.value)}
                      placeholder="Tuliskan situasi pengunjung, produk yang dominan laku, atau kendala POP pada shift pagi..."
                      className="w-full p-3 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                  </div>
                </div>
              )}

              {/* Step 3: Input Shift Siang */}
              {formShiftType === 'Siang' && (
                <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-4">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2.5">
                    <div className="flex items-center gap-2">
                      <Sunset className="w-4 h-4 text-indigo-600" />
                      <span className="text-xs font-bold text-slate-900">
                        DATA SHIFT SIANG ({activeFormDay?.tanggal})
                      </span>
                    </div>
                    <UserCheck className="w-4 h-4 text-slate-400" />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Nama Petugas Shift Siang
                    </label>
                    <input
                      type="text"
                      required
                      value={formStaffSiang}
                      onChange={(e) => setFormStaffSiang(e.target.value)}
                      placeholder="Ketik nama petugas..."
                      className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Sales Shift Siang (Rp)
                      </label>
                      <div className="relative flex items-center">
                        <span className="absolute left-3 text-xs font-bold text-slate-400 pointer-events-none select-none">
                          Rp
                        </span>
                        <input
                          type="text"
                          inputMode="numeric"
                          required
                          value={formSalesSiang}
                          onChange={(e) =>
                            setFormSalesSiang(formatAutoNominal(e.target.value))
                          }
                          placeholder="Contoh: 2.339.100"
                          className="w-full pl-8 pr-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Qty Terjual (Pcs)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        required
                        value={formQtySiang}
                        onChange={(e) =>
                          setFormQtySiang(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 27"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Jumlah Transaksi (Trx)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        required
                        value={formTrxSiang}
                        onChange={(e) =>
                          setFormTrxSiang(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 7"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-indigo-900 mb-1">
                        Traffic Pengunjung (Org)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={formTrafficSiang}
                        onChange={(e) =>
                          setFormTrafficSiang(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 60"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-indigo-50/40 border border-indigo-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Reason / Event Shift Siang
                    </label>
                    <textarea
                      rows={3}
                      value={formReasonSiang}
                      onChange={(e) => setFormReasonSiang(e.target.value)}
                      placeholder="Tuliskan situasi pengunjung sore/malam, event sekitar, atau kendala broken size..."
                      className="w-full p-3 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                  </div>
                </div>
              )}

              {/* Step 4: Input Shift Lembur */}
              {formShiftType === 'Lembur' && (
                <div className="p-4 bg-amber-50/60 border border-amber-200 rounded-xl space-y-4">
                  <div className="flex items-center justify-between border-b border-amber-200 pb-2.5">
                    <div className="flex items-center gap-2">
                      <Clock className="w-4 h-4 text-amber-700" />
                      <span className="text-xs font-bold text-slate-900">
                        DATA SHIFT LEMBUR ({activeFormDay?.tanggal})
                      </span>
                    </div>
                    <span className="text-[11px] font-semibold text-amber-900 bg-amber-200/80 px-2 py-0.5 rounded-full">
                      Pagi & Siang (Full Day)
                    </span>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Nama Petugas Shift Lembur
                    </label>
                    <input
                      type="text"
                      required
                      value={formStaffLembur || formStaffPagi || formStaffSiang}
                      onChange={(e) => {
                        const val = e.target.value;
                        setFormStaffLembur(val);
                        setFormStaffPagi(val);
                        setFormStaffSiang(val);
                      }}
                      placeholder="Ketik nama petugas lembur..."
                      className="w-full px-3 py-2 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                    {/* Sesi Pagi */}
                    <div className="p-3.5 bg-white border border-slate-200 rounded-lg space-y-3">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800 border-b border-slate-100 pb-2">
                        <Sun className="w-3.5 h-3.5 text-amber-500" />
                        <span>Sesi Pagi (10.00 – 16.00)</span>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Sales (Rp)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formSalesPagi}
                            onChange={(e) =>
                              setFormSalesPagi(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Qty (Pcs)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formQtyPagi}
                            onChange={(e) =>
                              setFormQtyPagi(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Trx
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formTrxPagi}
                            onChange={(e) =>
                              setFormTrxPagi(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-indigo-800 mb-1">
                            Traffic (Org)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formTrafficPagi}
                            onChange={(e) =>
                              setFormTrafficPagi(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num bg-indigo-50/40 border border-indigo-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                      </div>
                      <div>
                        <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                          Reason / Catatan Sesi Pagi
                        </label>
                        <textarea
                          rows={2}
                          value={formReasonPagi}
                          onChange={(e) => setFormReasonPagi(e.target.value)}
                          placeholder="Catatan sesi pagi..."
                          className="w-full p-2 text-xs border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                        />
                      </div>
                    </div>

                    {/* Sesi Siang */}
                    <div className="p-3.5 bg-white border border-slate-200 rounded-lg space-y-3">
                      <div className="flex items-center gap-1.5 text-xs font-bold text-slate-800 border-b border-slate-100 pb-2">
                        <Sunset className="w-3.5 h-3.5 text-indigo-600" />
                        <span>Sesi Siang (16.00 – Selesai)</span>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Sales (Rp)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formSalesSiang}
                            onChange={(e) =>
                              setFormSalesSiang(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Qty (Pcs)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formQtySiang}
                            onChange={(e) =>
                              setFormQtySiang(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                            Trx
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formTrxSiang}
                            onChange={(e) =>
                              setFormTrxSiang(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-indigo-800 mb-1">
                            Traffic (Org)
                          </label>
                          <input
                            type="text"
                            inputMode="numeric"
                            value={formTrafficSiang}
                            onChange={(e) =>
                              setFormTrafficSiang(formatAutoNominal(e.target.value))
                            }
                            placeholder="0"
                            className="w-full px-2 py-1.5 text-xs font-bold font-mono-num bg-indigo-50/40 border border-indigo-300 rounded focus:outline-none focus:border-indigo-600"
                          />
                        </div>
                      </div>
                      <div>
                        <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                          Reason / Catatan Sesi Siang
                        </label>
                        <textarea
                          rows={2}
                          value={formReasonSiang}
                          onChange={(e) => setFormReasonSiang(e.target.value)}
                          placeholder="Catatan sesi siang / malam..."
                          className="w-full p-2 text-xs border border-slate-300 rounded focus:outline-none focus:border-indigo-600"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Live Formula Calculation Preview */}
              <div className="p-4 bg-amber-50/70 border border-amber-200 rounded-xl flex flex-wrap items-center justify-between gap-4 text-xs">
                <div className="flex items-center gap-2 font-bold text-slate-900">
                  <Calculator className="w-4 h-4 text-amber-700" />
                  <span>Pratinjau Kalkulasi Otomatis:</span>
                </div>
                <div className="flex flex-wrap items-center gap-3 font-mono-num">
                  <span>
                    Sales: <strong>Rp {formatNumberId(previewSales) || '0'}</strong>
                  </span>
                  <span>·</span>
                  <span>
                    Price Point: <strong>Rp {formatNumberId(previewPP) || '0'}</strong>
                  </span>
                  <span>·</span>
                  <span>
                    Big Size: <strong>Rp {formatNumberId(previewBS) || '0'}</strong>
                  </span>
                  <span>·</span>
                  <span>
                    Traffic: <strong>{formatNumberId(previewTraffic) || '0'} org</strong>
                  </span>
                  <span>·</span>
                  <span className="text-indigo-800">
                    Est. Conv. Rate:{' '}
                    <strong>
                      {previewTraffic > 0 ? `${previewConvRate.toFixed(1)}%` : '—'}
                    </strong>
                  </span>
                </div>
              </div>

              <div className="pt-2 flex flex-wrap items-center justify-between gap-3">
                <div className="flex flex-wrap items-center gap-2">
                  <AutoDraftStatusBadge
                    menuLabel="Form Report DP"
                    lastDraftTimeStr={shiftDraftTimeStr}
                    isDirty={hasShiftDraftContent}
                    isOnline={isShiftOnline}
                  />
                </div>
                <button
                  type="submit"
                  className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 text-xs font-bold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl transition-colors shadow-xs"
                >
                  <Send className="w-4 h-4" />
                  Kirim &amp; Simpan Report DP
                </button>
              </div>
            </div>
          </form>
          </>
        )}
      </div>
    );
  }

  // RENDER 2: REVIEW MONITORING FILLING FORM
  if (activeView === 'review_form') {
    const resolvedReviewWeek = resolveCalendarDateToWeekly(
      formReviewCalendarIso,
      false
    );
    const activeReviewTable = resolvedReviewWeek.table;
    const activeReviewDay =
      activeReviewTable?.days[resolvedReviewWeek.dayIdx] ||
      activeReviewTable?.days[0];
    const autoReviewPeriodeText = getWeeklyPeriodeText(activeReviewTable);

    return (
      <div className="max-w-2xl mx-auto space-y-6">
        <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-200 rounded-xl p-3 shadow-2xs">
          <div className="flex items-center gap-2">
            {initialFormMode === 'none' && (
              <button
                type="button"
                onClick={() => {
                  setSubmitSuccessData(null);
                  setActiveView('hub');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-700"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                Daftar Link Report
              </button>
            )}
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-xs font-bold">
              <Lock className="w-3.5 h-3.5 text-emerald-600" />
              Khusus Menu Pengisian Report (Menu Lain Dikunci)
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('shift_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Report DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('review_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-emerald-700 text-white cursor-pointer"
            >
              Data DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('cv_report_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-purple-50 hover:bg-purple-100 text-purple-800 border border-purple-200 cursor-pointer"
            >
              Report CV
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('data_cv_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 cursor-pointer"
            >
              Data CV
            </button>
          </div>
        </div>

        {submitSuccessData ? (
          <div className="bg-white border border-emerald-200 rounded-xl p-6 sm:p-8 text-center space-y-4">
            <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto">
              <CheckCircle2 className="w-6 h-6" />
            </div>
            <h2 className="text-lg font-bold text-slate-900">
              Data DP (Review Monitoring Bazar) Berhasil Diperbarui!
            </h2>
            <p className="text-xs text-slate-500">
              Laporan evaluasi periode <strong>{submitSuccessData.weekTitle}</strong> telah tersimpan ke lembar kerja Data DP / Review Monitoring Bazar Tanjung.
            </p>
            <div className="flex flex-wrap items-center justify-center gap-2">
              <button
                type="button"
                onClick={() => {
                  const wk = resolveCalendarDateToWeekly(formReviewCalendarIso, false).table.weekNumber;
                  loadReviewIntoFormByWeekNumber(wk);
                  setSubmitSuccessData(null);
                }}
                className="px-5 py-2.5 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-lg"
              >
                Lihat / Edit Data DP Lagi
              </button>
              {initialFormMode === 'none' && (
                <button
                  type="button"
                  onClick={() => {
                    setSubmitSuccessData(null);
                    setActiveView('hub');
                  }}
                  className="px-4 py-2.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg"
                >
                  Kembali ke Menu Link Pengisian Report
                </button>
              )}
            </div>
          </div>
        ) : (
          <>
          <AutoDraftStatusBadge
            variant="banner"
            menuLabel="Form Data DP"
            lastDraftTimeStr={reviewDraftTimeStr}
            isDirty={hasReviewDraftContent}
            isOnline={isReviewOnline}
            hasRestorableDraft={hasReviewDraftContent}
            onClearDraft={() => resetReviewFormInputs(true)}
          />
          <form
            onSubmit={handleSubmitReviewReport}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
          >
            <div className="bg-emerald-900 text-white px-6 py-5">
              <div className="text-xs text-emerald-300 font-bold">
                REPORT BAZAR MATAHARI · Tanjung · PT MDS Retailing Tbk
              </div>
              <h2 className="text-lg font-bold mt-0.5">
                Pengisian Data DP (Review Monitoring Bazar)
              </h2>
              <p className="text-xs text-emerald-100 mt-1">
                Perbarui daftar produk Fast Moving, Slow Moving, Action Plan, Kendala & Need Support
              </p>
            </div>

            <div className="p-6 space-y-4">
              <div className="p-4 bg-emerald-50/70 border border-emerald-200 rounded-xl space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-1.5 text-xs font-bold text-emerald-950">
                    <Calendar className="w-4 h-4 text-emerald-600" />
                    <span>Filter Kalender Tanggal — Periode Weekly Otomatis</span>
                  </div>
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-700 text-white">
                    <Sparkles className="w-3 h-3" />
                    Otomatis: WEEKLY {activeReviewTable?.weekNumber} (W{activeReviewTable?.weekNumber})
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {/* 1. Pilih Tanggal (Kalender) */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      1. Pilih Tanggal (Kalender)
                    </label>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          handleSelectReviewCalendarDate(
                            shiftIsoDate(formReviewCalendarIso, -1)
                          )
                        }
                        title="Hari Sebelumnya"
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                      </button>
                      <input
                        type="date"
                        value={formReviewCalendarIso}
                        onChange={(e) =>
                          handleSelectReviewCalendarDate(e.target.value)
                        }
                        className="w-full px-2.5 py-1.5 text-xs font-bold font-mono-num bg-white border border-emerald-300 rounded-lg text-slate-900 focus:outline-none focus:border-emerald-600 cursor-pointer"
                      />
                      <button
                        type="button"
                        onClick={() =>
                          handleSelectReviewCalendarDate(
                            shiftIsoDate(formReviewCalendarIso, 1)
                          )
                        }
                        title="Hari Berikutnya"
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                      >
                        <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                    </div>
                    <div className="text-[11px] text-emerald-900 font-medium mt-1 font-mono-num">
                      {getIndoDayName(activeReviewDay?.tanggal || '')},{' '}
                      <strong>{activeReviewDay?.tanggal}</strong> (Hari ke-{resolvedReviewWeek.dayIdx + 1})
                    </div>
                  </div>

                  {/* 2. Periode Review Weekly (Otomatis) */}
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                      2. Periode Review Weekly (Otomatis)
                    </label>
                    <div className="px-3 py-2 bg-white border border-emerald-200 rounded-lg">
                      <div className="text-xs font-extrabold text-emerald-950 flex items-center justify-between">
                        <span>WEEKLY {activeReviewTable?.weekNumber}</span>
                        <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold">
                          Auto
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-500 truncate mt-0.5">
                        {autoReviewPeriodeText}
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* SECTION A: MERCHANDISE (Divisi 1 - 4, Fast Moving & Slow Moving) */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-3.5">
                <div className="text-xs font-extrabold text-slate-900 border-b border-slate-200 pb-2">
                  A. Merchandise
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 1 :
                    </label>
                    <textarea
                      rows={2}
                      value={formMerchDiv1}
                      onChange={(e) => setFormMerchDiv1(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 2 :
                    </label>
                    <textarea
                      rows={2}
                      value={formMerchDiv2}
                      onChange={(e) => setFormMerchDiv2(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 3 :
                    </label>
                    <textarea
                      rows={2}
                      value={formMerchDiv3}
                      onChange={(e) => setFormMerchDiv3(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 4 :
                    </label>
                    <textarea
                      rows={2}
                      value={formMerchDiv4}
                      onChange={(e) => setFormMerchDiv4(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Produk Fast Moving (Satu item per baris)
                    </label>
                    <textarea
                      rows={4}
                      value={formFastMoving}
                      onChange={(e) => setFormFastMoving(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Produk Slow Moving (Satu item per baris)
                    </label>
                    <textarea
                      rows={4}
                      value={formSlowMoving}
                      onChange={(e) => setFormSlowMoving(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                </div>
              </div>

              {/* SECTION B: PROMO (Divisi 1 - 4) */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl space-y-3.5">
                <div className="text-xs font-extrabold text-slate-900 border-b border-slate-200 pb-2">
                  B. Promo
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 1 :
                    </label>
                    <textarea
                      rows={2}
                      value={formPromoDiv1}
                      onChange={(e) => setFormPromoDiv1(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 2 :
                    </label>
                    <textarea
                      rows={2}
                      value={formPromoDiv2}
                      onChange={(e) => setFormPromoDiv2(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 3 :
                    </label>
                    <textarea
                      rows={2}
                      value={formPromoDiv3}
                      onChange={(e) => setFormPromoDiv3(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      - Divisi 4 :
                    </label>
                    <textarea
                      rows={2}
                      value={formPromoDiv4}
                      onChange={(e) => setFormPromoDiv4(e.target.value)}
                      className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-emerald-600"
                    />
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    C. Place / People (Temuan Customer)
                  </label>
                  <textarea
                    rows={3}
                    value={formCustomerNote}
                    onChange={(e) => setFormCustomerNote(e.target.value)}
                    className="w-full p-3 text-xs bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    D. Action Plan
                  </label>
                  <textarea
                    rows={3}
                    value={formActionPlan}
                    onChange={(e) => setFormActionPlan(e.target.value)}
                    className="w-full p-3 text-xs bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Reason / Kendala
                  </label>
                  <textarea
                    rows={3}
                    value={formKendala}
                    onChange={(e) => setFormKendala(e.target.value)}
                    className="w-full p-3 text-xs bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Need Support Barang
                  </label>
                  <textarea
                    rows={3}
                    value={formNeedSupport}
                    onChange={(e) => setFormNeedSupport(e.target.value)}
                    className="w-full p-3 text-xs bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              <div className="pt-2 flex flex-wrap items-center justify-between gap-3">
                <div className="flex flex-wrap items-center gap-2">
                  <AutoDraftStatusBadge
                    menuLabel="Form Data DP"
                    lastDraftTimeStr={reviewDraftTimeStr}
                    isDirty={hasReviewDraftContent}
                    isOnline={isReviewOnline}
                  />
                </div>
                <button
                  type="submit"
                  className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 text-xs font-bold text-white bg-emerald-700 hover:bg-emerald-800 rounded-xl transition-colors"
                >
                  <Send className="w-4 h-4" />
                  Simpan Data DP
                </button>
              </div>
            </div>
          </form>
          </>
        )}
      </div>
    );
  }

  // RENDER 2B: REPORT CV FILLING FORM (Divisi 1: 6201 Graphis & 6244 Simplicity)
  if (activeView === 'cv_report_form') {
    const cvSlashDate = isoToSlashDate(cvFormCalendarIso);
    const cvParts = cvFormCalendarIso.split('-');
    const cvMonthKey = cvParts[1] || '11';
    const cvMonthName = MONTH_NAMES_ID[cvMonthKey] || 'November';
    const cvMatchedWeek = ANNUAL_2026_WEEKS.find(
      (w) =>
        w.slashDates.includes(cvSlashDate) ||
        (cvFormCalendarIso >= w.startDateIso && cvFormCalendarIso <= w.endDateIso)
    );

    const gSalesPreview =
      cvFormBrandScope === 'GRAPHIS'
        ? parseCleanNumber(cvFormGraphisSales) ?? 0
        : 0;
    const gQtyPreview =
      cvFormBrandScope === 'GRAPHIS'
        ? parseCleanNumber(cvFormGraphisQty) ?? 0
        : 0;
    const gNotaPreview =
      cvFormBrandScope === 'GRAPHIS'
        ? parseCleanNumber(cvFormGraphisNota) ?? 0
        : 0;

    const sSalesPreview =
      cvFormBrandScope === 'SIMPLICITY'
        ? parseCleanNumber(cvFormSimplicitySales) ?? 0
        : 0;
    const sQtyPreview =
      cvFormBrandScope === 'SIMPLICITY'
        ? parseCleanNumber(cvFormSimplicityQty) ?? 0
        : 0;
    const sNotaPreview =
      cvFormBrandScope === 'SIMPLICITY'
        ? parseCleanNumber(cvFormSimplicityNota) ?? 0
        : 0;

    const totalCvSalesPreview =
      cvFormBrandScope === 'GRAPHIS' ? gSalesPreview : sSalesPreview;
    const totalCvQtyPreview =
      cvFormBrandScope === 'GRAPHIS' ? gQtyPreview : sQtyPreview;
    const totalCvNotaPreview =
      cvFormBrandScope === 'GRAPHIS' ? gNotaPreview : sNotaPreview;
    const cvPricePointPreview =
      totalCvQtyPreview > 0
        ? Math.round(totalCvSalesPreview / totalCvQtyPreview)
        : 0;

    return (
      <div className="max-w-2xl mx-auto space-y-6">
        {/* Top Switcher Bar */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 bg-white border border-slate-200 rounded-xl p-3 shadow-2xs">
          <div className="flex flex-wrap items-center gap-2">
            {initialFormMode === 'none' && (
              <button
                type="button"
                onClick={() => {
                  setSubmitSuccessData(null);
                  setActiveView('hub');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-700 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                Daftar Link Report
              </button>
            )}
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-purple-50 border border-purple-200 text-purple-900 text-xs font-bold">
              <Lock className="w-3.5 h-3.5 text-purple-700" />
              Form Pengisian Report CV
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('shift_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Report DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('review_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Data DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('cv_report_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-purple-700 text-white cursor-pointer"
            >
              Report CV
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('data_cv_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 cursor-pointer"
            >
              Data CV
            </button>
          </div>
        </div>

        {submitSuccessData && submitSuccessData.reportType === 'cv_report' ? (
          <div className="bg-white border border-emerald-200 rounded-xl p-6 sm:p-8 text-center space-y-5 shadow-xs">
            <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto">
              <CheckCircle2 className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">
                Report CV Divisi 1 Berhasil Tersimpan!
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Data penjualan <strong>{submitSuccessData.weekTitle}</strong> tanggal{' '}
                <strong>{submitSuccessData.tanggal}</strong> otomatis masuk ke Report CV &amp; Total DP+CV.
              </p>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 text-left text-xs space-y-2 font-mono-num">
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Petugas:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.submittedBy}
                </span>
              </div>
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Tanggal / Periode:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.tanggal} · {submitSuccessData.weekTitle}
                </span>
              </div>
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Total CV (Sales / Qty / Nota):</span>
                <span className="font-bold text-purple-800">
                  Rp {formatNumberId(submitSuccessData.salesAmount || 0)} ·{' '}
                  {submitSuccessData.qtyAmount || 0} pcs ·{' '}
                  {submitSuccessData.trxAmount || 0} nota
                </span>
              </div>
              <div className="pt-2 border-t border-slate-200 text-slate-700 font-sans">
                <strong>Rincian:</strong> {submitSuccessData.summaryNote}
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => setSubmitSuccessData(null)}
                className="px-5 py-2.5 text-xs font-bold text-white bg-purple-700 hover:bg-purple-800 rounded-xl shadow-xs cursor-pointer"
              >
                Isi / Edit Tanggal CV Lainnya
              </button>
              {initialFormMode === 'none' && (
                <button
                  type="button"
                  onClick={() => {
                    setSubmitSuccessData(null);
                    setActiveView('hub');
                  }}
                  className="px-4 py-2.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl cursor-pointer"
                >
                  Kembali ke Daftar Link
                </button>
              )}
            </div>
          </div>
        ) : (
          <form
            onSubmit={handleSubmitCvReport}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
          >
            {/* Header */}
            <div className="bg-gradient-to-r from-purple-950 via-slate-900 to-indigo-950 text-white px-5 py-5 border-b border-purple-900/50">
              <div className="text-[11px] text-amber-400 font-extrabold uppercase tracking-wide">
                KATEGORI CV · DIVISI 1 (YOUNG FASHION, LADIES &amp; INTIMATE)
              </div>
              <h2 className="text-base sm:text-lg font-bold mt-0.5 text-white">
                Pengisian Report Harian CV (Graphis 6201 &amp; Simplicity 6244)
              </h2>
              <p className="text-xs text-slate-300 mt-0.5">
                Input Sales (Rp), Qty (pcs), dan Nota harian masing-masing brand agar langsung terakumulasi ke Report CV &amp; Dashboard
              </p>
            </div>

            <div className="p-5 sm:p-6 space-y-5">
              {/* Step 1: Tanggal & Petugas */}
              <div className="p-4 bg-purple-50/60 border border-purple-200 rounded-xl space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2 text-xs font-bold text-purple-950">
                    <Calendar className="w-4 h-4 text-purple-700" />
                    <span>Pilih Tanggal Laporan CV</span>
                  </div>
                  <span className="text-[11px] font-bold text-purple-900 bg-white border border-purple-200 px-2.5 py-0.5 rounded-md">
                    {cvMatchedWeek
                      ? `Week ${cvMatchedWeek.weekNumber} · ${cvMonthName} 2026`
                      : `${cvMonthName} 2026`}
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Tanggal Penjualan CV
                    </label>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          setCvFormCalendarIso(shiftIsoDate(cvFormCalendarIso, -1))
                        }
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0 cursor-pointer"
                        title="Hari Sebelumnya"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                      </button>
                      <input
                        type="date"
                        value={cvFormCalendarIso}
                        onChange={(e) => {
                          if (e.target.value) setCvFormCalendarIso(e.target.value);
                        }}
                        className="w-full px-2.5 py-1.5 text-xs font-bold font-mono-num bg-white border border-purple-300 rounded-lg text-slate-900 focus:outline-none focus:border-purple-600 cursor-pointer"
                      />
                      <button
                        type="button"
                        onClick={() =>
                          setCvFormCalendarIso(shiftIsoDate(cvFormCalendarIso, 1))
                        }
                        className="p-2 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0 cursor-pointer"
                        title="Hari Berikutnya"
                      >
                        <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                    </div>
                    <div className="text-[11px] font-medium text-purple-900 mt-1 font-mono-num">
                      {getIndoDayName(cvSlashDate)}, <strong>{cvSlashDate}</strong>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Nama Petugas Pengisi
                    </label>
                    <input
                      type="text"
                      value={cvFormStaff}
                      onChange={(e) => setCvFormStaff(e.target.value)}
                      placeholder="Ketik nama petugas / PIC CV..."
                      className="w-full px-3 py-2 text-xs font-semibold bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-purple-600"
                    />
                  </div>
                </div>

                {/* Brand Scope Tabs */}
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Pilih Brand yang Diisi
                  </label>
                  <div className="grid grid-cols-2 gap-1.5 p-1 bg-white border border-purple-200 rounded-lg">
                    {(
                      [
                        { id: 'GRAPHIS', label: '6201 · Graphis' },
                        { id: 'SIMPLICITY', label: '6244 · Simplicity' },
                      ] as const
                    ).map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setCvFormBrandScope(opt.id)}
                        className={`py-1.5 px-2 rounded-md text-xs font-bold transition-colors cursor-pointer ${
                          cvFormBrandScope === opt.id
                            ? 'bg-purple-700 text-white shadow-2xs'
                            : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                        }`}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Input Brand 1: Dept 6201 Graphis */}
              {cvFormBrandScope === 'GRAPHIS' && (
                <div className="p-4 bg-rose-50/40 border border-rose-200 rounded-xl space-y-3">
                  <div className="flex items-center justify-between border-b border-rose-200 pb-2">
                    <div className="flex items-center gap-2">
                      <Tag className="w-4 h-4 text-rose-700" />
                      <span className="text-xs font-extrabold text-rose-950 uppercase">
                        Dept 6201 · Brand Graphis
                      </span>
                    </div>
                    <span className="text-[11px] font-mono-num font-bold text-rose-800">
                      ASP:{' '}
                      {gQtyPreview > 0
                        ? `Rp ${formatNumberId(Math.round(gSalesPreview / gQtyPreview))}`
                        : '—'}
                    </span>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Sales Graphis (Rp)
                      </label>
                      <div className="relative flex items-center">
                        <span className="absolute left-3 text-xs font-bold text-slate-400 pointer-events-none">
                          Rp
                        </span>
                        <input
                          type="text"
                          inputMode="numeric"
                          value={cvFormGraphisSales}
                          onChange={(e) =>
                            setCvFormGraphisSales(formatAutoNominal(e.target.value))
                          }
                          placeholder="Contoh: 1.250.000"
                          className="w-full pl-8 pr-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-rose-600"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Qty Graphis (Pcs)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={cvFormGraphisQty}
                        onChange={(e) =>
                          setCvFormGraphisQty(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 12"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-rose-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Jumlah Nota (Nota)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={cvFormGraphisNota}
                        onChange={(e) =>
                          setCvFormGraphisNota(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 8"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-rose-600"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* Input Brand 2: Dept 6244 Simplicity */}
              {cvFormBrandScope === 'SIMPLICITY' && (
                <div className="p-4 bg-indigo-50/40 border border-indigo-200 rounded-xl space-y-3">
                  <div className="flex items-center justify-between border-b border-indigo-200 pb-2">
                    <div className="flex items-center gap-2">
                      <Tag className="w-4 h-4 text-indigo-700" />
                      <span className="text-xs font-extrabold text-indigo-950 uppercase">
                        Dept 6244 · Brand Simplicity
                      </span>
                    </div>
                    <span className="text-[11px] font-mono-num font-bold text-indigo-800">
                      ASP:{' '}
                      {sQtyPreview > 0
                        ? `Rp ${formatNumberId(Math.round(sSalesPreview / sQtyPreview))}`
                        : '—'}
                    </span>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Sales Simplicity (Rp)
                      </label>
                      <div className="relative flex items-center">
                        <span className="absolute left-3 text-xs font-bold text-slate-400 pointer-events-none">
                          Rp
                        </span>
                        <input
                          type="text"
                          inputMode="numeric"
                          value={cvFormSimplicitySales}
                          onChange={(e) =>
                            setCvFormSimplicitySales(formatAutoNominal(e.target.value))
                          }
                          placeholder="Contoh: 950.000"
                          className="w-full pl-8 pr-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Qty Simplicity (Pcs)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={cvFormSimplicityQty}
                        onChange={(e) =>
                          setCvFormSimplicityQty(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 10"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Jumlah Nota (Nota)
                      </label>
                      <input
                        type="text"
                        inputMode="numeric"
                        value={cvFormSimplicityNota}
                        onChange={(e) =>
                          setCvFormSimplicityNota(formatAutoNominal(e.target.value))
                        }
                        placeholder="Contoh: 6"
                        className="w-full px-3 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-indigo-600"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* Catatan / Reason Event Harian */}
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Catatan / Reason Event Harian CV (Opsional)
                </label>
                <input
                  type="text"
                  value={cvFormReasonEvent}
                  onChange={(e) => setCvFormReasonEvent(e.target.value)}
                  placeholder="Cth: Item dress & atasan paling laris, size M habis..."
                  className="w-full px-3 py-2 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:border-purple-600"
                />
              </div>

              {/* Live Preview Box */}
              <div className="p-4 bg-slate-900 text-white rounded-xl flex flex-wrap items-center justify-between gap-3 text-xs">
                <div className="flex items-center gap-2 font-bold text-amber-300">
                  <Calculator className="w-4 h-4" />
                  <span>
                    Total Report {CV_BRAND_CATALOG[cvFormBrandScope].label} ({cvSlashDate}):
                  </span>
                </div>
                <div className="flex flex-wrap items-center gap-3 font-mono-num">
                  <span>
                    Sales: <strong>Rp {formatNumberId(totalCvSalesPreview) || '0'}</strong>
                  </span>
                  <span className="text-slate-500">·</span>
                  <span>
                    Qty: <strong>{totalCvQtyPreview} pcs</strong>
                  </span>
                  <span className="text-slate-500">·</span>
                  <span>
                    Nota: <strong>{totalCvNotaPreview} nota</strong>
                  </span>
                  <span className="text-slate-500">·</span>
                  <span className="text-amber-300">
                    Price Point:{' '}
                    <strong>Rp {formatNumberId(cvPricePointPreview) || '0'}</strong>
                  </span>
                </div>
              </div>

              <div className="pt-1 flex justify-end">
                <button
                  type="submit"
                  className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 text-xs font-bold text-white bg-purple-700 hover:bg-purple-800 rounded-xl transition-colors shadow-xs cursor-pointer"
                >
                  <Send className="w-4 h-4" />
                  <span>Kirim &amp; Simpan Report CV</span>
                </button>
              </div>
            </div>
          </form>
        )}
      </div>
    );
  }

  // RENDER 2C: DATA CV FILLING FORM (Inventory, Promo, Reason/Kendala & Action Plan)
  if (activeView === 'data_cv_form') {
    const activeWeekObj = ANNUAL_2026_WEEKS.find(
      (w) => w.weekNumber === dataCvFormWeekNum
    );

    const renderDataCvBrandEditor = (
      brand: CvInventoryBrandKey,
      draft: DataCvBrandDraft,
      setDraft: React.Dispatch<React.SetStateAction<DataCvBrandDraft>>
    ) => {
      const liveQty =
        dataCvFormPeriodType === 'weekly'
          ? liveCvSalesMap.weekly.get(dataCvFormWeekNum)?.[brand] || 0
          : liveCvSalesMap.monthly.get(dataCvFormMonthKey)?.[brand] || 0;
      const sAwal = Math.max(0, Number(draft.stockAwal) || 0);
      const bDatang = Math.max(0, Number(draft.barangDatang) || 0);
      const bKeluar = Math.max(0, Number(draft.barangKeluar) || 0);
      const sQty =
        draft.salesQtyMode === 'manual'
          ? Math.max(0, Number(draft.salesQtyManual) || 0)
          : liveQty;
      const sAkhir = sAwal + bDatang - sQty - bKeluar;
      const avail = sAwal + bDatang;
      const stPct = avail > 0 ? (sQty / avail) * 100 : 0;
      const isGraphis = brand === 'GRAPHIS';

      return (
        <div
          key={brand}
          className={`p-4 rounded-xl border space-y-4 ${
            isGraphis
              ? 'bg-rose-50/35 border-rose-200'
              : 'bg-indigo-50/35 border-indigo-200'
          }`}
        >
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2.5">
            <div className="flex items-center gap-2">
              <PackageCheck
                className={`w-4 h-4 ${
                  isGraphis ? 'text-rose-700' : 'text-indigo-700'
                }`}
              />
              <span className="text-xs font-extrabold text-slate-950 uppercase">
                Dept {CV_BRAND_CATALOG[brand].deptCode} · Brand{' '}
                {CV_BRAND_CATALOG[brand].label}
              </span>
            </div>
            <div className="flex items-center gap-2 font-mono-num text-xs">
              <span className="text-slate-500">Stock Akhir:</span>
              <span
                className={`font-extrabold ${
                  sAkhir < 0
                    ? 'text-rose-700'
                    : sAkhir <= 15
                      ? 'text-amber-700'
                      : 'text-emerald-700'
                }`}
              >
                {formatNumberId(sAkhir)} pcs
              </span>
              {avail > 0 && (
                <span className="text-[11px] text-slate-500">
                  (ST {stPct.toFixed(1)}%)
                </span>
              )}
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                Stock Awal (Pcs)
              </label>
              <input
                type="number"
                min="0"
                value={draft.stockAwal}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, stockAwal: e.target.value }))
                }
                className="w-full px-2.5 py-2 text-xs font-bold font-mono-num bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-sky-600"
              />
            </div>

            <div>
              <label className="block text-[11px] font-semibold text-emerald-800 mb-1">
                Barang Datang (+)
              </label>
              <input
                type="number"
                min="0"
                value={draft.barangDatang}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, barangDatang: e.target.value }))
                }
                className="w-full px-2.5 py-2 text-xs font-bold font-mono-num bg-white border border-emerald-300 rounded-lg text-emerald-900 focus:outline-none focus:border-emerald-600"
              />
            </div>

            <div>
              <div className="flex items-center justify-between mb-1">
                <label className="text-[11px] font-semibold text-sky-900">
                  Sales Qty (-)
                </label>
                <button
                  type="button"
                  onClick={() =>
                    setDraft((p) => ({
                      ...p,
                      salesQtyMode:
                        p.salesQtyMode === 'live' ? 'manual' : 'live',
                      salesQtyManual: String(liveQty),
                    }))
                  }
                  className="text-[10px] font-bold text-sky-700 hover:underline cursor-pointer"
                >
                  {draft.salesQtyMode === 'live' ? 'Ubah Manual' : 'Pakai Live'}
                </button>
              </div>
              {draft.salesQtyMode === 'live' ? (
                <div className="px-2.5 py-2 text-xs font-bold font-mono-num bg-sky-50 border border-sky-200 rounded-lg text-sky-900">
                  {formatNumberId(liveQty)} (Live)
                </div>
              ) : (
                <input
                  type="number"
                  min="0"
                  value={draft.salesQtyManual}
                  onChange={(e) =>
                    setDraft((p) => ({ ...p, salesQtyManual: e.target.value }))
                  }
                  className="w-full px-2.5 py-2 text-xs font-bold font-mono-num bg-white border border-sky-300 rounded-lg text-sky-900 focus:outline-none focus:border-sky-600"
                />
              )}
            </div>

            <div>
              <label className="block text-[11px] font-semibold text-amber-800 mb-1">
                Barang Keluar (-)
              </label>
              <input
                type="number"
                min="0"
                value={draft.barangKeluar}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, barangKeluar: e.target.value }))
                }
                className="w-full px-2.5 py-2 text-xs font-bold font-mono-num bg-white border border-amber-300 rounded-lg text-amber-900 focus:outline-none focus:border-amber-600"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Program Promo Berjalan
              </label>
              <textarea
                rows={3}
                value={draft.promo}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, promo: e.target.value }))
                }
                placeholder="Cth: Disc 50%+20%, Special Price 99rb..."
                className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-sky-600"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Reason / Kendala Lapangan
              </label>
              <textarea
                rows={3}
                value={draft.reasonKendala}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, reasonKendala: e.target.value }))
                }
                placeholder="Cth: Size L & XL kosong, butuh tambahan mannequin..."
                className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-sky-600"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                Action Plan Tindak Lanjut
              </label>
              <textarea
                rows={3}
                value={draft.actionPlan}
                onChange={(e) =>
                  setDraft((p) => ({ ...p, actionPlan: e.target.value }))
                }
                placeholder="Cth: Request DO size M-XL ke gudang pusat..."
                className="w-full p-2.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:border-sky-600"
              />
            </div>
          </div>
        </div>
      );
    };

    return (
      <div className="max-w-3xl mx-auto space-y-6">
        {/* Top Switcher Bar */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 bg-white border border-slate-200 rounded-xl p-3 shadow-2xs">
          <div className="flex flex-wrap items-center gap-2">
            {initialFormMode === 'none' && (
              <button
                type="button"
                onClick={() => {
                  setSubmitSuccessData(null);
                  setActiveView('hub');
                }}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-700 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                Daftar Link Report
              </button>
            )}
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-sky-50 border border-sky-200 text-sky-900 text-xs font-bold">
              <Lock className="w-3.5 h-3.5 text-sky-700" />
              Form Pengisian Data CV (Ops &amp; Stock)
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('shift_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Report DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('review_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer"
            >
              Data DP
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('cv_report_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-purple-50 hover:bg-purple-100 text-purple-800 border border-purple-200 cursor-pointer"
            >
              Report CV
            </button>
            <button
              type="button"
              onClick={() => {
                setSubmitSuccessData(null);
                setActiveView('data_cv_form');
              }}
              className="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-sky-700 text-white cursor-pointer"
            >
              Data CV
            </button>
          </div>
        </div>

        {submitSuccessData && submitSuccessData.reportType === 'data_cv' ? (
          <div className="bg-white border border-emerald-200 rounded-xl p-6 sm:p-8 text-center space-y-5 shadow-xs">
            <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto">
              <CheckCircle2 className="w-6 h-6" />
            </div>
            <div>
              <h2 className="text-lg font-bold text-slate-900">
                Data CV (Inventory, Promo &amp; Action Plan) Berhasil Disimpan!
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Data periode <strong>{submitSuccessData.weekTitle}</strong> telah disinkronkan ke menu Data CV &amp; ekspor Excel.
              </p>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 text-left text-xs space-y-2 font-mono-num">
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Petugas:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.submittedBy}
                </span>
              </div>
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Mode &amp; Periode:</span>
                <span className="font-bold text-slate-900">
                  {submitSuccessData.shift} · {submitSuccessData.weekTitle}
                </span>
              </div>
              <div className="flex justify-between gap-2">
                <span className="text-slate-500">Total Stock Akhir / Sales Qty:</span>
                <span className="font-bold text-sky-800">
                  {formatNumberId(submitSuccessData.salesAmount || 0)} pcs ·{' '}
                  {formatNumberId(submitSuccessData.qtyAmount || 0)} pcs terjual
                </span>
              </div>
              <div className="pt-2 border-t border-slate-200 text-slate-700 font-sans">
                <strong>Ringkasan Inventory:</strong> {submitSuccessData.summaryNote}
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => setSubmitSuccessData(null)}
                className="px-5 py-2.5 text-xs font-bold text-white bg-sky-700 hover:bg-sky-800 rounded-xl shadow-xs cursor-pointer"
              >
                Isi / Edit Data CV Periode Lain
              </button>
              {initialFormMode === 'none' && (
                <button
                  type="button"
                  onClick={() => {
                    setSubmitSuccessData(null);
                    setActiveView('hub');
                  }}
                  className="px-4 py-2.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-xl cursor-pointer"
                >
                  Kembali ke Daftar Link
                </button>
              )}
            </div>
          </div>
        ) : (
          <form
            onSubmit={handleSubmitDataCv}
            className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs"
          >
            {/* Header */}
            <div className="bg-slate-950 text-white px-5 py-5 border-b border-slate-800">
              <div className="text-[11px] text-amber-400 font-extrabold uppercase tracking-wide">
                DATA CV · DIVISI 1 (YOUNG FASHION, LADIES &amp; INTIMATE)
              </div>
              <h2 className="text-base sm:text-lg font-bold mt-0.5 text-white">
                Pengisian Data CV (Inventory, Promo, Reason/Kendala &amp; Action Plan)
              </h2>
              <p className="text-xs text-slate-300 mt-0.5">
                Stock Akhir dihitung otomatis: Stock Awal + Barang Datang - Sales Qty - Barang Keluar
              </p>
            </div>

            <div className="p-5 sm:p-6 space-y-5">
              {/* Step 1: Periode & Brand Scope */}
              <div className="p-4 bg-sky-50/70 border border-sky-200 rounded-xl space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      1. Mode Periode Data CV
                    </label>
                    <div className="grid grid-cols-2 gap-1 p-1 bg-white border border-sky-200 rounded-lg">
                      <button
                        type="button"
                        onClick={() => setDataCvFormPeriodType('weekly')}
                        className={`py-1.5 text-xs font-bold rounded-md cursor-pointer ${
                          dataCvFormPeriodType === 'weekly'
                            ? 'bg-sky-700 text-white'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Per Week
                      </button>
                      <button
                        type="button"
                        onClick={() => setDataCvFormPeriodType('monthly')}
                        className={`py-1.5 text-xs font-bold rounded-md cursor-pointer ${
                          dataCvFormPeriodType === 'monthly'
                            ? 'bg-sky-700 text-white'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        Per Monthly
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      {dataCvFormPeriodType === 'weekly'
                        ? '2. Pilih Minggu (Week 1–53)'
                        : '2. Pilih Bulan (2026)'}
                    </label>
                    {dataCvFormPeriodType === 'weekly' ? (
                      <select
                        value={dataCvFormWeekNum}
                        onChange={(e) =>
                          setDataCvFormWeekNum(Number(e.target.value))
                        }
                        className="w-full px-2.5 py-2 text-xs font-bold bg-white border border-sky-300 rounded-lg text-slate-900 focus:outline-none focus:border-sky-600"
                      >
                        {ANNUAL_2026_WEEKS.map((w) => (
                          <option key={w.weekNumber} value={w.weekNumber}>
                            W-{w.weekNumber} · {w.periodeText}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <select
                        value={dataCvFormMonthKey}
                        onChange={(e) => setDataCvFormMonthKey(e.target.value)}
                        className="w-full px-2.5 py-2 text-xs font-bold bg-white border border-sky-300 rounded-lg text-slate-900 focus:outline-none focus:border-sky-600"
                      >
                        {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                          <option key={mKey} value={mKey}>
                            {mName} 2026
                          </option>
                        ))}
                      </select>
                    )}
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      3. Nama Petugas Pengisi
                    </label>
                    <input
                      type="text"
                      value={dataCvFormStaff}
                      onChange={(e) => setDataCvFormStaff(e.target.value)}
                      placeholder="Nama petugas / SPV..."
                      className="w-full px-3 py-2 text-xs font-semibold bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-sky-600"
                    />
                  </div>
                </div>

                {/* Brand Scope Tabs */}
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Pilih Brand yang Diinput
                  </label>
                  <div className="grid grid-cols-2 gap-1.5 p-1 bg-white border border-sky-200 rounded-lg">
                    {(
                      [
                        { id: 'GRAPHIS', label: '6201 · Graphis' },
                        { id: 'SIMPLICITY', label: '6244 · Simplicity' },
                      ] as const
                    ).map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setDataCvFormBrandScope(opt.id)}
                        className={`py-1.5 px-2 rounded-md text-xs font-bold transition-colors cursor-pointer ${
                          dataCvFormBrandScope === opt.id
                            ? 'bg-sky-700 text-white'
                            : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                        }`}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>
                  {dataCvFormPeriodType === 'weekly' && activeWeekObj && (
                    <div className="text-[11px] font-medium text-sky-900 mt-1.5">
                      Periode aktif: <strong>Week {activeWeekObj.weekNumber}</strong> ({activeWeekObj.periodeText})
                    </div>
                  )}
                </div>
              </div>

              {/* Brand Draft Editors */}
              {dataCvFormBrandScope === 'GRAPHIS' &&
                renderDataCvBrandEditor(
                  'GRAPHIS',
                  dataCvGraphisDraft,
                  setDataCvGraphisDraft
                )}

              {dataCvFormBrandScope === 'SIMPLICITY' &&
                renderDataCvBrandEditor(
                  'SIMPLICITY',
                  dataCvSimplicityDraft,
                  setDataCvSimplicityDraft
                )}

              <div className="pt-1 flex justify-end">
                <button
                  type="submit"
                  className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-6 py-3 text-xs font-bold text-white bg-sky-700 hover:bg-sky-800 rounded-xl transition-colors shadow-xs cursor-pointer"
                >
                  <Send className="w-4 h-4" />
                  <span>Kirim &amp; Simpan Data CV</span>
                </button>
              </div>
            </div>
          </form>
        )}
      </div>
    );
  }

  // RENDER 3: HUB GENERATOR LINK & RIWAYAT PENGISIAN REPORT
  return (
    <div className="space-y-6">
      {/* Header Banner */}
      <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
          <div className="flex flex-wrap items-center gap-2 text-xs font-bold text-indigo-600">
            <Link2 className="w-4 h-4" />
            <span>REPORT BAZAR MATAHARI · Tanjung · PT MDS Retailing Tbk</span>
          </div>
          <h1 className="text-xl font-bold text-slate-900 mt-1">
            Link Pengisian Report Bazar Matahari — Tanjung
          </h1>
          <p className="text-xs text-slate-500 mt-0.5">
            Bagikan link pengisian ini kepada petugas shift melalui WhatsApp agar mereka dapat menginput laporan harian langsung dari HP.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {onToggleLockReportOnly && (
            <button
              type="button"
              onClick={() => onToggleLockReportOnly(!isLockedReportOnly)}
              className={`inline-flex items-center gap-1.5 px-3.5 py-2.5 text-xs font-bold rounded-lg transition-colors whitespace-nowrap border ${
                isLockedReportOnly
                  ? 'bg-amber-400 text-slate-950 border-amber-500 shadow-xs'
                  : 'bg-amber-50 text-amber-900 hover:bg-amber-100 border-amber-200'
              }`}
              title="Kunci aplikasi agar hanya bisa membuka Menu Pengisian Report dan menyembunyikan seluruh menu lainnya"
            >
              <Lock className="w-3.5 h-3.5" />
              {isLockedReportOnly
                ? 'Mode Terkunci Aktif (Hanya Pengisian Report)'
                : 'Kunci Hanya Menu Pengisian Report'}
            </button>
          )}
          <button
            type="button"
            onClick={() => {
              const dIdx = linkDayIdx === 'ANY' ? 0 : Number(linkDayIdx);
              if (!hasShiftDraftContent) {
                setFormCalendarIso(linkCalendarIso);
                setFormWeekId(linkWeekId);
                setFormDayIdx(dIdx);
              }
              setSubmitSuccessData(null);
              setActiveView('shift_form');
            }}
            className="inline-flex items-center gap-1.5 px-4 py-2.5 text-xs font-semibold text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg transition-colors whitespace-nowrap"
          >
            <ExternalLink className="w-3.5 h-3.5" />
            Buka Form Report DP
          </button>
          <button
            type="button"
            onClick={() => {
              setSubmitSuccessData(null);
              setActiveView('review_form');
            }}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap cursor-pointer"
          >
            <ClipboardList className="w-3.5 h-3.5" />
            Buka Form Data DP
          </button>
          <button
            type="button"
            onClick={() => {
              setSubmitSuccessData(null);
              setActiveView('cv_report_form');
            }}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-purple-700 hover:bg-purple-800 rounded-lg transition-colors whitespace-nowrap cursor-pointer"
          >
            <Tag className="w-3.5 h-3.5" />
            Buka Form Report CV
          </button>
          <button
            type="button"
            onClick={() => {
              setSubmitSuccessData(null);
              setActiveView('data_cv_form');
            }}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-sky-700 hover:bg-sky-800 rounded-lg transition-colors whitespace-nowrap cursor-pointer"
          >
            <Layers className="w-3.5 h-3.5" />
            Buka Form Data CV
          </button>
        </div>
      </div>

      {/* Info Banner: Link Google Publik (Buka Aplikasi & Pengisian Report) */}
      <div className="bg-emerald-50/90 border border-emerald-200 rounded-xl p-4 space-y-3 text-xs text-emerald-950">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="flex items-start gap-3">
            <ShieldCheck className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <div className="font-extrabold text-xs sm:text-sm text-emerald-950">
                Link Aplikasi &amp; Pengisian Report (Tersinkronisasi Otomatis)
              </div>
              <p className="text-xs text-emerald-900 leading-relaxed">
                Link Buka Aplikasi dan Link Pengisian Report (Report DP, Data DP, Report CV, dan Data CV) di bawah ini dapat dibuka langsung dan tersinkronisasi secara <em>real-time</em>.
              </p>
            </div>
          </div>

          <button
            type="button"
            onClick={() =>
              handleCopyText(
                'wa_all_google_public',
                [
                  '*LINK GOOGLE PUBLIK — REPORT BAZAR MATAHARI*',
                  '*Tanjung · PT MDS Retailing Tbk*',
                  '',
                  '1. *Link Buka Aplikasi (Google Publik / Hanya Lihat):*',
                  publicAppUrl,
                  '',
                  '2. *Link Pengisian Report DP (Shift Harian):*',
                  shiftReportUrl,
                  '',
                  '3. *Link Pengisian Data DP (Review Monitoring Bazar):*',
                  reviewReportUrl,
                  '',
                  '4. *Link Pengisian Report CV (Divisi 1 · Graphis 6201 & Simplicity 6244):*',
                  cvReportUrl,
                  '',
                  '5. *Link Pengisian Data CV (Inventory, Promo, Reason & Action Plan):*',
                  dataCvUrl,
                ].join('\n')
              )
            }
            className="px-3.5 py-2 rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-extrabold inline-flex items-center gap-1.5 shrink-0 cursor-pointer shadow-2xs"
          >
            {copiedKey === 'wa_all_google_public' ? (
              <>
                <Check className="w-3.5 h-3.5 text-amber-300" />
                <span>Semua Link Publik Tersalin!</span>
              </>
            ) : (
              <>
                <Share2 className="w-3.5 h-3.5" />
                <span>Salin Semua Link Google Publik (WA)</span>
              </>
            )}
          </button>
        </div>

        {/* Quick Box: Link Buka Aplikasi untuk Google Publik */}
        <div className="p-3 rounded-xl bg-white border border-emerald-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div className="min-w-0">
            <div className="text-[11px] font-extrabold text-sky-900">
              Link Buka Aplikasi untuk Google Publik (Mode Hanya Lihat / Tanpa Login):
            </div>
            <div className="text-xs font-mono text-slate-700 truncate select-all mt-0.5">
              {publicAppUrl}
            </div>
          </div>
          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={() => handleCopyText('public_app_url', publicAppUrl)}
              className="px-3 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold inline-flex items-center gap-1.5 cursor-pointer"
            >
              {copiedKey === 'public_app_url' ? (
                <>
                  <Check className="w-3.5 h-3.5" />
                  <span>Tersalin!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5" />
                  <span>Salin Link Buka Aplikasi</span>
                </>
              )}
            </button>
          </div>
        </div>
      </div>

      {/* Two Link Generator Cards */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Card 1: Link Pengisian Report Shift Harian (7 cols) */}
        <div className="lg:col-span-7 bg-white border border-slate-200 rounded-xl p-5 space-y-5 flex flex-col justify-between">
          <div className="space-y-4">
            <div className="flex items-start justify-between gap-3 border-b border-slate-100 pb-3.5">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-bold text-slate-900">
                    1. Link Pengisian Report DP (Sales, Qty, Trx &amp; Reason)
                  </h2>
                  <p className="text-xs text-slate-500">
                    Atur parameter link di bawah untuk mengarahkan petugas langsung ke minggu/shift tertentu
                  </p>
                </div>
              </div>
            </div>

            {/* Link Parameter Customizer with Calendar Date Filter -> Auto Weekly */}
            <div className="p-3.5 bg-indigo-50/60 border border-indigo-200 rounded-xl space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 text-xs font-bold text-indigo-950">
                  <Calendar className="w-4 h-4 text-indigo-600" />
                  <span>Filter Kalender Tanggal (Weekly Otomatis Terpilih)</span>
                </div>
                <div className="inline-flex items-center gap-1 bg-white p-0.5 rounded-lg border border-indigo-200 text-[11px]">
                  <button
                    type="button"
                    onClick={() => {
                      setLinkDateFilterMode('CALENDAR');
                      handleSelectLinkCalendarDate(linkCalendarIso);
                    }}
                    className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                      linkDateFilterMode === 'CALENDAR'
                        ? 'bg-indigo-600 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Pilih Tanggal Kalender
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setLinkDateFilterMode('ANY');
                      setLinkDayIdx('ANY');
                    }}
                    className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                      linkDateFilterMode === 'ANY'
                        ? 'bg-indigo-600 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Semua Tanggal (1 Minggu)
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                {/* 1. Filter Kalender Tanggal */}
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    1. Pilih Tanggal (Kalender)
                  </label>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        handleSelectLinkCalendarDate(shiftIsoDate(linkCalendarIso, -1))
                      }
                      title="Tanggal Sebelumnya"
                      className="p-1.5 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                    >
                      <ChevronLeft className="w-3.5 h-3.5" />
                    </button>
                    <input
                      type="date"
                      value={linkCalendarIso}
                      onChange={(e) => handleSelectLinkCalendarDate(e.target.value)}
                      className="w-full px-2 py-1.5 text-xs font-bold font-mono-num bg-white border border-indigo-300 rounded-lg text-slate-900 focus:outline-none focus:border-indigo-600 cursor-pointer"
                    />
                    <button
                      type="button"
                      onClick={() =>
                        handleSelectLinkCalendarDate(shiftIsoDate(linkCalendarIso, 1))
                      }
                      title="Tanggal Berikutnya"
                      className="p-1.5 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 shrink-0"
                    >
                      <ChevronRight className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>

                {/* 2. Target Weekly (Otomatis dari Tanggal Kalender) */}
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    2. Target Weekly (Otomatis)
                  </label>
                  <div className="px-2.5 py-1.5 bg-white border border-indigo-200 rounded-lg flex items-center justify-between gap-1.5">
                    <div className="min-w-0">
                      <div className="text-xs font-extrabold text-indigo-950 truncate">
                        {selectedLinkWeekObj?.title}
                      </div>
                      <div className="text-[10px] text-slate-500 truncate">
                        {selectedLinkWeekObj ? getWeeklyPeriodeText(selectedLinkWeekObj) : ''}
                      </div>
                    </div>
                    <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px] font-bold shrink-0">
                      <Sparkles className="w-2.5 h-2.5" />
                      Auto
                    </span>
                  </div>
                </div>

                {/* 3. Target Shift */}
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    3. Target Shift
                  </label>
                  <select
                    value={linkShift}
                    onChange={(e) =>
                      setLinkShift(e.target.value as 'ANY' | 'Pagi' | 'Siang' | 'Lembur')
                    }
                    className="w-full px-2.5 py-2 text-xs bg-white border border-slate-200 rounded-lg"
                  >
                    <option value="ANY">Semua Shift</option>
                    <option value="Pagi">Shift Pagi</option>
                    <option value="Siang">Shift Siang</option>
                    <option value="Lembur">Shift Lembur</option>
                  </select>
                </div>

                {/* 4. Petugas (Isi Manual) */}
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    4. Petugas
                  </label>
                  <input
                    type="text"
                    value={linkStaff === 'ANY' ? '' : linkStaff}
                    onChange={(e) => setLinkStaff(e.target.value)}
                    placeholder="Ketik nama petugas (manual)..."
                    className="w-full px-2.5 py-2 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 placeholder:text-slate-400"
                  />
                </div>
              </div>

              {/* Status Ringkasan Otomatis */}
              <div className="flex flex-wrap items-center justify-between gap-2 pt-1 text-[11px] text-indigo-950">
                <div className="flex items-center gap-1.5">
                  <Sparkles className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
                  <span>
                    Tanggal dipilih:{' '}
                    <strong className="font-mono-num">
                      {linkDateFilterMode === 'ANY'
                        ? `Semua Hari (${selectedLinkWeekObj?.days[0]?.tanggal} – ${selectedLinkWeekObj?.days[6]?.tanggal})`
                        : `${getIndoDayName(
                            selectedLinkWeekObj?.days[Number(linkDayIdx)]?.tanggal ||
                              isoToSlashDate(linkCalendarIso)
                          )}, ${
                            selectedLinkWeekObj?.days[Number(linkDayIdx)]?.tanggal ||
                            isoToSlashDate(linkCalendarIso)
                          }`}
                    </strong>{' '}
                    → Otomatis masuk ke <strong>{selectedLinkWeekObj?.title}</strong> (
                    {selectedLinkWeekObj ? getWeeklyPeriodeText(selectedLinkWeekObj) : ''})
                  </span>
                </div>
              </div>
            </div>

            {/* Generated URL Box */}
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
              <div className="text-[11px] font-semibold text-slate-500">
                Link Pengisian Report DP (Siap Dibagikan):
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  readOnly
                  value={shiftReportUrl}
                  className="flex-1 px-3 py-2 text-xs font-mono-num text-slate-800 bg-white border border-slate-200 rounded-md select-all"
                />
                <button
                  type="button"
                  onClick={() => handleCopyText('shift_url', shiftReportUrl)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-md transition-colors whitespace-nowrap"
                >
                  {copiedKey === 'shift_url' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-400" />
                      Link Tersalin!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      Salin Link
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
            <button
              type="button"
              onClick={() => handleCopyText('wa_invite', buildWhatsAppInviteText())}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors"
            >
              {copiedKey === 'wa_invite' ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  Pesan + Link WA Tersalin!
                </>
              ) : (
                <>
                  <Share2 className="w-3.5 h-3.5" />
                  Salin Format Undangan WhatsApp
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => {
                setFormCalendarIso(linkCalendarIso);
                setFormWeekId(linkWeekId);
                const dIdx = linkDayIdx === 'ANY' ? 0 : Number(linkDayIdx);
                setFormDayIdx(dIdx);
                setFormShiftType(linkShift === 'ANY' ? 'Pagi' : linkShift);
                resetShiftFormInputs();
                setSubmitSuccessData(null);
                if (linkStaff !== 'ANY') {
                  setFormStaffPagi(linkStaff);
                  setFormStaffSiang(linkStaff);
                }
                setActiveView('shift_form');
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 rounded-lg transition-colors"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              Isi / Uji Coba Link Ini Sekarang
            </button>
          </div>
        </div>

        {/* Card 2: Link Pengisian Review Monitoring Bazar (5 cols) */}
        <div className="lg:col-span-5 bg-white border border-slate-200 rounded-xl p-5 flex flex-col justify-between space-y-5">
          <div className="space-y-4">
            <div className="flex items-center gap-2.5 border-b border-slate-100 pb-3.5">
              <div className="w-9 h-9 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
                <ClipboardList className="w-5 h-5" />
              </div>
              <div>
                <h2 className="text-sm font-bold text-slate-900">
                  2. Link Pengisian Data DP (Review Monitoring Bazar)
                </h2>
                <p className="text-xs text-slate-500">
                  Formulir evaluasi Fast/Slow Moving, Promo, Action Plan &amp; Need Support
                </p>
              </div>
            </div>

            {/* Calendar Date Filter for Data DP Link */}
            <div className="p-3 bg-emerald-50/70 border border-emerald-200 rounded-xl space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="inline-flex items-center gap-1.5 text-xs font-bold text-emerald-950">
                  <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Filter Kalender Tanggal Link Data DP</span>
                </span>
                <div className="inline-flex items-center gap-1 bg-white p-0.5 rounded-lg border border-emerald-200 text-[10px]">
                  <button
                    type="button"
                    onClick={() => setLinkReviewDateMode('CALENDAR')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkReviewDateMode === 'CALENDAR'
                        ? 'bg-emerald-600 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tanggal
                  </button>
                  <button
                    type="button"
                    onClick={() => setLinkReviewDateMode('ANY')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkReviewDateMode === 'ANY'
                        ? 'bg-emerald-600 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Semua
                  </button>
                </div>
              </div>
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => {
                    setLinkReviewDateMode('CALENDAR');
                    setLinkReviewCalendarIso(shiftIsoDate(linkReviewCalendarIso, -1));
                  }}
                  title="Tanggal Sebelumnya"
                  className="p-1.5 rounded-lg bg-white border border-emerald-300 hover:bg-emerald-100 text-slate-700 shrink-0"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                </button>
                <input
                  type="date"
                  aria-label="Filter Kalender Tanggal Link Data DP"
                  value={linkReviewCalendarIso}
                  onChange={(e) => {
                    if (e.target.value) {
                      setLinkReviewDateMode('CALENDAR');
                      setLinkReviewCalendarIso(e.target.value);
                    }
                  }}
                  className="flex-1 px-2.5 py-1.5 text-xs font-bold font-mono-num bg-white border border-emerald-300 rounded-lg text-slate-900 focus:outline-none focus:border-emerald-600 cursor-pointer"
                />
                <button
                  type="button"
                  onClick={() => {
                    setLinkReviewDateMode('CALENDAR');
                    setLinkReviewCalendarIso(shiftIsoDate(linkReviewCalendarIso, 1));
                  }}
                  title="Tanggal Berikutnya"
                  className="p-1.5 rounded-lg bg-white border border-emerald-300 hover:bg-emerald-100 text-slate-700 shrink-0"
                >
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
              {linkReviewDateMode === 'CALENDAR' && (
                <div className="text-[11px] font-semibold text-emerald-900">
                  Target Weekly Data DP:{' '}
                  <strong>
                    {resolveCalendarDateToWeekly(linkReviewCalendarIso, false).table.title}
                  </strong>{' '}
                  ({isoToSlashDate(linkReviewCalendarIso)})
                </div>
              )}
            </div>

            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
              <div className="text-[11px] font-semibold text-slate-500">
                Link Pengisian Data DP (Siap Dibagikan):
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  readOnly
                  value={reviewReportUrl}
                  className="flex-1 px-3 py-2 text-xs font-mono-num text-slate-800 bg-white border border-slate-200 rounded-md select-all"
                />
                <button
                  type="button"
                  onClick={() => handleCopyText('review_url', reviewReportUrl)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-md transition-colors whitespace-nowrap"
                >
                  {copiedKey === 'review_url' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-400" />
                      Tersalin!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      Salin Link
                    </>
                  )}
                </button>
              </div>
            </div>

            <div className="text-xs text-slate-600 space-y-1 bg-slate-50/70 p-3 rounded-lg border border-slate-100">
              <div className="font-semibold text-slate-800">Cakupan Formulir Data DP:</div>
              <div>• Update daftar barang Fast Moving &amp; Slow Moving</div>
              <div>• Evaluasi C. Place/People &amp; D. Action Plan</div>
              <div>• Laporan Reason/Kendala &amp; Need Support Barang</div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-100 flex items-center justify-end">
            <button
              type="button"
              onClick={() => {
                const targetIso = linkReviewDateMode === 'CALENDAR' ? linkReviewCalendarIso : formReviewCalendarIso;
                setFormReviewCalendarIso(targetIso);
                const wk = resolveCalendarDateToWeekly(targetIso, false).table.weekNumber;
                loadReviewIntoFormByWeekNumber(wk);
                setSubmitSuccessData(null);
                setActiveView('review_form');
              }}
              className="w-full inline-flex items-center justify-center gap-1.5 px-4 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              Buka Form Data DP
            </button>
          </div>
        </div>
      </div>

      {/* Two CV Link Generator Cards: 3. Link Report CV & 4. Link Data CV */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Card 3: Link Pengisian Report CV (Divisi 1: 6201 Graphis & 6244 Simplicity) */}
        <div className="lg:col-span-6 bg-white border border-purple-200 rounded-xl p-5 flex flex-col justify-between space-y-5">
          <div className="space-y-4">
            <div className="flex items-start justify-between gap-3 border-b border-slate-100 pb-3.5">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-lg bg-purple-50 text-purple-700 flex items-center justify-center shrink-0">
                  <Tag className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-bold text-slate-900">
                    3. Link Pengisian Report CV (Sales Harian Divisi 1)
                  </h2>
                  <p className="text-xs text-slate-500">
                    Form khusus input Sales, Qty &amp; Nota harian Brand 6201 Graphis &amp; 6244 Simplicity
                  </p>
                </div>
              </div>
            </div>

            <div className="p-3.5 bg-purple-50/60 border border-purple-200 rounded-xl space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="inline-flex items-center gap-1.5 text-xs font-bold text-purple-950">
                  <Calendar className="w-4 h-4 text-purple-700" />
                  <span>Parameter Tanggal &amp; Brand Report CV</span>
                </span>
                <div className="inline-flex items-center gap-1 bg-white p-0.5 rounded-lg border border-purple-200 text-[10px]">
                  <button
                    type="button"
                    onClick={() => setLinkCvDateMode('CALENDAR')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkCvDateMode === 'CALENDAR'
                        ? 'bg-purple-700 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tanggal
                  </button>
                  <button
                    type="button"
                    onClick={() => setLinkCvDateMode('ANY')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkCvDateMode === 'ANY'
                        ? 'bg-purple-700 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Semua Tanggal
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Tanggal Laporan
                  </label>
                  <input
                    type="date"
                    value={linkCvCalendarIso}
                    onChange={(e) => {
                      if (e.target.value) {
                        setLinkCvDateMode('CALENDAR');
                        setLinkCvCalendarIso(e.target.value);
                      }
                    }}
                    className="w-full px-2 py-1.5 text-xs font-bold font-mono-num bg-white border border-purple-300 rounded-lg text-slate-900 focus:outline-none focus:border-purple-600 cursor-pointer"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Pilih Brand CV
                  </label>
                  <select
                    value={linkCvBrand}
                    onChange={(e) =>
                      setLinkCvBrand(e.target.value as CvInventoryBrandKey)
                    }
                    className="w-full px-2 py-1.5 text-xs font-semibold bg-white border border-slate-200 rounded-lg"
                  >
                    <option value="GRAPHIS">6201 · Graphis</option>
                    <option value="SIMPLICITY">6244 · Simplicity</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Petugas (Opsional)
                  </label>
                  <input
                    type="text"
                    value={linkCvStaff}
                    onChange={(e) => setLinkCvStaff(e.target.value)}
                    placeholder="Nama petugas CV..."
                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-purple-600"
                  />
                </div>
              </div>
            </div>

            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
              <div className="text-[11px] font-semibold text-slate-500">
                Link Pengisian Report CV (Siap Dibagikan):
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  readOnly
                  value={cvReportUrl}
                  onClick={(e) => (e.target as HTMLInputElement).select()}
                  className="flex-1 px-3 py-2 text-xs font-mono-num text-slate-800 bg-white border border-slate-200 rounded-md select-all"
                />
                <button
                  type="button"
                  onClick={() => handleCopyText('cv_report_url', cvReportUrl)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-md transition-colors whitespace-nowrap cursor-pointer"
                >
                  {copiedKey === 'cv_report_url' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-400" />
                      Tersalin!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      Salin Link
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
            <button
              type="button"
              onClick={() =>
                handleCopyText(
                  'wa_cv_report',
                  [
                    '*LINK PENGISIAN REPORT CV DIVISI 1 — BAZAR TANJUNG*',
                    `Halo ${linkCvStaff.trim() || `Tim ${CV_BRAND_CATALOG[linkCvBrand].label}`}, silakan isi laporan penjualan harian *Dept ${CV_BRAND_CATALOG[linkCvBrand].deptCode} · Brand ${CV_BRAND_CATALOG[linkCvBrand].label}* melalui link berikut:`,
                    '',
                    cvReportUrl,
                    '',
                    'Data *Sales, Qty, dan Nota* akan langsung terupdate ke Report CV, Total DP+CV & Dashboard. Terima kasih!',
                  ].join('\n')
                )
              }
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors cursor-pointer"
            >
              {copiedKey === 'wa_cv_report' ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  Undangan WA Tersalin!
                </>
              ) : (
                <>
                  <Share2 className="w-3.5 h-3.5" />
                  Salin Undangan WA
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => {
                if (linkCvDateMode === 'CALENDAR' && linkCvCalendarIso) {
                  setCvFormCalendarIso(linkCvCalendarIso);
                }
                setCvFormBrandScope(linkCvBrand);
                if (linkCvStaff.trim()) setCvFormStaff(linkCvStaff.trim());
                setSubmitSuccessData(null);
                setActiveView('cv_report_form');
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-purple-800 bg-purple-50 hover:bg-purple-100 border border-purple-200 rounded-lg transition-colors cursor-pointer"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              Buka Form Report CV
            </button>
          </div>
        </div>

        {/* Card 4: Link Pengisian Data CV (Inventory, Promo, Reason/Kendala & Action Plan) */}
        <div className="lg:col-span-6 bg-white border border-sky-200 rounded-xl p-5 flex flex-col justify-between space-y-5">
          <div className="space-y-4">
            <div className="flex items-start justify-between gap-3 border-b border-slate-100 pb-3.5">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-lg bg-sky-50 text-sky-700 flex items-center justify-center shrink-0">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-bold text-slate-900">
                    4. Link Pengisian Data CV (Inventory, Promo &amp; Action Plan)
                  </h2>
                  <p className="text-xs text-slate-500">
                    Form update Stock Awal, Barang Datang, Keluar, Promo, Kendala &amp; Plan Per Week / Per Monthly
                  </p>
                </div>
              </div>
            </div>

            <div className="p-3.5 bg-sky-50/60 border border-sky-200 rounded-xl space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="inline-flex items-center gap-1.5 text-xs font-bold text-sky-950">
                  <PackageCheck className="w-4 h-4 text-sky-700" />
                  <span>Parameter Periode &amp; Brand Data CV</span>
                </span>
                <div className="inline-flex items-center gap-1 bg-white p-0.5 rounded-lg border border-sky-200 text-[10px]">
                  <button
                    type="button"
                    onClick={() => setLinkDataCvPeriodType('weekly')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkDataCvPeriodType === 'weekly'
                        ? 'bg-sky-700 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Per Week
                  </button>
                  <button
                    type="button"
                    onClick={() => setLinkDataCvPeriodType('monthly')}
                    className={`px-2 py-0.5 rounded font-bold cursor-pointer ${
                      linkDataCvPeriodType === 'monthly'
                        ? 'bg-sky-700 text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Per Monthly
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    {linkDataCvPeriodType === 'weekly' ? 'Pilih Week' : 'Pilih Bulan'}
                  </label>
                  {linkDataCvPeriodType === 'weekly' ? (
                    <select
                      value={linkDataCvWeekNum}
                      onChange={(e) => setLinkDataCvWeekNum(Number(e.target.value))}
                      className="w-full px-2 py-1.5 text-xs font-bold bg-white border border-sky-300 rounded-lg text-slate-900"
                    >
                      {ANNUAL_2026_WEEKS.map((w) => (
                        <option key={w.weekNumber} value={w.weekNumber}>
                          W-{w.weekNumber} ({w.periodeText})
                        </option>
                      ))}
                    </select>
                  ) : (
                    <select
                      value={linkDataCvMonthKey}
                      onChange={(e) => setLinkDataCvMonthKey(e.target.value)}
                      className="w-full px-2 py-1.5 text-xs font-bold bg-white border border-sky-300 rounded-lg text-slate-900"
                    >
                      {Object.entries(MONTH_NAMES_ID).map(([mKey, mName]) => (
                        <option key={mKey} value={mKey}>
                          {mName} 2026
                        </option>
                      ))}
                    </select>
                  )}
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Pilih Brand CV
                  </label>
                  <select
                    value={linkDataCvBrand}
                    onChange={(e) =>
                      setLinkDataCvBrand(e.target.value as CvInventoryBrandKey)
                    }
                    className="w-full px-2 py-1.5 text-xs font-semibold bg-white border border-slate-200 rounded-lg"
                  >
                    <option value="GRAPHIS">6201 · Graphis</option>
                    <option value="SIMPLICITY">6244 · Simplicity</option>
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">
                    Petugas (Opsional)
                  </label>
                  <input
                    type="text"
                    value={linkDataCvStaff}
                    onChange={(e) => setLinkDataCvStaff(e.target.value)}
                    placeholder="Nama petugas..."
                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-sky-600"
                  />
                </div>
              </div>
            </div>

            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
              <div className="text-[11px] font-semibold text-slate-500">
                Link Pengisian Data CV (Siap Dibagikan):
              </div>
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  readOnly
                  value={dataCvUrl}
                  onClick={(e) => (e.target as HTMLInputElement).select()}
                  className="flex-1 px-3 py-2 text-xs font-mono-num text-slate-800 bg-white border border-slate-200 rounded-md select-all"
                />
                <button
                  type="button"
                  onClick={() => handleCopyText('data_cv_url', dataCvUrl)}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-slate-900 hover:bg-slate-800 rounded-md transition-colors whitespace-nowrap cursor-pointer"
                >
                  {copiedKey === 'data_cv_url' ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-400" />
                      Tersalin!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      Salin Link
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
            <button
              type="button"
              onClick={() =>
                handleCopyText(
                  'wa_data_cv',
                  [
                    '*LINK PENGISIAN DATA CV (STOCK, PROMO & PLAN) — BAZAR TANJUNG*',
                    `Halo ${linkDataCvStaff.trim() || `Tim ${CV_BRAND_CATALOG[linkDataCvBrand].label}`}, silakan update data operasional *Dept ${CV_BRAND_CATALOG[linkDataCvBrand].deptCode} · Brand ${CV_BRAND_CATALOG[linkDataCvBrand].label}* (*${
                      linkDataCvPeriodType === 'weekly'
                        ? `Week ${linkDataCvWeekNum}`
                        : MONTH_NAMES_ID[linkDataCvMonthKey]
                    }*) melalui link berikut:`,
                    '',
                    dataCvUrl,
                    '',
                    'Isi *Stock Awal, Barang Datang, Barang Keluar, Promo, Reason/Kendala & Action Plan*. Stock Akhir akan terhitung otomatis!',
                  ].join('\n')
                )
              }
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 rounded-lg transition-colors cursor-pointer"
            >
              {copiedKey === 'wa_data_cv' ? (
                <>
                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                  Undangan WA Tersalin!
                </>
              ) : (
                <>
                  <Share2 className="w-3.5 h-3.5" />
                  Salin Undangan WA
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => {
                setDataCvFormPeriodType(linkDataCvPeriodType);
                setDataCvFormWeekNum(linkDataCvWeekNum);
                setDataCvFormMonthKey(linkDataCvMonthKey);
                setDataCvFormBrandScope(linkDataCvBrand);
                if (linkDataCvStaff.trim()) setDataCvFormStaff(linkDataCvStaff.trim());
                setSubmitSuccessData(null);
                setActiveView('data_cv_form');
              }}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-sky-800 bg-sky-50 hover:bg-sky-100 border border-sky-200 rounded-lg transition-colors cursor-pointer"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              Buka Form Data CV
            </button>
          </div>
        </div>
      </div>

      {/* Riwayat Pengisian Report Masuk (Submission History Table) */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Clock className="w-4 h-4 text-slate-500" />
            <h2 className="text-sm font-bold text-slate-900">
              Riwayat Pengisian Report Masuk ({submissionLogs.length})
            </h2>
          </div>
          <span className="text-xs text-slate-500">
            Setiap pengisian melalui link otomatis tercatat di sini
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                <th className="py-2.5 px-4 font-semibold">Waktu Input</th>
                <th className="py-2.5 px-3 font-semibold">Petugas</th>
                <th className="py-2.5 px-3 font-semibold">Minggu / Tanggal</th>
                <th className="py-2.5 px-3 font-semibold">Kategori / Shift</th>
                <th className="py-2.5 px-3 font-semibold text-right">Sales / Qty / Trx / Stock</th>
                <th className="py-2.5 px-4 font-semibold">Catatan Reason / Ringkasan</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {submissionLogs.map((log) => (
                <tr key={log.id} className="hover:bg-slate-50/80">
                  <td className="py-3 px-4 font-mono-num text-slate-500 whitespace-nowrap">
                    {log.timestamp}
                  </td>
                  <td className="py-3 px-3 font-bold text-slate-900 whitespace-nowrap">
                    {log.submittedBy}
                  </td>
                  <td className="py-3 px-3 whitespace-nowrap">
                    <div className="font-semibold text-slate-800">{log.weekTitle}</div>
                    {log.tanggal && (
                      <div className="font-mono-num text-slate-500">{log.tanggal}</div>
                    )}
                  </td>
                  <td className="py-3 px-3 whitespace-nowrap font-medium text-slate-700">
                    {log.reportType === 'cv_report' ? (
                      <span className="font-bold text-purple-700">
                        {log.shift || 'Harian CV'}
                      </span>
                    ) : log.reportType === 'data_cv' ? (
                      <span className="font-bold text-sky-700">
                        {log.shift || 'Data CV Ops'}
                      </span>
                    ) : (
                      <span className="font-bold text-emerald-700">
                        {log.shift && log.shift !== 'Review Bazar' ? log.shift : 'Data DP'}
                      </span>
                    )}
                  </td>
                  <td className="py-3 px-3 text-right font-mono-num whitespace-nowrap">
                    {log.reportType === 'shift_sales' ? (
                      <div>
                        <span className="font-bold text-slate-900">
                          Rp {formatNumberId(log.salesAmount || 0)}
                        </span>
                        <div className="text-[11px] text-slate-500">
                          {log.qtyAmount || 0} pcs · {log.trxAmount || 0} trx
                          {(log.trafficAmount || 0) > 0 && (
                            <span className="text-indigo-700 font-semibold">
                              {' '}
                              · {formatNumberId(log.trafficAmount)} org
                            </span>
                          )}
                        </div>
                      </div>
                    ) : log.reportType === 'cv_report' ? (
                      <div>
                        <span className="font-bold text-purple-900">
                          Rp {formatNumberId(log.salesAmount || 0)}
                        </span>
                        <div className="text-[11px] text-slate-500">
                          {log.qtyAmount || 0} pcs · {log.trxAmount || 0} nota
                        </div>
                      </div>
                    ) : log.reportType === 'data_cv' ? (
                      <div>
                        <span className="font-bold text-sky-900">
                          Akhir: {formatNumberId(log.salesAmount || 0)} pcs
                        </span>
                        <div className="text-[11px] text-slate-500">
                          Sales {formatNumberId(log.qtyAmount || 0)} pcs
                        </div>
                      </div>
                    ) : (
                      <span className="text-emerald-700 font-semibold">
                        Update Data DP
                      </span>
                    )}
                  </td>
                  <td className="py-3 px-4 text-slate-600 max-w-md">
                    {log.summaryNote}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
