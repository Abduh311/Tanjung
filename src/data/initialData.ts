import { PERSISTED_STORE_SNAPSHOT } from './persistedStore.ts';

export interface ShiftEntry {
  nama: string;
  sales: number | null;
  qty: number | null;
  trx: number | null;
  traffic?: number | null;
}

export interface DailySalesRecord {
  id: string;
  tanggal: string;
  pagi: ShiftEntry;
  siang: ShiftEntry;
  reasonPagi: string;
  reasonSiang: string;
}

export interface WeeklySalesTable {
  id: string;
  weekNumber: number;
  title: string; // e.g. "TOTAL WEEKLY 41"
  periode?: string; // e.g. "01 Oktober - 07 Oktober 2026"
  startDateIso?: string; // e.g. "2026-10-01"
  reasonSiangLabel: string; // e.g. "Reason / Event Shift Pagi" (exact as screenshot) or "Reason / Event Shift Siang"
  days: DailySalesRecord[];
}

export interface ReviewDivisionRow {
  id: string;
  label: string;
  isItalic?: boolean;
  items: string[];
}

export interface ReviewSection {
  id: string;
  code: string; // e.g., "A. MERCHANDISE", "B. PROMO", "C. PLATE/PEOPLE", "D. ACTION PLAN", "REASON/KENDALA", "NEED SUPPORT"
  isMainHeader: boolean;
  directItems?: string[]; // when section row itself has items on the right column (like C, D, REASON, NEED SUPPORT)
  subRows?: ReviewDivisionRow[]; // when section has sub-rows (like A. MERCHANDISE and B. PROMO)
}

export interface ReviewMonitoringDoc {
  id: string;
  weekNumber?: number;
  title: string;
  periode: string;
  sections: ReviewSection[];
}

export interface ReportSubmissionLog {
  id: string;
  timestamp: string;
  reportType: 'shift_sales' | 'review_monitoring' | 'cv_report' | 'data_cv';
  submittedBy: string;
  weekTitle?: string;
  tanggal?: string;
  shift?:
    | 'Pagi'
    | 'Siang'
    | 'Pagi & Siang'
    | 'Lembur'
    | 'Harian CV'
    | 'Data CV Weekly'
    | 'Data CV Monthly'
    | string;
  salesAmount?: number;
  qtyAmount?: number;
  trxAmount?: number;
  trafficAmount?: number;
  summaryNote: string;
}

export const PERSISTED_STORE_UPDATED_AT: number = Number(
  PERSISTED_STORE_SNAPSHOT.updatedAt || 0
);

export const INITIAL_SUBMISSION_LOGS: ReportSubmissionLog[] =
  (PERSISTED_STORE_SNAPSHOT.submissionLogs as unknown as ReportSubmissionLog[]) || [];

export const INITIAL_WEEKLY_TABLES: WeeklySalesTable[] =
  (PERSISTED_STORE_SNAPSHOT.weeklyTables as unknown as WeeklySalesTable[]) || [];

export const INITIAL_REVIEW_DOCS: ReviewMonitoringDoc[] =
  (PERSISTED_STORE_SNAPSHOT.reviewDocs as unknown as ReviewMonitoringDoc[]) || [
  {
    id: 'review-sep-w4',
    weekNumber: 39,
    title: 'REPORT BAZAR MATAHARI — Tanjung (PT MDS Retailing Tbk)',
    periode: 'WEEKLY 39 (21 September - 27 September 2026)',
    sections: [
      {
        id: 'sec-a',
        code: 'A. MERCHANDISE',
        isMainHeader: true,
        subRows: [
          {
            id: 'merch-div1',
            label: 'Divisi 1',
            isItalic: true,
            items: [],
          },
          {
            id: 'merch-div2',
            label: 'Divisi 2',
            isItalic: true,
            items: [],
          },
          {
            id: 'merch-div3',
            label: 'Divisi 3',
            isItalic: true,
            items: [],
          },
          {
            id: 'merch-div4',
            label: 'Divisi 4',
            isItalic: true,
            items: [],
          },
          {
            id: 'merch-fast',
            label: 'FAST MOVING',
            isItalic: false,
            items: [],
          },
          {
            id: 'merch-slow',
            label: 'SLOW MOVING',
            isItalic: false,
            items: [],
          },
        ],
      },
      {
        id: 'sec-b',
        code: 'B.PROMO',
        isMainHeader: true,
        subRows: [
          {
            id: 'promo-div1',
            label: 'Divisi 1',
            isItalic: true,
            items: [],
          },
          {
            id: 'promo-div2',
            label: 'Divisi 2',
            isItalic: true,
            items: [],
          },
          {
            id: 'promo-div3',
            label: 'Divisi 3',
            isItalic: true,
            items: [],
          },
          {
            id: 'promo-div4',
            label: 'Divisi 4',
            isItalic: true,
            items: [],
          },
        ],
      },
      {
        id: 'sec-c',
        code: 'C. PLATE/PEOPLE',
        isMainHeader: true,
        directItems: [],
      },
      {
        id: 'sec-d',
        code: 'D. ACTION PLAN',
        isMainHeader: true,
        directItems: [],
      },
      {
        id: 'sec-reason',
        code: 'REASON/KENDALA',
        isMainHeader: true,
        directItems: [],
      },
      {
        id: 'sec-support',
        code: 'NEED SUPPORT',
        isMainHeader: true,
        directItems: [],
      },
    ],
  },
];

function isShiftFilled(shift: ShiftEntry | undefined): boolean {
  if (!shift) return false;
  return (
    (shift.sales !== null && shift.sales !== undefined && shift.sales > 0) ||
    (shift.qty !== null && shift.qty !== undefined && shift.qty > 0) ||
    (shift.trx !== null && shift.trx !== undefined && shift.trx > 0) ||
    (shift.traffic !== null && shift.traffic !== undefined && shift.traffic > 0)
  );
}

/**
 * Non-destructively merges two WeeklySalesTable arrays during automatic startup hydration
 * so that filled shift records from either source are never lost when AI Studio reloads.
 */
export function mergeWeeklyTablesNonDestructive(
  primary: WeeklySalesTable[],
  secondary: WeeklySalesTable[]
): WeeklySalesTable[] {
  const mapByWeek = new Map<number, WeeklySalesTable>();
  for (const sec of secondary || []) {
    if (sec && typeof sec.weekNumber === 'number') {
      mapByWeek.set(sec.weekNumber, JSON.parse(JSON.stringify(sec)));
    }
  }
  for (const pri of primary || []) {
    if (!pri || typeof pri.weekNumber !== 'number') continue;
    const existing = mapByWeek.get(pri.weekNumber);
    if (!existing) {
      mapByWeek.set(pri.weekNumber, JSON.parse(JSON.stringify(pri)));
      continue;
    }
    const mergedDays = (pri.days || []).map((pDay, idx) => {
      const sDay =
        existing.days.find((d) => d.tanggal === pDay.tanggal) || existing.days[idx];
      if (!sDay) return pDay;
      const pPagiFilled = isShiftFilled(pDay.pagi);
      const sPagiFilled = isShiftFilled(sDay.pagi);
      const pSiangFilled = isShiftFilled(pDay.siang);
      const sSiangFilled = isShiftFilled(sDay.siang);
      const mergedPagi = pPagiFilled
        ? {
            ...pDay.pagi,
            traffic:
              pDay.pagi.traffic !== undefined && pDay.pagi.traffic !== null
                ? pDay.pagi.traffic
                : (sDay.pagi?.traffic ?? null),
          }
        : sPagiFilled
          ? sDay.pagi
          : pDay.pagi;
      const mergedSiang = pSiangFilled
        ? {
            ...pDay.siang,
            traffic:
              pDay.siang.traffic !== undefined && pDay.siang.traffic !== null
                ? pDay.siang.traffic
                : (sDay.siang?.traffic ?? null),
          }
        : sSiangFilled
          ? sDay.siang
          : pDay.siang;
      return {
        ...pDay,
        pagi: mergedPagi,
        siang: mergedSiang,
        reasonPagi:
          pDay.reasonPagi && pDay.reasonPagi.trim() !== ''
            ? pDay.reasonPagi
            : sDay.reasonPagi || '',
        reasonSiang:
          pDay.reasonSiang && pDay.reasonSiang.trim() !== ''
            ? pDay.reasonSiang
            : sDay.reasonSiang || '',
      };
    });
    mapByWeek.set(pri.weekNumber, {
      ...pri,
      days: mergedDays,
    });
  }
  return normalizeAndRepairWeeklyTables(Array.from(mapByWeek.values()));
}

export function formatNumberId(val: number | null | undefined): string {
  if (val === null || val === undefined || Number.isNaN(val)) return '';
  return new Intl.NumberFormat('en-US', {
    maximumFractionDigits: 0,
  }).format(Math.round(val));
}

export function calculateDayMetrics(day: DailySalesRecord) {
  const pagiSales = day.pagi.sales ?? 0;
  const siangSales = day.siang.sales ?? 0;
  const totalSales = pagiSales + siangSales;

  const pagiQty = day.pagi.qty ?? 0;
  const siangQty = day.siang.qty ?? 0;
  const totalQty = pagiQty + siangQty;

  const pagiTrx = day.pagi.trx ?? 0;
  const siangTrx = day.siang.trx ?? 0;
  const totalTrx = pagiTrx + siangTrx;

  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : null;
  const bigSize = totalTrx > 0 ? Math.round(totalSales / totalTrx) : null;

  return {
    totalSales,
    totalQty,
    totalTrx,
    pricePoint,
    bigSize,
  };
}

export function calculateWeekMetrics(table: WeeklySalesTable) {
  let weeklySales = 0;
  let weeklyQty = 0;
  let weeklyTrx = 0;
  let weeklyPagiSales = 0;
  let weeklySiangSales = 0;

  for (const day of table.days) {
    const m = calculateDayMetrics(day);
    weeklySales += m.totalSales;
    weeklyQty += m.totalQty;
    weeklyTrx += m.totalTrx;
    weeklyPagiSales += day.pagi.sales ?? 0;
    weeklySiangSales += day.siang.sales ?? 0;
  }

  const weeklyPricePoint = weeklyQty > 0 ? Math.round(weeklySales / weeklyQty) : null;
  const weeklyBigSize = weeklyTrx > 0 ? Math.round(weeklySales / weeklyTrx) : null;

  return {
    weeklySales,
    weeklyQty,
    weeklyTrx,
    weeklyPagiSales,
    weeklySiangSales,
    weeklyPricePoint,
    weeklyBigSize,
  };
}

const INDO_MONTHS: Record<string, string> = {
  '01': 'Januari',
  '02': 'Februari',
  '03': 'Maret',
  '04': 'April',
  '05': 'Mei',
  '06': 'Juni',
  '07': 'Juli',
  '08': 'Agustus',
  '09': 'September',
  '10': 'Oktober',
  '11': 'November',
  '12': 'Desember',
};

export function formatSlashDateToIndo(slashDate: string): string {
  // e.g. "01/10/2026" -> "01 Oktober 2026"
  const parts = (slashDate || '').trim().split('/');
  if (parts.length !== 3) return slashDate;
  const dd = parts[0].padStart(2, '0');
  const mm = parts[1].padStart(2, '0');
  const yyyy = parts[2];
  return `${dd} ${INDO_MONTHS[mm] || mm} ${yyyy}`;
}

export function getWeeklyPeriodeText(table: WeeklySalesTable): string {
  const first = table.days?.[0]?.tanggal || '';
  const last = table.days?.[table.days.length - 1]?.tanggal || '';
  if (first && last) {
    const p1 = first.split('/');
    const p2 = last.split('/');
    if (p1.length === 3 && p2.length === 3) {
      const d1 = p1[0].padStart(2, '0');
      const m1 = INDO_MONTHS[p1[1].padStart(2, '0')] || p1[1];
      const y1 = p1[2];
      const d2 = p2[0].padStart(2, '0');
      const m2 = INDO_MONTHS[p2[1].padStart(2, '0')] || p2[1];
      const y2 = p2[2];
      return y1 !== y2
        ? `${d1} ${m1} ${y1} - ${d2} ${m2} ${y2}`
        : `${d1} ${m1} - ${d2} ${m2} ${y2}`;
    }
  }
  if (table.periode && table.periode.trim() !== '') {
    return table.periode;
  }
  if (!first || !last) return 'Periode Belum Diatur';
  return `${first} - ${last}`;
}

export function generateSevenSlashDatesFromIso(
  startDateIso: string,
  dayCount: number = 7
): {
  slashDates: string[];
  periodeText: string;
  dayCycleText: string;
} {
  // startDateIso e.g. "2026-10-01"
  const [y, m, d] = startDateIso.split('-').map(Number);
  const base = new Date(y, (m || 1) - 1, d || 1);
  const slashDates: string[] = [];
  const count = Math.max(1, Math.min(14, dayCount));

  for (let i = 0; i < count; i++) {
    const dt = new Date(base);
    dt.setDate(base.getDate() + i);
    const dd = String(dt.getDate()).padStart(2, '0');
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const yyyy = String(dt.getFullYear());
    slashDates.push(`${dd}/${mm}/${yyyy}`);
  }

  const firstParts = slashDates[0].split('/');
  const lastParts = slashDates[slashDates.length - 1].split('/');
  const m1 = INDO_MONTHS[firstParts[1]] || firstParts[1];
  const m2 = INDO_MONTHS[lastParts[1]] || lastParts[1];
  const periodeText =
    firstParts[2] !== lastParts[2]
      ? `${firstParts[0]} ${m1} ${firstParts[2]} - ${lastParts[0]} ${m2} ${lastParts[2]}`
      : `${firstParts[0]} ${m1} - ${lastParts[0]} ${m2} ${lastParts[2]}`;

  const firstDayName = getIndoDayName(slashDates[0]);
  const lastDayName = getIndoDayName(slashDates[slashDates.length - 1]);
  const dayCycleText = `${firstDayName} – ${lastDayName} (${slashDates.length} Hari)`;

  return { slashDates, periodeText, dayCycleText };
}

const INDO_DAY_NAMES = [
  'Minggu',
  'Senin',
  'Selasa',
  'Rabu',
  'Kamis',
  'Jumat',
  'Sabtu',
];

export function getIndoDayName(slashDate: string): string {
  // Expected DD/MM/YYYY
  const parts = (slashDate || '').trim().split('/');
  if (parts.length !== 3) return 'Hari';
  const dd = Number(parts[0]);
  const mm = Number(parts[1]);
  const yyyy = Number(parts[2]);
  if (!dd || !mm || !yyyy) return 'Hari';
  const dt = new Date(yyyy, mm - 1, dd);
  if (Number.isNaN(dt.getTime())) return 'Hari';
  return INDO_DAY_NAMES[dt.getDay()] || 'Hari';
}

export function isWeekendDate(slashDate: string): boolean {
  const dayName = getIndoDayName(slashDate);
  return dayName === 'Sabtu' || dayName === 'Minggu';
}

export function getWeeklyDayCycleText(table: WeeklySalesTable): string {
  if (!table.days || table.days.length === 0) return '0 Hari';
  const firstDay = getIndoDayName(table.days[0].tanggal);
  const lastDay = getIndoDayName(table.days[table.days.length - 1].tanggal);
  return `${firstDay} – ${lastDay} (${table.days.length} Hari)`;
}

export interface AnnualWeekScheduleItem {
  weekNumber: number;
  year: number;
  title: string; // e.g. "TOTAL WEEKLY 1"
  startDateIso: string; // Monday YYYY-MM-DD
  endDateIso: string; // Sunday YYYY-MM-DD
  seninSlash: string; // DD/MM/YYYY
  mingguSlash: string; // DD/MM/YYYY
  slashDates: string[]; // 7 dates [Senin, Selasa, Rabu, Kamis, Jumat, Sabtu, Minggu]
  periodeText: string; // e.g. "05 Oktober - 11 Oktober 2026"
  monthKey: string; // "01" .. "12" (based on Thursday of that week)
  monthName: string; // "Januari" .. "Desember"
  quarter: 'Q1' | 'Q2' | 'Q3' | 'Q4';
}

export function generateAnnualWeeksMondayToSunday(
  year: number = 2026
): AnnualWeekScheduleItem[] {
  // Find Monday of Week 1 (week containing Jan 4 of `year`)
  const jan4 = new Date(year, 0, 4);
  const dayOfWeek = jan4.getDay(); // 0=Sun, 1=Mon, ... 6=Sat
  const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
  const firstMonday = new Date(year, 0, 4 + diffToMonday);

  const schedule: AnnualWeekScheduleItem[] = [];
  const totalWeeks = 53;

  for (let w = 1; w <= totalWeeks; w++) {
    const monDate = new Date(firstMonday);
    monDate.setDate(firstMonday.getDate() + (w - 1) * 7);

    // Stop if week 53 starts in the next year after Jan 3
    if (w === 53 && monDate.getFullYear() > year) {
      break;
    }

    const startIso = `${monDate.getFullYear()}-${String(
      monDate.getMonth() + 1
    ).padStart(2, '0')}-${String(monDate.getDate()).padStart(2, '0')}`;

    const sunDate = new Date(monDate);
    sunDate.setDate(monDate.getDate() + 6);
    const endIso = `${sunDate.getFullYear()}-${String(
      sunDate.getMonth() + 1
    ).padStart(2, '0')}-${String(sunDate.getDate()).padStart(2, '0')}`;

    // Use Thursday of the week (or Sunday if week 1) to classify the primary month in `year`
    const thuDate = new Date(monDate);
    thuDate.setDate(monDate.getDate() + 3);
    const monthNum =
      thuDate.getFullYear() < year
        ? 1
        : thuDate.getFullYear() > year
          ? 12
          : thuDate.getMonth() + 1;
    const monthKey = String(monthNum).padStart(2, '0');
    const monthName = INDO_MONTHS[monthKey] || 'Januari';
    const quarter: 'Q1' | 'Q2' | 'Q3' | 'Q4' =
      monthNum <= 3
        ? 'Q1'
        : monthNum <= 6
          ? 'Q2'
          : monthNum <= 9
            ? 'Q3'
            : 'Q4';

    const { slashDates, periodeText } = generateSevenSlashDatesFromIso(startIso, 7);

    schedule.push({
      weekNumber: w,
      year,
      title: `TOTAL WEEKLY ${w}`,
      startDateIso: startIso,
      endDateIso: endIso,
      seninSlash: slashDates[0],
      mingguSlash: slashDates[6],
      slashDates,
      periodeText,
      monthKey,
      monthName,
      quarter,
    });
  }

  return schedule;
}

export const ANNUAL_2026_WEEKS: AnnualWeekScheduleItem[] =
  generateAnnualWeeksMondayToSunday(2026);

export function createWeeklyTableFromAnnualSchedule(
  item: AnnualWeekScheduleItem
): WeeklySalesTable {
  return {
    id: `weekly-${item.weekNumber}`,
    weekNumber: item.weekNumber,
    title: `TOTAL WEEKLY ${item.weekNumber}`,
    periode: item.periodeText,
    startDateIso: item.startDateIso,
    reasonSiangLabel: 'Reason / Event Shift Siang',
    days: item.slashDates.map((tgl, i) => ({
      id: `w${item.weekNumber}-d${i + 1}`,
      tanggal: tgl,
      pagi: {
        nama: '',
        sales: null,
        qty: null,
        trx: null,
        traffic: null,
      },
      siang: {
        nama: '',
        sales: null,
        qty: null,
        trx: null,
        traffic: null,
      },
      reasonPagi: '',
      reasonSiang: '',
    })),
  };
}

export function getTableMonthKey(table: WeeklySalesTable): string {
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === table.weekNumber);
  if (sched) return sched.monthKey;
  // Fallback from thursday (index 3) or first day
  const refSlash = table.days[3]?.tanggal || table.days[0]?.tanggal || '';
  const parts = refSlash.split('/');
  if (parts.length === 3) {
    return parts[1].padStart(2, '0');
  }
  return '10';
}

export function getTableMonthName(table: WeeklySalesTable): string {
  const mKey = getTableMonthKey(table);
  return INDO_MONTHS[mKey] || 'Oktober';
}

export interface TableWeekendInfo {
  weekendNumberInMonth: number; // 1 .. 5
  weekendMonthKey: string; // '01' .. '12'
  weekendMonthName: string; // 'Januari' .. 'Desember'
  sabtuTanggal: string; // DD/MM/YYYY
  mingguTanggal: string; // DD/MM/YYYY
  dateRangeText: string; // e.g. "03/10/2026 – 04/10/2026"
  shortDateRange: string; // e.g. "03/10 – 04/10/2026"
  shortBadge: string; // e.g. "Weekend Ke-1 Oktober"
  fullLabel: string; // e.g. "Weekend Ke-1 Oktober (03/10 – 04/10/2026) [W40]"
  sabtuSales: number;
  mingguSales: number;
  weekendTotalSales: number;
  weekendTotalQty: number;
  weekendTotalTrx: number;
}

export function getTableWeekendInfo(table: WeeklySalesTable): TableWeekendInfo {
  const days = Array.isArray(table.days) ? table.days : [];
  const sabtuDay =
    days.find((d) => getIndoDayName(d.tanggal) === 'Sabtu') || days[5] || days[days.length - 2];
  const mingguDay =
    days.find((d) => getIndoDayName(d.tanggal) === 'Minggu') || days[6] || days[days.length - 1];

  const refDate = sabtuDay?.tanggal || mingguDay?.tanggal || days[0]?.tanggal || '03/10/2026';
  const parts = refDate.trim().split('/');
  const dd = Number(parts[0]) || 1;
  const mm = (parts[1] || getTableMonthKey(table)).padStart(2, '0');
  const weekendNumberInMonth = Math.max(1, Math.min(5, Math.ceil(dd / 7)));
  const weekendMonthName = INDO_MONTHS[mm] || getTableMonthName(table);

  const sabtuTanggal = sabtuDay?.tanggal || '';
  const mingguTanggal = mingguDay?.tanggal || '';
  const dateRangeText =
    sabtuTanggal && mingguTanggal
      ? `${sabtuTanggal} – ${mingguTanggal}`
      : sabtuTanggal || mingguTanggal || '-';
  const shortDateRange =
    sabtuTanggal && mingguTanggal
      ? `${sabtuTanggal.slice(0, 5)} – ${mingguTanggal}`
      : dateRangeText;

  const sabtuSales = (sabtuDay?.pagi?.sales ?? 0) + (sabtuDay?.siang?.sales ?? 0);
  const mingguSales = (mingguDay?.pagi?.sales ?? 0) + (mingguDay?.siang?.sales ?? 0);
  const weekendTotalSales = sabtuSales + mingguSales;
  const weekendTotalQty =
    (sabtuDay?.pagi?.qty ?? 0) +
    (sabtuDay?.siang?.qty ?? 0) +
    (mingguDay?.pagi?.qty ?? 0) +
    (mingguDay?.siang?.qty ?? 0);
  const weekendTotalTrx =
    (sabtuDay?.pagi?.trx ?? 0) +
    (sabtuDay?.siang?.trx ?? 0) +
    (mingguDay?.pagi?.trx ?? 0) +
    (mingguDay?.siang?.trx ?? 0);

  const shortBadge = `Weekend Ke-${weekendNumberInMonth} ${weekendMonthName}`;
  const fullLabel = `Weekend Ke-${weekendNumberInMonth} ${weekendMonthName} (${shortDateRange}) [W${table.weekNumber}]`;

  return {
    weekendNumberInMonth,
    weekendMonthKey: mm,
    weekendMonthName,
    sabtuTanggal,
    mingguTanggal,
    dateRangeText,
    shortDateRange,
    shortBadge,
    fullLabel,
    sabtuSales,
    mingguSales,
    weekendTotalSales,
    weekendTotalQty,
    weekendTotalTrx,
  };
}

export function getAnnualWeekWeekendInfo(weekNumber: number): TableWeekendInfo | null {
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNumber);
  if (!sched) return null;
  const tempTable = createWeeklyTableFromAnnualSchedule(sched);
  return getTableWeekendInfo(tempTable);
}

export function createEmptyInitialWeeklyTables(): WeeklySalesTable[] {
  return ANNUAL_2026_WEEKS.map(createWeeklyTableFromAnnualSchedule);
}

export function normalizeAndRepairWeeklyTables(
  list: WeeklySalesTable[]
): WeeklySalesTable[] {
  if (!Array.isArray(list) || list.length === 0) {
    return INITIAL_WEEKLY_TABLES;
  }

  const scheduleMap = new Map<number, AnnualWeekScheduleItem>();
  ANNUAL_2026_WEEKS.forEach((w) => scheduleMap.set(w.weekNumber, w));

  let repaired: WeeklySalesTable[] = list.map((t): WeeklySalesTable => {
    const sched = scheduleMap.get(t.weekNumber);
    const baseSource = t;

    // Only align if table still uses legacy prototype Thursday dates ('01/10/2026' or '08/10/2026')
    const firstDateStr = (baseSource.days?.[0]?.tanggal || '').trim();
    const isLegacyThursdayProto =
      (baseSource.weekNumber === 40 && firstDateStr === '01/10/2026') ||
      (baseSource.weekNumber === 41 && firstDateStr === '08/10/2026');
    const needsMondayAlign = isLegacyThursdayProto && Boolean(sched);

    const sourceDays = Array.isArray(baseSource.days) ? baseSource.days : [];
    const isOldDemoSeedW41 =
      baseSource.weekNumber === 41 &&
      sourceDays[0]?.pagi?.sales === 705900 &&
      sourceDays[0]?.siang?.sales === 2339100 &&
      sourceDays[1]?.pagi?.sales === 1854250 &&
      !sourceDays[0]?.pagi?.nama;

    const nextDays = sourceDays.map((d, idx) => ({
      ...d,
      tanggal:
        needsMondayAlign && sched && sched.slashDates[idx]
          ? sched.slashDates[idx]
          : d.tanggal,
      pagi: isOldDemoSeedW41
        ? { nama: '', sales: null, qty: null, trx: null, traffic: null }
        : {
            nama: d.pagi?.nama ?? '',
            sales: d.pagi?.sales !== undefined ? d.pagi.sales : null,
            qty: d.pagi?.qty !== undefined ? d.pagi.qty : null,
            trx: d.pagi?.trx !== undefined ? d.pagi.trx : null,
            traffic: d.pagi?.traffic !== undefined ? d.pagi.traffic : null,
          },
      siang: isOldDemoSeedW41
        ? { nama: '', sales: null, qty: null, trx: null, traffic: null }
        : {
            nama: d.siang?.nama ?? '',
            sales: d.siang?.sales !== undefined ? d.siang.sales : null,
            qty: d.siang?.qty !== undefined ? d.siang.qty : null,
            trx: d.siang?.trx !== undefined ? d.siang.trx : null,
            traffic: d.siang?.traffic !== undefined ? d.siang.traffic : null,
          },
      reasonPagi: isOldDemoSeedW41 ? '' : (d.reasonPagi ?? ''),
      reasonSiang: isOldDemoSeedW41 ? '' : (d.reasonSiang ?? ''),
    }));

    return {
      ...baseSource,
      id: baseSource.id || `weekly-${baseSource.weekNumber}`,
      periode:
        needsMondayAlign && sched
          ? sched.periodeText
          : baseSource.periode || getWeeklyPeriodeText(baseSource),
      startDateIso:
        needsMondayAlign && sched
          ? sched.startDateIso
          : baseSource.startDateIso,
      reasonSiangLabel:
        !baseSource.reasonSiangLabel ||
        baseSource.reasonSiangLabel.trim() === 'Reason / Event Shift Pagi'
          ? 'Reason / Event Shift Siang'
          : baseSource.reasonSiangLabel,
      days: nextDays,
    };
  });

  // Deduplicate by weekNumber (keep latest entry)
  const byWeekNum = new Map<number, WeeklySalesTable>();
  for (const item of repaired) {
    byWeekNum.set(item.weekNumber, item);
  }
  repaired = Array.from(byWeekNum.values());

  return repaired.sort((a, b) => a.weekNumber - b.weekNumber);
}

export function formatReviewBulletItem(rawLine: string): string {
  const trimmed = (rawLine || '').trim();
  if (!trimmed) return '';
  const cleaned = trimmed
    .replace(
      /^\s*(?:\d+\s*(?:[.)](?!\d)\s*(?:[-•*]+\s*)?|[-•*]+\s*))+|^\s*[-•*]+\s*/,
      ''
    )
    .trim();
  if (!cleaned) return '';
  return `- ${cleaned}`;
}

export function normalizeReviewItemsList(items: string[] | undefined): string[] {
  if (!Array.isArray(items)) return [];
  const formatted = items
    .map((line) => formatReviewBulletItem(line))
    .filter((line) => line.length > 0);
  return formatted;
}

export function getReviewDocWeekNumber(doc: ReviewMonitoringDoc): number | null {
  if (typeof doc.weekNumber === 'number' && doc.weekNumber >= 1 && doc.weekNumber <= 53) {
    return doc.weekNumber;
  }
  const idMatch = (doc.id || '').match(/review-w(\d+)/i);
  if (idMatch) {
    const w = Number(idMatch[1]);
    if (w >= 1 && w <= 53) return w;
  }
  const perMatch = (doc.periode || '').match(/WEEKLY\s*(\d+)/i);
  if (perMatch) {
    const w = Number(perMatch[1]);
    if (w >= 1 && w <= 53) return w;
  }
  const perLower = (doc.periode || '').toLowerCase();
  for (const w of ANNUAL_2026_WEEKS) {
    if (perLower.includes(w.periodeText.toLowerCase())) {
      return w.weekNumber;
    }
  }
  if (doc.id === 'review-sep-w4') return 39;
  return null;
}

export function getReviewDocMonthKey(doc: ReviewMonitoringDoc): string {
  const wk = getReviewDocWeekNumber(doc);
  if (wk) {
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
    if (sched) return sched.monthKey;
  }
  const perLower = (doc.periode || '').toLowerCase();
  for (const [mKey, mName] of Object.entries(INDO_MONTHS)) {
    if (perLower.includes(mName.toLowerCase())) {
      return mKey;
    }
  }
  return '09';
}

export function getReviewDocWeeklyLabel(doc: ReviewMonitoringDoc): string {
  const wk = getReviewDocWeekNumber(doc);
  if (wk) {
    const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
    if (sched) {
      return `WEEKLY ${wk} (${sched.periodeText})`;
    }
    return `WEEKLY ${wk} (${doc.periode.replace(/^Periode\s*:\s*/i, '')})`;
  }
  return doc.periode.replace(/^Periode\s*:\s*/i, '');
}

export function countReviewDocItems(doc: ReviewMonitoringDoc | undefined): number {
  if (!doc || !Array.isArray(doc.sections)) return 0;
  let count = 0;
  for (const sec of doc.sections) {
    if (Array.isArray(sec.subRows)) {
      for (const sr of sec.subRows) {
        if (Array.isArray(sr.items)) {
          count += sr.items.filter((it) => (it || '').trim().replace(/^[-•*]+\s*/, '').trim().length > 0).length;
        }
      }
    }
    if (Array.isArray(sec.directItems)) {
      count += sec.directItems.filter((it) => (it || '').trim().replace(/^[-•*]+\s*/, '').trim().length > 0).length;
    }
  }
  return count;
}

export function isReviewDocFilled(doc: ReviewMonitoringDoc | undefined): boolean {
  return countReviewDocItems(doc) > 0;
}

export function createEmptyReviewSections(template?: ReviewSection[]): ReviewSection[] {
  const baseSections = template && template.length > 0 ? template : INITIAL_REVIEW_DOCS[0].sections;
  return baseSections.map((sec) => ({
    id: sec.id,
    code: sec.code,
    isMainHeader: sec.isMainHeader,
    directItems: sec.directItems !== undefined ? [] : undefined,
    subRows: sec.subRows
      ? sec.subRows.map((sr) => ({
          id: sr.id,
          label: sr.label,
          isItalic: sr.isItalic,
          items: [],
        }))
      : undefined,
  }));
}

export function createReviewDocForWeek(
  weekNumber: number,
  sourceTemplate?: ReviewMonitoringDoc,
  copyContent: boolean = false
): ReviewMonitoringDoc {
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === weekNumber);
  const periodeStr = sched
    ? `WEEKLY ${weekNumber} (${sched.periodeText})`
    : `WEEKLY ${weekNumber}`;
  const base = sourceTemplate || INITIAL_REVIEW_DOCS[0];

  return {
    id: `review-w${weekNumber}`,
    weekNumber,
    title:
      base?.title ||
      'REPORT BAZAR MATAHARI — Tanjung (PT MDS Retailing Tbk)',
    periode: periodeStr,
    sections: copyContent
      ? (base?.sections || []).map((sec) => ({
          ...sec,
          directItems: sec.directItems ? [...sec.directItems] : undefined,
          subRows: sec.subRows
            ? sec.subRows.map((sr) => ({
                ...sr,
                items: [...sr.items],
              }))
            : undefined,
        }))
      : createEmptyReviewSections(base?.sections),
  };
}

export function generateReviewDocFromWeeklySales(
  table: WeeklySalesTable,
  existingDoc?: ReviewMonitoringDoc
): ReviewMonitoringDoc {
  const wk = table.weekNumber;
  const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === wk);
  const periodeStr = sched
    ? `WEEKLY ${wk} (${sched.periodeText})`
    : `WEEKLY ${wk} (${getWeeklyPeriodeText(table)})`;

  const wm = calculateWeekMetrics(table);
  const dailyReasons: string[] = [];
  const fastKeywords = new Set<string>();
  const kendalaList: string[] = [];
  const customerNotes: string[] = [];

  table.days.forEach((d) => {
    const dName = getIndoDayName(d.tanggal);
    const rP = (d.reasonPagi || '').trim();
    const rS = (d.reasonSiang || '').trim();
    if (rP) {
      dailyReasons.push(`- ${dName} (${d.tanggal}) Pagi [${d.pagi.nama || 'Petugas'}]: ${rP}`);
      if (/broken size|ukuran|tidak tersedia|kabut asap|sepi/i.test(rP)) {
        kendalaList.push(`- ${dName} (${d.tanggal}) Shift Pagi: ${rP}`);
      }
      if (/produk yang|laku|promo|sepatu|sandal|children|b1g/i.test(rP)) {
        fastKeywords.add(`- ${dName} (${d.tanggal}): ${rP}`);
      }
    }
    if (rS) {
      dailyReasons.push(`- ${dName} (${d.tanggal}) Siang [${d.siang.nama || 'Petugas'}]: ${rS}`);
      if (/broken size|ukuran|tidak tersedia|kabut asap|sepi/i.test(rS)) {
        kendalaList.push(`- ${dName} (${d.tanggal}) Shift Siang: ${rS}`);
      }
      if (/produk yang|laku|promo|sepatu|sandal|children|b1g/i.test(rS)) {
        fastKeywords.add(`- ${dName} (${d.tanggal}): ${rS}`);
      }
    }
  });

  if (wm.weeklySales > 0) {
    customerNotes.push(
      `- Total pencapaian Sales Weekly ${wk}: Rp ${formatNumberId(wm.weeklySales)} (Qty: ${formatNumberId(wm.weeklyQty)} pcs, Trx: ${formatNumberId(wm.weeklyTrx)} trx)`
    );
    customerNotes.push(
      `- Kontribusi Shift Pagi: Rp ${formatNumberId(wm.weeklyPagiSales)} | Shift Siang: Rp ${formatNumberId(wm.weeklySiangSales)}`
    );
  }
  if (dailyReasons.length > 0) {
    customerNotes.push(...dailyReasons.slice(0, 4));
  }

  const baseDoc = existingDoc || createReviewDocForWeek(wk);
  return {
    ...baseDoc,
    id: baseDoc.id || `review-w${wk}`,
    weekNumber: wk,
    periode: periodeStr,
    sections: baseDoc.sections.map((sec) => {
      if (sec.code.includes('MERCHANDISE') && sec.subRows) {
        return {
          ...sec,
          subRows: sec.subRows.map((sr) => {
            if (sr.label.toUpperCase().includes('FAST MOVING') && fastKeywords.size > 0) {
              return {
                ...sr,
                items: Array.from(fastKeywords),
              };
            }
            return sr;
          }),
        };
      }
      if ((sec.code.includes('PLATE') || sec.code.includes('PLACE')) && customerNotes.length > 0) {
        return {
          ...sec,
          directItems: customerNotes,
        };
      }
      if (sec.code.includes('REASON') && (kendalaList.length > 0 || dailyReasons.length > 0)) {
        return {
          ...sec,
          directItems: kendalaList.length > 0 ? kendalaList : dailyReasons,
        };
      }
      return sec;
    }),
  };
}

function hasNonEmptyReviewLines(items: string[] | undefined): boolean {
  if (!Array.isArray(items)) return false;
  return items.some(
    (it) =>
      (it || '')
        .replace(/[\t\r]/g, '')
        .trim()
        .replace(/^[-•*]+\s*/, '')
        .trim().length > 0
  );
}

function cleanReviewLines(items: string[] | undefined): string[] {
  if (!Array.isArray(items)) return [];
  return items.map((it) => (it || '').replace(/[\t\r]+$/g, ''));
}

export function mergeTwoReviewDocsNonDestructive(
  primary: ReviewMonitoringDoc,
  secondary: ReviewMonitoringDoc
): ReviewMonitoringDoc {
  const templateSections = INITIAL_REVIEW_DOCS[0].sections;
  const pSecs = Array.isArray(primary.sections) && primary.sections.length > 0
    ? primary.sections
    : templateSections;
  const sSecs = Array.isArray(secondary.sections) ? secondary.sections : [];

  const mergedSections: ReviewSection[] = pSecs.map((pSec) => {
    const sSec = sSecs.find(
      (s) =>
        s.id === pSec.id ||
        (s.code || '').trim().toUpperCase() === (pSec.code || '').trim().toUpperCase()
    );
    if (!sSec) return pSec;

    if (Array.isArray(pSec.subRows)) {
      const sSubRows = Array.isArray(sSec.subRows) ? sSec.subRows : [];
      return {
        ...pSec,
        subRows: pSec.subRows.map((pSub) => {
          const sSub = sSubRows.find(
            (sr) =>
              sr.id === pSub.id ||
              (sr.label || '').trim().toLowerCase() ===
                (pSub.label || '').trim().toLowerCase()
          );
          const pHas = hasNonEmptyReviewLines(pSub.items);
          const sHas = sSub ? hasNonEmptyReviewLines(sSub.items) : false;
          return {
            ...pSub,
            items: pHas
              ? cleanReviewLines(pSub.items)
              : sHas && sSub
                ? cleanReviewLines(sSub.items)
                : cleanReviewLines(pSub.items),
          };
        }),
      };
    }

    const pDirectHas = hasNonEmptyReviewLines(pSec.directItems);
    const sDirectHas = hasNonEmptyReviewLines(sSec.directItems);
    return {
      ...pSec,
      directItems: pDirectHas
        ? cleanReviewLines(pSec.directItems)
        : sDirectHas
          ? cleanReviewLines(sSec.directItems)
          : cleanReviewLines(pSec.directItems),
    };
  });

  // Also include any custom sections present in secondary but not in primary
  for (const sSec of sSecs) {
    const alreadyExists = mergedSections.some(
      (m) =>
        m.id === sSec.id ||
        (m.code || '').trim().toUpperCase() === (sSec.code || '').trim().toUpperCase()
    );
    if (!alreadyExists) {
      mergedSections.push(sSec);
    }
  }

  return {
    ...secondary,
    ...primary,
    weekNumber:
      getReviewDocWeekNumber(primary) ??
      getReviewDocWeekNumber(secondary) ??
      primary.weekNumber,
    sections: mergedSections,
  };
}

export function mergeReviewDocsNonDestructive(
  primary: ReviewMonitoringDoc[],
  secondary: ReviewMonitoringDoc[]
): ReviewMonitoringDoc[] {
  const byWeek = new Map<number, ReviewMonitoringDoc>();
  const unnumbered: ReviewMonitoringDoc[] = [];

  for (const secDoc of secondary || []) {
    if (!secDoc) continue;
    const wk = getReviewDocWeekNumber(secDoc);
    if (wk) {
      const existing = byWeek.get(wk);
      byWeek.set(
        wk,
        existing ? mergeTwoReviewDocsNonDestructive(secDoc, existing) : secDoc
      );
    } else {
      unnumbered.push(secDoc);
    }
  }

  for (const priDoc of primary || []) {
    if (!priDoc) continue;
    const wk = getReviewDocWeekNumber(priDoc);
    if (wk) {
      const existing = byWeek.get(wk);
      byWeek.set(
        wk,
        existing ? mergeTwoReviewDocsNonDestructive(priDoc, existing) : priDoc
      );
    } else {
      unnumbered.push(priDoc);
    }
  }

  const combined = [
    ...Array.from(byWeek.values()).sort(
      (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
    ),
    ...unnumbered,
  ];
  return normalizeReviewDocs(combined);
}

export function normalizeReviewDocs(
  docs: ReviewMonitoringDoc[]
): ReviewMonitoringDoc[] {
  if (!Array.isArray(docs) || docs.length === 0) {
    return INITIAL_REVIEW_DOCS;
  }
  const templateSections = INITIAL_REVIEW_DOCS[0].sections;

  const normalizedList = docs.map((doc) => {
    const detectedWeek = getReviewDocWeekNumber(doc);
    const rawSections = Array.isArray(doc.sections) ? doc.sections : [];

    // Only migrate if doc is using the old legacy section code "1. DIVISI & PRODUK LARIS"
    const isAncientLegacyFormat = rawSections.some((s) =>
      (s.code || '').toUpperCase().includes('DIVISI & PRODUK LARIS')
    );

    if (isAncientLegacyFormat) {
      let fastItems: string[] = [];
      let slowItems: string[] = [];
      let promoItems: string[] = [];
      let kendalaItems: string[] = [];
      let actionItems: string[] = [];

      for (const s of rawSections) {
        const codeUpper = (s.code || '').toUpperCase();
        if (s.subRows) {
          for (const sr of s.subRows) {
            const lblUpper = (sr.label || '').toUpperCase();
            if (lblUpper.includes('FAST')) {
              fastItems.push(...normalizeReviewItemsList(sr.items));
            } else if (lblUpper.includes('SLOW')) {
              slowItems.push(...normalizeReviewItemsList(sr.items));
            }
          }
        }
        if (s.directItems) {
          if (codeUpper.includes('PROMO')) {
            promoItems.push(...normalizeReviewItemsList(s.directItems));
          } else if (codeUpper.includes('KENDALA') || codeUpper.includes('OPERASIONAL')) {
            kendalaItems.push(...normalizeReviewItemsList(s.directItems));
          } else if (codeUpper.includes('ACTION') || codeUpper.includes('TINDAK')) {
            actionItems.push(...normalizeReviewItemsList(s.directItems));
          }
        }
      }

      const migratedSections: ReviewSection[] = templateSections.map((tSec) => {
        if (tSec.id === 'sec-a' && tSec.subRows) {
          return {
            ...tSec,
            subRows: tSec.subRows.map((sr) => {
              if (sr.id === 'merch-fast' && fastItems.length > 0) {
                return { ...sr, items: fastItems };
              }
              if (sr.id === 'merch-slow' && slowItems.length > 0) {
                return { ...sr, items: slowItems };
              }
              return { ...sr, items: [...sr.items] };
            }),
          };
        }
        if (tSec.id === 'sec-b' && tSec.subRows) {
          return {
            ...tSec,
            subRows: tSec.subRows.map((sr, idx) =>
              idx === 0 && promoItems.length > 0
                ? { ...sr, items: promoItems }
                : { ...sr, items: [...sr.items] }
            ),
          };
        }
        if (tSec.id === 'sec-d' && actionItems.length > 0) {
          return { ...tSec, directItems: actionItems };
        }
        if (tSec.id === 'sec-reason' && kendalaItems.length > 0) {
          return { ...tSec, directItems: kendalaItems };
        }
        return {
          ...tSec,
          directItems: tSec.directItems ? [...tSec.directItems] : undefined,
          subRows: tSec.subRows
            ? tSec.subRows.map((sr) => ({ ...sr, items: [...sr.items] }))
            : undefined,
        };
      });

      return {
        ...doc,
        weekNumber: detectedWeek ?? doc.weekNumber,
        sections: migratedSections,
      };
    }

    // Ensure all standard sections exist while preserving every existing section's items
    const baseSections =
      rawSections.length > 0 ? rawSections : createEmptyReviewSections(templateSections);
    const completeSections: ReviewSection[] = baseSections.map((sec) => ({
      ...sec,
      directItems: Array.isArray(sec.directItems)
        ? cleanReviewLines(sec.directItems)
        : sec.subRows
          ? undefined
          : [],
      subRows: Array.isArray(sec.subRows)
        ? sec.subRows.map((sr) => ({
            ...sr,
            isItalic: sr.label.toLowerCase().startsWith('divisi')
              ? true
              : Boolean(sr.isItalic),
            items: Array.isArray(sr.items) ? cleanReviewLines(sr.items) : [],
          }))
        : undefined,
    }));

    for (const tmplSec of templateSections) {
      const hasSec = completeSections.some(
        (s) =>
          s.id === tmplSec.id ||
          (s.code || '').trim().toUpperCase() ===
            (tmplSec.code || '').trim().toUpperCase()
      );
      if (!hasSec) {
        completeSections.push({
          id: tmplSec.id,
          code: tmplSec.code,
          isMainHeader: tmplSec.isMainHeader,
          directItems: tmplSec.directItems !== undefined ? [] : undefined,
          subRows: tmplSec.subRows
            ? tmplSec.subRows.map((sr) => ({
                id: sr.id,
                label: sr.label,
                isItalic: sr.isItalic,
                items: [],
              }))
            : undefined,
        });
      }
    }

    return {
      ...doc,
      id: doc.id || (detectedWeek ? `review-w${detectedWeek}` : 'review-doc'),
      weekNumber: detectedWeek ?? doc.weekNumber,
      sections: completeSections,
    };
  });

  // Deduplicate by weekNumber non-destructively so a partial doc never shadows a filled doc
  const byWeek = new Map<number, ReviewMonitoringDoc>();
  const others: ReviewMonitoringDoc[] = [];
  for (const item of normalizedList) {
    const wk = getReviewDocWeekNumber(item);
    if (wk) {
      const prev = byWeek.get(wk);
      byWeek.set(wk, prev ? mergeTwoReviewDocsNonDestructive(item, prev) : item);
    } else {
      others.push(item);
    }
  }

  return [
    ...Array.from(byWeek.values()).sort(
      (a, b) => (getReviewDocWeekNumber(a) || 999) - (getReviewDocWeekNumber(b) || 999)
    ),
    ...others,
  ];
}

export interface MonthCalendarDayItem {
  dayOfMonth: number; // 1 .. 31
  tanggal: string; // DD/MM/YYYY
  dayName: string; // Senin .. Minggu
  isWeekend: boolean;
  weekNumber: number | null;
  weekTitle: string;
  pagiNama: string;
  pagiSales: number;
  pagiQty: number;
  pagiTrx: number;
  siangNama: string;
  siangSales: number;
  siangQty: number;
  siangTrx: number;
  totalSales: number;
  totalQty: number;
  totalTrx: number;
  pricePoint: number | null;
  bigSize: number | null;
  reasonPagi: string;
  reasonSiang: string;
}

export interface MonthCalendarSummaryItem {
  mm: string; // "01" .. "12"
  year: number;
  monthKey: string; // "01/2026" .. "12/2026"
  monthName: string; // "Januari" .. "Desember"
  lastDay: number; // 28, 30, or 31
  startSlash: string; // "01/MM/2026"
  endSlash: string; // "30/MM/2026" or "31/MM/2026"
  dateRangeLabel: string; // e.g. "01 – 31 Oktober 2026"
  shortRangeLabel: string; // e.g. "1–31 Okt"
  slashDateRange: string; // e.g. "01/10/2026 – 31/10/2026"
  pagiSales: number;
  siangSales: number;
  totalSales: number;
  pagiQty: number;
  siangQty: number;
  totalQty: number;
  pagiTrx: number;
  siangTrx: number;
  totalTrx: number;
  weekdaySales: number;
  weekendSales: number;
  activeDays: number;
  totalCalendarDays: number;
  pricePoint: number | null;
  bigSize: number | null;
  days: MonthCalendarDayItem[];
}

export const ALL_12_MONTH_KEYS = [
  '01',
  '02',
  '03',
  '04',
  '05',
  '06',
  '07',
  '08',
  '09',
  '10',
  '11',
  '12',
] as const;

export function getMonthLastDay(mm: string, year: number = 2026): number {
  const mNum = Number(mm);
  if (!mNum || mNum < 1 || mNum > 12) return 31;
  return new Date(year, mNum, 0).getDate();
}

export function buildMonthCalendarSummary(
  tables: WeeklySalesTable[],
  mmRaw: string,
  year: number = 2026
): MonthCalendarSummaryItem {
  const mm = mmRaw.split('/')[0].padStart(2, '0');
  const monthName = INDO_MONTHS[mm] || 'Oktober';
  const lastDay = getMonthLastDay(mm, year);
  const startSlash = `01/${mm}/${year}`;
  const endSlash = `${String(lastDay).padStart(2, '0')}/${mm}/${year}`;
  const dateRangeLabel = `01 – ${lastDay} ${monthName} ${year}`;
  const shortRangeLabel = `1–${lastDay} ${monthName.slice(0, 3)}`;
  const slashDateRange = `${startSlash} – ${endSlash}`;

  // Index existing daily records by date DD/MM/YYYY
  const recordByDate = new Map<
    string,
    { day: DailySalesRecord; weekNumber: number; weekTitle: string }
  >();
  for (const t of tables) {
    for (const d of t.days) {
      const key = (d.tanggal || '').trim();
      if (!key) continue;
      const existing = recordByDate.get(key);
      const dm = calculateDayMetrics(d);
      const existingSales = existing
        ? calculateDayMetrics(existing.day).totalSales
        : -1;
      if (!existing || dm.totalSales >= existingSales) {
        recordByDate.set(key, {
          day: d,
          weekNumber: t.weekNumber,
          weekTitle: t.title,
        });
      }
    }
  }

  // Also map date -> schedule weekNumber for dates not yet opened in tables
  const schedWeekByDate = new Map<string, number>();
  for (const w of ANNUAL_2026_WEEKS) {
    for (const dt of w.slashDates) {
      schedWeekByDate.set(dt, w.weekNumber);
    }
  }

  const days: MonthCalendarDayItem[] = [];
  let pagiSales = 0;
  let siangSales = 0;
  let pagiQty = 0;
  let siangQty = 0;
  let pagiTrx = 0;
  let siangTrx = 0;
  let weekdaySales = 0;
  let weekendSales = 0;
  let activeDays = 0;

  for (let dayNum = 1; dayNum <= lastDay; dayNum++) {
    const dd = String(dayNum).padStart(2, '0');
    const tanggal = `${dd}/${mm}/${year}`;
    const dayName = getIndoDayName(tanggal);
    const isWknd = isWeekendDate(tanggal);
    const found = recordByDate.get(tanggal);
    const schedWk = schedWeekByDate.get(tanggal) ?? null;
    const wkNum = found ? found.weekNumber : schedWk;
    const wkTitle = found
      ? found.weekTitle
      : wkNum
        ? `TOTAL WEEKLY ${wkNum}`
        : '-';

    const pS = found ? (found.day.pagi.sales ?? 0) : 0;
    const pQ = found ? (found.day.pagi.qty ?? 0) : 0;
    const pT = found ? (found.day.pagi.trx ?? 0) : 0;
    const sS = found ? (found.day.siang.sales ?? 0) : 0;
    const sQ = found ? (found.day.siang.qty ?? 0) : 0;
    const sT = found ? (found.day.siang.trx ?? 0) : 0;

    const totS = pS + sS;
    const totQ = pQ + sQ;
    const totT = pT + sT;
    const pp = totQ > 0 ? Math.round(totS / totQ) : null;
    const bs = totT > 0 ? Math.round(totS / totT) : null;

    pagiSales += pS;
    siangSales += sS;
    pagiQty += pQ;
    siangQty += sQ;
    pagiTrx += pT;
    siangTrx += sT;
    if (isWknd) weekendSales += totS;
    else weekdaySales += totS;
    if (totS > 0) activeDays += 1;

    days.push({
      dayOfMonth: dayNum,
      tanggal,
      dayName,
      isWeekend: isWknd,
      weekNumber: wkNum,
      weekTitle: wkTitle,
      pagiNama: found?.day.pagi.nama || '',
      pagiSales: pS,
      pagiQty: pQ,
      pagiTrx: pT,
      siangNama: found?.day.siang.nama || '',
      siangSales: sS,
      siangQty: sQ,
      siangTrx: sT,
      totalSales: totS,
      totalQty: totQ,
      totalTrx: totT,
      pricePoint: pp,
      bigSize: bs,
      reasonPagi: found?.day.reasonPagi || '',
      reasonSiang: found?.day.reasonSiang || '',
    });
  }

  const totalSales = pagiSales + siangSales;
  const totalQty = pagiQty + siangQty;
  const totalTrx = pagiTrx + siangTrx;
  const pricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : null;
  const bigSize = totalTrx > 0 ? Math.round(totalSales / totalTrx) : null;

  return {
    mm,
    year,
    monthKey: `${mm}/${year}`,
    monthName,
    lastDay,
    startSlash,
    endSlash,
    dateRangeLabel,
    shortRangeLabel,
    slashDateRange,
    pagiSales,
    siangSales,
    totalSales,
    pagiQty,
    siangQty,
    totalQty,
    pagiTrx,
    siangTrx,
    totalTrx,
    weekdaySales,
    weekendSales,
    activeDays,
    totalCalendarDays: lastDay,
    pricePoint,
    bigSize,
    days,
  };
}

export function buildAll12MonthsCalendarSummary(
  tables: WeeklySalesTable[],
  year: number = 2026
): {
  months: MonthCalendarSummaryItem[];
  grandTotals: {
    pagiSales: number;
    siangSales: number;
    totalSales: number;
    pagiQty: number;
    siangQty: number;
    totalQty: number;
    pagiTrx: number;
    siangTrx: number;
    totalTrx: number;
    weekdaySales: number;
    weekendSales: number;
    activeDays: number;
    totalCalendarDays: number;
    activeMonths: number;
    pricePoint: number | null;
    bigSize: number | null;
  };
} {
  const months = ALL_12_MONTH_KEYS.map((mm) =>
    buildMonthCalendarSummary(tables, mm, year)
  );
  const grandTotals = months.reduce(
    (acc, m) => {
      acc.pagiSales += m.pagiSales;
      acc.siangSales += m.siangSales;
      acc.totalSales += m.totalSales;
      acc.pagiQty += m.pagiQty;
      acc.siangQty += m.siangQty;
      acc.totalQty += m.totalQty;
      acc.pagiTrx += m.pagiTrx;
      acc.siangTrx += m.siangTrx;
      acc.totalTrx += m.totalTrx;
      acc.weekdaySales += m.weekdaySales;
      acc.weekendSales += m.weekendSales;
      acc.activeDays += m.activeDays;
      acc.totalCalendarDays += m.totalCalendarDays;
      if (m.totalSales > 0) acc.activeMonths += 1;
      return acc;
    },
    {
      pagiSales: 0,
      siangSales: 0,
      totalSales: 0,
      pagiQty: 0,
      siangQty: 0,
      totalQty: 0,
      pagiTrx: 0,
      siangTrx: 0,
      totalTrx: 0,
      weekdaySales: 0,
      weekendSales: 0,
      activeDays: 0,
      totalCalendarDays: 0,
      activeMonths: 0,
      pricePoint: null as number | null,
      bigSize: null as number | null,
    }
  );
  grandTotals.pricePoint =
    grandTotals.totalQty > 0
      ? Math.round(grandTotals.totalSales / grandTotals.totalQty)
      : null;
  grandTotals.bigSize =
    grandTotals.totalTrx > 0
      ? Math.round(grandTotals.totalSales / grandTotals.totalTrx)
      : null;

  return { months, grandTotals };
}





