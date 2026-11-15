/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 * 
 * Date formatting and calendar filter utilities for Bazar Matahari Tanjung
 */

import { getIndoDayName, ANNUAL_2026_WEEKS, WeeklySalesTable } from '../data/initialData';

/**
 * Converts 'DD/MM/YYYY' to 'YYYY-MM-DD' (HTML input type="date" value)
 */
export function slashDateToIso(slash: string): string {
  if (!slash) return '2026-10-05';
  const parts = slash.trim().split('/');
  if (parts.length !== 3) return '2026-10-05';
  const dd = parts[0].padStart(2, '0');
  const mm = parts[1].padStart(2, '0');
  const yyyy = parts[2].length === 4 ? parts[2] : `20${parts[2]}`;
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * Converts 'YYYY-MM-DD' (HTML input type="date") to 'DD/MM/YYYY'
 */
export function isoToSlashDate(iso: string): string {
  if (!iso) return '05/10/2026';
  const parts = iso.trim().split('-');
  if (parts.length !== 3) return '05/10/2026';
  const [yyyy, mm, dd] = parts;
  return `${dd.padStart(2, '0')}/${mm.padStart(2, '0')}/${yyyy}`;
}

/**
 * Shifts an ISO date by deltaDays (+1 for tomorrow, -1 for yesterday)
 */
export function shiftIsoDate(iso: string, deltaDays: number): string {
  const [y, m, d] = (iso || '2026-10-05').split('-').map(Number);
  const dt = new Date(y || 2026, (m || 10) - 1, d || 5);
  dt.setDate(dt.getDate() + deltaDays);
  const yyyy = String(dt.getFullYear());
  const mm = String(dt.getMonth() + 1).padStart(2, '0');
  const dd = String(dt.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

const MONTH_NAMES_INDO: Record<string, string> = {
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

/**
 * Returns e.g. "Senin, 05 Oktober 2026"
 */
export function formatIndoDateLong(dateStr: string): string {
  let slash = dateStr;
  if (dateStr.includes('-')) {
    slash = isoToSlashDate(dateStr);
  }
  const parts = slash.split('/');
  if (parts.length !== 3) return slash;
  const [dd, mm, yyyy] = parts;
  const dayName = getIndoDayName(slash);
  const monthName = MONTH_NAMES_INDO[mm] || mm;
  return `${dayName}, ${dd} ${monthName} ${yyyy}`;
}

/**
 * Helper to match a date to a weekly table or schedule item
 */
export function findWeekAndDayIndex(
  dateSlash: string,
  tables: WeeklySalesTable[]
): { tableIndex: number; dayIndex: number; weekNumber: number } | null {
  for (let tIdx = 0; tIdx < tables.length; tIdx++) {
    const table = tables[tIdx];
    const dIdx = table.days.findIndex((d) => d.tanggal === dateSlash);
    if (dIdx >= 0) {
      return { tableIndex: tIdx, dayIndex: dIdx, weekNumber: table.weekNumber };
    }
  }
  return null;
}
