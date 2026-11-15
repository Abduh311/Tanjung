/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * Tabel Inventory, Promo, Reason/Kendala & Action Plan — CV Divisi 1
 * Mendukung tampilan Per Week (Senin–Minggu, W1–W53) dan Per Monthly (Jan–Des 2026)
 * Divisi 1 (Young Fashion, Ladies & Intimate): Dept 6201 Graphis & Dept 6244 Simplicity
 */

import React, { useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  CvDailyRecord,
  CvInventoryBrandKey,
  CvInventoryOpsEntry,
  CvInventoryPeriodType,
  CV_BRAND_CATALOG,
  CV_DIVISI_DEFAULT,
  CV_DIVISI_FULL,
  MONTH_NAMES_ID,
  calcCvLiveSalesQtyMap,
  createDefaultCvOpsEntry,
  getCvOpsEntryId,
} from '../data/cvDiv1Data';
import { ANNUAL_2026_WEEKS, formatNumberId } from '../data/initialData';
import { isoToSlashDate } from '../utils/dateUtils';
import { copyTextToClipboard } from '../utils/clipboard';
import {
  AlertTriangle,
  ArrowRight,
  Calendar,
  CheckCircle2,
  ClipboardCopy,
  Layers,
  PackageCheck,
  PenSquare,
  RefreshCw,
  Save,
  Search,
  Sparkles,
  Tag,
  TrendingUp,
  X,
} from 'lucide-react';

interface CvInventoryOpsTableSectionProps {
  defaultPeriodType: CvInventoryPeriodType;
  records: CvDailyRecord[];
  entries: CvInventoryOpsEntry[];
  onUpdateEntries: (nextEntries: CvInventoryOpsEntry[]) => void;
  selectedMonthKey: string;
  onChangeMonthKey?: (monthKey: string) => void;
  selectedWeekNumber?: number;
  onSelectWeekNumber?: (weekNumber: number) => void;
  isAdmin: boolean;
  onToast?: (message: string) => void;
}

interface ComputedBrandRow {
  entry: CvInventoryOpsEntry;
  brand: CvInventoryBrandKey;
  brandLabel: string;
  deptCode: '6201' | '6244';
  divisi: string;
  liveSalesQty: number;
  effectiveSalesQty: number;
  isSalesOverridden: boolean;
  stockAwal: number;
  barangDatang: number;
  barangKeluar: number;
  stockAkhir: number;
  sellThroughPct: number;
}

interface ComputedPeriodGroup {
  periodKey: string; // String(weekNumber) or monthKey ('01'..'12')
  weekNumber?: number;
  monthKey: string;
  titleLabel: string; // "WEEK 1 (W-45)" or "November 2026"
  subtitleLabel: string; // "02 Nov – 08 Nov 2026" or "Bulan ke-11 · 30 Hari"
  divisi: string;
  graphis: ComputedBrandRow;
  simplicity: ComputedBrandRow;
  total: {
    stockAwal: number;
    barangDatang: number;
    effectiveSalesQty: number;
    barangKeluar: number;
    stockAkhir: number;
    sellThroughPct: number;
  };
}

function getStockHealthInfo(
  stockAwal: number,
  barangDatang: number,
  salesQty: number,
  barangKeluar: number,
  stockAkhir: number
): { label: string; toneClass: string; kind: 'ok' | 'low' | 'danger' | 'empty' } {
  const hasAnyMovement =
    stockAwal > 0 || barangDatang > 0 || salesQty > 0 || barangKeluar > 0;
  if (!hasAnyMovement) {
    return {
      label: 'Belum Diisi',
      toneClass: 'text-slate-500 bg-slate-100 border-slate-200',
      kind: 'empty',
    };
  }
  if (stockAkhir <= 0) {
    return {
      label: stockAkhir < 0 ? 'Minus / Selisih' : 'Stock Habis',
      toneClass: 'text-rose-700 bg-rose-50 border-rose-200',
      kind: 'danger',
    };
  }
  if (stockAkhir <= 15) {
    return {
      label: 'Menipis',
      toneClass: 'text-amber-700 bg-amber-50 border-amber-200',
      kind: 'low',
    };
  }
  return {
    label: 'Sehat',
    toneClass: 'text-emerald-700 bg-emerald-50 border-emerald-200',
    kind: 'ok',
  };
}

export const CvInventoryOpsTableSection: React.FC<CvInventoryOpsTableSectionProps> = ({
  defaultPeriodType,
  records,
  entries,
  onUpdateEntries,
  selectedMonthKey,
  onChangeMonthKey,
  selectedWeekNumber,
  onSelectWeekNumber,
  isAdmin,
  onToast,
}) => {
  const [periodMode, setPeriodMode] = useState<CvInventoryPeriodType>(defaultPeriodType);
  const [scopeMode, setScopeMode] = useState<'focused' | 'all'>('focused');
  const [brandFilter, setBrandFilter] = useState<'ALL' | CvInventoryBrandKey>('ALL');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [calendarIso, setCalendarIso] = useState<string>('');
  const [isInlineEdit, setIsInlineEdit] = useState<boolean>(false);

  // Modal state for comfortable mobile & full editing
  const [modalTarget, setModalTarget] = useState<{
    periodType: CvInventoryPeriodType;
    periodKey: string;
    brand: CvInventoryBrandKey;
  } | null>(null);
  const [modalDraft, setModalDraft] = useState<{
    divisi: string;
    stockAwal: string;
    barangDatang: string;
    salesQtyMode: 'live' | 'manual';
    salesQtyManual: string;
    barangKeluar: string;
    promo: string;
    reasonKendala: string;
    actionPlan: string;
  }>({
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
  const [modalCalendarIso, setModalCalendarIso] = useState<string>('');

  const liveSalesMap = useMemo(() => calcCvLiveSalesQtyMap(records), [records]);

  const entryMap = useMemo(() => {
    const map = new Map<string, CvInventoryOpsEntry>();
    entries.forEach((e) => map.set(e.id, e));
    return map;
  }, [entries]);

  const calendarResolvedFilter = useMemo(() => {
    if (!calendarIso) return null;
    const parts = calendarIso.split('-');
    const monthKey = parts.length === 3 ? parts[1] : '';
    const slash = isoToSlashDate(calendarIso);
    const matchedWeek = ANNUAL_2026_WEEKS.find(
      (w) =>
        w.slashDates.includes(slash) ||
        (calendarIso >= w.startDateIso && calendarIso <= w.endDateIso)
    );
    return {
      iso: calendarIso,
      monthKey,
      weekNumber: matchedWeek?.weekNumber,
    };
  }, [calendarIso]);

  const periodGroups = useMemo<ComputedPeriodGroup[]>(() => {
    const buildBrandRow = (
      pType: CvInventoryPeriodType,
      pKey: string | number,
      brand: CvInventoryBrandKey
    ): ComputedBrandRow => {
      const id = getCvOpsEntryId(pType, pKey, brand);
      const entry = entryMap.get(id) || createDefaultCvOpsEntry(pType, pKey, brand);
      const liveQty =
        pType === 'weekly'
          ? liveSalesMap.weekly.get(Number(pKey))?.[brand] || 0
          : liveSalesMap.monthly.get(String(pKey).padStart(2, '0'))?.[brand] || 0;
      const isSalesOverridden =
        entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined;
      const effectiveSalesQty = isSalesOverridden
        ? Number(entry.salesQtyOverride) || 0
        : liveQty;
      const stockAwal = Number(entry.stockAwal) || 0;
      const barangDatang = Number(entry.barangDatang) || 0;
      const barangKeluar = Number(entry.barangKeluar) || 0;
      const stockAkhir = stockAwal + barangDatang - effectiveSalesQty - barangKeluar;
      const available = stockAwal + barangDatang;
      const sellThroughPct = available > 0 ? (effectiveSalesQty / available) * 100 : 0;

      return {
        entry,
        brand,
        brandLabel: CV_BRAND_CATALOG[brand].label,
        deptCode: CV_BRAND_CATALOG[brand].deptCode,
        divisi: entry.divisi || CV_DIVISI_DEFAULT,
        liveSalesQty: liveQty,
        effectiveSalesQty,
        isSalesOverridden,
        stockAwal,
        barangDatang,
        barangKeluar,
        stockAkhir,
        sellThroughPct,
      };
    };

    const groups: ComputedPeriodGroup[] = [];

    if (periodMode === 'weekly') {
      const targetMonth = calendarResolvedFilter?.monthKey || selectedMonthKey;
      let targetWeeks = ANNUAL_2026_WEEKS;
      if (calendarResolvedFilter?.weekNumber) {
        targetWeeks = ANNUAL_2026_WEEKS.filter(
          (w) => w.weekNumber === calendarResolvedFilter.weekNumber
        );
      } else if (scopeMode === 'focused') {
        const filtered = ANNUAL_2026_WEEKS.filter(
          (w) =>
            w.monthKey === targetMonth ||
            w.slashDates.some((dt) => dt.split('/')[1] === targetMonth)
        );
        if (filtered.length > 0) targetWeeks = filtered;
      }

      targetWeeks.forEach((w, idx) => {
        const g = buildBrandRow('weekly', w.weekNumber, 'GRAPHIS');
        const s = buildBrandRow('weekly', w.weekNumber, 'SIMPLICITY');
        const totAwal = g.stockAwal + s.stockAwal;
        const totDatang = g.barangDatang + s.barangDatang;
        const totSales = g.effectiveSalesQty + s.effectiveSalesQty;
        const totKeluar = g.barangKeluar + s.barangKeluar;
        const totAkhir = g.stockAkhir + s.stockAkhir;
        const totAvail = totAwal + totDatang;
        groups.push({
          periodKey: String(w.weekNumber),
          weekNumber: w.weekNumber,
          monthKey: w.monthKey,
          titleLabel: `WEEK ${scopeMode === 'focused' ? idx + 1 : w.weekNumber} · W-${w.weekNumber}`,
          subtitleLabel: `${w.periodeText} (${MONTH_NAMES_ID[w.monthKey] || w.monthKey} 2026)`,
          divisi: CV_DIVISI_DEFAULT,
          graphis: g,
          simplicity: s,
          total: {
            stockAwal: totAwal,
            barangDatang: totDatang,
            effectiveSalesQty: totSales,
            barangKeluar: totKeluar,
            stockAkhir: totAkhir,
            sellThroughPct: totAvail > 0 ? (totSales / totAvail) * 100 : 0,
          },
        });
      });
    } else {
      const allMonths = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];
      let targetMonths = allMonths;
      if (calendarResolvedFilter?.monthKey && MONTH_NAMES_ID[calendarResolvedFilter.monthKey]) {
        targetMonths = [calendarResolvedFilter.monthKey];
      } else if (scopeMode === 'focused') {
        targetMonths = [selectedMonthKey];
      }

      targetMonths.forEach((mKey) => {
        const g = buildBrandRow('monthly', mKey, 'GRAPHIS');
        const s = buildBrandRow('monthly', mKey, 'SIMPLICITY');
        const totAwal = g.stockAwal + s.stockAwal;
        const totDatang = g.barangDatang + s.barangDatang;
        const totSales = g.effectiveSalesQty + s.effectiveSalesQty;
        const totKeluar = g.barangKeluar + s.barangKeluar;
        const totAkhir = g.stockAkhir + s.stockAkhir;
        const totAvail = totAwal + totDatang;

        const weeksInMonthCount = ANNUAL_2026_WEEKS.filter(
          (w) => w.monthKey === mKey || w.slashDates.some((dt) => dt.split('/')[1] === mKey)
        ).length;

        groups.push({
          periodKey: mKey,
          monthKey: mKey,
          titleLabel: `${MONTH_NAMES_ID[mKey] || mKey} 2026`,
          subtitleLabel: `Bulan ${mKey} · ${weeksInMonthCount} Week Kalender`,
          divisi: CV_DIVISI_DEFAULT,
          graphis: g,
          simplicity: s,
          total: {
            stockAwal: totAwal,
            barangDatang: totDatang,
            effectiveSalesQty: totSales,
            barangKeluar: totKeluar,
            stockAkhir: totAkhir,
            sellThroughPct: totAvail > 0 ? (totSales / totAvail) * 100 : 0,
          },
        });
      });
    }

    const q = searchQuery.trim().toLowerCase();
    if (!q) return groups;

    return groups.filter((grp) => {
      const hay = [
        grp.titleLabel,
        grp.subtitleLabel,
        grp.divisi,
        grp.graphis.entry.promo,
        grp.graphis.entry.reasonKendala,
        grp.graphis.entry.actionPlan,
        grp.simplicity.entry.promo,
        grp.simplicity.entry.reasonKendala,
        grp.simplicity.entry.actionPlan,
      ]
        .join(' ')
        .toLowerCase();
      return hay.includes(q);
    });
  }, [
    periodMode,
    scopeMode,
    selectedMonthKey,
    calendarResolvedFilter,
    entryMap,
    liveSalesMap,
    searchQuery,
  ]);

  const kpiSummary = useMemo(() => {
    let stockAwal = 0;
    let barangDatang = 0;
    let salesQty = 0;
    let barangKeluar = 0;
    let stockAkhir = 0;

    periodGroups.forEach((grp) => {
      if (brandFilter === 'GRAPHIS') {
        stockAwal += grp.graphis.stockAwal;
        barangDatang += grp.graphis.barangDatang;
        salesQty += grp.graphis.effectiveSalesQty;
        barangKeluar += grp.graphis.barangKeluar;
        stockAkhir += grp.graphis.stockAkhir;
      } else if (brandFilter === 'SIMPLICITY') {
        stockAwal += grp.simplicity.stockAwal;
        barangDatang += grp.simplicity.barangDatang;
        salesQty += grp.simplicity.effectiveSalesQty;
        barangKeluar += grp.simplicity.barangKeluar;
        stockAkhir += grp.simplicity.stockAkhir;
      } else {
        stockAwal += grp.total.stockAwal;
        barangDatang += grp.total.barangDatang;
        salesQty += grp.total.effectiveSalesQty;
        barangKeluar += grp.total.barangKeluar;
        stockAkhir += grp.total.stockAkhir;
      }
    });

    const totalAvail = stockAwal + barangDatang;
    const sellThroughPct = totalAvail > 0 ? (salesQty / totalAvail) * 100 : 0;

    return {
      stockAwal,
      barangDatang,
      salesQty,
      barangKeluar,
      stockAkhir,
      sellThroughPct,
    };
  }, [periodGroups, brandFilter]);

  const upsertEntryField = (
    pType: CvInventoryPeriodType,
    pKey: string,
    brand: CvInventoryBrandKey,
    patch: Partial<CvInventoryOpsEntry>
  ) => {
    const id = getCvOpsEntryId(pType, pKey, brand);
    const existing = entryMap.get(id) || createDefaultCvOpsEntry(pType, pKey, brand);
    const nextEntry: CvInventoryOpsEntry = {
      ...existing,
      ...patch,
      updatedAt: new Date().toISOString(),
    };
    const nextList = entries.some((e) => e.id === id)
      ? entries.map((e) => (e.id === id ? nextEntry : e))
      : [...entries, nextEntry];
    onUpdateEntries(nextList);
  };

  const handleCarryOverNext = (grp: ComputedPeriodGroup) => {
    if (!isAdmin) return;
    const brandsToCopy: CvInventoryBrandKey[] =
      brandFilter === 'ALL' ? ['GRAPHIS', 'SIMPLICITY'] : [brandFilter];

    let nextKey: string | null = null;
    if (periodMode === 'weekly' && grp.weekNumber && grp.weekNumber < 53) {
      nextKey = String(grp.weekNumber + 1);
    } else if (periodMode === 'monthly') {
      const mNum = parseInt(grp.monthKey, 10);
      if (mNum >= 1 && mNum < 12) {
        nextKey = String(mNum + 1).padStart(2, '0');
      }
    }

    if (!nextKey) {
      onToast?.('Sudah berada di periode akhir tahun 2026.');
      return;
    }

    let updated = [...entries];
    brandsToCopy.forEach((b) => {
      const currentRow = b === 'GRAPHIS' ? grp.graphis : grp.simplicity;
      const targetId = getCvOpsEntryId(periodMode, nextKey!, b);
      const existing =
        updated.find((e) => e.id === targetId) ||
        createDefaultCvOpsEntry(periodMode, nextKey!, b);
      const patched: CvInventoryOpsEntry = {
        ...existing,
        stockAwal: Math.max(0, currentRow.stockAkhir),
        updatedAt: new Date().toISOString(),
      };
      if (updated.some((e) => e.id === targetId)) {
        updated = updated.map((e) => (e.id === targetId ? patched : e));
      } else {
        updated.push(patched);
      }
    });

    onUpdateEntries(updated);
    onToast?.(
      periodMode === 'weekly'
        ? `Stock Akhir disalin ke Stock Awal W-${nextKey}!`
        : `Stock Akhir disalin ke Stock Awal Bulan ${MONTH_NAMES_ID[nextKey] || nextKey}!`
    );
  };

  const openEditModal = (
    pType: CvInventoryPeriodType,
    pKey: string,
    brand: CvInventoryBrandKey
  ) => {
    const id = getCvOpsEntryId(pType, pKey, brand);
    const entry = entryMap.get(id) || createDefaultCvOpsEntry(pType, pKey, brand);
    const liveQty =
      pType === 'weekly'
        ? liveSalesMap.weekly.get(Number(pKey))?.[brand] || 0
        : liveSalesMap.monthly.get(String(pKey).padStart(2, '0'))?.[brand] || 0;
    const isManual =
      entry.salesQtyOverride !== null && entry.salesQtyOverride !== undefined;

    if (pType === 'weekly') {
      const sched = ANNUAL_2026_WEEKS.find((w) => w.weekNumber === Number(pKey));
      setModalCalendarIso(sched?.startDateIso || `2026-${selectedMonthKey}-01`);
    } else {
      setModalCalendarIso(`2026-${String(pKey).padStart(2, '0')}-01`);
    }

    setModalTarget({ periodType: pType, periodKey: pKey, brand });
    setModalDraft({
      divisi: entry.divisi || CV_DIVISI_DEFAULT,
      stockAwal: String(entry.stockAwal ?? 0),
      barangDatang: String(entry.barangDatang ?? 0),
      salesQtyMode: isManual ? 'manual' : 'live',
      salesQtyManual: String(isManual ? entry.salesQtyOverride : liveQty),
      barangKeluar: String(entry.barangKeluar ?? 0),
      promo: entry.promo || '',
      reasonKendala: entry.reasonKendala || '',
      actionPlan: entry.actionPlan || '',
    });
  };

  const handleModalCalendarJump = (isoVal: string) => {
    setModalCalendarIso(isoVal);
    if (!isoVal || !modalTarget) return;
    const parts = isoVal.split('-');
    if (parts.length !== 3) return;
    const slash = isoToSlashDate(isoVal);
    if (modalTarget.periodType === 'weekly') {
      const matched = ANNUAL_2026_WEEKS.find(
        (w) => w.slashDates.includes(slash) || (isoVal >= w.startDateIso && isoVal <= w.endDateIso)
      );
      if (matched) {
        openEditModal('weekly', String(matched.weekNumber), modalTarget.brand);
        setModalCalendarIso(isoVal);
      }
    } else if (MONTH_NAMES_ID[parts[1]]) {
      openEditModal('monthly', parts[1], modalTarget.brand);
      setModalCalendarIso(isoVal);
    }
  };

  const handleSaveModal = (e: React.FormEvent) => {
    e.preventDefault();
    if (!modalTarget) return;
    upsertEntryField(modalTarget.periodType, modalTarget.periodKey, modalTarget.brand, {
      divisi: modalDraft.divisi.trim() || CV_DIVISI_DEFAULT,
      stockAwal: Math.max(0, Number(modalDraft.stockAwal) || 0),
      barangDatang: Math.max(0, Number(modalDraft.barangDatang) || 0),
      barangKeluar: Math.max(0, Number(modalDraft.barangKeluar) || 0),
      salesQtyOverride:
        modalDraft.salesQtyMode === 'manual'
          ? Math.max(0, Number(modalDraft.salesQtyManual) || 0)
          : null,
      promo: modalDraft.promo.trim(),
      reasonKendala: modalDraft.reasonKendala.trim(),
      actionPlan: modalDraft.actionPlan.trim(),
    });
    setModalTarget(null);
    onToast?.('Data Inventory, Promo & Evaluasi CV berhasil disimpan!');
  };

  const handleCopyWhatsAppInventory = () => {
    const modeTitle =
      periodMode === 'weekly'
        ? `PER WEEK (${scopeMode === 'focused' ? MONTH_NAMES_ID[selectedMonthKey] : 'SETAHUN 2026'})`
        : `PER MONTHLY (${scopeMode === 'focused' ? MONTH_NAMES_ID[selectedMonthKey] : 'JAN–DES 2026'})`;

    const lines: string[] = [
      `*REPORT INVENTORY & OPS CV — DIVISI 1*`,
      `*Bazar Tanjung · Matahari 311*`,
      `Mode: *${modeTitle}*`,
      `Brand: Dept 6201 Graphis & Dept 6244 Simplicity`,
      `--------------------------------`,
      `Ringkasan Total:`,
      `• Stock Awal   : ${formatNumberId(kpiSummary.stockAwal)} pcs`,
      `• Barang Datang: +${formatNumberId(kpiSummary.barangDatang)} pcs`,
      `• Sales Qty    : -${formatNumberId(kpiSummary.salesQty)} pcs`,
      `• Barang Keluar: -${formatNumberId(kpiSummary.barangKeluar)} pcs`,
      `• *Stock Akhir*: *${formatNumberId(kpiSummary.stockAkhir)} pcs* (ST ${kpiSummary.sellThroughPct.toFixed(1)}%)`,
      ``,
    ];

    const sliceGroups = periodGroups.slice(0, 6);
    sliceGroups.forEach((grp) => {
      lines.push(`*[${grp.titleLabel}]* — ${grp.subtitleLabel}`);
      const brands: ComputedBrandRow[] =
        brandFilter === 'GRAPHIS'
          ? [grp.graphis]
          : brandFilter === 'SIMPLICITY'
          ? [grp.simplicity]
          : [grp.graphis, grp.simplicity];
      brands.forEach((br) => {
        lines.push(
          `  ${br.brandLabel} (${br.deptCode}): Awal ${br.stockAwal} | Dtg +${br.barangDatang} | Sls -${br.effectiveSalesQty} | Klr -${br.barangKeluar} => *Akhir ${br.stockAkhir} pcs*`
        );
        if (br.entry.promo) lines.push(`    Promo: ${br.entry.promo}`);
        if (br.entry.reasonKendala) lines.push(`    Kendala: ${br.entry.reasonKendala}`);
        if (br.entry.actionPlan) lines.push(`    Plan: ${br.entry.actionPlan}`);
      });
    });

    void copyTextToClipboard(lines.join('\n')).then((ok) => {
      onToast?.(
        ok
          ? 'Ringkasan Inventory & Action Plan CV disalin ke WhatsApp!'
          : 'Gagal menyalin teks ke clipboard.'
      );
    });
  };

  const renderBrandRow = (
    grp: ComputedPeriodGroup,
    br: ComputedBrandRow,
    showPeriodCells: boolean,
    rowSpanCount: number,
    isFirstInGroup: boolean
  ) => {
    const health = getStockHealthInfo(
      br.stockAwal,
      br.barangDatang,
      br.effectiveSalesQty,
      br.barangKeluar,
      br.stockAkhir
    );
    const isGraphis = br.brand === 'GRAPHIS';

    return (
      <tr
        key={`${grp.periodKey}-${br.brand}`}
        onClick={() => {
          if (grp.weekNumber && onSelectWeekNumber) {
            onSelectWeekNumber(grp.weekNumber);
          }
        }}
        className={`group border-b border-slate-200 transition-colors hover:bg-slate-50/90 ${
          isGraphis ? 'bg-white' : 'bg-slate-50/40'
        }`}
      >
        {/* 1. Periode (Merged on split mode) */}
        {showPeriodCells && (
          <td
            rowSpan={rowSpanCount}
            className="py-3 px-3.5 align-top border-r border-slate-200 bg-white sticky left-0 z-10 min-w-[155px]"
          >
            <div className="font-bold text-slate-900 text-xs tracking-tight leading-snug">
              {grp.titleLabel}
            </div>
            <div className="text-[11px] text-slate-500 mt-0.5 leading-snug">
              {grp.subtitleLabel}
            </div>
            {isAdmin && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  handleCarryOverNext(grp);
                }}
                className="mt-2 inline-flex items-center gap-1 text-[10px] font-semibold text-indigo-700 hover:text-indigo-900 hover:underline cursor-pointer"
                title="Salin Stock Akhir periode ini menjadi Stock Awal periode berikutnya"
              >
                <ArrowRight className="w-3 h-3" />
                <span>Salin Stock ke Berikutnya</span>
              </button>
            )}
          </td>
        )}

        {/* 2. Divisi */}
        <td className="py-2.5 px-3 align-top border-r border-slate-200 min-w-[130px]">
          <div className="text-xs font-semibold text-slate-900">{br.divisi}</div>
          <div className="text-[10px] text-slate-500 mt-0.5 leading-tight">
            Young Fashion · Ladies
          </div>
        </td>

        {/* 3. Brand */}
        <td className="py-2.5 px-3 align-top border-r border-slate-200 min-w-[135px]">
          <div className="flex items-center gap-2">
            <span
              className={`w-2 h-2 rounded-full shrink-0 ${
                isGraphis ? 'bg-rose-600' : 'bg-indigo-600'
              }`}
            />
            <div>
              <div className="text-xs font-bold text-slate-950 uppercase tracking-wide">
                {br.brandLabel}
              </div>
              <div className="text-[10px] font-mono-num text-slate-500">
                Dept {br.deptCode}
              </div>
            </div>
          </div>
        </td>

        {/* 4a. Stock Awal */}
        <td className="py-2.5 px-2.5 text-right align-top border-r border-slate-100 font-mono-num min-w-[92px]">
          {isAdmin && isInlineEdit ? (
            <input
              type="number"
              min="0"
              aria-label={`Stock Awal ${br.brandLabel} ${grp.titleLabel}`}
              value={br.stockAwal}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  stockAwal: Math.max(0, Number(e.target.value) || 0),
                })
              }
              className="w-20 px-1.5 py-1 text-right text-xs font-bold text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
          ) : (
            <span className="text-xs font-semibold text-slate-800">
              {formatNumberId(br.stockAwal)}
            </span>
          )}
        </td>

        {/* 4b. Barang Datang (+) */}
        <td className="py-2.5 px-2.5 text-right align-top border-r border-slate-100 font-mono-num min-w-[92px]">
          {isAdmin && isInlineEdit ? (
            <input
              type="number"
              min="0"
              aria-label={`Barang Datang ${br.brandLabel} ${grp.titleLabel}`}
              value={br.barangDatang}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  barangDatang: Math.max(0, Number(e.target.value) || 0),
                })
              }
              className="w-20 px-1.5 py-1 text-right text-xs font-bold text-emerald-800 bg-emerald-50/40 border border-emerald-300 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500"
            />
          ) : (
            <span
              className={`text-xs font-semibold ${
                br.barangDatang > 0 ? 'text-emerald-700' : 'text-slate-400'
              }`}
            >
              {br.barangDatang > 0 ? `+${formatNumberId(br.barangDatang)}` : '0'}
            </span>
          )}
        </td>

        {/* 4c. Sales Qty (-) */}
        <td className="py-2.5 px-2.5 text-right align-top border-r border-slate-100 font-mono-num min-w-[100px]">
          {isAdmin && isInlineEdit ? (
            <div className="flex flex-col items-end gap-1" onClick={(e) => e.stopPropagation()}>
              <input
                type="number"
                min="0"
                aria-label={`Sales Qty ${br.brandLabel} ${grp.titleLabel}`}
                value={br.effectiveSalesQty}
                onChange={(e) =>
                  upsertEntryField(periodMode, grp.periodKey, br.brand, {
                    salesQtyOverride: Math.max(0, Number(e.target.value) || 0),
                  })
                }
                className="w-20 px-1.5 py-1 text-right text-xs font-bold text-sky-900 bg-sky-50/50 border border-sky-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-500"
              />
              {br.isSalesOverridden && (
                <button
                  type="button"
                  onClick={() =>
                    upsertEntryField(periodMode, grp.periodKey, br.brand, {
                      salesQtyOverride: null,
                    })
                  }
                  className="text-[9.5px] font-semibold text-sky-700 hover:underline inline-flex items-center gap-0.5 cursor-pointer"
                  title={`Sinkron ulang ke Sales Qty Harian (${br.liveSalesQty} pcs)`}
                >
                  <RefreshCw className="w-2.5 h-2.5" />
                  <span>Live ({br.liveSalesQty})</span>
                </button>
              )}
            </div>
          ) : (
            <div className="flex flex-col items-end">
              <span className="text-xs font-bold text-slate-900">
                {br.effectiveSalesQty > 0 ? `-${formatNumberId(br.effectiveSalesQty)}` : '0'}
              </span>
              <span className="text-[9.5px] text-slate-400 font-sans">
                {br.isSalesOverridden ? 'Manual' : 'Live Sales'}
              </span>
            </div>
          )}
        </td>

        {/* 4d. Barang Keluar (-) */}
        <td className="py-2.5 px-2.5 text-right align-top border-r border-slate-100 font-mono-num min-w-[92px]">
          {isAdmin && isInlineEdit ? (
            <input
              type="number"
              min="0"
              aria-label={`Barang Keluar ${br.brandLabel} ${grp.titleLabel}`}
              value={br.barangKeluar}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  barangKeluar: Math.max(0, Number(e.target.value) || 0),
                })
              }
              className="w-20 px-1.5 py-1 text-right text-xs font-bold text-amber-900 bg-amber-50/40 border border-amber-300 rounded-md focus:outline-none focus:ring-2 focus:ring-amber-500"
            />
          ) : (
            <span
              className={`text-xs font-semibold ${
                br.barangKeluar > 0 ? 'text-amber-700' : 'text-slate-400'
              }`}
            >
              {br.barangKeluar > 0 ? `-${formatNumberId(br.barangKeluar)}` : '0'}
            </span>
          )}
        </td>

        {/* 4e. Stock Akhir (Auto Calculated) */}
        <td className="py-2.5 px-3 text-right align-top border-r border-slate-200 bg-slate-50/60 font-mono-num min-w-[118px]">
          <div className="text-xs font-bold text-slate-950">
            {formatNumberId(br.stockAkhir)} pcs
          </div>
          <div className="mt-1 flex items-center justify-end gap-1 font-sans">
            <span
              className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded border text-[9.5px] font-semibold ${health.toneClass}`}
            >
              {health.kind === 'danger' && <AlertTriangle className="w-2.5 h-2.5" />}
              {health.kind === 'ok' && <CheckCircle2 className="w-2.5 h-2.5" />}
              <span>{health.label}</span>
            </span>
          </div>
          {(br.stockAwal + br.barangDatang) > 0 && (
            <div className="text-[9.5px] text-slate-500 mt-0.5">
              ST {br.sellThroughPct.toFixed(1)}%
            </div>
          )}
        </td>

        {/* 5. Promo */}
        <td className="py-2.5 px-3 align-top border-r border-slate-200 min-w-[180px] max-w-[230px]">
          {isAdmin && isInlineEdit ? (
            <textarea
              rows={2}
              placeholder="Cth: Disc 50%+20%, POP Bazar..."
              aria-label={`Promo ${br.brandLabel} ${grp.titleLabel}`}
              value={br.entry.promo}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  promo: e.target.value,
                })
              }
              className="w-full px-2 py-1 text-[11px] text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 resize-y"
            />
          ) : br.entry.promo ? (
            <div className="text-xs text-slate-800 leading-relaxed whitespace-pre-wrap">
              {br.entry.promo}
            </div>
          ) : (
            <span className="text-[11px] text-slate-400 italic">Belum ada catatan promo</span>
          )}
        </td>

        {/* 6. Reason / Kendala */}
        <td className="py-2.5 px-3 align-top border-r border-slate-200 min-w-[200px] max-w-[260px]">
          {isAdmin && isInlineEdit ? (
            <textarea
              rows={2}
              placeholder="Cth: Size L/XL cepat habis, traffic sepi..."
              aria-label={`Reason Kendala ${br.brandLabel} ${grp.titleLabel}`}
              value={br.entry.reasonKendala}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  reasonKendala: e.target.value,
                })
              }
              className="w-full px-2 py-1 text-[11px] text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 resize-y"
            />
          ) : br.entry.reasonKendala ? (
            <div className="text-xs text-slate-800 leading-relaxed whitespace-pre-wrap">
              {br.entry.reasonKendala}
            </div>
          ) : (
            <span className="text-[11px] text-slate-400 italic">Belum ada kendala tercatat</span>
          )}
        </td>

        {/* 7. Action Plan */}
        <td className="py-2.5 px-3 align-top border-r border-slate-200 min-w-[200px] max-w-[260px]">
          {isAdmin && isInlineEdit ? (
            <textarea
              rows={2}
              placeholder="Cth: Request DO size M-XL, refresh display..."
              aria-label={`Action Plan ${br.brandLabel} ${grp.titleLabel}`}
              value={br.entry.actionPlan}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) =>
                upsertEntryField(periodMode, grp.periodKey, br.brand, {
                  actionPlan: e.target.value,
                })
              }
              className="w-full px-2 py-1 text-[11px] text-slate-900 bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 resize-y"
            />
          ) : br.entry.actionPlan ? (
            <div className="text-xs text-slate-800 leading-relaxed whitespace-pre-wrap">
              {br.entry.actionPlan}
            </div>
          ) : (
            <span className="text-[11px] text-slate-400 italic">Belum ada action plan</span>
          )}
        </td>

        {/* 8. Aksi */}
        <td className="py-2.5 px-2.5 text-center align-top min-w-[80px]">
          {isAdmin ? (
            <button
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                openEditModal(periodMode, grp.periodKey, br.brand);
              }}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-[11px] font-semibold cursor-pointer whitespace-nowrap shrink-0 transition-colors"
              title={`Edit Inventory & Evaluasi ${br.brandLabel}`}
            >
              <PenSquare className="w-3 h-3" />
              <span>Edit</span>
            </button>
          ) : (
            <span className="text-[10px] text-slate-400 font-medium">Read-only</span>
          )}
        </td>
      </tr>
    );
  };

  // Preview Stock Akhir inside Edit Modal
  const modalPreview = useMemo(() => {
    if (!modalTarget) return null;
    const liveQty =
      modalTarget.periodType === 'weekly'
        ? liveSalesMap.weekly.get(Number(modalTarget.periodKey))?.[modalTarget.brand] || 0
        : liveSalesMap.monthly.get(String(modalTarget.periodKey).padStart(2, '0'))?.[
            modalTarget.brand
          ] || 0;
    const sAwal = Math.max(0, Number(modalDraft.stockAwal) || 0);
    const bDatang = Math.max(0, Number(modalDraft.barangDatang) || 0);
    const bKeluar = Math.max(0, Number(modalDraft.barangKeluar) || 0);
    const sQty =
      modalDraft.salesQtyMode === 'manual'
        ? Math.max(0, Number(modalDraft.salesQtyManual) || 0)
        : liveQty;
    const sAkhir = sAwal + bDatang - sQty - bKeluar;
    const avail = sAwal + bDatang;
    const st = avail > 0 ? (sQty / avail) * 100 : 0;
    return { liveQty, sAwal, bDatang, bKeluar, sQty, sAkhir, st };
  }, [modalTarget, modalDraft, liveSalesMap]);

  return (
    <section className="bg-white border border-slate-200 rounded-2xl overflow-hidden">
      {/* Top Executive Dark Slate Bar */}
      <div className="px-4 sm:px-5 py-4 bg-slate-950 text-white border-b border-slate-800 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div>
          <div className="text-[11px] font-semibold text-amber-400 tracking-wide uppercase">
            {CV_DIVISI_FULL} · Dept 6201 Graphis &amp; Dept 6244 Simplicity
          </div>
          <h3 className="text-sm sm:text-base font-bold text-white tracking-tight mt-0.5 flex items-center gap-2">
            <Layers className="w-4 h-4 text-sky-400 shrink-0" />
            <span>
              DATA CV INVENTORY, PROMO, REASON/KENDALA &amp; ACTION PLAN (
              {periodMode === 'weekly' ? 'PER WEEK' : 'PER MONTHLY'})
            </span>
          </h3>
        </div>

        {/* Period Switcher + Quick Actions */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="inline-flex p-0.5 bg-slate-900 border border-slate-700 rounded-lg">
            <button
              type="button"
              onClick={() => setPeriodMode('weekly')}
              className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                periodMode === 'weekly'
                  ? 'bg-amber-400 text-slate-950'
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              Per Week (W1–W53)
            </button>
            <button
              type="button"
              onClick={() => setPeriodMode('monthly')}
              className={`px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap shrink-0 ${
                periodMode === 'monthly'
                  ? 'bg-sky-400 text-slate-950'
                  : 'text-slate-300 hover:text-white'
              }`}
            >
              Per Monthly (12 Bulan)
            </button>
          </div>

          <button
            type="button"
            onClick={handleCopyWhatsAppInventory}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer whitespace-nowrap shrink-0 transition-colors"
            title="Salin ringkasan Inventory, Promo, Kendala & Action Plan ke clipboard"
          >
            <ClipboardCopy className="w-3.5 h-3.5" />
            <span>Salin WA Ops</span>
          </button>

          {isAdmin && (
            <button
              type="button"
              onClick={() => {
                const defaultKey =
                  periodMode === 'weekly'
                    ? String(
                        selectedWeekNumber ||
                          periodGroups[0]?.weekNumber ||
                          ANNUAL_2026_WEEKS[0]?.weekNumber ||
                          45
                      )
                    : selectedMonthKey;
                openEditModal(
                  periodMode,
                  defaultKey,
                  brandFilter === 'SIMPLICITY' ? 'SIMPLICITY' : 'GRAPHIS'
                );
              }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white text-slate-950 hover:bg-slate-100 cursor-pointer whitespace-nowrap shrink-0 transition-colors"
            >
              <PenSquare className="w-3.5 h-3.5" />
              <span>Input / Edit Data</span>
            </button>
          )}
        </div>
      </div>

      {/* Unified Inventory Ledger Strip (Single Elevation, 1px Dividers) */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 divide-y sm:divide-y-0 sm:divide-x divide-slate-200 border-b border-slate-200 bg-slate-50/70">
        <div className="p-3.5">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-slate-500">
            Total Stock Awal
          </div>
          <div className="text-lg font-bold text-slate-900 font-mono-num mt-0.5">
            {formatNumberId(kpiSummary.stockAwal)} <span className="text-xs font-medium text-slate-500">pcs</span>
          </div>
        </div>

        <div className="p-3.5">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
            + Barang Datang
          </div>
          <div className="text-lg font-bold text-emerald-700 font-mono-num mt-0.5">
            +{formatNumberId(kpiSummary.barangDatang)} <span className="text-xs font-medium text-emerald-600/80">pcs</span>
          </div>
        </div>

        <div className="p-3.5">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-sky-700">
            - Sales Qty (CV)
          </div>
          <div className="text-lg font-bold text-sky-800 font-mono-num mt-0.5">
            -{formatNumberId(kpiSummary.salesQty)} <span className="text-xs font-medium text-sky-600/80">pcs</span>
          </div>
        </div>

        <div className="p-3.5">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-amber-700">
            - Barang Keluar
          </div>
          <div className="text-lg font-bold text-amber-700 font-mono-num mt-0.5">
            -{formatNumberId(kpiSummary.barangKeluar)} <span className="text-xs font-medium text-amber-600/80">pcs</span>
          </div>
        </div>

        <div className="p-3.5 bg-white">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-slate-700">
            = Net Stock Akhir
          </div>
          <div
            className={`text-lg font-bold font-mono-num mt-0.5 ${
              kpiSummary.stockAkhir < 0 ? 'text-rose-700' : 'text-slate-950'
            }`}
          >
            {formatNumberId(kpiSummary.stockAkhir)} <span className="text-xs font-medium text-slate-500">pcs</span>
          </div>
        </div>

        <div className="p-3.5 bg-white">
          <div className="text-[10px] font-semibold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <TrendingUp className="w-3 h-3 text-indigo-600" />
            <span>Sell-Through</span>
          </div>
          <div className="text-lg font-bold text-indigo-700 font-mono-num mt-0.5">
            {kpiSummary.sellThroughPct.toFixed(1)}%
          </div>
        </div>
      </div>

      {/* Filter & Inline Edit Toolbar */}
      <div className="px-4 py-3 border-b border-slate-200 bg-white flex flex-wrap items-center justify-between gap-2.5 text-xs">
        <div className="flex flex-wrap items-center gap-2">
          {/* Calendar Date Filter */}
          <div className="flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-300">
            <Calendar className="w-3.5 h-3.5 text-slate-700 shrink-0" />
            <label htmlFor={`cv-ops-cal-${defaultPeriodType}`} className="font-semibold text-slate-700 shrink-0">
              Tanggal:
            </label>
            <input
              id={`cv-ops-cal-${defaultPeriodType}`}
              type="date"
              value={calendarIso}
              onChange={(e) => {
                const val = e.target.value;
                setCalendarIso(val);
                if (val) {
                  const parts = val.split('-');
                  if (parts.length === 3 && MONTH_NAMES_ID[parts[1]] && onChangeMonthKey) {
                    onChangeMonthKey(parts[1]);
                  }
                }
              }}
              className="text-xs font-semibold bg-transparent text-slate-900 focus:outline-none cursor-pointer"
              title="Filter langsung ke Week atau Bulan dari kalender tanggal"
            />
            {calendarIso && (
              <button
                type="button"
                onClick={() => setCalendarIso('')}
                className="text-[10px] font-semibold text-slate-700 hover:text-slate-950 px-1.5 py-0.5 rounded bg-slate-200 cursor-pointer"
              >
                Reset
              </button>
            )}
          </div>

          {/* Scope Selector */}
          <div className="inline-flex p-0.5 bg-slate-100 rounded-lg border border-slate-200">
            <button
              type="button"
              onClick={() => setScopeMode('focused')}
              className={`px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                scopeMode === 'focused'
                  ? 'bg-white text-slate-950 border border-slate-200'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {periodMode === 'weekly'
                ? `Week di ${MONTH_NAMES_ID[selectedMonthKey] || selectedMonthKey}`
                : `Bulan ${MONTH_NAMES_ID[selectedMonthKey] || selectedMonthKey}`}
            </button>
            <button
              type="button"
              onClick={() => setScopeMode('all')}
              className={`px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                scopeMode === 'all'
                  ? 'bg-white text-slate-950 border border-slate-200'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {periodMode === 'weekly' ? 'Semua 53 Week' : '12 Bulan (Jan–Des)'}
            </button>
          </div>

          {/* Brand Filter */}
          <div className="inline-flex p-0.5 bg-slate-100 rounded-lg border border-slate-200">
            <button
              type="button"
              onClick={() => setBrandFilter('ALL')}
              className={`px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                brandFilter === 'ALL'
                  ? 'bg-slate-900 text-white'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Semua Brand + Total
            </button>
            <button
              type="button"
              onClick={() => setBrandFilter('GRAPHIS')}
              className={`px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                brandFilter === 'GRAPHIS'
                  ? 'bg-rose-600 text-white'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Graphis (6201)
            </button>
            <button
              type="button"
              onClick={() => setBrandFilter('SIMPLICITY')}
              className={`px-2.5 py-1 rounded-md text-[11px] font-semibold cursor-pointer transition-colors ${
                brandFilter === 'SIMPLICITY'
                  ? 'bg-indigo-600 text-white'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Simplicity (6244)
            </button>
          </div>
        </div>

        {/* Search + Inline Edit Switch */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="relative">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari promo, kendala, plan..."
              className="pl-8 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-slate-900 w-48 sm:w-56"
            />
          </div>

          {isAdmin && (
            <button
              type="button"
              onClick={() => setIsInlineEdit((prev) => !prev)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border cursor-pointer transition-colors whitespace-nowrap shrink-0 ${
                isInlineEdit
                  ? 'bg-amber-400 border-amber-500 text-slate-950'
                  : 'bg-white border-slate-300 text-slate-800 hover:bg-slate-50'
              }`}
            >
              <PackageCheck className="w-3.5 h-3.5" />
              <span>{isInlineEdit ? 'Mode Edit Tabel Aktif' : 'Edit Langsung di Tabel'}</span>
            </button>
          )}
        </div>
      </div>

      {/* Main High-Density Table */}
      {periodGroups.length === 0 ? (
        <div className="p-8 text-center space-y-2 bg-white">
          <div className="text-sm font-semibold text-slate-800">
            Tidak ada data CV yang sesuai dengan filter saat ini
          </div>
          <p className="text-xs text-slate-500">
            Ubah kata kunci pencarian, tanggal kalender, atau pilih tampilkan seluruh periode.
          </p>
          <button
            type="button"
            onClick={() => {
              setSearchQuery('');
              setCalendarIso('');
              setScopeMode('all');
            }}
            className="mt-2 px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-semibold cursor-pointer"
          >
            Tampilkan Semua Data
          </button>
        </div>
      ) : (
        <div className="overflow-x-auto max-h-[680px]">
          <table className="w-full text-left border-collapse text-xs">
            <thead className="sticky top-0 z-20 bg-slate-900 text-white border-b border-slate-700">
              {/* Group Headers */}
              <tr className="text-[10px] font-bold uppercase tracking-wider text-slate-300 border-b border-slate-800">
                <th
                  colSpan={3}
                  className="py-2 px-3 border-r border-slate-800 bg-slate-950 text-left sticky left-0 z-30"
                >
                  IDENTITAS PERIODE, DIVISI &amp; BRAND
                </th>
                <th
                  colSpan={5}
                  className="py-2 px-3 border-r border-slate-800 bg-slate-900 text-center text-sky-300"
                >
                  INVENTORY MOVEMENT (PCS) — RUMUS: AWAL + DATANG − SALES − KELUAR = AKHIR
                </th>
                <th
                  colSpan={4}
                  className="py-2 px-3 bg-slate-950 text-center text-amber-300"
                >
                  OPERASIONAL, EVALUASI &amp; TINDAK LANJUT
                </th>
              </tr>

              {/* Leaf Column Headers */}
              <tr className="text-[10.5px] font-bold uppercase tracking-wider text-slate-100">
                <th className="py-2.5 px-3.5 border-r border-slate-800 bg-slate-900 sticky left-0 z-30 min-w-[155px]">
                  {periodMode === 'weekly' ? 'Week & Periode' : 'Bulan & Periode'}
                </th>
                <th className="py-2.5 px-3 border-r border-slate-800 min-w-[130px]">Divisi</th>
                <th className="py-2.5 px-3 border-r border-slate-800 min-w-[135px]">Brand (Dept)</th>
                <th className="py-2.5 px-2.5 text-right border-r border-slate-800 min-w-[92px]">
                  Stock Awal
                </th>
                <th className="py-2.5 px-2.5 text-right border-r border-slate-800 text-emerald-300 min-w-[92px]">
                  Barang Datang
                </th>
                <th className="py-2.5 px-2.5 text-right border-r border-slate-800 text-sky-300 min-w-[100px]">
                  Sales Qty
                </th>
                <th className="py-2.5 px-2.5 text-right border-r border-slate-800 text-amber-300 min-w-[92px]">
                  Barang Keluar
                </th>
                <th className="py-2.5 px-3 text-right border-r border-slate-800 bg-slate-800/80 text-white min-w-[118px]">
                  Stock Akhir
                </th>
                <th className="py-2.5 px-3 border-r border-slate-800 min-w-[180px]">Promo</th>
                <th className="py-2.5 px-3 border-r border-slate-800 min-w-[200px]">
                  Reason / Kendala
                </th>
                <th className="py-2.5 px-3 border-r border-slate-800 min-w-[200px]">
                  Action Plan
                </th>
                <th className="py-2.5 px-2.5 text-center min-w-[80px]">Aksi</th>
              </tr>
            </thead>

            <tbody>
              {periodGroups.map((grp) => {
                if (brandFilter === 'GRAPHIS') {
                  return renderBrandRow(grp, grp.graphis, true, 1, true);
                }
                if (brandFilter === 'SIMPLICITY') {
                  return renderBrandRow(grp, grp.simplicity, true, 1, true);
                }

                // Default ALL: 2 Brand rows + 1 clean Total Divisi 1 rollup row (rowSpan=3 on Period cell)
                const totalHealth = getStockHealthInfo(
                  grp.total.stockAwal,
                  grp.total.barangDatang,
                  grp.total.effectiveSalesQty,
                  grp.total.barangKeluar,
                  grp.total.stockAkhir
                );

                return (
                  <React.Fragment key={grp.periodKey}>
                    {renderBrandRow(grp, grp.graphis, true, 3, true)}
                    {renderBrandRow(grp, grp.simplicity, false, 3, false)}
                    <tr className="bg-slate-100/90 border-b-2 border-slate-300 text-slate-950 font-semibold">
                      <td className="py-2 px-3 border-r border-slate-200 text-[11px] text-slate-700">
                        Total {grp.divisi}
                      </td>
                      <td className="py-2 px-3 border-r border-slate-200 text-[11px] font-bold text-slate-900">
                        TOTAL CV (6201 + 6244)
                      </td>
                      <td className="py-2 px-2.5 text-right border-r border-slate-200 font-mono-num text-xs">
                        {formatNumberId(grp.total.stockAwal)}
                      </td>
                      <td className="py-2 px-2.5 text-right border-r border-slate-200 font-mono-num text-xs text-emerald-700">
                        {grp.total.barangDatang > 0 ? `+${formatNumberId(grp.total.barangDatang)}` : '0'}
                      </td>
                      <td className="py-2 px-2.5 text-right border-r border-slate-200 font-mono-num text-xs text-sky-800">
                        {grp.total.effectiveSalesQty > 0
                          ? `-${formatNumberId(grp.total.effectiveSalesQty)}`
                          : '0'}
                      </td>
                      <td className="py-2 px-2.5 text-right border-r border-slate-200 font-mono-num text-xs text-amber-700">
                        {grp.total.barangKeluar > 0 ? `-${formatNumberId(grp.total.barangKeluar)}` : '0'}
                      </td>
                      <td className="py-2 px-3 text-right border-r border-slate-200 font-mono-num text-xs font-bold bg-slate-200/70">
                        <span>{formatNumberId(grp.total.stockAkhir)} pcs</span>
                        <span className="block text-[9.5px] font-normal text-slate-600 font-sans">
                          {totalHealth.label} · ST {grp.total.sellThroughPct.toFixed(1)}%
                        </span>
                      </td>
                      <td colSpan={4} className="py-2 px-3 text-[11px] text-slate-600">
                        Konsolidasi Divisi 1 (Graphis Dept 6201 + Simplicity Dept 6244)
                      </td>
                    </tr>
                  </React.Fragment>
                );
              })}
            </tbody>

            {/* Sticky Grand Total Footer */}
            <tfoot className="sticky bottom-0 z-20 bg-slate-950 text-white border-t-2 border-amber-400 font-semibold">
              <tr>
                <td colSpan={3} className="py-3 px-3.5 text-xs uppercase tracking-wider text-amber-300 border-r border-slate-800">
                  TOTAL KESELURUHAN ({periodGroups.length}{' '}
                  {periodMode === 'weekly' ? 'WEEK' : 'BULAN'})
                </td>
                <td className="py-3 px-2.5 text-right font-mono-num text-xs border-r border-slate-800">
                  {formatNumberId(kpiSummary.stockAwal)}
                </td>
                <td className="py-3 px-2.5 text-right font-mono-num text-xs text-emerald-300 border-r border-slate-800">
                  +{formatNumberId(kpiSummary.barangDatang)}
                </td>
                <td className="py-3 px-2.5 text-right font-mono-num text-xs text-sky-300 border-r border-slate-800">
                  -{formatNumberId(kpiSummary.salesQty)}
                </td>
                <td className="py-3 px-2.5 text-right font-mono-num text-xs text-amber-300 border-r border-slate-800">
                  -{formatNumberId(kpiSummary.barangKeluar)}
                </td>
                <td className="py-3 px-3 text-right font-mono-num text-xs font-bold text-amber-300 bg-slate-900 border-r border-slate-800">
                  {formatNumberId(kpiSummary.stockAkhir)} pcs
                </td>
                <td colSpan={4} className="py-3 px-3 text-xs text-slate-300">
                  Rata-rata Sell-Through:{' '}
                  <span className="font-mono-num font-bold text-white">
                    {kpiSummary.sellThroughPct.toFixed(1)}%
                  </span>
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}

      {/* ============================================================ */}
      {/* DETAIL EDIT MODAL (WITH CALENDAR DATE PICKER)                */}
      {/* ============================================================ */}
      {modalTarget &&
        modalPreview &&
        typeof document !== 'undefined' &&
        createPortal(
          <div
            className="fixed inset-0 z-[9999] overflow-y-auto bg-slate-950/80 backdrop-blur-xs p-3 sm:p-4"
            onClick={(e) => {
              if (e.target === e.currentTarget) setModalTarget(null);
            }}
          >
            <div className="flex min-h-full items-center justify-center py-4">
              <div
                className="w-full max-w-xl bg-white border border-slate-200 rounded-2xl shadow-2xl overflow-hidden"
                onClick={(e) => e.stopPropagation()}
              >
                {/* Modal Header */}
                <div className="px-5 py-3.5 bg-slate-950 text-white flex items-center justify-between border-b border-slate-800">
                  <div>
                    <div className="text-[10.5px] font-semibold uppercase tracking-wider text-amber-400">
                      {modalTarget.periodType === 'weekly'
                        ? `Periode Weekly · W-${modalTarget.periodKey}`
                        : `Periode Bulanan · ${MONTH_NAMES_ID[modalTarget.periodKey] || modalTarget.periodKey} 2026`}
                    </div>
                    <h4 className="text-sm font-bold text-white mt-0.5 flex items-center gap-2">
                      <Tag className="w-4 h-4 text-sky-400" />
                      <span>
                        Edit Inventory &amp; Evaluasi — {CV_BRAND_CATALOG[modalTarget.brand].label} (Dept{' '}
                        {CV_BRAND_CATALOG[modalTarget.brand].deptCode})
                      </span>
                    </h4>
                  </div>
                  <button
                    type="button"
                    onClick={() => setModalTarget(null)}
                    className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 cursor-pointer"
                    title="Tutup"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>

                <form onSubmit={handleSaveModal} className="p-4 sm:p-5 space-y-4 text-xs">
                  {/* Period Switcher, Brand Switcher & Calendar Date Filter */}
                  <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5 p-3 rounded-xl bg-slate-50 border border-slate-200">
                    <div className="sm:col-span-6">
                      <label className="block text-[10.5px] font-bold uppercase text-slate-600 mb-1">
                        Pindah Periode via Kalender Tanggal
                      </label>
                      <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg px-2.5 h-9">
                        <Calendar className="w-3.5 h-3.5 text-slate-700 shrink-0" />
                        <input
                          type="date"
                          aria-label="Filter Kalender Tanggal Edit Inventory CV"
                          value={modalCalendarIso}
                          onChange={(e) => handleModalCalendarJump(e.target.value)}
                          className="w-full text-xs font-semibold text-slate-900 bg-transparent focus:outline-none cursor-pointer"
                        />
                      </div>
                    </div>

                    <div className="sm:col-span-6">
                      <label className="block text-[10.5px] font-bold uppercase text-slate-600 mb-1">
                        Pilih Brand
                      </label>
                      <div className="grid grid-cols-2 gap-1 p-0.5 bg-slate-200 rounded-lg h-9">
                        <button
                          type="button"
                          onClick={() =>
                            openEditModal(modalTarget.periodType, modalTarget.periodKey, 'GRAPHIS')
                          }
                          className={`rounded-md text-xs font-bold cursor-pointer transition-colors ${
                            modalTarget.brand === 'GRAPHIS'
                              ? 'bg-rose-600 text-white'
                              : 'text-slate-700 hover:bg-white/70'
                          }`}
                        >
                          Graphis (6201)
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            openEditModal(
                              modalTarget.periodType,
                              modalTarget.periodKey,
                              'SIMPLICITY'
                            )
                          }
                          className={`rounded-md text-xs font-bold cursor-pointer transition-colors ${
                            modalTarget.brand === 'SIMPLICITY'
                              ? 'bg-indigo-600 text-white'
                              : 'text-slate-700 hover:bg-white/70'
                          }`}
                        >
                          Simplicity (6244)
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Divisi Input */}
                  <div>
                    <label className="block text-[10.5px] font-bold uppercase text-slate-700 mb-1">
                      Divisi
                    </label>
                    <input
                      type="text"
                      value={modalDraft.divisi}
                      onChange={(e) => setModalDraft({ ...modalDraft, divisi: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-slate-300 font-semibold text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900"
                    />
                  </div>

                  {/* Inventory Inputs Grid */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-bold uppercase text-slate-800">
                        Mutasi Inventory (pcs)
                      </span>
                      <span className="text-[10.5px] text-slate-500">
                        Stock Akhir dihitung otomatis
                      </span>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                      <div>
                        <label className="block text-[10px] font-bold uppercase text-slate-600 mb-1">
                          Stock Awal
                        </label>
                        <input
                          type="number"
                          min="0"
                          value={modalDraft.stockAwal}
                          onFocus={(e) => e.target.select()}
                          onChange={(e) =>
                            setModalDraft({ ...modalDraft, stockAwal: e.target.value })
                          }
                          className="w-full h-10 px-2.5 rounded-lg border border-slate-300 text-right font-mono-num font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900"
                        />
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold uppercase text-emerald-700 mb-1">
                          Barang Datang (+)
                        </label>
                        <input
                          type="number"
                          min="0"
                          value={modalDraft.barangDatang}
                          onFocus={(e) => e.target.select()}
                          onChange={(e) =>
                            setModalDraft({ ...modalDraft, barangDatang: e.target.value })
                          }
                          className="w-full h-10 px-2.5 rounded-lg border border-emerald-300 bg-emerald-50/30 text-right font-mono-num font-bold text-emerald-900 focus:outline-none focus:ring-2 focus:ring-emerald-600"
                        />
                      </div>

                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="text-[10px] font-bold uppercase text-sky-700">
                            Sales Qty (-)
                          </label>
                          <button
                            type="button"
                            onClick={() =>
                              setModalDraft((prev) => ({
                                ...prev,
                                salesQtyMode: prev.salesQtyMode === 'live' ? 'manual' : 'live',
                                salesQtyManual: String(modalPreview.liveQty),
                              }))
                            }
                            className="text-[9.5px] font-semibold text-sky-700 hover:underline cursor-pointer"
                          >
                            {modalDraft.salesQtyMode === 'live' ? 'Ubah Manual' : 'Pakai Live'}
                          </button>
                        </div>
                        <input
                          type="number"
                          min="0"
                          readOnly={modalDraft.salesQtyMode === 'live'}
                          value={
                            modalDraft.salesQtyMode === 'live'
                              ? String(modalPreview.liveQty)
                              : modalDraft.salesQtyManual
                          }
                          onFocus={(e) => e.target.select()}
                          onChange={(e) =>
                            setModalDraft({ ...modalDraft, salesQtyManual: e.target.value })
                          }
                          className={`w-full h-10 px-2.5 rounded-lg border text-right font-mono-num font-bold focus:outline-none ${
                            modalDraft.salesQtyMode === 'live'
                              ? 'border-sky-200 bg-sky-50 text-sky-900'
                              : 'border-sky-400 bg-white text-slate-950 focus:ring-2 focus:ring-sky-600'
                          }`}
                        />
                      </div>

                      <div>
                        <label className="block text-[10px] font-bold uppercase text-amber-700 mb-1">
                          Barang Keluar (-)
                        </label>
                        <input
                          type="number"
                          min="0"
                          value={modalDraft.barangKeluar}
                          onFocus={(e) => e.target.select()}
                          onChange={(e) =>
                            setModalDraft({ ...modalDraft, barangKeluar: e.target.value })
                          }
                          className="w-full h-10 px-2.5 rounded-lg border border-amber-300 bg-amber-50/30 text-right font-mono-num font-bold text-amber-900 focus:outline-none focus:ring-2 focus:ring-amber-600"
                        />
                      </div>
                    </div>

                    {/* Live Stock Calculation Bar */}
                    <div className="p-3 rounded-xl bg-slate-950 text-white flex flex-wrap items-center justify-between gap-2">
                      <div className="text-[11px] font-mono-num text-slate-300">
                        {modalPreview.sAwal} + {modalPreview.bDatang} − {modalPreview.sQty} −{' '}
                        {modalPreview.bKeluar}
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="text-[11px] text-slate-400">
                          ST {modalPreview.st.toFixed(1)}%
                        </span>
                        <span
                          className={`text-sm font-bold font-mono-num ${
                            modalPreview.sAkhir < 0 ? 'text-rose-400' : 'text-amber-300'
                          }`}
                        >
                          Stock Akhir: {formatNumberId(modalPreview.sAkhir)} pcs
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Operational Evaluation Fields */}
                  <div className="space-y-3">
                    <div>
                      <div className="flex items-center justify-between mb-1">
                        <label className="text-[10.5px] font-bold uppercase text-slate-700">
                          Promo Berjalan
                        </label>
                        <button
                          type="button"
                          onClick={() =>
                            setModalDraft((prev) => ({
                              ...prev,
                              promo: prev.promo
                                ? `${prev.promo} · Special Price & Disc 50%+20%`
                                : 'Special Price · Disc up to 50%+20% Member',
                            }))
                          }
                          className="text-[10px] font-semibold text-indigo-700 hover:underline inline-flex items-center gap-1 cursor-pointer"
                        >
                          <Sparkles className="w-3 h-3" />
                          <span>Sisipkan Template Promo</span>
                        </button>
                      </div>
                      <textarea
                        rows={2}
                        placeholder="Tuliskan mekanisme promo aktif, POP display, atau program diskon..."
                        value={modalDraft.promo}
                        onChange={(e) => setModalDraft({ ...modalDraft, promo: e.target.value })}
                        className="w-full p-2.5 rounded-lg border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900"
                      />
                    </div>

                    <div>
                      <label className="block text-[10.5px] font-bold uppercase text-slate-700 mb-1">
                        Reason / Kendala
                      </label>
                      <textarea
                        rows={2}
                        placeholder="Jelaskan faktor pendorong sales atau kendala (broken size, stock kosong, traffic mall, cuaca)..."
                        value={modalDraft.reasonKendala}
                        onChange={(e) =>
                          setModalDraft({ ...modalDraft, reasonKendala: e.target.value })
                        }
                        className="w-full p-2.5 rounded-lg border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900"
                      />
                    </div>

                    <div>
                      <label className="block text-[10.5px] font-bold uppercase text-slate-700 mb-1">
                        Action Plan
                      </label>
                      <textarea
                        rows={2}
                        placeholder="Langkah konkrit berikutnya (request DO replenishment, rotasi manekin, push add-on kasir)..."
                        value={modalDraft.actionPlan}
                        onChange={(e) =>
                          setModalDraft({ ...modalDraft, actionPlan: e.target.value })
                        }
                        className="w-full p-2.5 rounded-lg border border-slate-300 text-slate-900 focus:outline-none focus:ring-2 focus:ring-slate-900"
                      />
                    </div>
                  </div>

                  {/* Modal Footer Actions */}
                  <div className="pt-2 border-t border-slate-200 flex items-center justify-end gap-2">
                    <button
                      type="button"
                      onClick={() => setModalTarget(null)}
                      className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-700 hover:bg-slate-100 cursor-pointer"
                    >
                      Batal
                    </button>
                    <button
                      type="submit"
                      className="inline-flex items-center gap-1.5 px-5 py-2 rounded-lg text-xs font-bold bg-slate-950 hover:bg-slate-800 text-white cursor-pointer"
                    >
                      <Save className="w-3.5 h-3.5 text-amber-400" />
                      <span>Simpan Data Inventory &amp; Ops</span>
                    </button>
                  </div>
                </form>
              </div>
            </div>
          </div>,
          document.body
        )}
    </section>
  );
};
