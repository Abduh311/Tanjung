import React, { useState } from 'react';
import { formatNumberId } from '../data/initialData';
import { TrendingUp, PieChart, ArrowUp } from 'lucide-react';

export interface GrowthContributionBarItem {
  key: string;
  label: string; // e.g. "Sen (H1)", "W40", "Okt"
  subLabel?: string; // e.g. "28/09", "28/09–04/10"
  totalSales: number;
  pagiSales?: number;
  siangSales?: number;
  prevSales?: number | null; // Previous period sales for Growth calculation
  prevLabel?: string; // e.g. "vs H-1" or "vs W39"
  isWeekend?: boolean;
  isSelected?: boolean;
  onClick?: () => void;
}

interface GrowthAndContributionBarDiagramProps {
  headerTitle: string;
  headerSubtitle?: string;
  periodBadgeText?: string;
  items: GrowthContributionBarItem[];
  compactX?: boolean;
  initialFilter?: 'BOTH' | 'GROWTH' | 'KONTRIBUSI';
}

export const GrowthAndContributionBarDiagram: React.FC<
  GrowthAndContributionBarDiagramProps
> = ({
  headerTitle,
  headerSubtitle = 'Visualisasi Diagram Batang Pertumbuhan (Growth %) & Porsi Kontribusi Sales (%)',
  periodBadgeText,
  items,
  compactX = false,
  initialFilter = 'BOTH',
}) => {
  const [viewFilter, setViewFilter] = useState<'BOTH' | 'GROWTH' | 'KONTRIBUSI'>(initialFilter);

  const grandTotalSales = items.reduce((acc, it) => acc + (it.totalSales || 0), 0);

  // Enrich items with Contribution % and Growth %
  const enrichedItems = items.map((item, idx) => {
    const contributionPct =
      grandTotalSales > 0 ? (item.totalSales / grandTotalSales) * 100 : 0;

    // Determine previous sales: if explicitly provided use it, otherwise use previous item in array
    const effectivePrevSales =
      item.prevSales !== undefined
        ? item.prevSales
        : idx > 0
          ? items[idx - 1].totalSales
          : null;

    let growthPct: number = 0;
    let growthDiff: number = 0;
    let growthStatus: 'UP' | 'DOWN' | 'FLAT' | 'BASE' | 'NEW' = 'FLAT';

    if (effectivePrevSales === null) {
      growthPct = 0;
      growthDiff = 0;
      growthStatus = 'BASE';
    } else {
      growthDiff = item.totalSales - effectivePrevSales;
      if (effectivePrevSales <= 0 && item.totalSales <= 0) {
        growthPct = 0;
        growthStatus = 'FLAT';
      } else if (effectivePrevSales <= 0 && item.totalSales > 0) {
        growthPct = 100;
        growthStatus = 'NEW';
      } else {
        growthPct = ((item.totalSales - effectivePrevSales) / effectivePrevSales) * 100;
        if (Math.abs(growthPct) < 0.05) {
          growthPct = 0;
          growthStatus = 'FLAT';
        } else {
          growthStatus = growthPct > 0 ? 'UP' : 'DOWN';
        }
      }
    }

    return {
      ...item,
      contributionPct,
      effectivePrevSales,
      growthPct,
      growthDiff,
      growthStatus,
    };
  });

  // Max values for scaling bars
  const maxContribPct = Math.max(
    10,
    ...enrichedItems.map((it) => it.contributionPct)
  );
  const avgContribPct =
    enrichedItems.length > 0 && grandTotalSales > 0
      ? 100 / enrichedItems.length
      : 0;

  const maxAbsGrowthPct = Math.max(
    15,
    ...enrichedItems.map((it) => Math.min(200, Math.abs(it.growthPct)))
  );

  // Summary highlights
  const topContribItem = [...enrichedItems].sort(
    (a, b) => b.contributionPct - a.contributionPct
  )[0];
  const topGrowthItem = [...enrichedItems]
    .filter((it) => it.growthStatus === 'UP' || it.growthStatus === 'NEW')
    .sort((a, b) => b.growthPct - a.growthPct)[0];

  const isManyBars = compactX || items.length > 10;
  const minWidthPx = isManyBars ? `${Math.max(660, items.length * 52)}px` : '100%';

  const formatShortRp = (val: number) => {
    if (val >= 1000000) return `${(val / 1000000).toFixed(1)}jt`;
    if (val > 0) return `${Math.round(val / 1000)}rb`;
    return '0';
  };

  return (
    <div className="rounded-xl overflow-hidden border border-slate-300 bg-slate-100 shadow-xs">
      {/* Dark Header Banner */}
      <div className="bg-slate-900 px-4 py-3 sm:px-5 space-y-1.5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <span className="px-2 py-0.5 rounded bg-emerald-500 text-slate-950 text-[10px] font-extrabold uppercase tracking-wider">
              GROWTH &amp; KONTRIBUSI
            </span>
            <h3 className="text-xs sm:text-sm font-extrabold text-amber-400 tracking-tight uppercase">
              {headerTitle}
            </h3>
          </div>

          <div className="flex flex-wrap items-center gap-1.5 no-print">
            {periodBadgeText && (
              <span className="px-2.5 py-0.5 rounded bg-slate-800 border border-amber-400/50 text-amber-300 text-[10px] font-bold font-mono-num">
                {periodBadgeText}
              </span>
            )}
            <div className="inline-flex items-center bg-slate-800 p-0.5 rounded-lg border border-slate-700 text-[10px] font-bold">
              {[
                { id: 'BOTH', label: 'Growth & Kontribusi' },
                { id: 'GROWTH', label: 'Diagram Growth' },
                { id: 'KONTRIBUSI', label: 'Diagram Kontribusi' },
              ].map((tab) => (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() =>
                    setViewFilter(tab.id as 'BOTH' | 'GROWTH' | 'KONTRIBUSI')
                  }
                  className={`px-2 py-1 rounded-md transition-colors cursor-pointer ${
                    viewFilter === tab.id
                      ? 'bg-amber-400 text-slate-950 font-extrabold'
                      : 'text-slate-300 hover:text-white'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>
          </div>
        </div>

        <div className="text-xs font-medium text-slate-300">{headerSubtitle}</div>

        {/* Legend Strip */}
        <div className="flex flex-wrap items-center gap-x-4 gap-y-1 pt-0.5 text-xs font-bold">
          <span className="inline-flex items-center gap-1.5 text-emerald-300">
            <span className="w-2.5 h-2.5 rounded-xs bg-emerald-500 inline-block shrink-0" />
            Growth Naik (▲ +%)
          </span>
          <span className="inline-flex items-center gap-1.5 text-rose-300">
            <span className="w-2.5 h-2.5 rounded-xs bg-rose-500 inline-block shrink-0" />
            Growth Turun (▼ -%)
          </span>
          <span className="inline-flex items-center gap-1.5 text-sky-300">
            <span className="w-2.5 h-2.5 rounded-xs bg-sky-500 inline-block shrink-0" />
            Kontribusi terhadap Total (%)
          </span>
        </div>
      </div>

      {/* Charts Grid */}
      <div className="p-3 sm:p-4">
        <div
          className={`grid grid-cols-1 ${
            viewFilter === 'BOTH' ? 'xl:grid-cols-2' : ''
          } gap-4`}
        >
          {/* =================================================================
              CHART 1: GRAFIK DIAGRAM BATANG GROWTH (%)
             ================================================================= */}
          {(viewFilter === 'BOTH' || viewFilter === 'GROWTH') && (
            <div className="w-full bg-white border border-slate-300 rounded-xl p-3.5 sm:p-4 shadow-2xs space-y-3 flex flex-col justify-between">
              <div className="flex flex-wrap items-start justify-between gap-2">
                <div className="space-y-1">
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-extrabold tracking-wide uppercase">
                    <TrendingUp className="w-3 h-3" />
                    DIAGRAM BATANG GROWTH (%)
                  </span>
                  <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                    Grafik Batang Pertumbuhan Sales (Naik ▲ / Turun ▼)
                  </div>
                </div>
                <div className="text-right">
                  <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-slate-100 border border-slate-200 text-slate-700 text-[11px] font-bold font-mono-num">
                    Total: Rp {formatNumberId(grandTotalSales) || '0'}
                  </span>
                </div>
              </div>

              {/* Growth Bar Plot Area */}
              <div className="overflow-x-auto pb-1">
                <div
                  style={{ minWidth: minWidthPx }}
                  className="pt-9 pb-12 pl-14 pr-2 select-none"
                >
                  <div className="relative h-44 w-full">
                    {/* Y-Axis Gridlines (100%, 50%, 0%) */}
                    {[1, 0.5, 0].map((ratio) => {
                      const val = Math.round(maxAbsGrowthPct * ratio);
                      return (
                        <div
                          key={ratio}
                          style={{ bottom: `${ratio * 100}%` }}
                          className="absolute -left-14 right-0 flex items-center pointer-events-none translate-y-1/2"
                        >
                          <span className="w-13 pr-2 text-right text-xs font-mono-num font-bold text-slate-600">
                            {val === 0 ? '0%' : `±${val}%`}
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

                    {/* SVG Growth Trend Line Overlay */}
                    {enrichedItems.length > 1 && (
                      <svg
                        viewBox={`0 0 ${enrichedItems.length * 100} 200`}
                        preserveAspectRatio="none"
                        className="absolute inset-0 w-full h-full pointer-events-none z-10 overflow-visible"
                      >
                        {(() => {
                          const pts = enrichedItems.map((it, i) => {
                            const x = (i + 0.5) * 100;
                            const cappedAbs = Math.min(
                              maxAbsGrowthPct,
                              Math.abs(it.growthPct)
                            );
                            const y =
                              200 -
                              Math.max(
                                0,
                                Math.min(200, (cappedAbs / maxAbsGrowthPct) * 200)
                              );
                            return { x, y, status: it.growthStatus };
                          });
                          const poly = pts.map((p) => `${p.x},${p.y}`).join(' ');
                          return (
                            <>
                              <polyline
                                fill="none"
                                stroke="#0d9488"
                                strokeWidth="2.2"
                                strokeDasharray="4 3"
                                strokeLinecap="round"
                                strokeLinejoin="round"
                                points={poly}
                              />
                              {pts.map((p, idx) => (
                                <circle
                                  key={idx}
                                  cx={p.x}
                                  cy={p.y}
                                  r="3.8"
                                  fill="#ffffff"
                                  stroke={
                                    p.status === 'DOWN'
                                      ? '#e11d48'
                                      : p.status === 'UP' || p.status === 'NEW'
                                        ? '#059669'
                                        : '#64748b'
                                  }
                                  strokeWidth="2.2"
                                />
                              ))}
                            </>
                          );
                        })()}
                      </svg>
                    )}

                    {/* Vertical Growth Bars */}
                    <div
                      style={{
                        gridTemplateColumns: `repeat(${Math.max(
                          1,
                          enrichedItems.length
                        )}, minmax(0, 1fr))`,
                      }}
                      className="relative z-0 h-full grid gap-1.5 sm:gap-2 items-end"
                    >
                      {enrichedItems.map((it) => {
                        const cappedAbs = Math.min(
                          maxAbsGrowthPct,
                          Math.abs(it.growthPct)
                        );
                        const hPct =
                          maxAbsGrowthPct > 0
                            ? Math.max(
                                it.growthStatus === 'UP' ||
                                  it.growthStatus === 'DOWN' ||
                                  it.growthStatus === 'NEW'
                                  ? 10
                                  : 4,
                                Math.round((cappedAbs / maxAbsGrowthPct) * 100)
                              )
                            : 4;

                        const isUp =
                          it.growthStatus === 'UP' || it.growthStatus === 'NEW';
                        const isDown = it.growthStatus === 'DOWN';
                        const isBase = it.growthStatus === 'BASE';

                        const barColorClass = isUp
                          ? 'bg-gradient-to-t from-emerald-600 to-emerald-400 border-emerald-700/30'
                          : isDown
                            ? 'bg-gradient-to-t from-rose-600 to-rose-400 border-rose-700/30'
                            : isBase
                              ? 'bg-amber-400/80 border-amber-500/40'
                              : 'bg-slate-300 border-slate-400/40';

                        const growthText = isBase
                          ? 'Basis'
                          : isUp
                            ? `+${it.growthPct.toFixed(1)}%`
                            : isDown
                              ? `${it.growthPct.toFixed(1)}%`
                              : '0%';

                        const diffSignText =
                          it.growthDiff > 0
                            ? `+Rp ${formatNumberId(it.growthDiff)}`
                            : it.growthDiff < 0
                              ? `-Rp ${formatNumberId(Math.abs(it.growthDiff))}`
                              : 'Rp 0';

                        return (
                          <div
                            key={it.key}
                            onClick={it.onClick}
                            className={`h-full flex flex-col justify-end items-center group relative ${
                              it.onClick ? 'cursor-pointer' : ''
                            }`}
                          >
                            <div
                              style={{ height: `${hPct}%` }}
                              className={`w-full ${
                                isManyBars ? 'max-w-[26px]' : 'max-w-[42px]'
                              } rounded-t-md border transition-all relative ${barColorClass} ${
                                it.isSelected
                                  ? 'ring-2 ring-amber-400 ring-offset-1'
                                  : ''
                              }`}
                              title={`${it.label}${
                                it.subLabel ? ` (${it.subLabel})` : ''
                              } — Growth: ${growthText} (${diffSignText}) | Sales: Rp ${formatNumberId(
                                it.totalSales
                              )}`}
                            />

                            {/* Value Label Above Bar */}
                            <div
                              style={{ bottom: `calc(${hPct}% + 6px)` }}
                              className="absolute inset-x-0 text-center pointer-events-none z-20"
                            >
                              <span
                                className={`${
                                  isManyBars
                                    ? 'text-[11px] sm:text-xs'
                                    : 'text-xs sm:text-sm'
                                } font-mono-num font-extrabold whitespace-nowrap px-1 py-0.5 rounded ${
                                  isUp
                                    ? 'text-emerald-800 bg-emerald-50/95'
                                    : isDown
                                      ? 'text-rose-800 bg-rose-50/95'
                                      : isBase
                                        ? 'text-amber-900 bg-amber-50/95'
                                        : 'text-slate-700'
                                }`}
                              >
                                {isUp ? `▲${growthText}` : isDown ? `▼${growthText}` : growthText}
                              </span>
                            </div>

                            {/* X-Axis Label */}
                            <div className="absolute -bottom-10 inset-x-0 text-center">
                              <div
                                className={`${
                                  isManyBars ? 'text-[11px]' : 'text-xs'
                                } font-extrabold truncate ${
                                  it.isWeekend
                                    ? 'text-rose-700'
                                    : it.isSelected
                                      ? 'text-indigo-700 underline'
                                      : 'text-slate-800'
                                }`}
                              >
                                {it.label}
                              </div>
                              <div className="text-[10px] font-mono-num font-semibold text-slate-500 truncate">
                                {formatShortRp(it.totalSales)}
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
          )}

          {/* =================================================================
              CHART 2: GRAFIK DIAGRAM BATANG KONTRIBUSI (%)
             ================================================================= */}
          {(viewFilter === 'BOTH' || viewFilter === 'KONTRIBUSI') && (
            <div className="w-full bg-white border border-slate-300 rounded-xl p-3.5 sm:p-4 shadow-2xs space-y-3 flex flex-col justify-between">
              <div className="flex flex-wrap items-start justify-between gap-2">
                <div className="space-y-1">
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded bg-sky-600 text-white text-[10px] font-extrabold tracking-wide uppercase">
                    <PieChart className="w-3 h-3" />
                    DIAGRAM BATANG KONTRIBUSI (%)
                  </span>
                  <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                    Grafik Batang Porsi Kontribusi terhadap Total Sales (100%)
                  </div>
                </div>
                <span className="px-2.5 py-1 rounded bg-sky-50 border border-sky-200 text-sky-900 text-[11px] font-extrabold font-mono-num">
                  Total 100% (Rp {formatNumberId(grandTotalSales) || '0'})
                </span>
              </div>

              {/* Contribution Bar Plot Area */}
              <div className="overflow-x-auto pb-1">
                <div
                  style={{ minWidth: minWidthPx }}
                  className="pt-9 pb-12 pl-14 pr-2 select-none"
                >
                  <div className="relative h-44 w-full">
                    {/* Y-Axis Gridlines (100%, 50%, 0% of maxContribPct) */}
                    {[1, 0.5, 0].map((ratio) => {
                      const val = Math.round(maxContribPct * ratio);
                      return (
                        <div
                          key={ratio}
                          style={{ bottom: `${ratio * 100}%` }}
                          className="absolute -left-14 right-0 flex items-center pointer-events-none translate-y-1/2"
                        >
                          <span className="w-13 pr-2 text-right text-xs font-mono-num font-bold text-slate-600">
                            {val}%
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

                    {/* Average Contribution Reference Line */}
                    {avgContribPct > 0 && avgContribPct <= maxContribPct && (
                      <div
                        style={{
                          bottom: `${Math.min(
                            100,
                            (avgContribPct / maxContribPct) * 100
                          )}%`,
                        }}
                        className="absolute left-0 right-0 border-b border-dashed border-amber-500/80 pointer-events-none z-10 flex justify-end"
                      >
                        <span className="px-2 py-0.5 bg-amber-100 text-amber-900 text-xs font-mono-num font-extrabold rounded -translate-y-3">
                          Rata2: {avgContribPct.toFixed(1)}%
                        </span>
                      </div>
                    )}

                    {/* SVG Contribution Curve Overlay */}
                    {enrichedItems.length > 1 && (
                      <svg
                        viewBox={`0 0 ${enrichedItems.length * 100} 200`}
                        preserveAspectRatio="none"
                        className="absolute inset-0 w-full h-full pointer-events-none z-10 overflow-visible"
                      >
                        {(() => {
                          const pts = enrichedItems.map((it, i) => {
                            const x = (i + 0.5) * 100;
                            const y =
                              200 -
                              Math.max(
                                0,
                                Math.min(
                                  200,
                                  (it.contributionPct / maxContribPct) * 200
                                )
                              );
                            return { x, y };
                          });
                          const poly = pts.map((p) => `${p.x},${p.y}`).join(' ');
                          return (
                            <>
                              <polyline
                                fill="none"
                                stroke="#0284c7"
                                strokeWidth="2.4"
                                strokeLinecap="round"
                                strokeLinejoin="round"
                                points={poly}
                              />
                              {pts.map((p, idx) => (
                                <circle
                                  key={idx}
                                  cx={p.x}
                                  cy={p.y}
                                  r="3.8"
                                  fill="#ffffff"
                                  stroke="#0369a1"
                                  strokeWidth="2.2"
                                />
                              ))}
                            </>
                          );
                        })()}
                      </svg>
                    )}

                    {/* Vertical Contribution Bars */}
                    <div
                      style={{
                        gridTemplateColumns: `repeat(${Math.max(
                          1,
                          enrichedItems.length
                        )}, minmax(0, 1fr))`,
                      }}
                      className="relative z-0 h-full grid gap-1.5 sm:gap-2 items-end"
                    >
                      {enrichedItems.map((it) => {
                        const hPct =
                          maxContribPct > 0
                            ? Math.max(
                                it.totalSales > 0 ? 8 : 2,
                                Math.round(
                                  (it.contributionPct / maxContribPct) * 100
                                )
                              )
                            : 2;

                        const hasShiftSplit =
                          (it.pagiSales ?? 0) + (it.siangSales ?? 0) > 0;
                        const pShare =
                          hasShiftSplit && it.totalSales > 0
                            ? Math.round(
                                ((it.pagiSales ?? 0) / it.totalSales) * 100
                              )
                            : 0;
                        const sShare = hasShiftSplit ? 100 - pShare : 100;

                        return (
                          <div
                            key={it.key}
                            onClick={it.onClick}
                            className={`h-full flex flex-col justify-end items-center group relative ${
                              it.onClick ? 'cursor-pointer' : ''
                            }`}
                          >
                            <div
                              style={{ height: `${hPct}%` }}
                              className={`w-full ${
                                isManyBars ? 'max-w-[26px]' : 'max-w-[42px]'
                              } rounded-t-md overflow-hidden flex flex-col border border-sky-700/25 transition-all relative ${
                                it.isSelected
                                  ? 'ring-2 ring-amber-400 ring-offset-1'
                                  : ''
                              }`}
                              title={`${it.label}${
                                it.subLabel ? ` (${it.subLabel})` : ''
                              } — Kontribusi: ${it.contributionPct.toFixed(
                                1
                              )}% (Rp ${formatNumberId(it.totalSales)})`}
                            >
                              {hasShiftSplit ? (
                                <>
                                  <div
                                    style={{ height: `${sShare}%` }}
                                    className="w-full bg-sky-600 transition-all"
                                  />
                                  <div
                                    style={{ height: `${pShare}%` }}
                                    className="w-full bg-cyan-400 transition-all"
                                  />
                                </>
                              ) : (
                                <div className="w-full h-full bg-gradient-to-t from-sky-600 to-cyan-400" />
                              )}
                            </div>

                            {/* Percentage Label Above Bar */}
                            <div
                              style={{ bottom: `calc(${hPct}% + 6px)` }}
                              className="absolute inset-x-0 text-center pointer-events-none z-20"
                            >
                              <span
                                className={`${
                                  isManyBars
                                    ? 'text-[11px] sm:text-xs'
                                    : 'text-xs sm:text-sm'
                                } font-mono-num font-extrabold text-sky-950 bg-sky-50/95 px-1 py-0.5 rounded whitespace-nowrap`}
                              >
                                {it.contributionPct.toFixed(1)}%
                              </span>
                            </div>

                            {/* X-Axis Label */}
                            <div className="absolute -bottom-10 inset-x-0 text-center">
                              <div
                                className={`${
                                  isManyBars ? 'text-[11px]' : 'text-xs'
                                } font-extrabold truncate ${
                                  it.isWeekend
                                    ? 'text-rose-700'
                                    : it.isSelected
                                      ? 'text-indigo-700 underline'
                                      : 'text-slate-800'
                                }`}
                              >
                                {it.label}
                              </div>
                              <div className="text-[10px] font-mono-num font-semibold text-slate-500 truncate">
                                {formatShortRp(it.totalSales)}
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
          )}
        </div>
      </div>
    </div>
  );
};
