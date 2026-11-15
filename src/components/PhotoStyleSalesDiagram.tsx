import React from 'react';
import { formatNumberId } from '../data/initialData';

export interface SalesDiagramBarItem {
  key: string;
  label: string; // e.g. "Sen (H1)" or "W40" or "Okt"
  subLabel?: string; // e.g. "28/09" or "28/09–04/10"
  pagiSales: number;
  siangSales: number;
  totalSales: number;
  isWeekend?: boolean;
  isSelected?: boolean;
  onClick?: () => void;
}

interface PhotoStyleSalesDiagramProps {
  headerTitle: string; // e.g. "DIAGRAM SALES WEEKLY — TOTAL WEEKLY 40 (28 SEPTEMBER - 04 OKTOBER 2026)"
  headerSubtitle?: string; // e.g. "Diagram Batang Vertikal 7 Hari & Garis Tren Penjualan Shift Pagi vs Shift Siang"
  badgeText?: string; // default: "DIAGRAM SALES (OMZET RUPIAH)"
  cardTitle?: string; // default: "Grafik Batang Shift Pagi vs Siang & Garis Tren"
  primaryLabel?: string; // default: "Shift Pagi"
  secondaryLabel?: string; // default: "Shift Siang"
  totalSales: number;
  totalPagi: number;
  totalSiang: number;
  bars: SalesDiagramBarItem[];
  headerRightElement?: React.ReactNode;
  compactX?: boolean;
}

export const PhotoStyleSalesDiagram: React.FC<PhotoStyleSalesDiagramProps> = ({
  headerTitle,
  headerSubtitle = 'Diagram Batang Vertikal & Garis Tren Penjualan Shift Pagi vs Shift Siang',
  badgeText = 'DIAGRAM SALES (OMZET RUPIAH)',
  cardTitle = 'Grafik Batang Shift Pagi vs Siang & Garis Tren',
  primaryLabel = 'Shift Pagi',
  secondaryLabel = 'Shift Siang',
  totalSales,
  totalPagi,
  totalSiang,
  bars,
  headerRightElement,
  compactX = false,
}) => {
  const pagiPct = totalSales > 0 ? Math.round((totalPagi / totalSales) * 100) : 0;
  const siangPct = totalSales > 0 ? 100 - pagiPct : 0;

  const rawMax = Math.max(0, ...bars.map((b) => b.totalSales));
  const maxVal = rawMax > 0 ? rawMax : 2000000;

  const isManyBars = compactX || bars.length > 12;
  const minWidthPx = isManyBars ? `${Math.max(620, bars.length * 34)}px` : '100%';
  const svgW = Math.max(1, bars.length) * 100;

  const formatShortRp = (val: number) => {
    if (val >= 1000000) return `${(val / 1000000).toFixed(1)}jt`;
    if (val > 0) return `${Math.round(val / 1000)}rb`;
    return '0';
  };

  return (
    <div className="rounded-lg overflow-hidden border border-slate-300 bg-slate-100 shadow-xs">
      {/* Dark Navy Header Banner matching reference photo */}
      <div className="bg-slate-900 px-4 py-3 sm:px-5 space-y-1">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="text-xs sm:text-sm font-extrabold text-amber-400 tracking-tight uppercase">
            {headerTitle}
          </div>
          {headerRightElement && <div className="no-print">{headerRightElement}</div>}
        </div>
        <div className="text-[11px] font-bold text-slate-300">{headerSubtitle}</div>
        <div className="flex flex-wrap items-center gap-4 pt-1 text-[11px] font-bold">
          <span className="inline-flex items-center gap-1.5 text-amber-200">
            <span className="w-2.5 h-2.5 bg-amber-500 inline-block shrink-0" />
            {primaryLabel}: Rp {formatNumberId(totalPagi) || '0'} ({pagiPct}%)
          </span>
          <span className="inline-flex items-center gap-1.5 text-indigo-200">
            <span className="w-2.5 h-2.5 bg-indigo-600 inline-block shrink-0" />
            {secondaryLabel}: Rp {formatNumberId(totalSiang) || '0'} ({siangPct}%)
          </span>
          <span className="inline-flex items-center gap-1.5 text-emerald-300">
            <span className="w-3.5 h-0.5 bg-emerald-500 inline-block shrink-0" />
            Garis Tren Total Sales
          </span>
        </div>
      </div>

      {/* Inner White Card matching reference photo */}
      <div className="p-3 sm:p-4">
        <div className="w-full bg-white border border-slate-300 rounded-lg p-3.5 sm:p-5 shadow-2xs space-y-3">
          <div className="flex flex-wrap items-start justify-between gap-2">
            <div className="space-y-1.5">
              <span className="inline-block px-2.5 py-0.5 rounded bg-amber-600 text-white text-[10px] font-extrabold tracking-wide uppercase">
                {badgeText}
              </span>
              <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                {cardTitle}
              </div>
            </div>
            <span className="px-3 py-1 rounded bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-extrabold font-mono-num">
              Total: Rp {formatNumberId(totalSales) || '0'}
            </span>
          </div>

          {/* Plot Area */}
          <div className="overflow-x-auto pb-1">
            <div
              style={{ minWidth: minWidthPx }}
              className="pt-5 pb-9 pl-12 pr-2 select-none"
            >
              <div className="relative h-40 w-full">
                {/* Y-Axis Gridlines (100%, 50%, 0%) */}
                {[1, 0.5, 0].map((ratio) => {
                  const val = Math.round(maxVal * ratio);
                  return (
                    <div
                      key={ratio}
                      style={{ bottom: `${ratio * 100}%` }}
                      className="absolute -left-12 right-0 flex items-center pointer-events-none translate-y-1/2"
                    >
                      <span className="w-11 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-500">
                        {formatShortRp(val)}
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

                {/* SVG Green Trend Line Overlay */}
                {bars.length > 1 && (
                  <svg
                    viewBox={`0 0 ${svgW} 200`}
                    preserveAspectRatio="none"
                    className="absolute inset-0 w-full h-full pointer-events-none z-10 overflow-visible"
                  >
                    {(() => {
                      const pts = bars.map((b, i) => {
                        const x = (i + 0.5) * 100;
                        const y =
                          200 -
                          Math.max(0, Math.min(200, (b.totalSales / maxVal) * 200));
                        return { x, y, val: b.totalSales };
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
                )}

                {/* Vertical Stacked Columns */}
                <div
                  style={{
                    gridTemplateColumns: `repeat(${Math.max(1, bars.length)}, minmax(0, 1fr))`,
                  }}
                  className="relative z-0 h-full grid gap-1.5 sm:gap-2 items-end"
                >
                  {bars.map((b) => {
                    const hPct =
                      maxVal > 0
                        ? Math.max(
                            b.totalSales > 0 ? 6 : 2,
                            Math.round((b.totalSales / maxVal) * 100)
                          )
                        : 2;
                    const pagiShare =
                      b.totalSales > 0 ? Math.round((b.pagiSales / b.totalSales) * 100) : 0;
                    const siangShare = b.totalSales > 0 ? 100 - pagiShare : 0;
                    const compactNominal = formatShortRp(b.totalSales);

                    return (
                      <div
                        key={b.key}
                        onClick={b.onClick}
                        className={`h-full flex flex-col justify-end items-center group relative ${
                          b.onClick ? 'cursor-pointer' : ''
                        }`}
                      >
                        <div
                          style={{ height: `${hPct}%` }}
                          className={`w-full ${
                            isManyBars ? 'max-w-[26px]' : 'max-w-[44px]'
                          } rounded-t-xs overflow-hidden flex flex-col transition-all relative ${
                            b.isSelected ? 'ring-2 ring-amber-400 ring-offset-1' : ''
                          }`}
                          title={`${b.label}${b.subLabel ? ` (${b.subLabel})` : ''} — Total: Rp ${formatNumberId(b.totalSales)} (${primaryLabel}: Rp ${formatNumberId(b.pagiSales)} | ${secondaryLabel}: Rp ${formatNumberId(b.siangSales)})`}
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

                        {/* Top Value Label positioned cleanly above bar top */}
                        <div
                          style={{ bottom: `calc(${hPct}% + 5px)` }}
                          className="absolute inset-x-0 text-center pointer-events-none z-20"
                        >
                          <span
                            className={`${
                              isManyBars ? 'text-[8.5px]' : 'text-[10px]'
                            } font-mono-num font-extrabold text-slate-900 whitespace-nowrap`}
                          >
                            {compactNominal}
                          </span>
                        </div>

                        {/* X-Axis Label */}
                        <div className="absolute -bottom-8 inset-x-0 text-center">
                          <div
                            className={`${
                              isManyBars ? 'text-[9px]' : 'text-[10px]'
                            } font-extrabold truncate ${
                              b.isWeekend
                                ? 'text-rose-700'
                                : b.isSelected
                                  ? 'text-indigo-700 underline'
                                  : 'text-slate-700'
                            }`}
                          >
                            {b.label}
                          </div>
                          {b.subLabel && (
                            <div className="text-[9px] font-mono-num font-semibold text-slate-500 truncate">
                              {b.subLabel}
                            </div>
                          )}
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
    </div>
  );
};
