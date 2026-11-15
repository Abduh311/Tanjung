import React from 'react';
import { formatNumberId } from '../data/initialData';

export interface MultiMetricPoint {
  key: string;
  label: string;
  subLabel?: string;
  pagiSales: number;
  siangSales: number;
  totalSales: number;
  pagiQty: number;
  siangQty: number;
  totalQty: number;
  pagiTrx: number;
  siangTrx: number;
  totalTrx: number;
  pricePoint: number; // Sales / Qty
  bigSize: number; // Sales / Trx
  isWeekend?: boolean;
}

interface SalesMultiMetricDiagramsProps {
  points: MultiMetricPoint[];
  periodLabel: string;
  totalQty: number;
  pagiQty: number;
  siangQty: number;
  totalTrx: number;
  pagiTrx: number;
  siangTrx: number;
  avgPricePoint: number;
  avgBigSize: number;
}

export const SalesMultiMetricDiagrams: React.FC<SalesMultiMetricDiagramsProps> = ({
  points,
  periodLabel,
  totalQty,
  pagiQty,
  siangQty,
  totalTrx,
  pagiTrx,
  siangTrx,
  avgPricePoint,
  avgBigSize,
}) => {
  const rawMaxQty = Math.max(0, ...points.map((p) => p.totalQty));
  const maxQty = rawMaxQty > 0 ? rawMaxQty : 50;

  const rawMaxTrx = Math.max(0, ...points.map((p) => p.totalTrx));
  const maxTrx = rawMaxTrx > 0 ? rawMaxTrx : 40;

  const rawMaxPP = Math.max(0, ...points.map((p) => p.pricePoint));
  const maxPP = rawMaxPP > 0 ? rawMaxPP : 100000;

  const rawMaxBS = Math.max(0, ...points.map((p) => p.bigSize));
  const maxBS = rawMaxBS > 0 ? rawMaxBS : 150000;

  const isCompactX = points.length > 10;
  const minChartWidth = isCompactX ? `${Math.max(520, points.length * 28)}px` : '100%';

  return (
    <div className="pt-2 space-y-5">
      {/* Sub-header for the 4 Additional Metric Diagrams */}
      <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-200/80">
        <div>
          <span className="text-xs sm:text-sm font-extrabold text-slate-900">
            Rincian Grafik Diagram: Qty (pcs), Trx (trx), Price Point (APC) &amp; Big Size (Basket Size)
          </span>
          <p className="text-xs text-slate-500 mt-0.5">
            Visualisasi lengkap 4 indikator utama pendamping Sales ({periodLabel})
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2 text-[11px] font-bold">
          <span className="px-2.5 py-1 rounded-md bg-cyan-50 border border-cyan-200 text-cyan-900 font-mono-num">
            Qty: {formatNumberId(totalQty) || '0'} pcs
          </span>
          <span className="px-2.5 py-1 rounded-md bg-violet-50 border border-violet-200 text-violet-900 font-mono-num">
            Trx: {formatNumberId(totalTrx) || '0'} trx
          </span>
          <span className="px-2.5 py-1 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-900 font-mono-num">
            Price Point: Rp {formatNumberId(avgPricePoint) || '0'}
          </span>
          <span className="px-2.5 py-1 rounded-md bg-rose-50 border border-rose-200 text-rose-900 font-mono-num">
            Big Size: Rp {formatNumberId(avgBigSize) || '0'}
          </span>
        </div>
      </div>

      {/* 2x2 Spacious Grid for Qty, Trx, Price Point, and Big Size */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* =================================================================
            DIAGRAM 2: QTY / BARANG TERJUAL (PCS)
           ================================================================= */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-2xs flex flex-col justify-between">
          <div className="flex flex-wrap items-start justify-between gap-2 mb-5">
            <div>
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-cyan-600 text-white text-[10px] font-extrabold uppercase tracking-wide mb-1">
                2. Diagram Qty (Barang Terjual)
              </div>
              <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                Grafik Quantity Terjual — Shift Pagi vs Siang (pcs)
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5">
                Pagi: <strong className="text-amber-800">{formatNumberId(pagiQty) || '0'} pcs</strong> · Siang:{' '}
                <strong className="text-indigo-800">{formatNumberId(siangQty) || '0'} pcs</strong>
              </div>
            </div>
            <span className="px-3 py-1.5 rounded-lg bg-cyan-50 border border-cyan-200 text-cyan-900 text-xs font-extrabold font-mono-num">
              Total Qty: {formatNumberId(totalQty) || '0'} pcs
            </span>
          </div>

          <div className="overflow-x-auto pb-2">
            <div
              style={{ minWidth: minChartWidth }}
              className="relative h-60 pt-7 pb-11 pl-12 pr-4 select-none"
            >
              {[1, 0.75, 0.5, 0.25, 0].map((ratio) => {
                const val = Math.round(maxQty * ratio);
                return (
                  <div
                    key={ratio}
                    style={{ bottom: `${ratio * 100}%` }}
                    className="absolute left-0 right-0 flex items-center pointer-events-none"
                  >
                    <span className="w-10 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-400">
                      {val}
                    </span>
                    <div className="flex-1 border-b border-dashed border-slate-200" />
                  </div>
                );
              })}

              {points.length > 1 && (
                <svg
                  viewBox={`0 0 ${points.length * 100} 200`}
                  preserveAspectRatio="none"
                  className="absolute inset-0 left-12 right-4 top-7 bottom-11 w-[calc(100%-4rem)] h-[calc(100%-4.5rem)] pointer-events-none z-10 overflow-visible"
                >
                  {(() => {
                    const pts = points.map((p, i) => {
                      const x = (i + 0.5) * 100;
                      const y =
                        200 - Math.max(0, Math.min(200, (p.totalQty / maxQty) * 200));
                      return { x, y, val: p.totalQty };
                    });
                    const poly = pts.map((pt) => `${pt.x},${pt.y}`).join(' ');
                    return (
                      <>
                        <polyline
                          fill="none"
                          stroke="#0891b2"
                          strokeWidth="2.75"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          points={poly}
                        />
                        {pts.map((pt, idx) => (
                          <circle
                            key={idx}
                            cx={pt.x}
                            cy={pt.y}
                            r="4"
                            fill="#ffffff"
                            stroke="#0e7490"
                            strokeWidth="2.2"
                          />
                        ))}
                      </>
                    );
                  })()}
                </svg>
              )}

              <div
                style={{
                  gridTemplateColumns: `repeat(${Math.max(1, points.length)}, minmax(0, 1fr))`,
                }}
                className="relative z-0 h-full grid gap-2 items-end"
              >
                {points.map((p) => {
                  const hPct =
                    maxQty > 0
                      ? Math.max(p.totalQty > 0 ? 6 : 1, Math.round((p.totalQty / maxQty) * 100))
                      : 0;
                  const pShare =
                    p.totalQty > 0 ? Math.round((p.pagiQty / p.totalQty) * 100) : 0;
                  const sShare = p.totalQty > 0 ? 100 - pShare : 0;

                  return (
                    <div
                      key={p.key}
                      className="h-full flex flex-col justify-end items-center group relative"
                    >
                      <div className="mb-1 text-[10px] font-mono-num font-extrabold text-cyan-950 whitespace-nowrap">
                        {p.totalQty > 0 ? `${p.totalQty}` : '-'}
                      </div>
                      <div
                        style={{ height: `${hPct}%` }}
                        className="w-full max-w-[38px] rounded-t-lg overflow-hidden flex flex-col border border-slate-900/15 shadow-2xs"
                        title={`${p.label} — Total Qty: ${p.totalQty} pcs (Pagi: ${p.pagiQty} | Siang: ${p.siangQty})`}
                      >
                        <div style={{ height: `${sShare}%` }} className="w-full bg-indigo-600" />
                        <div style={{ height: `${pShare}%` }} className="w-full bg-amber-400" />
                      </div>
                      <div className="absolute -bottom-9 inset-x-0 text-center">
                        <div
                          className={`text-[10px] font-extrabold truncate ${
                            p.isWeekend ? 'text-rose-700' : 'text-slate-800'
                          }`}
                        >
                          {p.label}
                        </div>
                        {p.subLabel && !isCompactX && (
                          <div className="text-[9px] font-mono-num text-slate-500 mt-0.5">
                            {p.subLabel}
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

        {/* =================================================================
            DIAGRAM 3: TRX / TRANSAKSI (TRX)
           ================================================================= */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-2xs flex flex-col justify-between">
          <div className="flex flex-wrap items-start justify-between gap-2 mb-5">
            <div>
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-violet-600 text-white text-[10px] font-extrabold uppercase tracking-wide mb-1">
                3. Diagram Trx (Jumlah Transaksi)
              </div>
              <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                Grafik Transaksi — Shift Pagi vs Siang (trx)
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5">
                Pagi: <strong className="text-amber-800">{formatNumberId(pagiTrx) || '0'} trx</strong> · Siang:{' '}
                <strong className="text-indigo-800">{formatNumberId(siangTrx) || '0'} trx</strong>
              </div>
            </div>
            <span className="px-3 py-1.5 rounded-lg bg-violet-50 border border-violet-200 text-violet-900 text-xs font-extrabold font-mono-num">
              Total Trx: {formatNumberId(totalTrx) || '0'} trx
            </span>
          </div>

          <div className="overflow-x-auto pb-2">
            <div
              style={{ minWidth: minChartWidth }}
              className="relative h-60 pt-7 pb-11 pl-12 pr-4 select-none"
            >
              {[1, 0.75, 0.5, 0.25, 0].map((ratio) => {
                const val = Math.round(maxTrx * ratio);
                return (
                  <div
                    key={ratio}
                    style={{ bottom: `${ratio * 100}%` }}
                    className="absolute left-0 right-0 flex items-center pointer-events-none"
                  >
                    <span className="w-10 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-400">
                      {val}
                    </span>
                    <div className="flex-1 border-b border-dashed border-slate-200" />
                  </div>
                );
              })}

              {points.length > 1 && (
                <svg
                  viewBox={`0 0 ${points.length * 100} 200`}
                  preserveAspectRatio="none"
                  className="absolute inset-0 left-12 right-4 top-7 bottom-11 w-[calc(100%-4rem)] h-[calc(100%-4.5rem)] pointer-events-none z-10 overflow-visible"
                >
                  {(() => {
                    const pts = points.map((p, i) => {
                      const x = (i + 0.5) * 100;
                      const y =
                        200 - Math.max(0, Math.min(200, (p.totalTrx / maxTrx) * 200));
                      return { x, y, val: p.totalTrx };
                    });
                    const poly = pts.map((pt) => `${pt.x},${pt.y}`).join(' ');
                    return (
                      <>
                        <polyline
                          fill="none"
                          stroke="#7c3aed"
                          strokeWidth="2.75"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          points={poly}
                        />
                        {pts.map((pt, idx) => (
                          <circle
                            key={idx}
                            cx={pt.x}
                            cy={pt.y}
                            r="4"
                            fill="#ffffff"
                            stroke="#6d28d9"
                            strokeWidth="2.2"
                          />
                        ))}
                      </>
                    );
                  })()}
                </svg>
              )}

              <div
                style={{
                  gridTemplateColumns: `repeat(${Math.max(1, points.length)}, minmax(0, 1fr))`,
                }}
                className="relative z-0 h-full grid gap-2 items-end"
              >
                {points.map((p) => {
                  const hPct =
                    maxTrx > 0
                      ? Math.max(p.totalTrx > 0 ? 6 : 1, Math.round((p.totalTrx / maxTrx) * 100))
                      : 0;
                  const pShare =
                    p.totalTrx > 0 ? Math.round((p.pagiTrx / p.totalTrx) * 100) : 0;
                  const sShare = p.totalTrx > 0 ? 100 - pShare : 0;

                  return (
                    <div
                      key={p.key}
                      className="h-full flex flex-col justify-end items-center group relative"
                    >
                      <div className="mb-1 text-[10px] font-mono-num font-extrabold text-violet-950 whitespace-nowrap">
                        {p.totalTrx > 0 ? `${p.totalTrx}` : '-'}
                      </div>
                      <div
                        style={{ height: `${hPct}%` }}
                        className="w-full max-w-[38px] rounded-t-lg overflow-hidden flex flex-col border border-slate-900/15 shadow-2xs"
                        title={`${p.label} — Total Trx: ${p.totalTrx} trx (Pagi: ${p.pagiTrx} | Siang: ${p.siangTrx})`}
                      >
                        <div style={{ height: `${sShare}%` }} className="w-full bg-indigo-600" />
                        <div style={{ height: `${pShare}%` }} className="w-full bg-amber-400" />
                      </div>
                      <div className="absolute -bottom-9 inset-x-0 text-center">
                        <div
                          className={`text-[10px] font-extrabold truncate ${
                            p.isWeekend ? 'text-rose-700' : 'text-slate-800'
                          }`}
                        >
                          {p.label}
                        </div>
                        {p.subLabel && !isCompactX && (
                          <div className="text-[9px] font-mono-num text-slate-500 mt-0.5">
                            {p.subLabel}
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

        {/* =================================================================
            DIAGRAM 4: PRICE POINT / APC (RP / PCS)
           ================================================================= */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-2xs flex flex-col justify-between">
          <div className="flex flex-wrap items-start justify-between gap-2 mb-5">
            <div>
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-extrabold uppercase tracking-wide mb-1">
                4. Diagram Price Point / APC (Rp/pcs)
              </div>
              <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                Grafik Harga Rata-Rata Per Pcs (Sales ÷ Qty)
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5">
                Mengukur nilai rata-rata harga per barang yang terjual
              </div>
            </div>
            <span className="px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 text-xs font-extrabold font-mono-num">
              Rata-rata: Rp {formatNumberId(avgPricePoint) || '0'}
            </span>
          </div>

          <div className="overflow-x-auto pb-2">
            <div
              style={{ minWidth: minChartWidth }}
              className="relative h-60 pt-7 pb-11 pl-14 pr-4 select-none"
            >
              {[1, 0.75, 0.5, 0.25, 0].map((ratio) => {
                const val = Math.round(maxPP * ratio);
                return (
                  <div
                    key={ratio}
                    style={{ bottom: `${ratio * 100}%` }}
                    className="absolute left-0 right-0 flex items-center pointer-events-none"
                  >
                    <span className="w-12 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-400">
                      {val >= 1000 ? `${Math.round(val / 1000)}rb` : val}
                    </span>
                    <div className="flex-1 border-b border-dashed border-slate-200" />
                  </div>
                );
              })}

              {points.length > 1 && (
                <svg
                  viewBox={`0 0 ${points.length * 100} 200`}
                  preserveAspectRatio="none"
                  className="absolute inset-0 left-14 right-4 top-7 bottom-11 w-[calc(100%-4.5rem)] h-[calc(100%-4.5rem)] pointer-events-none z-10 overflow-visible"
                >
                  {(() => {
                    const pts = points.map((p, i) => {
                      const x = (i + 0.5) * 100;
                      const y =
                        200 - Math.max(0, Math.min(200, (p.pricePoint / maxPP) * 200));
                      return { x, y, val: p.pricePoint };
                    });
                    const poly = pts.map((pt) => `${pt.x},${pt.y}`).join(' ');
                    return (
                      <>
                        <polyline
                          fill="none"
                          stroke="#059669"
                          strokeWidth="2.75"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          points={poly}
                        />
                        {pts.map((pt, idx) => (
                          <circle
                            key={idx}
                            cx={pt.x}
                            cy={pt.y}
                            r="4"
                            fill="#ffffff"
                            stroke="#047857"
                            strokeWidth="2.2"
                          />
                        ))}
                      </>
                    );
                  })()}
                </svg>
              )}

              <div
                style={{
                  gridTemplateColumns: `repeat(${Math.max(1, points.length)}, minmax(0, 1fr))`,
                }}
                className="relative z-0 h-full grid gap-2 items-end"
              >
                {points.map((p) => {
                  const hPct =
                    maxPP > 0
                      ? Math.max(
                          p.pricePoint > 0 ? 6 : 1,
                          Math.round((p.pricePoint / maxPP) * 100)
                        )
                      : 0;

                  return (
                    <div
                      key={p.key}
                      className="h-full flex flex-col justify-end items-center group relative"
                    >
                      <div className="mb-1 text-[10px] font-mono-num font-extrabold text-emerald-900 whitespace-nowrap">
                        {p.pricePoint > 0 ? `${Math.round(p.pricePoint / 1000)}k` : '-'}
                      </div>
                      <div
                        style={{ height: `${hPct}%` }}
                        className="w-full max-w-[38px] rounded-t-lg bg-gradient-to-t from-emerald-600 to-teal-400 border border-emerald-900/15 shadow-2xs"
                        title={`${p.label} — Price Point (APC): Rp ${formatNumberId(p.pricePoint)} / pcs`}
                      />
                      <div className="absolute -bottom-9 inset-x-0 text-center">
                        <div
                          className={`text-[10px] font-extrabold truncate ${
                            p.isWeekend ? 'text-rose-700' : 'text-slate-800'
                          }`}
                        >
                          {p.label}
                        </div>
                        {p.subLabel && !isCompactX && (
                          <div className="text-[9px] font-mono-num text-slate-500 mt-0.5">
                            {p.subLabel}
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

        {/* =================================================================
            DIAGRAM 5: BIG SIZE / BASKET SIZE (RP / TRX)
           ================================================================= */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 shadow-2xs flex flex-col justify-between">
          <div className="flex flex-wrap items-start justify-between gap-2 mb-5">
            <div>
              <div className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded bg-rose-600 text-white text-[10px] font-extrabold uppercase tracking-wide mb-1">
                5. Diagram Big Size / Basket Size (Rp/trx)
              </div>
              <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                Grafik Nilai Belanja Rata-Rata Per Transaksi (Sales ÷ Trx)
              </div>
              <div className="text-[11px] text-slate-500 mt-0.5">
                Mengukur besarnya nominal belanja rata-rata setiap transaksi
              </div>
            </div>
            <span className="px-3 py-1.5 rounded-lg bg-rose-50 border border-rose-200 text-rose-900 text-xs font-extrabold font-mono-num">
              Rata-rata: Rp {formatNumberId(avgBigSize) || '0'}
            </span>
          </div>

          <div className="overflow-x-auto pb-2">
            <div
              style={{ minWidth: minChartWidth }}
              className="relative h-60 pt-7 pb-11 pl-14 pr-4 select-none"
            >
              {[1, 0.75, 0.5, 0.25, 0].map((ratio) => {
                const val = Math.round(maxBS * ratio);
                return (
                  <div
                    key={ratio}
                    style={{ bottom: `${ratio * 100}%` }}
                    className="absolute left-0 right-0 flex items-center pointer-events-none"
                  >
                    <span className="w-12 pr-2 text-right text-[10px] font-mono-num font-bold text-slate-400">
                      {val >= 1000 ? `${Math.round(val / 1000)}rb` : val}
                    </span>
                    <div className="flex-1 border-b border-dashed border-slate-200" />
                  </div>
                );
              })}

              {points.length > 1 && (
                <svg
                  viewBox={`0 0 ${points.length * 100} 200`}
                  preserveAspectRatio="none"
                  className="absolute inset-0 left-14 right-4 top-7 bottom-11 w-[calc(100%-4.5rem)] h-[calc(100%-4.5rem)] pointer-events-none z-10 overflow-visible"
                >
                  {(() => {
                    const pts = points.map((p, i) => {
                      const x = (i + 0.5) * 100;
                      const y =
                        200 - Math.max(0, Math.min(200, (p.bigSize / maxBS) * 200));
                      return { x, y, val: p.bigSize };
                    });
                    const poly = pts.map((pt) => `${pt.x},${pt.y}`).join(' ');
                    return (
                      <>
                        <polyline
                          fill="none"
                          stroke="#e11d48"
                          strokeWidth="2.75"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          points={poly}
                        />
                        {pts.map((pt, idx) => (
                          <circle
                            key={idx}
                            cx={pt.x}
                            cy={pt.y}
                            r="4"
                            fill="#ffffff"
                            stroke="#be123c"
                            strokeWidth="2.2"
                          />
                        ))}
                      </>
                    );
                  })()}
                </svg>
              )}

              <div
                style={{
                  gridTemplateColumns: `repeat(${Math.max(1, points.length)}, minmax(0, 1fr))`,
                }}
                className="relative z-0 h-full grid gap-2 items-end"
              >
                {points.map((p) => {
                  const hPct =
                    maxBS > 0
                      ? Math.max(p.bigSize > 0 ? 6 : 1, Math.round((p.bigSize / maxBS) * 100))
                      : 0;

                  return (
                    <div
                      key={p.key}
                      className="h-full flex flex-col justify-end items-center group relative"
                    >
                      <div className="mb-1 text-[10px] font-mono-num font-extrabold text-rose-900 whitespace-nowrap">
                        {p.bigSize > 0 ? `${Math.round(p.bigSize / 1000)}k` : '-'}
                      </div>
                      <div
                        style={{ height: `${hPct}%` }}
                        className="w-full max-w-[38px] rounded-t-lg bg-gradient-to-t from-rose-600 to-orange-400 border border-rose-900/15 shadow-2xs"
                        title={`${p.label} — Big Size (Basket Size): Rp ${formatNumberId(p.bigSize)} / trx`}
                      />
                      <div className="absolute -bottom-9 inset-x-0 text-center">
                        <div
                          className={`text-[10px] font-extrabold truncate ${
                            p.isWeekend ? 'text-rose-700' : 'text-slate-800'
                          }`}
                        >
                          {p.label}
                        </div>
                        {p.subLabel && !isCompactX && (
                          <div className="text-[9px] font-mono-num text-slate-500 mt-0.5">
                            {p.subLabel}
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
  );
};
