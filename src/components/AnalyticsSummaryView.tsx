import React from 'react';
import {
  WeeklySalesTable,
  calculateDayMetrics,
  calculateWeekMetrics,
  formatNumberId,
  getIndoDayName,
  isWeekendDate,
  getWeeklyPeriodeText,
} from '../data/initialData';
import { Filter } from 'lucide-react';
import { PhotoStyleSalesDiagram } from './PhotoStyleSalesDiagram';

interface AnalyticsSummaryViewProps {
  tables: WeeklySalesTable[];
  selectedWeeklyFilter: string; // 'ALL' or table.id
  onSelectWeeklyFilter: (filterId: string) => void;
}

interface StaffAgg {
  nama: string;
  pagiShifts: number;
  siangShifts: number;
  totalSales: number;
  totalQty: number;
  totalTrx: number;
}

export const AnalyticsSummaryView: React.FC<AnalyticsSummaryViewProps> = ({
  tables,
  selectedWeeklyFilter,
  onSelectWeeklyFilter,
}) => {
  const activeTables =
    selectedWeeklyFilter === 'ALL'
      ? tables
      : tables.filter((t) => t.id === selectedWeeklyFilter);

  const validTables = activeTables.length > 0 ? activeTables : tables;

  // Aggregate weekly metrics across validTables
  let totalSales = 0;
  let totalQty = 0;
  let totalTrx = 0;

  for (const t of validTables) {
    const wm = calculateWeekMetrics(t);
    totalSales += wm.weeklySales;
    totalQty += wm.weeklyQty;
    totalTrx += wm.weeklyTrx;
  }

  const avgPricePoint = totalQty > 0 ? Math.round(totalSales / totalQty) : null;
  const avgBigSize = totalTrx > 0 ? Math.round(totalSales / totalTrx) : null;

  // Aggregate by Staff across validTables
  const staffMap = new Map<string, StaffAgg>();
  let totalPagiSales = 0;
  let totalPagiQty = 0;
  let totalPagiTrx = 0;

  let totalSiangSales = 0;
  let totalSiangQty = 0;
  let totalSiangTrx = 0;

  for (const table of validTables) {
    for (const day of table.days) {
      // Pagi
      const pName = (day.pagi.nama || 'Tanpa Nama').trim();
      const pSales = day.pagi.sales ?? 0;
      const pQty = day.pagi.qty ?? 0;
      const pTrx = day.pagi.trx ?? 0;

      totalPagiSales += pSales;
      totalPagiQty += pQty;
      totalPagiTrx += pTrx;

      const existingP = staffMap.get(pName) || {
        nama: pName,
        pagiShifts: 0,
        siangShifts: 0,
        totalSales: 0,
        totalQty: 0,
        totalTrx: 0,
      };
      existingP.pagiShifts += 1;
      existingP.totalSales += pSales;
      existingP.totalQty += pQty;
      existingP.totalTrx += pTrx;
      staffMap.set(pName, existingP);

      // Siang
      const sName = (day.siang.nama || 'Tanpa Nama').trim();
      const sSales = day.siang.sales ?? 0;
      const sQty = day.siang.qty ?? 0;
      const sTrx = day.siang.trx ?? 0;

      totalSiangSales += sSales;
      totalSiangQty += sQty;
      totalSiangTrx += sTrx;

      const existingS = staffMap.get(sName) || {
        nama: sName,
        pagiShifts: 0,
        siangShifts: 0,
        totalSales: 0,
        totalQty: 0,
        totalTrx: 0,
      };
      existingS.siangShifts += 1;
      existingS.totalSales += sSales;
      existingS.totalQty += sQty;
      existingS.totalTrx += sTrx;
      staffMap.set(sName, existingS);
    }
  }

  const staffList = Array.from(staffMap.values()).sort(
    (a, b) => b.totalSales - a.totalSales
  );

  const primaryChartTable = validTables[0];
  const maxDaySales = Math.max(
    1,
    ...primaryChartTable.days.map((d) => calculateDayMetrics(d).totalSales)
  );

  const filterTitle =
    selectedWeeklyFilter === 'ALL'
      ? `Semua Weekly (${tables.length} Minggu)`
      : primaryChartTable.title;

  return (
    <div className="space-y-6">
      {/* Filter Weekly Selector Bar */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-white border border-slate-200 rounded-xl p-3.5 sm:px-5 w-full max-w-full overflow-hidden">
        <div className="flex flex-wrap items-center gap-2 min-w-0 max-w-full">
          <div className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-800 mr-1 shrink-0">
            <Filter className="w-4 h-4 text-indigo-600" />
            <span>Filter Weekly:</span>
          </div>

          <select
            aria-label="Filter Weekly Petugas"
            value={selectedWeeklyFilter}
            onChange={(e) => onSelectWeeklyFilter(e.target.value)}
            className="px-2.5 py-1.5 text-xs font-bold bg-amber-50 border border-amber-300 text-amber-950 rounded-lg focus:outline-none min-w-0 max-w-[220px] truncate"
          >
            <option value="ALL">Semua Weekly ({tables.length} Minggu)</option>
            {tables.map((t) => (
              <option key={t.id} value={t.id}>
                Weekly {t.weekNumber} ({t.periode})
              </option>
            ))}
          </select>

          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-lg max-w-full">
            <button
              type="button"
              onClick={() => onSelectWeeklyFilter('ALL')}
              className={`px-2.5 py-1 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                selectedWeeklyFilter === 'ALL'
                  ? 'bg-slate-900 text-white shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Semua ({tables.length})
            </button>
            {tables.slice(0, 6).map((t) => (
              <button
                key={t.id}
                type="button"
                onClick={() => onSelectWeeklyFilter(t.id)}
                className={`px-2.5 py-1 text-xs font-semibold rounded-md transition-colors whitespace-nowrap ${
                  selectedWeeklyFilter === t.id
                    ? 'bg-slate-900 text-white shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                W{t.weekNumber}
              </button>
            ))}
          </div>
        </div>

        <div className="text-xs text-slate-500 truncate max-w-full">
          Menampilkan analitik untuk: <strong className="text-slate-900">{filterTitle}</strong>
        </div>
      </div>

      {/* KPI Summary Strip */}
      <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <div className="text-xs text-slate-500">Total Sales ({filterTitle})</div>
          <div className="mt-1 text-xl font-bold text-slate-900 font-mono-num">
            Rp {formatNumberId(totalSales) || '0'}
          </div>
          <div className="mt-1 text-xs text-slate-500">
            {validTables.length * 7} Hari · {validTables.length * 14} Shift
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <div className="text-xs text-slate-500">Total Qty Terjual</div>
          <div className="mt-1 text-xl font-bold text-slate-900 font-mono-num">
            {formatNumberId(totalQty) || '0'} pcs
          </div>
          <div className="mt-1 text-xs text-slate-500 font-mono-num">
            Rata-rata {Math.round(totalQty / (validTables.length * 7))} pcs / hari
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <div className="text-xs text-slate-500">Total Transaksi (Trx)</div>
          <div className="mt-1 text-xl font-bold text-slate-900 font-mono-num">
            {formatNumberId(totalTrx) || '0'} trx
          </div>
          <div className="mt-1 text-xs text-slate-500 font-mono-num">
            Rata-rata {(totalTrx / (validTables.length * 7)).toFixed(1)} trx / hari
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <div className="text-xs text-slate-500">Price Point (Sales / Qty)</div>
          <div className="mt-1 text-xl font-bold text-slate-900 font-mono-num">
            Rp {avgPricePoint ? formatNumberId(avgPricePoint) : '0'}
          </div>
          <div className="mt-1 text-xs text-slate-500">
            Harga rata-rata per barang
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <div className="text-xs text-slate-500">Big Size (Sales / Trx)</div>
          <div className="mt-1 text-xl font-bold text-slate-900 font-mono-num">
            Rp {avgBigSize ? formatNumberId(avgBigSize) : '0'}
          </div>
          <div className="mt-1 text-xs text-slate-500">
            Nilai rata-rata per transaksi (Trx)
          </div>
        </div>
      </div>

      {/* Staff Performance Table */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
          <h3 className="text-sm font-bold text-slate-900">
            Rekap Performa Petugas ({filterTitle})
          </h3>
          <span className="text-xs text-slate-500">Diurutkan berdasarkan Total Sales</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-50 border-b border-slate-200 text-slate-600">
                <th className="py-2.5 px-4 font-semibold">Nama Petugas</th>
                <th className="py-2.5 px-3 font-semibold text-right">Jml Shift</th>
                <th className="py-2.5 px-3 font-semibold text-right">Total Sales</th>
                <th className="py-2.5 px-3 font-semibold text-right">Rata-rata / Shift</th>
                <th className="py-2.5 px-3 font-semibold text-right">Qty</th>
                <th className="py-2.5 px-3 font-semibold text-right">Trx</th>
                <th className="py-2.5 px-3 font-semibold text-right">Price Point</th>
                <th className="py-2.5 px-4 font-semibold text-right">Big Size</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {staffList.map((st) => {
                const totalShifts = st.pagiShifts + st.siangShifts;
                const avgShift = totalShifts > 0 ? Math.round(st.totalSales / totalShifts) : 0;
                const pp = st.totalQty > 0 ? Math.round(st.totalSales / st.totalQty) : 0;
                const bs = st.totalTrx > 0 ? Math.round(st.totalSales / st.totalTrx) : 0;
                const sharePct =
                  totalSales > 0
                    ? ((st.totalSales / totalSales) * 100).toFixed(1)
                    : '0.0';

                return (
                  <tr key={st.nama} className="hover:bg-slate-50">
                    <td className="py-2.5 px-4 font-semibold text-slate-900">
                      <div>{st.nama}</div>
                      <div className="text-[11px] font-normal text-slate-500">
                        {st.pagiShifts} Pagi · {st.siangShifts} Siang ({sharePct}% omzet)
                      </div>
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                      {totalShifts}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num font-semibold text-slate-900">
                      {formatNumberId(st.totalSales)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                      {formatNumberId(avgShift)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                      {formatNumberId(st.totalQty)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                      {formatNumberId(st.totalTrx)}
                    </td>
                    <td className="py-2.5 px-3 text-right font-mono-num text-slate-700">
                      {formatNumberId(pp)}
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono-num text-slate-700">
                      {formatNumberId(bs)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Daily Trend Bar Chart (Sesuai Foto) */}
      <PhotoStyleSalesDiagram
        headerTitle={`DIAGRAM SALES WEEKLY — ${primaryChartTable.title} (${getWeeklyPeriodeText(primaryChartTable)})`}
        headerSubtitle="Diagram Batang Vertikal 7 Hari & Garis Tren Penjualan Shift Pagi vs Shift Siang"
        badgeText="DIAGRAM SALES (OMZET RUPIAH)"
        cardTitle="Grafik Batang Shift Pagi vs Siang & Garis Tren"
        totalSales={primaryChartTable.days.reduce(
          (s, d) => s + calculateDayMetrics(d).totalSales,
          0
        )}
        totalPagi={primaryChartTable.days.reduce((s, d) => s + (d.pagi.sales ?? 0), 0)}
        totalSiang={primaryChartTable.days.reduce((s, d) => s + (d.siang.sales ?? 0), 0)}
        bars={primaryChartTable.days.map((day, idx) => {
          const m = calculateDayMetrics(day);
          return {
            key: day.id,
            label: `${getIndoDayName(day.tanggal).slice(0, 3)} (H${idx + 1})`,
            subLabel: day.tanggal.slice(0, 5),
            pagiSales: day.pagi.sales ?? 0,
            siangSales: day.siang.sales ?? 0,
            totalSales: m.totalSales,
            isWeekend: isWeekendDate(day.tanggal),
          };
        })}
      />
    </div>
  );
};
