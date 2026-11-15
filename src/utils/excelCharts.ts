/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 *
 * Shared high-DPI HTML5 Canvas chart diagram renderers for ExcelJS (.xlsx) exports,
 * matching the visual aesthetic of Report DP (Stacked Bars + Trendline + Growth % & Contribution %).
 */

import ExcelJS from 'exceljs';

export interface ExcelChartBarItem {
  label: string;
  subLabel?: string;
  seriesAVal: number;
  seriesBVal: number;
  totalVal: number;
  isWeekend?: boolean;
}

export interface ExcelChartSeriesConfig {
  seriesALabel: string;
  seriesAColor: string;
  seriesALegendColor: string;
  seriesBLabel: string;
  seriesBColor: string;
  seriesBLegendColor: string;
  badgeText?: string;
  badgeBg?: string;
  cardSubtitle?: string;
}

export interface ExcelTrafficChartItem {
  label: string;
  subLabel?: string;
  trafficPagi: number;
  trafficSiang: number;
  totalTraffic: number;
  totalTrxNota: number;
  convRatePct: number; // 0..100
  isWeekend?: boolean;
}

function drawRoundedRect(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  r: number
) {
  const radius = Math.max(0, Math.min(r, w / 2, h / 2));
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + w - radius, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
  ctx.lineTo(x + w, y + h - radius);
  ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
  ctx.lineTo(x + radius, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function formatCompactNominal(val: number): string {
  if (val >= 1000000) return `${(val / 1000000).toFixed(1)}jt`;
  if (val > 0) return `${Math.round(val / 1000)}rb`;
  return '0';
}

function generateDualSalesDiagramBase64(
  chartTitle: string,
  chartSubtitle: string,
  bars: ExcelChartBarItem[],
  cfg: ExcelChartSeriesConfig,
  rightFilterBadgeText?: string
): string | null {
  if (typeof document === 'undefined' || bars.length === 0) return null;
  const dpr = 2;
  const isWide = bars.length > 14;
  const logicalW = isWide ? Math.max(980, bars.length * 32 + 90) : 760;
  const logicalH = 314;
  const canvas = document.createElement('canvas');
  canvas.width = logicalW * dpr;
  canvas.height = logicalH * dpr;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.scale(dpr, dpr);

  const width = logicalW;
  const height = logicalH;

  const totalA = bars.reduce((s, b) => s + (b.seriesAVal || 0), 0);
  const totalB = bars.reduce((s, b) => s + (b.seriesBVal || 0), 0);
  const grandTotal = totalA + totalB;
  const pctA = grandTotal > 0 ? Math.round((totalA / grandTotal) * 100) : 0;
  const pctB = grandTotal > 0 ? 100 - pctA : 0;

  // Outer Card Background
  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(0.75, 0.75, width - 1.5, height - 1.5);

  // Top Dark Navy Header Banner
  const headerH = 62;
  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, width, headerH);

  ctx.textAlign = 'left';
  ctx.fillStyle = '#FBBF24';
  ctx.font = 'bold 12px Inter, Calibri, sans-serif';
  ctx.fillText(chartTitle.toUpperCase(), 16, 20);

  if (rightFilterBadgeText) {
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const fbW = ctx.measureText(rightFilterBadgeText).width + 18;
    const fbX = width - 14 - fbW;
    const fbY = 7;
    ctx.fillStyle = '#1E293B';
    drawRoundedRect(ctx, fbX, fbY, fbW, 19, 4);
    ctx.fill();
    ctx.strokeStyle = '#FBBF24';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#FCD34D';
    ctx.textAlign = 'center';
    ctx.fillText(rightFilterBadgeText, fbX + fbW / 2, fbY + 13);
    ctx.textAlign = 'left';
  }

  ctx.fillStyle = '#CBD5E1';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillText(chartSubtitle, 16, 36);

  // Legend Line 3
  const legA = `${cfg.seriesALabel}: Rp ${totalA.toLocaleString('id-ID')} (${pctA}%)`;
  const legB = `${cfg.seriesBLabel}: Rp ${totalB.toLocaleString('id-ID')} (${pctB}%)`;
  const legTrend = 'Garis Tren Total Sales';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';

  ctx.fillStyle = cfg.seriesAColor;
  ctx.fillRect(16, 44, 9, 9);
  ctx.fillStyle = cfg.seriesALegendColor;
  ctx.fillText(legA, 29, 52);

  const bStartX = Math.max(255, 42 + ctx.measureText(legA).width + 20);
  ctx.fillStyle = cfg.seriesBColor;
  ctx.fillRect(bStartX, 44, 9, 9);
  ctx.fillStyle = cfg.seriesBLegendColor;
  ctx.fillText(legB, bStartX + 13, 52);

  const tStartX = Math.max(bStartX + 225, bStartX + 26 + ctx.measureText(legB).width + 20);
  if (tStartX + 130 < width) {
    ctx.strokeStyle = '#10B981';
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    ctx.moveTo(tStartX, 48.5);
    ctx.lineTo(tStartX + 15, 48.5);
    ctx.stroke();
    ctx.fillStyle = '#6EE7B7';
    ctx.fillText(legTrend, tStartX + 20, 52);
  }

  // Inner White Plot Card
  const cardX = 12;
  const cardY = headerH + 10;
  const cardW = width - 24;
  const cardH = height - headerH - 20;

  ctx.fillStyle = '#FFFFFF';
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, 8);
  ctx.fill();
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.2;
  ctx.stroke();

  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  const bText = cfg.badgeText || 'DIAGRAM SALES (OMZET RUPIAH)';
  const badgeW = Math.max(178, ctx.measureText(bText).width + 20);
  ctx.fillStyle = cfg.badgeBg || '#D97706';
  drawRoundedRect(ctx, cardX + 14, cardY + 12, badgeW, 20, 4);
  ctx.fill();
  ctx.fillStyle = '#FFFFFF';
  ctx.textAlign = 'left';
  ctx.fillText(bText, cardX + 22, cardY + 25.5);

  ctx.fillStyle = '#0F172A';
  ctx.font = 'bold 11px Inter, Calibri, sans-serif';
  ctx.fillText(
    cfg.cardSubtitle || `Grafik Batang ${cfg.seriesALabel} vs ${cfg.seriesBLabel} & Garis Tren`,
    cardX + 14,
    cardY + 47
  );

  const summaryRightText = `Total: Rp ${grandTotal.toLocaleString('id-ID')}`;
  ctx.font = 'bold 10.5px Inter, Calibri, sans-serif';
  const sumW = ctx.measureText(summaryRightText).width + 22;
  const pillX = cardX + cardW - 14 - sumW;
  const pillY = cardY + 11;
  ctx.fillStyle = '#ECFDF5';
  drawRoundedRect(ctx, pillX, pillY, sumW, 23, 5);
  ctx.fill();
  ctx.strokeStyle = '#A7F3D0';
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.textAlign = 'center';
  ctx.fillStyle = '#047857';
  ctx.fillText(summaryRightText, pillX + sumW / 2, pillY + 15.5);

  const plotX = cardX + 52;
  const plotY = cardY + 72;
  const plotW = cardW - 68;
  const plotH = cardH - 112;

  const rawMax = Math.max(0, ...bars.map((b) => b.totalVal));
  const maxVal = rawMax > 0 ? rawMax : 2000000;

  [1, 0.5, 0].forEach((ratio) => {
    const gy = plotY + plotH * (1 - ratio);
    const v = Math.round(maxVal * ratio);
    ctx.save();
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1;
    if (ratio > 0) ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(plotX, gy);
    ctx.lineTo(plotX + plotW, gy);
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = '#64748B';
    ctx.font = 'bold 10px Inter, Calibri, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(formatCompactNominal(v), plotX - 10, gy + 3.5);
  });

  const count = Math.max(1, bars.length);
  const isCompactMany = count > 14;
  const slotW = plotW / count;
  const barW = isCompactMany
    ? Math.min(22, Math.max(10, slotW * 0.58))
    : Math.min(38, Math.max(16, slotW * 0.42));
  const trendPoints: Array<{ x: number; y: number; bar: ExcelChartBarItem }> = [];

  bars.forEach((b, idx) => {
    const cx = plotX + slotW * (idx + 0.5);
    const hRatio =
      maxVal > 0
        ? Math.max(b.totalVal > 0 ? 0.06 : 0.02, b.totalVal / maxVal)
        : 0.02;
    const totalH = Math.min(plotH, hRatio * plotH);
    const shareA = b.totalVal > 0 ? b.seriesAVal / b.totalVal : 0;
    const shareB = b.totalVal > 0 ? 1 - shareA : 0;
    const hB = totalH * shareB;
    const hA = totalH * shareA;

    const byTop = plotY + plotH - totalH;
    trendPoints.push({ x: cx, y: byTop, bar: b });

    ctx.save();
    drawRoundedRect(ctx, cx - barW / 2, byTop, barW, totalH, 2);
    ctx.clip();

    if (b.totalVal > 0) {
      if (hB > 0) {
        ctx.fillStyle = cfg.seriesBColor;
        ctx.fillRect(cx - barW / 2, byTop, barW, hB + 0.5);
      }
      if (hA > 0) {
        ctx.fillStyle = cfg.seriesAColor;
        ctx.fillRect(cx - barW / 2, byTop + hB, barW, hA + 0.5);
      }
    } else {
      ctx.fillStyle = '#E2E8F0';
      ctx.fillRect(cx - barW / 2, byTop, barW, totalH);
    }
    ctx.restore();

    ctx.fillStyle = b.isWeekend ? '#BE123C' : '#334155';
    ctx.font = isCompactMany
      ? 'bold 8px Inter, Calibri, sans-serif'
      : 'bold 9.5px Inter, Calibri, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(b.label, cx, plotY + plotH + 15);
    if (b.subLabel) {
      ctx.fillStyle = '#64748B';
      ctx.font = isCompactMany
        ? '7.5px Inter, Calibri, sans-serif'
        : '9px Inter, Calibri, sans-serif';
      ctx.fillText(b.subLabel, cx, plotY + plotH + 26);
    }
  });

  if (trendPoints.length > 1) {
    ctx.strokeStyle = '#10B981';
    ctx.lineWidth = 2.2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    trendPoints.forEach((pt, i) => {
      if (i === 0) ctx.moveTo(pt.x, pt.y);
      else ctx.lineTo(pt.x, pt.y);
    });
    ctx.stroke();

    trendPoints.forEach((pt) => {
      ctx.fillStyle = '#FFFFFF';
      ctx.strokeStyle = '#059669';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(pt.x, pt.y, isCompactMany ? 3 : 3.6, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
    });
  }

  trendPoints.forEach((pt) => {
    const labelText = formatCompactNominal(pt.bar.totalVal);
    ctx.font = isCompactMany
      ? 'bold 8px Inter, Calibri, sans-serif'
      : 'bold 9.5px Inter, Calibri, sans-serif';
    ctx.fillStyle = '#0F172A';
    ctx.textAlign = 'center';
    ctx.fillText(labelText, pt.x, Math.max(cardY + 56, pt.y - 7));
  });

  return canvas.toDataURL('image/png');
}

function generateGrowthContributionBase64(
  chartTitle: string,
  chartSubtitle: string,
  bars: ExcelChartBarItem[],
  rightFilterBadgeText?: string
): string | null {
  if (typeof document === 'undefined' || bars.length === 0) return null;
  const dpr = 2;
  const isWide = bars.length > 14;
  const logicalW = isWide ? Math.max(980, bars.length * 34 + 90) : 760;
  const logicalH = isWide ? 540 : 326;
  const canvas = document.createElement('canvas');
  canvas.width = logicalW * dpr;
  canvas.height = logicalH * dpr;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.scale(dpr, dpr);

  const width = logicalW;
  const height = logicalH;
  const grandTotalSales = bars.reduce((s, b) => s + (b.totalVal || 0), 0);

  const enriched = bars.map((b, idx) => {
    const contribPct =
      grandTotalSales > 0 ? (b.totalVal / grandTotalSales) * 100 : 0;
    const prevSales = idx > 0 ? bars[idx - 1].totalVal : null;
    let growthPct = 0;
    let growthStatus: 'UP' | 'DOWN' | 'FLAT' | 'BASE' | 'NEW' = 'FLAT';

    if (prevSales === null) {
      growthPct = 0;
      growthStatus = 'BASE';
    } else if (prevSales <= 0 && b.totalVal <= 0) {
      growthPct = 0;
      growthStatus = 'FLAT';
    } else if (prevSales <= 0 && b.totalVal > 0) {
      growthPct = 100;
      growthStatus = 'NEW';
    } else {
      growthPct = ((b.totalVal - prevSales) / prevSales) * 100;
      if (Math.abs(growthPct) < 0.05) {
        growthPct = 0;
        growthStatus = 'FLAT';
      } else {
        growthStatus = growthPct > 0 ? 'UP' : 'DOWN';
      }
    }
    return { ...b, contribPct, growthPct, growthStatus };
  });

  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(0.75, 0.75, width - 1.5, height - 1.5);

  const headerH = 62;
  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, width, headerH);

  ctx.fillStyle = '#10B981';
  drawRoundedRect(ctx, 16, 7, 130, 17, 4);
  ctx.fill();
  ctx.fillStyle = '#022C22';
  ctx.font = 'bold 9px Inter, Calibri, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('GROWTH & KONTRIBUSI', 16 + 65, 18.5);

  ctx.textAlign = 'left';
  ctx.fillStyle = '#FBBF24';
  ctx.font = 'bold 11.5px Inter, Calibri, sans-serif';
  ctx.fillText(chartTitle.toUpperCase(), 154, 19.5);

  if (rightFilterBadgeText) {
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const fbW = ctx.measureText(rightFilterBadgeText).width + 18;
    const fbX = width - 14 - fbW;
    const fbY = 7;
    ctx.fillStyle = '#1E293B';
    drawRoundedRect(ctx, fbX, fbY, fbW, 19, 4);
    ctx.fill();
    ctx.strokeStyle = '#FBBF24';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#FCD34D';
    ctx.textAlign = 'center';
    ctx.fillText(rightFilterBadgeText, fbX + fbW / 2, fbY + 13);
    ctx.textAlign = 'left';
  }

  ctx.fillStyle = '#CBD5E1';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillText(
    chartSubtitle ||
      'Visualisasi Diagram Batang Pertumbuhan (Growth %) & Porsi Kontribusi Sales (%)',
    16,
    36
  );

  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillStyle = '#10B981';
  ctx.fillRect(16, 44, 9, 9);
  ctx.fillStyle = '#6EE7B7';
  ctx.fillText('Growth Naik (▲ +%)', 29, 52);

  ctx.fillStyle = '#F43F5E';
  ctx.fillRect(160, 44, 9, 9);
  ctx.fillStyle = '#FDA4AF';
  ctx.fillText('Growth Turun (▼ -%)', 173, 52);

  ctx.fillStyle = '#0284C7';
  ctx.fillRect(310, 44, 9, 9);
  ctx.fillStyle = '#7DD3FC';
  ctx.fillText('Kontribusi terhadap Total (%)', 323, 52);

  const topContrib = [...enriched].sort((a, b) => b.contribPct - a.contribPct)[0];
  if (topContrib && topContrib.totalVal > 0) {
    ctx.fillStyle = '#FDE68A';
    ctx.fillText(
      `Tertinggi: ${topContrib.label} (${topContrib.contribPct.toFixed(1)}%)`,
      510,
      52
    );
  }

  const drawSubChart = (
    cardX: number,
    cardY: number,
    cardW: number,
    cardH: number,
    type: 'GROWTH' | 'KONTRIBUSI'
  ) => {
    ctx.fillStyle = '#FFFFFF';
    drawRoundedRect(ctx, cardX, cardY, cardW, cardH, 8);
    ctx.fill();
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1.2;
    ctx.stroke();

    const isGrowth = type === 'GROWTH';
    const badgeTxt = isGrowth
      ? 'DIAGRAM BATANG GROWTH (%)'
      : 'DIAGRAM BATANG KONTRIBUSI (%)';
    const badgeColor = isGrowth ? '#059669' : '#0284C7';
    const subTitleTxt = isGrowth
      ? 'Grafik Batang Pertumbuhan Sales (Naik ▲ / Turun ▼)'
      : 'Grafik Batang Porsi Kontribusi terhadap Total Sales (100%)';

    ctx.font = 'bold 9px Inter, Calibri, sans-serif';
    const bW = Math.max(150, ctx.measureText(badgeTxt).width + 16);
    ctx.fillStyle = badgeColor;
    drawRoundedRect(ctx, cardX + 12, cardY + 10, bW, 18, 4);
    ctx.fill();
    ctx.fillStyle = '#FFFFFF';
    ctx.textAlign = 'left';
    ctx.fillText(badgeTxt, cardX + 20, cardY + 22);

    ctx.fillStyle = '#0F172A';
    ctx.font = 'bold 10px Inter, Calibri, sans-serif';
    ctx.fillText(subTitleTxt, cardX + 12, cardY + 42);

    const topGrowthItem = [...enriched]
      .filter((it) => it.growthStatus === 'UP' || it.growthStatus === 'NEW')
      .sort((a, b) => b.growthPct - a.growthPct)[0];
    const rightTxt = isGrowth
      ? topGrowthItem
        ? `Tertinggi: ${topGrowthItem.label} (+${topGrowthItem.growthPct.toFixed(1)}%)`
        : `Total: Rp ${grandTotalSales.toLocaleString('id-ID')}`
      : `Total 100% (Rp ${grandTotalSales.toLocaleString('id-ID')})`;

    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const rW = ctx.measureText(rightTxt).width + 16;
    const rX = cardX + cardW - 12 - rW;
    const rY = cardY + 10;
    ctx.fillStyle = isGrowth ? '#ECFDF5' : '#F0F9FF';
    drawRoundedRect(ctx, rX, rY, rW, 20, 4);
    ctx.fill();
    ctx.strokeStyle = isGrowth ? '#A7F3D0' : '#BAE6FD';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = isGrowth ? '#047857' : '#0369A1';
    ctx.textAlign = 'center';
    ctx.fillText(rightTxt, rX + rW / 2, rY + 13.5);

    const plotX = cardX + 44;
    const plotY = cardY + 62;
    const plotW = cardW - 56;
    const plotH = cardH - 96;

    const maxScale = isGrowth
      ? Math.max(15, ...enriched.map((it) => Math.min(200, Math.abs(it.growthPct))))
      : Math.max(10, ...enriched.map((it) => it.contribPct));

    [1, 0.5, 0].forEach((ratio) => {
      const gy = plotY + plotH * (1 - ratio);
      const val = Math.round(maxScale * ratio);
      ctx.save();
      ctx.strokeStyle = '#CBD5E1';
      ctx.lineWidth = 1;
      if (ratio > 0) ctx.setLineDash([4, 4]);
      ctx.beginPath();
      ctx.moveTo(plotX, gy);
      ctx.lineTo(plotX + plotW, gy);
      ctx.stroke();
      ctx.restore();

      ctx.fillStyle = '#64748B';
      ctx.font = 'bold 9px Inter, Calibri, sans-serif';
      ctx.textAlign = 'right';
      const yLbl = isGrowth ? (val === 0 ? '0%' : `±${val}%`) : `${val}%`;
      ctx.fillText(yLbl, plotX - 6, gy + 3);
    });

    if (!isGrowth && enriched.length > 0 && grandTotalSales > 0) {
      const avgPct = 100 / enriched.length;
      if (avgPct <= maxScale) {
        const avgY = plotY + plotH * (1 - avgPct / maxScale);
        ctx.save();
        ctx.strokeStyle = '#F59E0B';
        ctx.lineWidth = 1.2;
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.moveTo(plotX, avgY);
        ctx.lineTo(plotX + plotW, avgY);
        ctx.stroke();
        ctx.restore();
      }
    }

    const count = Math.max(1, enriched.length);
    const isCompactMany = count > 10;
    const slotW = plotW / count;
    const barW = isCompactMany
      ? Math.min(20, Math.max(9, slotW * 0.56))
      : Math.min(32, Math.max(14, slotW * 0.44));

    const pts: Array<{ x: number; y: number; item: (typeof enriched)[number] }> = [];

    enriched.forEach((it, idx) => {
      const cx = plotX + slotW * (idx + 0.5);
      const val = isGrowth
        ? Math.min(maxScale, Math.abs(it.growthPct))
        : it.contribPct;
      const minRatio = isGrowth
        ? it.growthStatus === 'UP' || it.growthStatus === 'DOWN' || it.growthStatus === 'NEW'
          ? 0.08
          : 0.03
        : it.totalVal > 0
          ? 0.06
          : 0.02;
      const hRatio = maxScale > 0 ? Math.max(minRatio, val / maxScale) : 0.03;
      const barH = Math.min(plotH, hRatio * plotH);
      const byTop = plotY + plotH - barH;
      pts.push({ x: cx, y: byTop, item: it });

      ctx.save();
      drawRoundedRect(ctx, cx - barW / 2, byTop, barW, barH, 2.5);
      ctx.clip();

      if (isGrowth) {
        if (it.growthStatus === 'UP' || it.growthStatus === 'NEW') {
          ctx.fillStyle = '#10B981';
        } else if (it.growthStatus === 'DOWN') {
          ctx.fillStyle = '#F43F5E';
        } else if (it.growthStatus === 'BASE') {
          ctx.fillStyle = '#F59E0B';
        } else {
          ctx.fillStyle = '#CBD5E1';
        }
        ctx.fillRect(cx - barW / 2, byTop, barW, barH);
      } else {
        if (it.totalVal > 0) {
          const shareA = it.totalVal > 0 ? it.seriesAVal / it.totalVal : 0;
          const shareB = 1 - shareA;
          ctx.fillStyle = '#0284C7';
          ctx.fillRect(cx - barW / 2, byTop, barW, barH * shareB + 0.5);
          ctx.fillStyle = '#38BDF8';
          ctx.fillRect(cx - barW / 2, byTop + barH * shareB, barW, barH * shareA + 0.5);
        } else {
          ctx.fillStyle = '#E2E8F0';
          ctx.fillRect(cx - barW / 2, byTop, barW, barH);
        }
      }
      ctx.restore();

      ctx.fillStyle = it.isWeekend ? '#BE123C' : '#334155';
      ctx.font = isCompactMany
        ? 'bold 7.5px Inter, Calibri, sans-serif'
        : 'bold 9px Inter, Calibri, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(it.label, cx, plotY + plotH + 14);
      if (it.subLabel && !isCompactMany) {
        ctx.fillStyle = '#64748B';
        ctx.font = '8px Inter, Calibri, sans-serif';
        ctx.fillText(it.subLabel, cx, plotY + plotH + 24);
      }
    });

    if (pts.length > 1) {
      ctx.save();
      ctx.strokeStyle = isGrowth ? '#0D9488' : '#0284C7';
      ctx.lineWidth = 2;
      if (isGrowth) ctx.setLineDash([4, 3]);
      ctx.beginPath();
      pts.forEach((p, i) => {
        if (i === 0) ctx.moveTo(p.x, p.y);
        else ctx.lineTo(p.x, p.y);
      });
      ctx.stroke();
      ctx.restore();

      pts.forEach((p) => {
        ctx.fillStyle = '#FFFFFF';
        ctx.strokeStyle = isGrowth
          ? p.item.growthStatus === 'DOWN'
            ? '#E11D48'
            : p.item.growthStatus === 'UP' || p.item.growthStatus === 'NEW'
              ? '#059669'
              : '#64748B'
          : '#0369A1';
        ctx.lineWidth = 1.8;
        ctx.beginPath();
        ctx.arc(p.x, p.y, isCompactMany ? 2.6 : 3.2, 0, Math.PI * 2);
        ctx.fill();
        ctx.stroke();
      });
    }

    pts.forEach((p) => {
      const it = p.item;
      let valTxt = '';
      if (isGrowth) {
        if (it.growthStatus === 'BASE') valTxt = 'Basis';
        else if (it.growthStatus === 'UP' || it.growthStatus === 'NEW') valTxt = `+${it.growthPct.toFixed(1)}%`;
        else if (it.growthStatus === 'DOWN') valTxt = `${it.growthPct.toFixed(1)}%`;
        else valTxt = '0%';
      } else {
        valTxt = `${it.contribPct.toFixed(1)}%`;
      }

      ctx.font = isCompactMany
        ? 'bold 7.5px Inter, Calibri, sans-serif'
        : 'bold 8.5px Inter, Calibri, sans-serif';
      ctx.fillStyle = isGrowth
        ? it.growthStatus === 'DOWN'
          ? '#BE123C'
          : it.growthStatus === 'UP' || it.growthStatus === 'NEW'
            ? '#047857'
            : '#475569'
        : '#0369A1';
      ctx.textAlign = 'center';
      ctx.fillText(valTxt, p.x, Math.max(cardY + 50, p.y - 6));
    });
  };

  if (isWide) {
    const eachH = Math.floor((height - headerH - 28) / 2);
    drawSubChart(12, headerH + 8, width - 24, eachH, 'GROWTH');
    drawSubChart(12, headerH + 16 + eachH, width - 24, eachH, 'KONTRIBUSI');
  } else {
    const eachW = Math.floor((width - 32) / 2);
    const subH = height - headerH - 18;
    drawSubChart(12, headerH + 9, eachW, subH, 'GROWTH');
    drawSubChart(20 + eachW, headerH + 9, eachW, subH, 'KONTRIBUSI');
  }

  return canvas.toDataURL('image/png');
}

function generateTrafficConversionDiagramBase64(
  chartTitle: string,
  chartSubtitle: string,
  items: ExcelTrafficChartItem[],
  rightFilterBadgeText?: string
): string | null {
  if (typeof document === 'undefined' || items.length === 0) return null;
  const dpr = 2;
  const isWide = items.length > 14;
  const logicalW = isWide ? Math.max(980, items.length * 32 + 90) : 760;
  const logicalH = 310;
  const canvas = document.createElement('canvas');
  canvas.width = logicalW * dpr;
  canvas.height = logicalH * dpr;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.scale(dpr, dpr);

  const width = logicalW;
  const height = logicalH;

  const totPagi = items.reduce((s, it) => s + it.trafficPagi, 0);
  const totSiang = items.reduce((s, it) => s + it.trafficSiang, 0);
  const totTraffic = totPagi + totSiang;
  const totTrxNota = items.reduce((s, it) => s + it.totalTrxNota, 0);
  const overallConv = totTraffic > 0 ? (totTrxNota / totTraffic) * 100 : 0;

  ctx.fillStyle = '#F1F5F9';
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.5;
  ctx.strokeRect(0.75, 0.75, width - 1.5, height - 1.5);

  const headerH = 62;
  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, width, headerH);

  ctx.fillStyle = '#14B8A6';
  drawRoundedRect(ctx, 16, 7, 150, 17, 4);
  ctx.fill();
  ctx.fillStyle = '#042F2E';
  ctx.font = 'bold 9px Inter, Calibri, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('TRAFFIC & CONV. RATE', 16 + 75, 18.5);

  ctx.textAlign = 'left';
  ctx.fillStyle = '#FBBF24';
  ctx.font = 'bold 11.5px Inter, Calibri, sans-serif';
  ctx.fillText(chartTitle.toUpperCase(), 174, 19.5);

  if (rightFilterBadgeText) {
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    const fbW = ctx.measureText(rightFilterBadgeText).width + 18;
    const fbX = width - 14 - fbW;
    const fbY = 7;
    ctx.fillStyle = '#1E293B';
    drawRoundedRect(ctx, fbX, fbY, fbW, 19, 4);
    ctx.fill();
    ctx.strokeStyle = '#2DD4BF';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#99F6E4';
    ctx.textAlign = 'center';
    ctx.fillText(rightFilterBadgeText, fbX + fbW / 2, fbY + 13);
    ctx.textAlign = 'left';
  }

  ctx.fillStyle = '#CBD5E1';
  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  ctx.fillText(chartSubtitle, 16, 36);

  ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
  const legPagi = `Traffic Kasir Pagi: ${totPagi.toLocaleString('id-ID')} org`;
  const legSiang = `Traffic Kasir Siang: ${totSiang.toLocaleString('id-ID')} org`;
  ctx.fillStyle = '#0EA5E9';
  ctx.fillRect(16, 44, 9, 9);
  ctx.fillStyle = '#BAE6FD';
  ctx.fillText(legPagi, 29, 52);

  const sx = 240;
  ctx.fillStyle = '#6366F1';
  ctx.fillRect(sx, 44, 9, 9);
  ctx.fillStyle = '#C7D2FE';
  ctx.fillText(legSiang, sx + 13, 52);

  const cxLeg = 470;
  ctx.strokeStyle = '#F59E0B';
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  ctx.moveTo(cxLeg, 48.5);
  ctx.lineTo(cxLeg + 15, 48.5);
  ctx.stroke();
  ctx.fillStyle = '#FDE68A';
  ctx.fillText(`Conversion Rate % (Rata-rata: ${overallConv.toFixed(1)}%)`, cxLeg + 20, 52);

  // Inner Card
  const cardX = 12;
  const cardY = headerH + 10;
  const cardW = width - 24;
  const cardH = height - headerH - 20;

  ctx.fillStyle = '#FFFFFF';
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, 8);
  ctx.fill();
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1.2;
  ctx.stroke();

  ctx.font = 'bold 9px Inter, Calibri, sans-serif';
  ctx.fillStyle = '#0F766E';
  drawRoundedRect(ctx, cardX + 14, cardY + 11, 215, 19, 4);
  ctx.fill();
  ctx.fillStyle = '#FFFFFF';
  ctx.textAlign = 'left';
  ctx.fillText('DIAGRAM TRAFFIC KASIR & CONV. RATE', cardX + 22, cardY + 24);

  const rightPill = `Total: ${totTraffic.toLocaleString('id-ID')} Org · ${totTrxNota.toLocaleString('id-ID')} Trx+Nota · CR ${overallConv.toFixed(1)}%`;
  ctx.font = 'bold 10px Inter, Calibri, sans-serif';
  const rW = ctx.measureText(rightPill).width + 20;
  const rX = cardX + cardW - 14 - rW;
  const rY = cardY + 10;
  ctx.fillStyle = '#F0FDFA';
  drawRoundedRect(ctx, rX, rY, rW, 22, 5);
  ctx.fill();
  ctx.strokeStyle = '#99F6E4';
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.textAlign = 'center';
  ctx.fillStyle = '#0F766E';
  ctx.fillText(rightPill, rX + rW / 2, rY + 14.5);

  const plotX = cardX + 48;
  const plotY = cardY + 56;
  const plotW = cardW - 92;
  const plotH = cardH - 94;

  const maxTraffic = Math.max(20, ...items.map((it) => it.totalTraffic));
  const maxConv = Math.max(50, Math.min(100, Math.ceil(Math.max(...items.map((it) => it.convRatePct), 10) / 10) * 10));

  [1, 0.5, 0].forEach((ratio) => {
    const gy = plotY + plotH * (1 - ratio);
    const vTraf = Math.round(maxTraffic * ratio);
    const vConv = Math.round(maxConv * ratio);

    ctx.save();
    ctx.strokeStyle = '#CBD5E1';
    ctx.lineWidth = 1;
    if (ratio > 0) ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(plotX, gy);
    ctx.lineTo(plotX + plotW, gy);
    ctx.stroke();
    ctx.restore();

    ctx.fillStyle = '#0284C7';
    ctx.font = 'bold 9.5px Inter, Calibri, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(`${vTraf}`, plotX - 8, gy + 3);

    ctx.fillStyle = '#D97706';
    ctx.textAlign = 'left';
    ctx.fillText(`${vConv}%`, plotX + plotW + 8, gy + 3);
  });

  const count = Math.max(1, items.length);
  const isCompactMany = count > 14;
  const slotW = plotW / count;
  const barW = isCompactMany
    ? Math.min(22, Math.max(10, slotW * 0.56))
    : Math.min(36, Math.max(15, slotW * 0.42));

  const convPts: Array<{ x: number; y: number; item: ExcelTrafficChartItem }> = [];

  items.forEach((it, idx) => {
    const cx = plotX + slotW * (idx + 0.5);
    const hRatio = maxTraffic > 0 ? Math.max(it.totalTraffic > 0 ? 0.05 : 0.02, it.totalTraffic / maxTraffic) : 0.02;
    const totalH = Math.min(plotH, hRatio * plotH);
    const pagiShare = it.totalTraffic > 0 ? it.trafficPagi / it.totalTraffic : 0;
    const siangShare = it.totalTraffic > 0 ? 1 - pagiShare : 0;
    const byTop = plotY + plotH - totalH;

    ctx.save();
    drawRoundedRect(ctx, cx - barW / 2, byTop, barW, totalH, 2);
    ctx.clip();
    if (it.totalTraffic > 0) {
      ctx.fillStyle = '#6366F1';
      ctx.fillRect(cx - barW / 2, byTop, barW, totalH * siangShare + 0.5);
      ctx.fillStyle = '#0EA5E9';
      ctx.fillRect(cx - barW / 2, byTop + totalH * siangShare, barW, totalH * pagiShare + 0.5);
    } else {
      ctx.fillStyle = '#E2E8F0';
      ctx.fillRect(cx - barW / 2, byTop, barW, totalH);
    }
    ctx.restore();

    const convRatio = maxConv > 0 ? Math.min(1, Math.max(0, it.convRatePct / maxConv)) : 0;
    const convY = plotY + plotH - convRatio * plotH;
    convPts.push({ x: cx, y: convY, item: it });

    // Traffic Count inside or above bar if space, X label at bottom
    ctx.fillStyle = it.isWeekend ? '#BE123C' : '#334155';
    ctx.font = isCompactMany ? 'bold 8px Inter, Calibri, sans-serif' : 'bold 9px Inter, Calibri, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(it.label, cx, plotY + plotH + 15);
    if (it.subLabel && !isCompactMany) {
      ctx.fillStyle = '#64748B';
      ctx.font = '8.5px Inter, Calibri, sans-serif';
      ctx.fillText(it.subLabel, cx, plotY + plotH + 26);
    }
  });

  if (convPts.length > 1) {
    ctx.strokeStyle = '#F59E0B';
    ctx.lineWidth = 2.2;
    ctx.beginPath();
    convPts.forEach((p, i) => {
      if (i === 0) ctx.moveTo(p.x, p.y);
      else ctx.lineTo(p.x, p.y);
    });
    ctx.stroke();

    convPts.forEach((p) => {
      ctx.fillStyle = '#FFFFFF';
      ctx.strokeStyle = '#D97706';
      ctx.lineWidth = 1.8;
      ctx.beginPath();
      ctx.arc(p.x, p.y, isCompactMany ? 2.8 : 3.4, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();

      const lbl = p.item.totalTraffic > 0 ? `${p.item.convRatePct.toFixed(0)}%` : `${p.item.totalTraffic}`;
      ctx.font = isCompactMany ? 'bold 7.5px Inter, Calibri, sans-serif' : 'bold 8.5px Inter, Calibri, sans-serif';
      ctx.fillStyle = '#92400E';
      ctx.textAlign = 'center';
      ctx.fillText(lbl, p.x, Math.max(cardY + 42, p.y - 6));
    });
  }

  return canvas.toDataURL('image/png');
}

function addImageBlock(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  base64Img: string | null,
  imgRowsCount: number,
  widthPx: number,
  heightPx: number
) {
  if (!base64Img) return;
  const startRowIdx = ws.lastRow ? ws.lastRow.number + 1 : 1;
  for (let r = 0; r < imgRowsCount; r++) {
    const row = ws.addRow([]);
    row.height = 25;
  }
  const imageId = workbook.addImage({
    base64: base64Img,
    extension: 'png',
  });
  ws.addImage(imageId, {
    tl: { col: 1.02, row: startRowIdx - 1 + 0.04 },
    ext: { width: widthPx, height: heightPx },
    editAs: 'oneCell',
  });
}

export function embedExcelSalesDiagramPair(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  params: {
    salesTitle: string;
    salesSubtitle: string;
    growthTitle: string;
    growthSubtitle: string;
    bars: ExcelChartBarItem[];
    config: ExcelChartSeriesConfig;
    rightFilterBadgeText?: string;
  }
): void {
  if (params.bars.length === 0) return;

  const sp1 = ws.addRow([]);
  sp1.height = 8;

  const isWide = params.bars.length > 14;
  const salesImg = generateDualSalesDiagramBase64(
    params.salesTitle,
    params.salesSubtitle,
    params.bars,
    params.config,
    params.rightFilterBadgeText
  );
  addImageBlock(
    workbook,
    ws,
    salesImg,
    isWide ? 11 : 10,
    isWide ? 820 : 620,
    isWide ? 265 : 250
  );

  if (salesImg) {
    const spMid = ws.addRow([]);
    spMid.height = 8;
  }

  const growthImg = generateGrowthContributionBase64(
    params.growthTitle,
    params.growthSubtitle,
    params.bars,
    params.rightFilterBadgeText
  );
  addImageBlock(
    workbook,
    ws,
    growthImg,
    isWide ? 18 : 11,
    isWide ? 820 : 620,
    isWide ? 430 : 260
  );

  if (growthImg) {
    const spEnd = ws.addRow([]);
    spEnd.height = 8;
  }
}

export function embedExcelTrafficConversionDiagram(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  params: {
    title: string;
    subtitle: string;
    items: ExcelTrafficChartItem[];
    rightFilterBadgeText?: string;
  }
): void {
  if (params.items.length === 0) return;

  const isWide = params.items.length > 14;
  const img = generateTrafficConversionDiagramBase64(
    params.title,
    params.subtitle,
    params.items,
    params.rightFilterBadgeText
  );
  addImageBlock(
    workbook,
    ws,
    img,
    isWide ? 11 : 10,
    isWide ? 820 : 620,
    isWide ? 260 : 245
  );

  if (img) {
    const spEnd = ws.addRow([]);
    spEnd.height = 8;
  }
}
