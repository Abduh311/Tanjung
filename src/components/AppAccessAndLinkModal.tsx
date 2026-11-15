import React, { useState, useEffect } from 'react';
import {
  Lock,
  Unlock,
  Eye,
  Link2,
  Copy,
  Check,
  ShieldCheck,
  Users,
  KeyRound,
  Share2,
  X,
  ExternalLink,
  AlertCircle,
  LogIn,
  LogOut,
  EyeOff,
  Globe,
  FileSpreadsheet,
  ClipboardList,
  Sparkles,
  ArrowRight,
  Shield,
  Zap,
  Calendar,
  Tag,
  Layers,
} from 'lucide-react';
import {
  getPublicShareableBaseUrl,
  getGooglePublicAppUrl,
  getGooglePublicGateUrl,
  getGooglePublicShiftReportUrl,
  getGooglePublicReviewReportUrl,
  getGooglePublicCvReportUrl,
  getGooglePublicDataCvUrl,
  getSavedCustomPublicOrigin,
  setSavedCustomPublicOrigin,
  setAutoDetectedPublicOrigin,
} from '../utils/shareUrl';
import { ANNUAL_2026_WEEKS } from '../data/initialData';
import { isoToSlashDate, formatIndoDateLong } from '../utils/dateUtils';
import { copyTextToClipboard } from '../utils/clipboard';

export type AppAccessRole = 'public' | 'admin';

interface AppEntryGateProps {
  onEnterPublic: () => void;
  onEnterAdmin: (passwordInput: string) => boolean;
  onOpenReportForm?: (formType: 'shift' | 'review' | 'cv_report' | 'data_cv') => void;
  initialFocusAdmin?: boolean;
}

export const AppEntryAccessGate: React.FC<AppEntryGateProps> = ({
  onEnterPublic,
  onEnterAdmin,
  onOpenReportForm,
  initialFocusAdmin = true,
}) => {
  const [selectedTab, setSelectedTab] = useState<'admin' | 'public'>(
    initialFocusAdmin ? 'admin' : 'public'
  );
  const [passwordInput, setPasswordInput] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isShaking, setIsShaking] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [copiedLink, setCopiedLink] = useState<
    'public' | 'gate' | 'shift' | 'review' | 'cv_report' | 'data_cv' | 'all_wa' | null
  >(null);
  const [gateLinkDateIso, setGateLinkDateIso] = useState<string>('');
  const [, setTunnelTick] = useState(0);

  useEffect(() => {
    let active = true;
    const checkTunnel = () => {
      fetch('/api/public-url', { cache: 'no-store' })
        .then((r) => (r.ok ? r.json() : null))
        .then((d) => {
          if (active && d && d.publicTunnelUrl) {
            setAutoDetectedPublicOrigin(d.publicTunnelUrl);
            setTunnelTick((n) => n + 1);
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

  const gateMatchedSched = gateLinkDateIso
    ? ANNUAL_2026_WEEKS.find(
        (w) =>
          w.slashDates.includes(isoToSlashDate(gateLinkDateIso)) ||
          (gateLinkDateIso >= w.startDateIso && gateLinkDateIso <= w.endDateIso)
      )
    : undefined;
  const gateDateParams = gateLinkDateIso
    ? {
        date: gateLinkDateIso,
        ...(gateMatchedSched ? { week: `weekly-${gateMatchedSched.weekNumber}` } : {}),
      }
    : undefined;

  const publicShareUrl = getGooglePublicAppUrl(undefined, gateDateParams);
  const gateShareUrl = getGooglePublicGateUrl(undefined, gateDateParams);
  const shiftReportPublicUrl = getGooglePublicShiftReportUrl(gateDateParams);
  const reviewReportPublicUrl = getGooglePublicReviewReportUrl(
    gateMatchedSched ? `review-w${gateMatchedSched.weekNumber}` : undefined,
    undefined,
    gateDateParams
  );
  const cvReportPublicUrl = getGooglePublicCvReportUrl(
    gateLinkDateIso ? { date: gateLinkDateIso } : undefined
  );
  const dataCvPublicUrl = getGooglePublicDataCvUrl(
    gateLinkDateIso ? { date: gateLinkDateIso } : undefined
  );

  const handleCopyUrl = (
    url: string,
    type: 'public' | 'gate' | 'shift' | 'review' | 'cv_report' | 'data_cv' | 'all_wa'
  ) => {
    void copyTextToClipboard(url).then((ok) => {
      if (ok) {
        setCopiedLink(type);
        setTimeout(() => {
          setCopiedLink((prev) => (prev === type ? null : prev));
        }, 2200);
      }
    });
  };

  const buildAllGooglePublicLinksText = () => {
    return [
      '*GERBANG AKSES RESMI — REPORT BAZAR MATAHARI*',
      '*Tanjung · PT MDS Retailing Tbk*',
      gateLinkDateIso
        ? `*Filter Tanggal Link:* ${formatIndoDateLong(gateLinkDateIso)}${gateMatchedSched ? ` (Weekly ${gateMatchedSched.weekNumber})` : ''}`
        : null,
      '',
      '1. *Link Aplikasi Utama (Gerbang Login & Mode Publik):*',
      publicShareUrl,
      '',
      '2. *Link Pengisian Report DP (Shift Harian):*',
      shiftReportPublicUrl,
      '',
      '3. *Link Pengisian Data DP (Review Monitoring Bazar):*',
      reviewReportPublicUrl,
      '',
      '4. *Link Pengisian Report CV (Divisi 1 · Graphis & Simplicity):*',
      cvReportPublicUrl,
      '',
      '5. *Link Pengisian Data CV (Inventory, Promo, Kendala & Action Plan):*',
      dataCvPublicUrl,
    ]
      .filter((line) => line !== null)
      .join('\n');
  };

  const handleAdminSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setIsSubmitting(true);
    const ok = onEnterAdmin(passwordInput.trim());
    if (!ok) {
      setErrorMsg('Sandi Admin tidak sesuai. Silakan periksa kembali sandi Anda.');
      setIsShaking(true);
      setTimeout(() => setIsShaking(false), 500);
      setIsSubmitting(false);
    } else {
      setErrorMsg(null);
      setIsSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen w-full bg-[#060A12] text-slate-100 flex flex-col justify-between relative overflow-x-hidden selection:bg-amber-400 selection:text-slate-950 font-sans">
      {/* Dynamic Animated Ambient Background */}
      <div
        aria-hidden="true"
        className="pointer-events-none fixed inset-0 overflow-hidden z-0"
      >
        {/* Subtle grid pattern overlay */}
        <div className="absolute inset-0 bg-[linear-gradient(to_right,#ffffff08_1px,transparent_1px),linear-gradient(to_bottom,#ffffff08_1px,transparent_1px)] bg-[size:36px_36px] [mask-image:radial-gradient(ellipse_75%_65%_at_50%_40%,#000_60%,transparent_100%)] opacity-70" />

        {/* Floating Ambient Glowing Light Beams */}
        <div className="absolute -top-32 left-1/2 -translate-x-1/2 w-[850px] h-[400px] rounded-full bg-gradient-to-tr from-amber-500/15 via-orange-500/10 to-amber-600/15 blur-[120px] animate-float-slow" />
        <div className="absolute -bottom-40 left-10 w-[600px] h-[400px] rounded-full bg-gradient-to-br from-indigo-600/15 via-sky-500/10 to-transparent blur-[110px] animate-float-reverse" />
        <div className="absolute top-1/3 -right-20 w-[500px] h-[350px] rounded-full bg-gradient-to-bl from-emerald-500/10 via-amber-500/5 to-transparent blur-[100px] animate-float-slow" />

        {/* Architectural Pillars (Desktop subtle lines) */}
        <div className="hidden xl:block absolute inset-y-0 left-12 w-px bg-gradient-to-b from-transparent via-amber-400/20 to-transparent" />
        <div className="hidden xl:block absolute inset-y-0 right-12 w-px bg-gradient-to-b from-transparent via-amber-400/20 to-transparent" />
      </div>

      {/* Top Glassmorphic Navigation Bar */}
      <header className="relative z-10 w-full border-b border-slate-800/80 bg-slate-950/75 backdrop-blur-xl">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="relative group">
              <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-amber-400 via-amber-500 to-orange-500 text-slate-950 font-black text-base flex items-center justify-center shadow-lg shadow-amber-500/25 ring-2 ring-amber-400/30 transform group-hover:scale-105 transition-transform duration-300">
                M
              </div>
              <span className="absolute -bottom-0.5 -right-0.5 w-2.5 h-2.5 rounded-full bg-emerald-400 border-2 border-slate-950" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[11px] font-black tracking-widest uppercase text-amber-400">
                  PORTAL RESMI BAZAR
                </span>
                <span className="text-[9px] font-extrabold px-1.5 py-0.5 rounded bg-amber-400/10 border border-amber-400/30 text-amber-300">
                  v5.0
                </span>
              </div>
              <div className="text-[11px] text-slate-400 font-medium truncate max-w-[220px] sm:max-w-none">
                Tanjung &nbsp;·&nbsp; PT MDS Retailing Tbk
              </div>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-slate-900/90 border border-slate-800 shadow-sm backdrop-blur-md">
              <span className="relative flex h-2.5 w-2.5">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75" />
                <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500 shadow-[0_0_8px_#10B981]" />
              </span>
              <span className="text-emerald-400 font-bold text-xs tracking-wide">
                Sinkronisasi Cloud Aktif
              </span>
            </div>
          </div>
        </div>
      </header>

      {/* Main Interactive Login Gate */}
      <main className="relative z-10 flex-1 flex items-center justify-center px-3.5 py-6 sm:px-6 sm:py-10">
        <div className="w-full max-w-4xl">
          {/* Grand Archway Frame with Glassmorphism */}
          <div className="rounded-3xl bg-slate-900/85 border border-amber-500/30 shadow-[0_20px_60px_-15px_rgba(0,0,0,0.8),0_0_40px_-10px_rgba(245,158,11,0.15)] overflow-hidden backdrop-blur-2xl transition-all duration-300">
            {/* Animated Golden Crown Header Beam */}
            <div className="h-1.5 w-full bg-gradient-to-r from-amber-600 via-amber-300 to-amber-600 animate-pulse-glow" />

            {/* Gate Crest & Title Banner */}
            <div className="px-5 pt-7 pb-6 sm:px-8 sm:pt-9 sm:pb-7 border-b border-slate-800/90 bg-gradient-to-b from-slate-900/90 via-slate-900/95 to-slate-950 text-center relative overflow-hidden">
              {/* Subtle radial crest glow */}
              <div className="absolute top-0 left-1/2 -translate-x-1/2 w-80 h-36 bg-amber-500/10 blur-2xl pointer-events-none" />

              {/* Animated Crest Icon */}
              <div className="relative inline-flex items-center justify-center mb-3">
                {/* Rotating dashed gold ring */}
                <div className="absolute -inset-2.5 rounded-full border border-dashed border-amber-400/30 animate-spin-slow pointer-events-none" />
                {/* Outer glow ring */}
                <div className="absolute -inset-1 rounded-2xl bg-amber-500/25 blur-md animate-pulse-glow pointer-events-none" />
                {/* Main badge */}
                <div className="relative w-16 h-16 sm:w-18 sm:h-18 rounded-2xl bg-gradient-to-b from-amber-400 via-amber-500 to-orange-600 p-0.5 shadow-xl shadow-amber-500/30 group">
                  <div className="w-full h-full rounded-[14px] bg-slate-950 flex flex-col items-center justify-center relative overflow-hidden">
                    {/* Shimmer sweep inside badge */}
                    <div className="absolute inset-0 bg-gradient-to-r from-transparent via-amber-400/20 to-transparent -translate-x-full animate-shimmer-sweep pointer-events-none" />
                    <KeyRound className="w-7 h-7 sm:w-8 sm:h-8 text-amber-400 drop-shadow-[0_0_8px_rgba(251,191,36,0.5)] transform group-hover:rotate-12 transition-transform duration-300" />
                    <Sparkles className="w-3.5 h-3.5 text-amber-300 absolute top-2 right-2 animate-bounce" />
                  </div>
                </div>
              </div>

              {/* Badge Tag */}
              <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-500/10 border border-amber-400/30 text-[10px] sm:text-[11px] font-black uppercase tracking-[0.2em] text-amber-300 backdrop-blur-md mb-2 shadow-sm">
                <ShieldCheck className="w-3.5 h-3.5 text-amber-400" />
                <span>GERBANG RESMI · SISTEM KEAMANAN TERPADU</span>
              </div>

              {/* Main Title with metallic gradient text */}
              <h1 className="text-2xl sm:text-3xl md:text-4xl font-black tracking-tight text-white drop-shadow-[0_4px_16px_rgba(245,158,11,0.25)]">
                REPORT BAZAR MATAHARI
              </h1>
              <p className="mt-1.5 text-xs sm:text-sm text-slate-300 font-medium">
                Tanjung &nbsp;·&nbsp; PT MDS Retailing Tbk &nbsp;·&nbsp; Dashboard &amp; Laporan Harian
              </p>

              {/* Dual Gate Mode Switcher Tabs */}
              <div className="mt-6 max-w-md mx-auto grid grid-cols-2 gap-2 p-1.5 bg-slate-950/90 rounded-2xl border border-slate-800 shadow-inner">
                <button
                  type="button"
                  onClick={() => {
                    setSelectedTab('admin');
                    setErrorMsg(null);
                  }}
                  className={`flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl text-xs sm:text-sm font-black transition-all duration-300 cursor-pointer ${
                    selectedTab === 'admin'
                      ? 'bg-gradient-to-r from-amber-400 to-amber-500 text-slate-950 shadow-lg shadow-amber-500/25 scale-[1.02]'
                      : 'text-slate-400 hover:text-white hover:bg-slate-900/60'
                  }`}
                >
                  <KeyRound className="w-4 h-4 shrink-0" />
                  <span>Pintu 01 · Admin</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSelectedTab('public');
                    setErrorMsg(null);
                  }}
                  className={`flex items-center justify-center gap-2 py-2.5 px-4 rounded-xl text-xs sm:text-sm font-black transition-all duration-300 cursor-pointer ${
                    selectedTab === 'public'
                      ? 'bg-gradient-to-r from-sky-400 to-sky-500 text-slate-950 shadow-lg shadow-sky-500/25 scale-[1.02]'
                      : 'text-slate-400 hover:text-white hover:bg-slate-900/60'
                  }`}
                >
                  <Users className="w-4 h-4 shrink-0" />
                  <span>Pintu 02 · Publik</span>
                </button>
              </div>
            </div>

            {/* Dual Gate Body Panels */}
            <div className="grid grid-cols-1 lg:grid-cols-12 divide-y lg:divide-y-0 lg:divide-x divide-slate-800/80">
              {/* PINTU 1: GERBANG LOGIN ADMIN (LEFT / PRIMARY GATE) */}
              <div
                className={`lg:col-span-7 p-6 sm:p-8 bg-gradient-to-br from-slate-900/70 to-slate-950/90 flex flex-col justify-between transition-all duration-300 ${
                  selectedTab === 'public' ? 'hidden lg:flex' : 'flex'
                }`}
              >
                <div className="space-y-5">
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-xl bg-amber-400/15 border border-amber-400/30 flex items-center justify-center text-amber-400 shrink-0 shadow-sm">
                        <ShieldCheck className="w-5 h-5" />
                      </div>
                      <div>
                        <div className="text-[10px] font-black uppercase tracking-wider text-amber-400 flex items-center gap-1.5">
                          <span>PINTU GERBANG 01</span>
                          <span className="w-1 h-1 rounded-full bg-amber-400" />
                          <span>OTORISASI PENUH</span>
                        </div>
                        <h2 className="text-base sm:text-lg font-black text-white">
                          Login Administrator Bazar
                        </h2>
                      </div>
                    </div>
                    <span className="hidden sm:inline-flex px-2.5 py-1 rounded-md bg-amber-400/10 border border-amber-400/25 text-amber-300 text-[10px] font-black uppercase tracking-wide">
                      Full Control
                    </span>
                  </div>

                  <p className="text-xs text-slate-300 leading-relaxed">
                    Masukkan <strong>Sandi Admin</strong> untuk membuka kunci edit tabel
                    sales mingguan/bulanan/tahunan, pencatatan barang hilang, serta fitur simpan &amp;
                    export laporan.
                  </p>

                  <form onSubmit={handleAdminSubmit} className="space-y-4 pt-1">
                    <div>
                      <div className="flex items-center justify-between mb-1.5">
                        <label className="block text-xs font-black text-slate-200">
                          Sandi Gerbang Admin
                        </label>
                        <span className="text-[11px] text-amber-400/80 font-medium">
                          Diperlukan untuk akses edit
                        </span>
                      </div>
                      <div className="relative group">
                        <div className="w-5 h-5 text-amber-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none group-focus-within:text-amber-300 transition-colors">
                          <Lock className="w-4 h-4" />
                        </div>
                        <input
                          type={showPassword ? 'text' : 'password'}
                          value={passwordInput}
                          onChange={(e) => {
                            setPasswordInput(e.target.value);
                            if (errorMsg) setErrorMsg(null);
                          }}
                          placeholder="Masukkan Sandi Admin..."
                          autoFocus
                          className={`w-full pl-11 pr-12 py-3.5 rounded-xl bg-slate-950 border text-sm font-bold text-white placeholder:text-slate-500 focus:outline-none transition-all duration-200 shadow-inner ${
                            isShaking
                              ? 'animate-error-shake border-rose-500 ring-2 ring-rose-500/30'
                              : 'border-slate-700/80 focus:border-amber-400 focus:ring-4 focus:ring-amber-500/20'
                          }`}
                        />
                        <button
                          type="button"
                          onClick={() => setShowPassword((prev) => !prev)}
                          className="absolute right-3 top-1/2 -translate-y-1/2 p-1.5 rounded-lg text-slate-400 hover:text-amber-300 hover:bg-slate-900 transition-colors cursor-pointer"
                          title={showPassword ? 'Sembunyikan Sandi' : 'Lihat Sandi'}
                        >
                          {showPassword ? (
                            <EyeOff className="w-4 h-4" />
                          ) : (
                            <Eye className="w-4 h-4" />
                          )}
                        </button>
                      </div>

                      {/* Error Alert Box with Shake & Glow */}
                      {errorMsg && (
                        <div className="mt-2.5 p-3 rounded-xl bg-rose-950/80 border border-rose-500/40 flex items-center gap-2.5 text-xs font-bold text-rose-200 shadow-lg shadow-rose-950/40">
                          <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
                          <span>{errorMsg}</span>
                        </div>
                      )}
                    </div>

                    {/* Animated Unlock Button */}
                    <button
                      type="submit"
                      disabled={isSubmitting}
                      className="relative overflow-hidden w-full py-3.5 px-4 rounded-xl bg-gradient-to-r from-amber-400 via-amber-500 to-orange-500 hover:from-amber-300 hover:to-orange-400 text-slate-950 font-black text-xs sm:text-sm flex items-center justify-center gap-2.5 shadow-xl shadow-amber-500/20 hover:shadow-amber-500/35 hover:scale-[1.01] active:scale-[0.99] transition-all duration-200 cursor-pointer group"
                    >
                      {/* Shimmer sweep effect */}
                      <div className="absolute inset-0 bg-gradient-to-r from-transparent via-white/30 to-transparent -translate-x-full group-hover:animate-shimmer-sweep pointer-events-none" />
                      <Unlock className="w-4 h-4 stroke-[2.5] text-slate-950 group-hover:rotate-12 transition-transform duration-200" />
                      <span>MASUK GERBANG ADMIN (BUKA AKSES EDIT)</span>
                      <ArrowRight className="w-4 h-4 text-slate-950 group-hover:translate-x-1 transition-transform duration-200" />
                    </button>
                  </form>

                  {/* Quick Alternative for Mobile Visitors */}
                  <div className="pt-4 border-t border-slate-800/80 lg:hidden space-y-2.5">
                    <div className="text-[11px] font-bold text-slate-400 text-center flex items-center justify-center gap-2">
                      <span className="h-px w-6 bg-slate-800" />
                      <span>Atau masuk tanpa sandi:</span>
                      <span className="h-px w-6 bg-slate-800" />
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                      <button
                        type="button"
                        onClick={onEnterPublic}
                        className="w-full py-2.5 px-3 rounded-xl bg-slate-800/90 hover:bg-slate-700 text-sky-300 border border-slate-700/80 font-extrabold text-xs flex items-center justify-center gap-1.5 transition-all cursor-pointer"
                      >
                        <Eye className="w-3.5 h-3.5 text-sky-400 shrink-0" />
                        <span>Mode Publik (Hanya Lihat)</span>
                      </button>
                      {onOpenReportForm && (
                        <button
                          type="button"
                          onClick={() => onOpenReportForm('shift')}
                          className="w-full py-2.5 px-3 rounded-xl bg-indigo-950/80 hover:bg-indigo-900 text-indigo-200 border border-indigo-500/30 font-extrabold text-xs flex items-center justify-center gap-1.5 transition-all cursor-pointer"
                        >
                          <FileSpreadsheet className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                          <span>Isi Form Report DP</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                <div className="hidden lg:flex items-center justify-between pt-4 mt-6 border-t border-slate-800/80 text-[11px] text-slate-400">
                  <span className="flex items-center gap-1.5">
                    <Shield className="w-3.5 h-3.5 text-amber-400" />
                    <span>Hak Akses: Edit Sales, Barang Hilang, Export &amp; Reset</span>
                  </span>
                  <span className="text-amber-400 font-black">Full Otorisasi</span>
                </div>
              </div>

              {/* PINTU 2: GERBANG PUBLIK & FORM REPORT (RIGHT / SECONDARY GATE) */}
              <div
                className={`lg:col-span-5 p-6 sm:p-8 bg-slate-950/80 flex flex-col justify-between transition-all duration-300 ${
                  selectedTab === 'admin' ? 'hidden lg:flex' : 'flex'
                }`}
              >
                <div className="space-y-4">
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-xl bg-sky-500/15 border border-sky-400/30 flex items-center justify-center text-sky-400 shrink-0 shadow-sm">
                        <Users className="w-5 h-5" />
                      </div>
                      <div>
                        <div className="text-[10px] font-black uppercase tracking-wider text-sky-400 flex items-center gap-1.5">
                          <span>PINTU GERBANG 02</span>
                          <span className="w-1 h-1 rounded-full bg-sky-400" />
                          <span>TANPA SANDI</span>
                        </div>
                        <h2 className="text-base sm:text-lg font-black text-white">
                          Akses Publik &amp; Form Laporan
                        </h2>
                      </div>
                    </div>
                    <span className="hidden sm:inline-flex px-2.5 py-1 rounded-md bg-sky-400/10 border border-sky-400/25 text-sky-300 text-[10px] font-black uppercase tracking-wide">
                      Tanpa Sandi
                    </span>
                  </div>

                  <p className="text-xs text-slate-300 leading-relaxed">
                    Bagi pengunjung umum atau petugas shift bazar harian, pilih gerbang akses langsung
                    di bawah ini tanpa memerlukan otorisasi sandi:
                  </p>

                  <div className="space-y-2.5 pt-1">
                    {/* Public View Card */}
                    <button
                      type="button"
                      onClick={onEnterPublic}
                      className="group w-full p-3.5 rounded-xl bg-gradient-to-r from-sky-500 to-sky-600 hover:from-sky-400 hover:to-sky-500 text-slate-950 font-black text-xs sm:text-sm flex items-center justify-between shadow-lg shadow-sky-500/15 hover:shadow-sky-500/30 hover:scale-[1.01] active:scale-[0.99] transition-all duration-200 cursor-pointer"
                    >
                      <div className="flex items-center gap-2.5">
                        <Eye className="w-4 h-4 stroke-[2.5] shrink-0" />
                        <span>MASUK GERBANG PUBLIK (HANYA LIHAT)</span>
                      </div>
                      <ArrowRight className="w-4 h-4 group-hover:translate-x-1 transition-transform duration-200 shrink-0" />
                    </button>

                    {onOpenReportForm && (
                      <>
                        {/* Daily Shift Report Card */}
                        <button
                          type="button"
                          onClick={() => onOpenReportForm('shift')}
                          className="group w-full p-3.5 rounded-xl bg-gradient-to-r from-indigo-900/90 to-indigo-800/90 hover:from-indigo-800 hover:to-indigo-700 text-white font-extrabold text-xs flex items-center justify-between border border-indigo-400/30 shadow-md hover:shadow-indigo-500/20 hover:scale-[1.01] active:scale-[0.99] transition-all duration-200 cursor-pointer"
                        >
                          <div className="flex items-center gap-2.5">
                            <FileSpreadsheet className="w-4 h-4 text-amber-300 shrink-0" />
                            <div className="text-left">
                              <div>Form Isi Report DP</div>
                              <div className="text-[10px] text-indigo-300 font-normal">
                                Sales, Qty, Trx &amp; Reason Kendala
                              </div>
                            </div>
                          </div>
                          <span className="text-[10px] px-2 py-0.5 rounded bg-indigo-500/30 border border-indigo-400/40 text-indigo-200 shrink-0 group-hover:translate-x-1 transition-transform">
                            Buka Form →
                          </span>
                        </button>

                        {/* Data DP (Review Monitoring) Form Card */}
                        <button
                          type="button"
                          onClick={() => onOpenReportForm('review')}
                          className="group w-full p-3 rounded-xl bg-gradient-to-r from-emerald-950/90 to-emerald-900/90 hover:from-emerald-900 hover:to-emerald-800 text-white font-extrabold text-xs flex items-center justify-between border border-emerald-400/30 shadow-md hover:shadow-emerald-500/20 hover:scale-[1.01] active:scale-[0.99] transition-all duration-200 cursor-pointer"
                        >
                          <div className="flex items-center gap-2.5">
                            <ClipboardList className="w-4 h-4 text-emerald-300 shrink-0" />
                            <div className="text-left">
                              <div>Form Isi Data DP</div>
                              <div className="text-[10px] text-emerald-300 font-normal">
                                Review Monitoring, Fast/Slow Moving &amp; Evaluasi
                              </div>
                            </div>
                          </div>
                          <span className="text-[10px] px-2 py-0.5 rounded bg-emerald-500/30 border border-emerald-400/40 text-emerald-200 shrink-0 group-hover:translate-x-1 transition-transform">
                            Buka Form →
                          </span>
                        </button>

                        {/* Report CV & Data CV Quick Buttons */}
                        <div className="grid grid-cols-2 gap-2">
                          <button
                            type="button"
                            onClick={() => onOpenReportForm('cv_report')}
                            className="p-2.5 rounded-xl bg-purple-950/90 hover:bg-purple-900 text-white font-bold text-xs flex items-center gap-2 border border-purple-400/30 transition-colors cursor-pointer"
                          >
                            <Tag className="w-3.5 h-3.5 text-purple-300 shrink-0" />
                            <span className="truncate">Form Report CV</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => onOpenReportForm('data_cv')}
                            className="p-2.5 rounded-xl bg-sky-950/90 hover:bg-sky-900 text-white font-bold text-xs flex items-center gap-2 border border-sky-400/30 transition-colors cursor-pointer"
                          >
                            <Layers className="w-3.5 h-3.5 text-sky-300 shrink-0" />
                            <span className="truncate">Form Data CV</span>
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                </div>

                <div className="pt-4 mt-6 border-t border-slate-800/80 text-[11px] text-slate-400 flex items-center justify-between">
                  <span>Mode Baca &amp; Kirim Laporan Shift</span>
                  <span className="text-sky-400 font-black">Tanpa Sandi</span>
                </div>
              </div>
            </div>

            {/* Footer Gate Panel: Official Connected Link Hub */}
            <div className="px-5 py-5 sm:px-8 sm:py-6 bg-slate-950/95 border-t border-slate-800/90 space-y-3.5">
              <div className="flex flex-wrap items-center justify-between gap-2.5">
                <div className="flex items-center gap-2 text-xs font-black text-slate-200">
                  <Globe className="w-4 h-4 text-emerald-400 shrink-0 animate-pulse" />
                  <span>Tautan Resmi Gerbang Aplikasi &amp; Pengisian Report:</span>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl bg-slate-900 border border-amber-400/35 text-[11px]">
                    <Calendar className="w-3.5 h-3.5 text-amber-400 shrink-0" />
                    <span className="text-amber-300 font-bold hidden sm:inline">Filter Tanggal:</span>
                    <input
                      type="date"
                      aria-label="Filter Kalender Tanggal Tautan Gerbang"
                      value={gateLinkDateIso}
                      onChange={(e) => setGateLinkDateIso(e.target.value)}
                      className="bg-transparent text-white font-bold font-mono-num focus:outline-none cursor-pointer"
                    />
                    {gateLinkDateIso && (
                      <button
                        type="button"
                        onClick={() => setGateLinkDateIso('')}
                        className="px-1.5 py-0.5 rounded bg-amber-400/20 hover:bg-amber-400/30 text-amber-200 text-[10px] font-bold cursor-pointer"
                      >
                        Semua
                      </button>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={() => handleCopyUrl(buildAllGooglePublicLinksText(), 'all_wa')}
                    className="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 active:scale-95 text-white text-[11px] font-black inline-flex items-center gap-1.5 shadow-md shadow-emerald-600/20 hover:shadow-emerald-600/40 transition-all duration-200 cursor-pointer"
                  >
                    {copiedLink === 'all_wa' ? (
                      <>
                        <Check className="w-3.5 h-3.5 text-white" />
                        <span>Semua Link Berhasil Disalin!</span>
                      </>
                    ) : (
                      <>
                        <Share2 className="w-3.5 h-3.5" />
                        <span>Salin Paket Link (WhatsApp)</span>
                      </>
                    )}
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {/* 1. Link Aplikasi Utama */}
                <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800/90 space-y-1.5">
                  <div className="flex items-center justify-between text-[11px] font-bold text-slate-300">
                    <span className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-amber-400" />
                      <span>1. Link Aplikasi Utama:</span>
                    </span>
                    <span className="text-[10px] text-amber-400 font-bold px-1.5 py-0.5 rounded bg-amber-400/10">
                      Publik &amp; Admin
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      readOnly
                      value={publicShareUrl}
                      onClick={(e) => (e.target as HTMLInputElement).select()}
                      className="flex-1 min-w-0 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-800 text-[11px] font-mono text-slate-200 focus:outline-none focus:border-amber-400"
                    />
                    <button
                      type="button"
                      onClick={() => handleCopyUrl(publicShareUrl, 'public')}
                      className={`px-3 py-2 rounded-lg text-[11px] font-black inline-flex items-center gap-1 shrink-0 transition-all duration-200 cursor-pointer ${
                        copiedLink === 'public'
                          ? 'bg-emerald-600 text-white shadow-sm'
                          : 'bg-slate-800 hover:bg-slate-700 text-amber-300 border border-slate-700/80'
                      }`}
                    >
                      {copiedLink === 'public' ? (
                        <>
                          <Check className="w-3.5 h-3.5" />
                          <span>Tersalin</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Salin</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* 2. Link Pengisian Report DP */}
                <div className="p-3 rounded-xl bg-slate-900/90 border border-slate-800/90 space-y-1.5">
                  <div className="flex items-center justify-between text-[11px] font-bold text-slate-300">
                    <span className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-indigo-400" />
                      <span>2. Link Form Report DP:</span>
                    </span>
                    <button
                      type="button"
                      onClick={() => handleCopyUrl(reviewReportPublicUrl, 'review')}
                      className="text-[10px] font-bold text-emerald-400 hover:text-emerald-300 hover:underline cursor-pointer"
                    >
                      {copiedLink === 'review' ? '✓ Link Data DP Tersalin' : '+ Salin Form Data DP'}
                    </button>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      readOnly
                      value={shiftReportPublicUrl}
                      onClick={(e) => (e.target as HTMLInputElement).select()}
                      className="flex-1 min-w-0 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-800 text-[11px] font-mono text-indigo-300 focus:outline-none focus:border-indigo-400"
                    />
                    <button
                      type="button"
                      onClick={() => handleCopyUrl(shiftReportPublicUrl, 'shift')}
                      className={`px-3 py-2 rounded-lg text-[11px] font-black inline-flex items-center gap-1 shrink-0 transition-all duration-200 cursor-pointer ${
                        copiedLink === 'shift'
                          ? 'bg-emerald-600 text-white shadow-sm'
                          : 'bg-indigo-600 hover:bg-indigo-500 text-white'
                      }`}
                    >
                      {copiedLink === 'shift' ? (
                        <>
                          <Check className="w-3.5 h-3.5" />
                          <span>Tersalin</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Salin</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* 3. Link Pengisian Report CV */}
                <div className="p-3 rounded-xl bg-slate-900/90 border border-purple-500/25 space-y-1.5">
                  <div className="flex items-center justify-between text-[11px] font-bold text-slate-300">
                    <span className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-purple-400" />
                      <span>3. Link Form Report CV (6201 &amp; 6244):</span>
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      readOnly
                      value={cvReportPublicUrl}
                      onClick={(e) => (e.target as HTMLInputElement).select()}
                      className="flex-1 min-w-0 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-800 text-[11px] font-mono text-purple-300 focus:outline-none focus:border-purple-400"
                    />
                    <button
                      type="button"
                      onClick={() => handleCopyUrl(cvReportPublicUrl, 'cv_report')}
                      className={`px-3 py-2 rounded-lg text-[11px] font-black inline-flex items-center gap-1 shrink-0 transition-all duration-200 cursor-pointer ${
                        copiedLink === 'cv_report'
                          ? 'bg-emerald-600 text-white shadow-sm'
                          : 'bg-purple-600 hover:bg-purple-500 text-white'
                      }`}
                    >
                      {copiedLink === 'cv_report' ? (
                        <>
                          <Check className="w-3.5 h-3.5" />
                          <span>Tersalin</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Salin</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>

                {/* 4. Link Pengisian Data CV */}
                <div className="p-3 rounded-xl bg-slate-900/90 border border-sky-500/25 space-y-1.5">
                  <div className="flex items-center justify-between text-[11px] font-bold text-slate-300">
                    <span className="flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-sky-400" />
                      <span>4. Link Form Data CV (Stock, Promo &amp; Plan):</span>
                    </span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      readOnly
                      value={dataCvPublicUrl}
                      onClick={(e) => (e.target as HTMLInputElement).select()}
                      className="flex-1 min-w-0 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-800 text-[11px] font-mono text-sky-300 focus:outline-none focus:border-sky-400"
                    />
                    <button
                      type="button"
                      onClick={() => handleCopyUrl(dataCvPublicUrl, 'data_cv')}
                      className={`px-3 py-2 rounded-lg text-[11px] font-black inline-flex items-center gap-1 shrink-0 transition-all duration-200 cursor-pointer ${
                        copiedLink === 'data_cv'
                          ? 'bg-emerald-600 text-white shadow-sm'
                          : 'bg-sky-600 hover:bg-sky-500 text-white'
                      }`}
                    >
                      {copiedLink === 'data_cv' ? (
                        <>
                          <Check className="w-3.5 h-3.5" />
                          <span>Tersalin</span>
                        </>
                      ) : (
                        <>
                          <Copy className="w-3.5 h-3.5" />
                          <span>Salin</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Modern Architectural Footer */}
      <footer className="relative z-10 py-3.5 px-4 text-center text-[11px] text-slate-500 border-t border-slate-900 bg-slate-950/90 backdrop-blur-md">
        Gerbang Keamanan Data Terpadu &nbsp;·&nbsp; Report Bazar Matahari Tanjung &nbsp;·&nbsp; PT MDS Retailing Tbk
      </footer>
    </div>
  );
};

interface AppShareLinkModalProps {
  isOpen: boolean;
  onClose: () => void;
  accessRole: AppAccessRole;
  adminPassword: string;
  onLoginAdmin: (passwordInput: string) => boolean;
  onLogoutToPublic: () => void;
  onLogoutToEntryGate?: () => void;
  onChangeAdminPassword: (newPassword: string) => void;
  onOpenReportForm?: (formType: 'shift' | 'review' | 'cv_report' | 'data_cv') => void;
}

export const AppShareLinkModal: React.FC<AppShareLinkModalProps> = ({
  isOpen,
  onClose,
  accessRole,
  onLoginAdmin,
  onLogoutToPublic,
  onLogoutToEntryGate,
  onChangeAdminPassword,
  onOpenReportForm,
}) => {
  const [copiedType, setCopiedType] = useState<string | null>(null);
  const [passwordInput, setPasswordInput] = useState('');
  const [loginError, setLoginError] = useState<string | null>(null);
  const [loginSuccess, setLoginSuccess] = useState<string | null>(null);

  const [newPassInput, setNewPassInput] = useState('');
  const [passSavedNotice, setPassSavedNotice] = useState(false);
  const [shareCalendarIso, setShareCalendarIso] = useState<string>('');
  const [customDomainInput] = useState<string>(() =>
    getSavedCustomPublicOrigin()
  );
  const [, setModalTunnelTick] = useState(0);

  useEffect(() => {
    if (!isOpen) return;
    let active = true;
    const checkTunnel = () => {
      fetch('/api/public-url', { cache: 'no-store' })
        .then((r) => (r.ok ? r.json() : null))
        .then((d) => {
          if (active && d && d.publicTunnelUrl) {
            setAutoDetectedPublicOrigin(d.publicTunnelUrl);
            setModalTunnelTick((n) => n + 1);
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
  }, [isOpen]);

  if (!isOpen) return null;

  const baseUrl = getPublicShareableBaseUrl(customDomainInput);
  const matchedShareSched = shareCalendarIso
    ? ANNUAL_2026_WEEKS.find(
        (w) =>
          w.slashDates.includes(isoToSlashDate(shareCalendarIso)) ||
          (shareCalendarIso >= w.startDateIso && shareCalendarIso <= w.endDateIso)
      )
    : undefined;
  const shareDateParams = shareCalendarIso
    ? {
        date: shareCalendarIso,
        ...(matchedShareSched ? { week: `weekly-${matchedShareSched.weekNumber}` } : {}),
      }
    : undefined;

  const publicViewUrl = getGooglePublicAppUrl(customDomainInput, shareDateParams);
  const gateEntryUrl = getGooglePublicGateUrl(customDomainInput, shareDateParams);
  const adminLoginUrl = shareCalendarIso
    ? `${baseUrl}?mode=admin&date=${encodeURIComponent(shareCalendarIso)}${
        matchedShareSched ? `&week=weekly-${matchedShareSched.weekNumber}` : ''
      }`
    : `${baseUrl}?mode=admin`;
  const shiftReportUrl = getGooglePublicShiftReportUrl(shareDateParams, customDomainInput);
  const reviewReportUrl = getGooglePublicReviewReportUrl(
    matchedShareSched ? `review-w${matchedShareSched.weekNumber}` : undefined,
    customDomainInput,
    shareDateParams
  );
  const cvReportUrl = getGooglePublicCvReportUrl(
    shareCalendarIso ? { date: shareCalendarIso } : undefined,
    customDomainInput
  );
  const dataCvUrl = getGooglePublicDataCvUrl(
    shareCalendarIso ? { date: shareCalendarIso } : undefined,
    customDomainInput
  );

  const handleCopy = (text: string, key: string) => {
    void copyTextToClipboard(text).then((ok) => {
      if (ok) {
        setCopiedType(key);
        setTimeout(() => {
          setCopiedType((prev) => (prev === key ? null : prev));
        }, 2200);
      }
    });
  };

  const handleShareWhatsAppAllLinks = () => {
    const text = [
      '*LINK GOOGLE PUBLIK — REPORT BAZAR MATAHARI*',
      '*Tanjung* — PT MDS Retailing Tbk',
      shareCalendarIso
        ? `*Filter Tanggal:* ${formatIndoDateLong(shareCalendarIso)}${matchedShareSched ? ` (Weekly ${matchedShareSched.weekNumber})` : ''}`
        : null,
      '',
      '1. *Link Buka Aplikasi (Google Publik / Hanya Lihat):*',
      publicViewUrl,
      '',
      '2. *Link Pengisian Report DP (Shift Harian):*',
      shiftReportUrl,
      '',
      '3. *Link Pengisian Data DP (Review Monitoring Bazar):*',
      reviewReportUrl,
      '',
      '4. *Link Pengisian Report CV (Divisi 1 · Graphis & Simplicity):*',
      cvReportUrl,
      '',
      '5. *Link Pengisian Data CV (Inventory, Promo, Kendala & Action Plan):*',
      dataCvUrl,
    ]
      .filter((line) => line !== null)
      .join('\n');
    void copyTextToClipboard(text).then((ok) => {
      if (ok) {
        setCopiedType('wa_all');
        setTimeout(() => setCopiedType(null), 2500);
      }
    });
  };

  const handleAdminUnlock = (e: React.FormEvent) => {
    e.preventDefault();
    const ok = onLoginAdmin(passwordInput.trim());
    if (!ok) {
      setLoginError('Sandi Admin salah. Silakan coba lagi.');
      setLoginSuccess(null);
    } else {
      setLoginError(null);
      setLoginSuccess('Berhasil masuk sebagai Admin! Sekarang Anda dapat mengedit seluruh data.');
      setPasswordInput('');
    }
  };

  const handleSaveNewPassword = (e: React.FormEvent) => {
    e.preventDefault();
    const cleaned = newPassInput.trim();
    if (cleaned.length < 3) return;
    onChangeAdminPassword(cleaned);
    setPassSavedNotice(true);
    setTimeout(() => setPassSavedNotice(false), 2500);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/60 p-4 overflow-y-auto">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-2xl max-w-2xl w-full overflow-hidden my-auto">
        {/* Header */}
        <div className="bg-slate-900 text-white px-5 py-4 flex items-center justify-between gap-3 border-b border-slate-800">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-xl bg-amber-400 text-slate-950 flex items-center justify-center shrink-0 font-black">
              <Globe className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h3 className="text-sm sm:text-base font-extrabold truncate">
                Link Google Publik (Buka Aplikasi &amp; Link Pengisian Report)
              </h3>
              <p className="text-[11px] text-amber-300 truncate">
                REPORT BAZAR MATAHARI · Tanjung · PT MDS Retailing Tbk
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 sm:p-6 space-y-4 max-h-[82vh] overflow-y-auto">
          {/* Current Access Status Banner */}
          <div
            className={`p-3.5 rounded-xl border flex flex-wrap items-center justify-between gap-3 ${
              accessRole === 'admin'
                ? 'bg-emerald-50 border-emerald-200'
                : 'bg-amber-50 border-amber-200'
            }`}
          >
            <div className="flex items-center gap-3">
              <div
                className={`w-9 h-9 rounded-xl flex items-center justify-center shrink-0 ${
                  accessRole === 'admin'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-amber-500 text-slate-950'
                }`}
              >
                {accessRole === 'admin' ? (
                  <ShieldCheck className="w-5 h-5" />
                ) : (
                  <Eye className="w-5 h-5" />
                )}
              </div>
              <div>
                <div className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Status Akses Saat Ini:{' '}
                  {accessRole === 'admin' ? (
                    <span className="text-emerald-700">MODE ADMIN (Bisa Edit Penuh)</span>
                  ) : (
                    <span className="text-amber-800">MODE GOOGLE PUBLIK (Hanya Lihat)</span>
                  )}
                </div>
                <div className="text-[11px] text-slate-600 mt-0.5">
                  {accessRole === 'admin'
                    ? 'Akses penuh untuk mengedit tabel sales, menambah minggu, dan mengelola data. Fitur Screenshot & Export dapat diakses oleh semua pengguna.'
                    : 'Mode Publik (Read-Only). Pengunjung tetap dapat membuka Screenshot, Export Excel, dan mengisi report melalui Link Form di bawah.'}
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 shrink-0">
              {accessRole === 'admin' && (
                <button
                  type="button"
                  onClick={onLogoutToPublic}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white hover:bg-amber-50 text-amber-900 border border-amber-300 text-xs font-extrabold transition-colors shrink-0 cursor-pointer"
                >
                  <Lock className="w-3.5 h-3.5" />
                  <span>Kunci ke Mode Publik</span>
                </button>
              )}
              {onLogoutToEntryGate && (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onLogoutToEntryGate();
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-extrabold transition-colors shrink-0 cursor-pointer"
                >
                  <LogOut className="w-3.5 h-3.5" />
                  <span>Logout (Keluar)</span>
                </button>
              )}
            </div>
          </div>

          {/* FILTER KALENDER TANGGAL UNTUK SEMUA LINK */}
          <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-300 space-y-2">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2 text-xs font-extrabold text-amber-950">
                <Calendar className="w-4 h-4 text-amber-700 shrink-0" />
                <span>Filter Kalender Tanggal Link (Opsional)</span>
              </div>
              {shareCalendarIso && (
                <button
                  type="button"
                  onClick={() => setShareCalendarIso('')}
                  className="px-2 py-0.5 rounded bg-amber-200 hover:bg-amber-300 text-amber-950 text-[11px] font-extrabold cursor-pointer"
                >
                  Reset (Semua Tanggal)
                </button>
              )}
            </div>
            <div className="flex flex-col sm:flex-row sm:items-center gap-2">
              <input
                type="date"
                aria-label="Filter Kalender Tanggal untuk Link Publik"
                value={shareCalendarIso}
                onChange={(e) => setShareCalendarIso(e.target.value)}
                className="px-3 py-1.5 rounded-lg bg-white border border-amber-300 text-xs font-bold font-mono-num text-slate-900 focus:outline-none focus:border-amber-600 cursor-pointer"
              />
              <div className="text-[11px] font-semibold text-amber-900">
                {shareCalendarIso ? (
                  <span>
                    Target aktif:{' '}
                    <strong>{formatIndoDateLong(shareCalendarIso)}</strong>
                    {matchedShareSched ? ` → Weekly ${matchedShareSched.weekNumber}` : ''}
                  </span>
                ) : (
                  <span>
                    Pilih tanggal dari kalender agar setiap link langsung menuju tanggal &amp; minggu terkait.
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* TOMBOL SALIN PAKET LENGKAP LINK GOOGLE PUBLIK */}
          <div className="p-4 rounded-xl bg-gradient-to-r from-emerald-900 via-teal-900 to-slate-900 text-white flex flex-wrap items-center justify-between gap-3">
            <div className="space-y-0.5">
              <div className="text-xs sm:text-sm font-extrabold flex items-center gap-1.5 text-amber-300">
                <Share2 className="w-4 h-4" />
                <span>Paket Lengkap Link Google Publik (Buka Aplikasi + Isi Report)</span>
              </div>
              <p className="text-[11px] text-slate-200">
                Salin sekaligus Link Buka Aplikasi Publik, Link Report DP, Data DP, Report CV, dan Data CV untuk dikirim ke grup WhatsApp.
              </p>
            </div>
            <button
              type="button"
              onClick={handleShareWhatsAppAllLinks}
              className="px-3.5 py-2 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 text-xs font-black inline-flex items-center gap-1.5 shrink-0 cursor-pointer shadow-sm"
            >
              {copiedType === 'wa_all' ? (
                <>
                  <Check className="w-4 h-4" />
                  <span>Semua Link Tersalin!</span>
                </>
              ) : (
                <>
                  <Copy className="w-4 h-4" />
                  <span>Salin Semua Link Publik (WA)</span>
                </>
              )}
            </button>
          </div>

          {/* 1. LINK BUKA APLIKASI UNTUK GOOGLE PUBLIK */}
          <div className="p-4 rounded-xl bg-sky-50/70 border border-sky-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-sky-600 text-white text-[10px] font-extrabold uppercase">
                  1. Buka Aplikasi (Publik)
                </span>
                <h4 className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Link Buka Aplikasi untuk Google Publik (Hanya Lihat / Tanpa Login)
                </h4>
              </div>
            </div>
            <p className="text-xs text-slate-600">
              Siapa pun yang membuka link ini di Google / Browser HP langsung masuk ke aplikasi dalam{' '}
              <strong>Mode Publik (Hanya Lihat)</strong> tanpa diminta sandi admin.
            </p>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                readOnly
                value={publicViewUrl}
                onClick={(e) => (e.target as HTMLInputElement).select()}
                className="flex-1 px-3 py-2 rounded-lg bg-white border border-sky-300 text-xs font-mono text-slate-800 select-all"
              />
              <button
                type="button"
                onClick={() => handleCopy(publicViewUrl, 'public')}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-xs font-extrabold shrink-0 cursor-pointer"
              >
                {copiedType === 'public' ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Salin Link Aplikasi</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* 2. LINK PENGISIAN REPORT DP HARIAN UNTUK GOOGLE PUBLIK */}
          <div className="p-4 rounded-xl bg-indigo-50/70 border border-indigo-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-indigo-600 text-white text-[10px] font-extrabold uppercase">
                  2. Link Pengisian Report DP
                </span>
                <h4 className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Link Pengisian Report DP (Shift Harian · Google Publik)
                </h4>
              </div>
              {onOpenReportForm && (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onOpenReportForm('shift');
                  }}
                  className="text-xs font-extrabold text-indigo-700 hover:text-indigo-900 inline-flex items-center gap-1 underline cursor-pointer"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  <span>Buka Form Report DP</span>
                </button>
              )}
            </div>
            <p className="text-xs text-slate-600">
              Link khusus bagi petugas untuk mengisi <strong>Sales, Qty, Trx, dan Reason/Event</strong>{' '}
              Shift Pagi &amp; Siang. Saat diisi, otomatis tersimpan langsung ke aplikasi Admin &amp; Pelihat.
            </p>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                readOnly
                value={shiftReportUrl}
                onClick={(e) => (e.target as HTMLInputElement).select()}
                className="flex-1 px-3 py-2 rounded-lg bg-white border border-indigo-300 text-xs font-mono text-slate-800 select-all"
              />
              <button
                type="button"
                onClick={() => handleCopy(shiftReportUrl, 'shift_report')}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-extrabold shrink-0 cursor-pointer"
              >
                {copiedType === 'shift_report' ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Salin Link Report DP</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* 3. LINK PENGISIAN DATA DP (REVIEW MONITORING) UNTUK GOOGLE PUBLIK */}
          <div className="p-4 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-extrabold uppercase">
                  3. Link Pengisian Data DP
                </span>
                <h4 className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Link Pengisian Data DP (Review Monitoring Bazar · Google Publik)
                </h4>
              </div>
              {onOpenReportForm && (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onOpenReportForm('review');
                  }}
                  className="text-xs font-extrabold text-emerald-700 hover:text-emerald-900 inline-flex items-center gap-1 underline cursor-pointer"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  <span>Buka Form Data DP</span>
                </button>
              )}
            </div>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                readOnly
                value={reviewReportUrl}
                onClick={(e) => (e.target as HTMLInputElement).select()}
                className="flex-1 px-3 py-2 rounded-lg bg-white border border-emerald-300 text-xs font-mono text-slate-800 select-all"
              />
              <button
                type="button"
                onClick={() => handleCopy(reviewReportUrl, 'review_report')}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold shrink-0 cursor-pointer"
              >
                {copiedType === 'review_report' ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Salin Link Data DP</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* 4. LINK PENGISIAN REPORT CV */}
          <div className="p-4 rounded-xl bg-purple-50/70 border border-purple-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-purple-700 text-white text-[10px] font-extrabold uppercase">
                  4. Link Pengisian Report CV
                </span>
                <h4 className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Link Pengisian Report CV Divisi 1 (Graphis 6201 &amp; Simplicity 6244)
                </h4>
              </div>
              {onOpenReportForm && (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onOpenReportForm('cv_report');
                  }}
                  className="text-xs font-extrabold text-purple-700 hover:text-purple-900 inline-flex items-center gap-1 underline cursor-pointer"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  <span>Buka Form Report CV</span>
                </button>
              )}
            </div>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                readOnly
                value={cvReportUrl}
                onClick={(e) => (e.target as HTMLInputElement).select()}
                className="flex-1 px-3 py-2 rounded-lg bg-white border border-purple-300 text-xs font-mono text-slate-800 select-all"
              />
              <button
                type="button"
                onClick={() => handleCopy(cvReportUrl, 'cv_report')}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-purple-700 hover:bg-purple-800 text-white text-xs font-extrabold shrink-0 cursor-pointer"
              >
                {copiedType === 'cv_report' ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Salin Link Report CV</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* 5. LINK PENGISIAN DATA CV */}
          <div className="p-4 rounded-xl bg-sky-50/70 border border-sky-200 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-sky-700 text-white text-[10px] font-extrabold uppercase">
                  5. Link Pengisian Data CV
                </span>
                <h4 className="text-xs sm:text-sm font-extrabold text-slate-900">
                  Link Pengisian Data CV (Inventory, Promo, Reason &amp; Action Plan)
                </h4>
              </div>
              {onOpenReportForm && (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onOpenReportForm('data_cv');
                  }}
                  className="text-xs font-extrabold text-sky-700 hover:text-sky-900 inline-flex items-center gap-1 underline cursor-pointer"
                >
                  <ExternalLink className="w-3.5 h-3.5" />
                  <span>Buka Form Data CV</span>
                </button>
              )}
            </div>
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
              <input
                type="text"
                readOnly
                value={dataCvUrl}
                onClick={(e) => (e.target as HTMLInputElement).select()}
                className="flex-1 px-3 py-2 rounded-lg bg-white border border-sky-300 text-xs font-mono text-slate-800 select-all"
              />
              <button
                type="button"
                onClick={() => handleCopy(dataCvUrl, 'data_cv')}
                className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-lg bg-sky-700 hover:bg-sky-800 text-white text-xs font-extrabold shrink-0 cursor-pointer"
              >
                {copiedType === 'data_cv' ? (
                  <>
                    <Check className="w-3.5 h-3.5" />
                    <span>Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Salin Link Data CV</span>
                  </>
                )}
              </button>
            </div>
          </div>

          {/* SECTION: LINK GERBANG MASUK & LINK ADMIN */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
              <div className="text-xs font-extrabold text-slate-900 flex items-center gap-1.5">
                <Link2 className="w-3.5 h-3.5 text-indigo-600" />
                <span>Link Gerbang Masuk (Pilih Publik / Admin)</span>
              </div>
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  readOnly
                  value={gateEntryUrl}
                  className="flex-1 min-w-0 px-2.5 py-1.5 rounded bg-white border border-slate-300 text-[11px] font-mono text-slate-700"
                />
                <button
                  type="button"
                  onClick={() => handleCopy(gateEntryUrl, 'gate')}
                  className="px-2.5 py-1.5 rounded bg-slate-800 hover:bg-slate-900 text-white text-[11px] font-bold shrink-0"
                >
                  {copiedType === 'gate' ? 'Tersalin!' : 'Salin'}
                </button>
              </div>
            </div>

            <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
              <div className="text-xs font-extrabold text-slate-900 flex items-center gap-1.5">
                <KeyRound className="w-3.5 h-3.5 text-amber-600" />
                <span>Link Khusus Login Admin</span>
              </div>
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  readOnly
                  value={adminLoginUrl}
                  className="flex-1 min-w-0 px-2.5 py-1.5 rounded bg-white border border-slate-300 text-[11px] font-mono text-slate-700"
                />
                <button
                  type="button"
                  onClick={() => handleCopy(adminLoginUrl, 'admin_link')}
                  className="px-2.5 py-1.5 rounded bg-indigo-600 hover:bg-indigo-700 text-white text-[11px] font-bold shrink-0"
                >
                  {copiedType === 'admin_link' ? 'Tersalin!' : 'Salin'}
                </button>
              </div>
            </div>
          </div>

          {/* SECTION 3: MASUK ADMIN ATAU KELOLA SANDI ADMIN */}
          {accessRole === 'public' ? (
            <form
              onSubmit={handleAdminUnlock}
              className="p-4 rounded-xl bg-indigo-50/70 border border-indigo-200 space-y-3"
            >
              <div className="flex items-center gap-2 text-indigo-950 font-extrabold text-xs sm:text-sm">
                <Lock className="w-4 h-4 text-indigo-600" />
                <span>Masuk sebagai Admin untuk Mengedit Data</span>
              </div>
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                <div className="relative flex-1">
                  <input
                    type="password"
                    value={passwordInput}
                    onChange={(e) => {
                      setPasswordInput(e.target.value);
                      if (loginError) setLoginError(null);
                    }}
                    placeholder="Masukkan Sandi Admin..."
                    className="w-full px-3 py-2 rounded-lg bg-white border border-slate-300 text-xs sm:text-sm font-bold text-slate-900 focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <button
                  type="submit"
                  className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-extrabold flex items-center justify-center gap-1.5 shrink-0"
                >
                  <Unlock className="w-3.5 h-3.5" />
                  <span>Buka Akses Edit (Admin)</span>
                </button>
              </div>
              {loginError && (
                <div className="text-xs font-bold text-rose-600">{loginError}</div>
              )}
              {loginSuccess && (
                <div className="text-xs font-bold text-emerald-700">{loginSuccess}</div>
              )}
            </form>
          ) : (
            <form
              onSubmit={handleSaveNewPassword}
              className="p-4 rounded-xl bg-amber-50/70 border border-amber-200 space-y-3"
            >
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2 text-amber-950 font-extrabold text-xs sm:text-sm">
                  <KeyRound className="w-4 h-4 text-amber-600" />
                  <span>Pengaturan Sandi Admin</span>
                </div>
                <span className="px-2.5 py-0.5 rounded bg-amber-200/80 text-amber-950 text-xs font-extrabold">
                  Sandi Admin Terlindungi (Disembunyikan)
                </span>
              </div>
              <p className="text-xs text-amber-900">
                Sandi Admin disembunyikan demi keamanan. Anda dapat mengganti Sandi Admin baru di bawah ini:
              </p>
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
                <input
                  type="password"
                  value={newPassInput}
                  onChange={(e) => setNewPassInput(e.target.value)}
                  placeholder="Ketik Sandi Admin Baru..."
                  className="flex-1 px-3 py-2 rounded-lg bg-white border border-amber-300 text-xs sm:text-sm font-bold text-slate-900"
                />
                <button
                  type="submit"
                  className="px-4 py-2 rounded-lg bg-amber-500 hover:bg-amber-600 text-slate-950 text-xs font-extrabold shrink-0"
                >
                  Simpan Sandi Baru
                </button>
              </div>
              {passSavedNotice && (
                <div className="text-xs font-bold text-emerald-700">
                  ✓ Sandi Admin berhasil diperbarui.
                </div>
              )}
            </form>
          )}
        </div>

        {/* Footer */}
        <div className="bg-slate-50 px-5 py-3 border-t border-slate-200 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-extrabold"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
};

interface AppLoginLogoutModalProps {
  isOpen: boolean;
  onClose: () => void;
  accessRole: AppAccessRole;
  onLoginAdmin: (passwordInput: string) => boolean;
  onEnterPublic: () => void;
  onLogoutToPublic: () => void;
  onLogoutToEntryGate: () => void;
  onChangeAdminPassword: (newPassword: string) => void;
  onOpenReportForm?: (formType: 'shift' | 'review' | 'cv_report' | 'data_cv') => void;
}

export const AppLoginLogoutModal: React.FC<AppLoginLogoutModalProps> = ({
  isOpen,
  onClose,
  accessRole,
  onLoginAdmin,
  onEnterPublic,
  onLogoutToPublic,
  onLogoutToEntryGate,
  onChangeAdminPassword,
  onOpenReportForm,
}) => {
  const [passwordInput, setPasswordInput] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loginError, setLoginError] = useState<string | null>(null);
  const [loginSuccess, setLoginSuccess] = useState<string | null>(null);
  const [newPassInput, setNewPassInput] = useState('');
  const [passSavedNotice, setPassSavedNotice] = useState(false);

  if (!isOpen) return null;

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const ok = onLoginAdmin(passwordInput.trim());
    if (!ok) {
      setLoginError('Sandi Admin tidak sesuai. Silakan periksa kembali sandi Anda.');
      setLoginSuccess(null);
    } else {
      setLoginError(null);
      setLoginSuccess('Login Admin berhasil! Seluruh fitur edit kini aktif.');
      setPasswordInput('');
      setTimeout(() => {
        setLoginSuccess(null);
        onClose();
      }, 900);
    }
  };

  const handleSaveNewPass = (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newPassInput.trim();
    if (clean.length < 3) return;
    onChangeAdminPassword(clean);
    setNewPassInput('');
    setPassSavedNotice(true);
    setTimeout(() => setPassSavedNotice(false), 2500);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/65 p-4 overflow-y-auto">
      <div className="bg-white border border-slate-200 rounded-2xl shadow-2xl max-w-md w-full overflow-hidden my-auto">
        {/* Header */}
        <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white px-5 py-4 flex items-center justify-between gap-3 border-b border-slate-800">
          <div className="flex items-center gap-2.5 min-w-0">
            <div
              className={`w-9 h-9 rounded-xl flex items-center justify-center shrink-0 font-black ${
                accessRole === 'admin'
                  ? 'bg-emerald-500 text-slate-950'
                  : 'bg-amber-400 text-slate-950'
              }`}
            >
              {accessRole === 'admin' ? (
                <ShieldCheck className="w-5 h-5" />
              ) : (
                <LogIn className="w-5 h-5" />
              )}
            </div>
            <div className="min-w-0">
              <h3 className="text-sm sm:text-base font-extrabold truncate">
                Menu Login &amp; Logout Aplikasi
              </h3>
              <p className="text-[11px] text-amber-300 truncate">
                REPORT BAZAR MATAHARI · Tanjung
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 space-y-4">
          {/* Status Sesi Saat Ini */}
          <div
            className={`p-3.5 rounded-xl border flex items-center gap-3 ${
              accessRole === 'admin'
                ? 'bg-emerald-50 border-emerald-200'
                : 'bg-sky-50 border-sky-200'
            }`}
          >
            <div
              className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${
                accessRole === 'admin'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-sky-600 text-white'
              }`}
            >
              {accessRole === 'admin' ? (
                <ShieldCheck className="w-5 h-5" />
              ) : (
                <Eye className="w-5 h-5" />
              )}
            </div>
            <div className="min-w-0 flex-1">
              <div className="text-[11px] font-bold uppercase tracking-wider text-slate-500">
                Status Login Saat Ini
              </div>
              <div className="text-xs sm:text-sm font-black text-slate-900">
                {accessRole === 'admin' ? (
                  <span className="text-emerald-700">LOGIN ADMIN AKTIF (Bisa Edit)</span>
                ) : (
                  <span className="text-sky-800">MODE PUBLIK (Hanya Lihat)</span>
                )}
              </div>
              <div className="text-[11px] text-slate-600 mt-0.5">
                {accessRole === 'admin'
                  ? 'Anda memiliki akses penuh untuk mengedit tabel sales, simpan, dan export.'
                  : 'Silakan Login menggunakan Sandi Admin jika ingin mengedit data.'}
              </div>
            </div>
          </div>

          {/* JIKA BELUM LOGIN ADMIN -> TAMPILKAN FORM LOGIN ADMIN */}
          {accessRole === 'public' ? (
            <form onSubmit={handleLoginSubmit} className="space-y-3.5">
              <div className="p-4 rounded-xl bg-indigo-50/80 border border-indigo-200 space-y-3">
                <div className="flex items-center gap-2 text-indigo-950 font-extrabold text-xs sm:text-sm">
                  <LogIn className="w-4 h-4 text-indigo-600 shrink-0" />
                  <span>Login Admin (Buka Akses Edit)</span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Sandi Admin
                  </label>
                  <div className="relative">
                    <input
                      type={showPassword ? 'text' : 'password'}
                      value={passwordInput}
                      onChange={(e) => {
                        setPasswordInput(e.target.value);
                        if (loginError) setLoginError(null);
                      }}
                      placeholder="Masukkan Sandi Admin..."
                      autoFocus
                      className="w-full pl-3.5 pr-10 py-2.5 rounded-xl bg-white border border-slate-300 text-sm font-bold text-slate-900 focus:outline-none focus:border-indigo-600 focus:ring-2 focus:ring-indigo-500/20"
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword((v) => !v)}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-700 cursor-pointer"
                      title={showPassword ? 'Sembunyikan Sandi' : 'Lihat Sandi'}
                    >
                      {showPassword ? (
                        <EyeOff className="w-4 h-4" />
                      ) : (
                        <Eye className="w-4 h-4" />
                      )}
                    </button>
                  </div>
                </div>

                {loginError && (
                  <div className="flex items-center gap-1.5 text-xs font-bold text-rose-600">
                    <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                    <span>{loginError}</span>
                  </div>
                )}

                {loginSuccess && (
                  <div className="flex items-center gap-1.5 text-xs font-bold text-emerald-700">
                    <Check className="w-3.5 h-3.5 shrink-0" />
                    <span>{loginSuccess}</span>
                  </div>
                )}

                <button
                  type="submit"
                  className="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs sm:text-sm flex items-center justify-center gap-2 shadow-sm transition-colors cursor-pointer"
                >
                  <LogIn className="w-4 h-4" />
                  <span>Login sebagai Admin Sekarang</span>
                </button>
              </div>
            </form>
          ) : (
            /* JIKA SUDAH LOGIN ADMIN -> TAMPILKAN OPSI LOGOUT & GANTI SANDI */
            <div className="space-y-3.5">
              <div className="p-4 rounded-xl bg-rose-50/80 border border-rose-200 space-y-2.5">
                <div className="flex items-center gap-2 text-rose-950 font-extrabold text-xs sm:text-sm">
                  <LogOut className="w-4 h-4 text-rose-600 shrink-0" />
                  <span>Pilih Mode Logout (Keluar Sesi Admin)</span>
                </div>
                <p className="text-xs text-rose-900">
                  Pilih apakah ingin mengunci kembali ke Mode Publik (Hanya Lihat) atau keluar ke layar Login utama:
                </p>
                <div className="space-y-2 pt-1">
                  <button
                    type="button"
                    onClick={() => {
                      onLogoutToPublic();
                      onClose();
                    }}
                    className="w-full py-2.5 px-3.5 rounded-xl bg-amber-500 hover:bg-amber-600 text-slate-950 font-extrabold text-xs flex items-center justify-center gap-2 transition-colors cursor-pointer"
                  >
                    <Lock className="w-4 h-4" />
                    <span>Logout ke Mode Publik (Hanya Lihat)</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      onClose();
                      onLogoutToEntryGate();
                    }}
                    className="w-full py-2.5 px-3.5 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-extrabold text-xs flex items-center justify-center gap-2 transition-colors cursor-pointer shadow-xs"
                  >
                    <LogOut className="w-4 h-4" />
                    <span>Logout &amp; Kembali ke Layar Login Awal</span>
                  </button>
                </div>
              </div>

              {/* Pengaturan Ganti Sandi Admin */}
              <form
                onSubmit={handleSaveNewPass}
                className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5"
              >
                <div className="flex items-center justify-between">
                  <span className="text-xs font-extrabold text-slate-800 flex items-center gap-1.5">
                    <KeyRound className="w-3.5 h-3.5 text-amber-600" />
                    Ganti Sandi Admin Baru
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="password"
                    value={newPassInput}
                    onChange={(e) => setNewPassInput(e.target.value)}
                    placeholder="Ketik sandi baru..."
                    className="flex-1 px-3 py-2 rounded-lg bg-white border border-slate-300 text-xs font-bold text-slate-900"
                  />
                  <button
                    type="submit"
                    className="px-3 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-extrabold shrink-0 cursor-pointer"
                  >
                    Simpan
                  </button>
                </div>
                {passSavedNotice && (
                  <div className="text-xs font-bold text-emerald-700">
                    ✓ Sandi Admin berhasil diperbarui.
                  </div>
                )}
              </form>
            </div>
          )}

          {/* Tombol Keluar ke Layar Masuk & Akses Cepat saat Mode Publik */}
          <div className="pt-3 border-t border-slate-200 space-y-2">
            {accessRole === 'public' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={() => {
                    onEnterPublic();
                    onClose();
                  }}
                  className="py-2 px-3 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <Eye className="w-3.5 h-3.5 text-sky-600" />
                  <span>Tetap di Mode Publik</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onLogoutToEntryGate();
                  }}
                  className="py-2 px-3 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-extrabold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <LogOut className="w-3.5 h-3.5" />
                  <span>Logout ke Layar Awal</span>
                </button>
              </div>
            )}

            {onOpenReportForm && (
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenReportForm('shift');
                }}
                className="w-full py-2 px-3 rounded-xl bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 font-bold text-xs flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-indigo-600" />
                <span>Buka Form Pengisian Report DP</span>
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

