const STORAGE_KEY_CUSTOM_PUBLIC_ORIGIN = 'bazar_kapuas_public_origin_v1';

export const DEFAULT_SHARED_APP_ORIGIN =
  (typeof import.meta !== 'undefined' &&
    (import.meta as { env?: Record<string, string | undefined> }).env?.VITE_APP_URL &&
    String((import.meta as { env?: Record<string, string | undefined> }).env?.VITE_APP_URL)
      .trim()
      .replace(/\/+$/, '')) ||
  'https://ais-pre-uklvsl4nqhfutwisngvitw-780718281765.asia-east1.run.app';

let runtimeAutoTunnelOrigin = '';

export function setAutoDetectedPublicOrigin(tunnelUrl?: string): void {
  if (!tunnelUrl || !tunnelUrl.trim().startsWith('https://')) return;
  const clean = tunnelUrl.trim().replace(/\/+$/, '');
  runtimeAutoTunnelOrigin = clean;
}

export function getAutoDetectedPublicOrigin(): string {
  return runtimeAutoTunnelOrigin;
}

/**
 * Returns the clean origin (without trailing slash) for Link Aplikasi & Link Pengisian Report.
 */
export function getPublicShareableOrigin(customOriginOverride?: string): string {
  if (customOriginOverride && customOriginOverride.trim()) {
    return customOriginOverride.trim().replace(/\/+$/, '');
  }

  if (
    typeof window !== 'undefined' &&
    window.location &&
    window.location.origin
  ) {
    const currentOrigin = window.location.origin.replace(/\/+$/, '');
    // Convert private ais-dev-* URL to public Shared App URL ais-pre-*
    if (currentOrigin.includes('ais-dev-')) {
      return currentOrigin.replace('ais-dev-', 'ais-pre-');
    }
    return currentOrigin;
  }

  if (runtimeAutoTunnelOrigin) {
    return runtimeAutoTunnelOrigin;
  }

  return DEFAULT_SHARED_APP_ORIGIN;
}

/**
 * Returns the shareable base URL with trailing slash (including GitHub Pages subpath if applicable).
 */
export function getPublicShareableBaseUrl(customOriginOverride?: string): string {
  const origin = getPublicShareableOrigin(customOriginOverride);
  if (customOriginOverride && customOriginOverride.trim()) {
    try {
      const parsed = new URL(origin);
      if (parsed.pathname && parsed.pathname !== '/') {
        return `${origin}/`;
      }
    } catch {}
  }
  const rawBase =
    (typeof import.meta !== 'undefined' &&
      (import.meta as { env?: Record<string, string | undefined> }).env?.BASE_URL) ||
    '/';
  const normalizedBase = rawBase.startsWith('/') ? rawBase : `/${rawBase}`;
  const cleanBase = normalizedBase.endsWith('/') ? normalizedBase : `${normalizedBase}/`;
  return `${origin}${cleanBase}`;
}

/**
 * Direct link to open the application (`/?mode=public`).
 */
export function getGooglePublicAppUrl(
  customOriginOverride?: string,
  extraParams?: Record<string, string>
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('mode', 'public');
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

/**
 * Direct link to the Entry Gate (`/?mode=gate`).
 */
export function getGooglePublicGateUrl(
  customOriginOverride?: string,
  extraParams?: Record<string, string>
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('mode', 'gate');
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

/**
 * Direct link to fill the Daily Shift Report (`/?form=shift&lock=10`).
 */
export function getGooglePublicShiftReportUrl(
  extraParams?: Record<string, string>,
  customOriginOverride?: string
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('form', 'shift');
  params.set('lock', '10');
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

/**
 * Direct link to fill the Review Monitoring Bazar form (`/?form=review&lock=10`).
 */
export function getGooglePublicReviewReportUrl(
  docId?: string,
  customOriginOverride?: string,
  extraParams?: Record<string, string>
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('form', 'review');
  params.set('lock', '10');
  if (docId) {
    params.set('doc', docId);
  }
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

/**
 * Direct link to fill Report CV (Divisi 1: 6201 Graphis & 6244 Simplicity) (`/?form=cv_report&lock=10`).
 */
export function getGooglePublicCvReportUrl(
  extraParams?: Record<string, string>,
  customOriginOverride?: string
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('form', 'cv_report');
  params.set('lock', '10');
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

/**
 * Direct link to fill Data CV (Inventory, Promo, Reason/Kendala & Action Plan) (`/?form=data_cv&lock=10`).
 */
export function getGooglePublicDataCvUrl(
  extraParams?: Record<string, string>,
  customOriginOverride?: string
): string {
  const baseUrl = getPublicShareableBaseUrl(customOriginOverride);
  const params = new URLSearchParams();
  params.set('form', 'data_cv');
  params.set('lock', '10');
  if (extraParams) {
    Object.entries(extraParams).forEach(([k, v]) => {
      if (v !== undefined && v !== '') {
        params.set(k, v);
      }
    });
  }
  return `${baseUrl}?${params.toString()}`;
}

export function getSavedCustomPublicOrigin(): string {
  try {
    return localStorage.getItem(STORAGE_KEY_CUSTOM_PUBLIC_ORIGIN) || '';
  } catch {
    return '';
  }
}

export function setSavedCustomPublicOrigin(originUrl: string): void {
  try {
    const clean = (originUrl || '').trim();
    if (!clean) {
      localStorage.removeItem(STORAGE_KEY_CUSTOM_PUBLIC_ORIGIN);
    } else {
      localStorage.setItem(STORAGE_KEY_CUSTOM_PUBLIC_ORIGIN, clean);
    }
  } catch {}
}
