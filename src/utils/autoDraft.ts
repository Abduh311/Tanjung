import { useState, useEffect, useRef, useCallback } from 'react';

export interface AutoDraftEnvelope<T> {
  key: string;
  menuName: string;
  savedAt: number;
  savedTimeStr: string;
  isDirty: boolean;
  data: T;
}

export const AUTO_DRAFT_KEYS = {
  GLOBAL_WORKSPACE: 'bazar_autodraft_global_workspace_v1',
  DAILY_INPUT: 'bazar_autodraft_daily_input_v1',
  SALES_MATRIX: 'bazar_autodraft_sales_matrix_v1',
  QUICK_DAY_MODAL: 'bazar_autodraft_quick_day_modal_v1',
  REASON_EVALUATION: 'bazar_autodraft_reason_evaluation_v1',
  REVIEW_MONITORING: 'bazar_autodraft_review_monitoring_v1',
  REPORT_FORM_SHIFT: 'bazar_autodraft_report_form_shift_v1',
  REPORT_FORM_REVIEW: 'bazar_autodraft_report_form_review_v1',
  REPORT_FORM_CV: 'bazar_autodraft_report_form_cv_v1',
  REPORT_FORM_DATA_CV: 'bazar_autodraft_report_form_data_cv_v1',
} as const;

export function formatDraftTime(ts: number = Date.now()): string {
  try {
    return new Date(ts).toLocaleTimeString('id-ID', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
  } catch {
    return '';
  }
}

export function saveAutoDraftToStorage<T>(
  _key: string,
  _menuName: string,
  _data: T,
  _isDirty: boolean = true
): AutoDraftEnvelope<T> | null {
  return null;
}

export function loadAutoDraftFromStorage<T>(
  _key: string
): AutoDraftEnvelope<T> | null {
  return null;
}

export function clearAutoDraftFromStorage(key: string): void {
  try {
    localStorage.removeItem(key);
  } catch {}
}

export function markAutoDraftClean(_key: string): void {}

export function usePeriodicAutoDraft<T>(_opts: {
  storageKey: string;
  menuName: string;
  data: T;
  isDirty: boolean;
  enabled?: boolean;
  intervalMs?: number;
}) {
  const flushDraftNow = useCallback((_overrideDirty?: boolean) => null, []);

  return {
    lastDraftTimeStr: '',
    isOnline: true,
    justAutoSaved: false,
    flushDraftNow,
  };
}

