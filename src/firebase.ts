import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut } from 'firebase/auth';
import {
  getFirestore,
  initializeFirestore,
  doc,
  getDoc,
  getDocFromServer,
  setDoc,
  onSnapshot,
  Unsubscribe,
} from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';
import {
  WeeklySalesTable,
  ReviewMonitoringDoc,
  ReportSubmissionLog,
  normalizeAndRepairWeeklyTables,
  normalizeReviewDocs,
} from './data/initialData';
import {
  CvDailyRecord,
  CvInventoryOpsEntry,
  normalizeCvInventoryOpsList,
} from './data/cvDiv1Data';

// Initialize Firebase SDK with provisioned configuration.
// Force HTTP long polling so Firestore works reliably inside iframes, Cloud Run proxies,
// and restricted networks where streaming WebChannel handshakes can stall for 10s.
export const app =
  getApps().length > 0 ? getApp() : initializeApp(firebaseConfig);
export const db = (() => {
  try {
    return initializeFirestore(
      app,
      {
        experimentalForceLongPolling: true,
        ignoreUndefinedProperties: true,
      },
      firebaseConfig.firestoreDatabaseId
    );
  } catch {
    return getFirestore(app, firebaseConfig.firestoreDatabaseId);
  }
})();
export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

export function isPermissionError(error: unknown): boolean {
  if (!error) return false;
  const msg = error instanceof Error ? error.message : String(error);
  const code = (error as { code?: string })?.code;
  return (
    code === 'permission-denied' ||
    msg.toLowerCase().includes('permission') ||
    msg.toLowerCase().includes('insufficient')
  );
}

export function handleFirestoreError(
  error: unknown,
  operationType: OperationType,
  path: string | null
): never {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo:
        auth.currentUser?.providerData?.map((provider) => ({
          providerId: provider.providerId,
          email: provider.email,
        })) || [],
    },
    operationType,
    path,
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Validate connection to Firestore on boot
async function testConnection() {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    return;
  }
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      try {
        await new Promise((resolve) => setTimeout(resolve, 2000));
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (retryError) {
        if (
          retryError instanceof Error &&
          retryError.message.includes('the client is offline')
        ) {
          console.error('Please check your Firebase configuration.');
        }
      }
    }
  }
}
testConnection();

const WORKSPACE_COLLECTION = 'bazarWorkspaces';
const WORKSPACE_DOC_ID = 'main';
const CV_WORKSPACE_DOC_ID = 'cv';
const SUBMISSIONS_COLLECTION = 'reportSubmissions';

export interface BazarFirestoreWorkspacePayload {
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  submissionLogs: ReportSubmissionLog[];
  updatedAtMs: number;
  updatedByRole: 'admin' | 'report_link' | 'public_sync';
}

export interface BazarFirestoreCvPayload {
  cvDiv1Records: CvDailyRecord[];
  cvInventoryOps: CvInventoryOpsEntry[];
  updatedAtMs: number;
  updatedByRole: 'admin' | 'report_link' | 'public_sync';
}

function sanitizeIdForFirestore(rawId: string): string {
  const cleaned = (rawId || 'item')
    .replace(/[^a-zA-Z0-9_-]/g, '-')
    .slice(0, 128);
  return cleaned || 'item-default';
}

let latestKnownFirestoreUpdatedAtMs = 0;
let latestKnownCvFirestoreUpdatedAtMs = 0;

/**
 * Saves the synchronized Bazar workspace state to Firestore (`/bazarWorkspaces/main`)
 * while strictly adhering to the blueprint schema & size limits.
 */
export async function saveBazarWorkspaceToFirestore(
  weeklyTables: WeeklySalesTable[],
  reviewDocs: ReviewMonitoringDoc[],
  submissionLogs: ReportSubmissionLog[],
  updatedAtMs: number = Date.now(),
  updatedByRole: 'admin' | 'report_link' | 'public_sync' = 'admin'
): Promise<number> {
  const path = `${WORKSPACE_COLLECTION}/${WORKSPACE_DOC_ID}`;
  const weeklyTablesJson = JSON.stringify(weeklyTables || []).slice(0, 900000);
  const reviewDocsJson = JSON.stringify(reviewDocs || []).slice(0, 500000);
  const submissionLogsJson = JSON.stringify(
    (submissionLogs || []).slice(0, 250)
  ).slice(0, 500000);

  const safeUpdatedAtMs = Math.max(
    Date.now(),
    Number(updatedAtMs) || 0,
    latestKnownFirestoreUpdatedAtMs + 1
  );
  latestKnownFirestoreUpdatedAtMs = safeUpdatedAtMs;

  const payload = {
    workspaceId: WORKSPACE_DOC_ID,
    weeklyTablesJson: weeklyTablesJson.length >= 2 ? weeklyTablesJson : '[]',
    reviewDocsJson: reviewDocsJson.length >= 2 ? reviewDocsJson : '[]',
    submissionLogsJson:
      submissionLogsJson.length >= 2 ? submissionLogsJson : '[]',
    updatedAtMs: safeUpdatedAtMs,
    updatedByRole,
  };

  try {
    await setDoc(doc(db, WORKSPACE_COLLECTION, WORKSPACE_DOC_ID), payload);
    return safeUpdatedAtMs;
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Saves the synchronized CV Divisi 1 & Inventory Ops state to Firestore (`/bazarWorkspaces/cv`).
 */
export async function saveCvWorkspaceToFirestore(
  cvDiv1Records: CvDailyRecord[],
  cvInventoryOps: CvInventoryOpsEntry[],
  updatedAtMs: number = Date.now(),
  updatedByRole: 'admin' | 'report_link' | 'public_sync' = 'admin'
): Promise<number> {
  const path = `${WORKSPACE_COLLECTION}/${CV_WORKSPACE_DOC_ID}`;
  const cvDiv1RecordsJson = JSON.stringify(cvDiv1Records || []).slice(0, 900000);
  const cvInventoryOpsJson = JSON.stringify(cvInventoryOps || []).slice(0, 500000);

  const safeUpdatedAtMs = Math.max(
    Date.now(),
    Number(updatedAtMs) || 0,
    latestKnownCvFirestoreUpdatedAtMs + 1
  );
  latestKnownCvFirestoreUpdatedAtMs = safeUpdatedAtMs;

  const payload = {
    workspaceId: CV_WORKSPACE_DOC_ID,
    cvDiv1RecordsJson: cvDiv1RecordsJson.length >= 2 ? cvDiv1RecordsJson : '[]',
    cvInventoryOpsJson: cvInventoryOpsJson.length >= 2 ? cvInventoryOpsJson : '[]',
    updatedAtMs: safeUpdatedAtMs,
    updatedByRole,
  };

  try {
    await setDoc(doc(db, WORKSPACE_COLLECTION, CV_WORKSPACE_DOC_ID), payload);
    return safeUpdatedAtMs;
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Saves an individual shift or review report submission to `/reportSubmissions/{submissionId}`.
 */
export async function saveReportSubmissionToFirestore(
  log: ReportSubmissionLog
): Promise<void> {
  const safeId = sanitizeIdForFirestore(log.id);
  const path = `${SUBMISSIONS_COLLECTION}/${safeId}`;

  const payload = {
    id: safeId,
    timestamp: String(log.timestamp || '-').slice(0, 64),
    reportType:
      log.reportType === 'review_monitoring' ||
      log.reportType === 'cv_report' ||
      log.reportType === 'data_cv'
        ? log.reportType
        : 'shift_sales',
    submittedBy: String(log.submittedBy || 'Petugas').slice(0, 120),
    weekTitle: String(log.weekTitle || 'Weekly').slice(0, 120),
    tanggal: String(log.tanggal || '-').slice(0, 32),
    shift: String(log.shift || 'Review Bazar').slice(0, 32),
    salesAmount: Math.max(0, Number(log.salesAmount) || 0),
    qtyAmount: Math.max(0, Number(log.qtyAmount) || 0),
    trxAmount: Math.max(0, Number(log.trxAmount) || 0),
    summaryNote: String(log.summaryNote || '').slice(0, 2000),
    createdAtMs: Date.now(),
  };

  try {
    await setDoc(doc(db, SUBMISSIONS_COLLECTION, safeId), payload);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

/**
 * Fetches the current Bazar workspace state from Firestore server directly.
 */
export async function fetchBazarWorkspaceFromFirestore(): Promise<BazarFirestoreWorkspacePayload | null> {
  const path = `${WORKSPACE_COLLECTION}/${WORKSPACE_DOC_ID}`;
  const docRef = doc(db, WORKSPACE_COLLECTION, WORKSPACE_DOC_ID);
  try {
    let snap;
    try {
      snap = await getDocFromServer(docRef);
    } catch {
      snap = await getDoc(docRef);
    }
    if (!snap.exists()) return null;
    const data = snap.data();
    if (!data) return null;

    const parsedTables = JSON.parse(data.weeklyTablesJson || '[]');
    const parsedReviews = JSON.parse(data.reviewDocsJson || '[]');
    const parsedLogs = JSON.parse(data.submissionLogsJson || '[]');

    if (!Array.isArray(parsedTables) || parsedTables.length === 0) {
      return null;
    }

    const remoteTs = Number(data.updatedAtMs) || 0;
    if (remoteTs > latestKnownFirestoreUpdatedAtMs) {
      latestKnownFirestoreUpdatedAtMs = remoteTs;
    }

    return {
      weeklyTables: normalizeAndRepairWeeklyTables(parsedTables),
      reviewDocs: Array.isArray(parsedReviews)
        ? normalizeReviewDocs(parsedReviews)
        : [],
      submissionLogs: Array.isArray(parsedLogs) ? parsedLogs : [],
      updatedAtMs: remoteTs,
      updatedByRole:
        data.updatedByRole === 'report_link' ||
        data.updatedByRole === 'public_sync'
          ? data.updatedByRole
          : 'admin',
    };
  } catch (error) {
    if (isPermissionError(error)) {
      handleFirestoreError(error, OperationType.GET, path);
    }
    return null;
  }
}

/**
 * Fetches the CV Divisi 1 & Inventory Ops state from Firestore (`/bazarWorkspaces/cv`).
 */
export async function fetchCvWorkspaceFromFirestore(): Promise<BazarFirestoreCvPayload | null> {
  const path = `${WORKSPACE_COLLECTION}/${CV_WORKSPACE_DOC_ID}`;
  const docRef = doc(db, WORKSPACE_COLLECTION, CV_WORKSPACE_DOC_ID);
  try {
    let snap;
    try {
      snap = await getDocFromServer(docRef);
    } catch {
      snap = await getDoc(docRef);
    }
    if (!snap.exists()) return null;
    const data = snap.data();
    if (!data) return null;

    const parsedRecords = JSON.parse(data.cvDiv1RecordsJson || '[]');
    const parsedOps = JSON.parse(data.cvInventoryOpsJson || '[]');

    if (!Array.isArray(parsedRecords) || parsedRecords.length === 0) {
      return null;
    }

    const remoteTs = Number(data.updatedAtMs) || 0;
    if (remoteTs > latestKnownCvFirestoreUpdatedAtMs) {
      latestKnownCvFirestoreUpdatedAtMs = remoteTs;
    }

    return {
      cvDiv1Records: parsedRecords,
      cvInventoryOps: normalizeCvInventoryOpsList(parsedOps),
      updatedAtMs: remoteTs,
      updatedByRole:
        data.updatedByRole === 'report_link' ||
        data.updatedByRole === 'public_sync'
          ? data.updatedByRole
          : 'admin',
    };
  } catch (error) {
    if (isPermissionError(error)) {
      handleFirestoreError(error, OperationType.GET, path);
    }
    return null;
  }
}

/**
 * Subscribes in real time (`onSnapshot`) to `/bazarWorkspaces/main` so that
 * whenever someone fills a Report Link or Admin edits data in AI Studio or Vercel,
 * all connected browsers receive the exact update instantaneously via Firebase Firestore.
 */
export function subscribeToBazarWorkspaceFirestore(
  onUpdate: (payload: BazarFirestoreWorkspacePayload) => void
): Unsubscribe {
  const path = `${WORKSPACE_COLLECTION}/${WORKSPACE_DOC_ID}`;
  return onSnapshot(
    doc(db, WORKSPACE_COLLECTION, WORKSPACE_DOC_ID),
    (snap) => {
      if (!snap.exists()) return;
      const data = snap.data();
      if (!data) return;
      try {
        const parsedTables = JSON.parse(data.weeklyTablesJson || '[]');
        const parsedReviews = JSON.parse(data.reviewDocsJson || '[]');
        const parsedLogs = JSON.parse(data.submissionLogsJson || '[]');
        if (!Array.isArray(parsedTables) || parsedTables.length === 0) return;

        const remoteTs = Number(data.updatedAtMs) || 0;
        if (remoteTs > latestKnownFirestoreUpdatedAtMs) {
          latestKnownFirestoreUpdatedAtMs = remoteTs;
        }

        onUpdate({
          weeklyTables: normalizeAndRepairWeeklyTables(parsedTables),
          reviewDocs: Array.isArray(parsedReviews)
            ? normalizeReviewDocs(parsedReviews)
            : [],
          submissionLogs: Array.isArray(parsedLogs) ? parsedLogs : [],
          updatedAtMs: remoteTs,
          updatedByRole:
            data.updatedByRole === 'report_link' ||
            data.updatedByRole === 'public_sync'
              ? data.updatedByRole
              : 'admin',
        });
      } catch (err) {
        console.error('Error parsing Firestore workspace snapshot:', err);
      }
    },
    (error) => {
      if (isPermissionError(error)) {
        handleFirestoreError(error, OperationType.GET, path);
      } else {
        console.warn('Firestore real-time sync is waiting for connection/offline mode:', error?.message);
      }
    }
  );
}

/**
 * Subscribes in real time (`onSnapshot`) to `/bazarWorkspaces/cv`.
 */
export function subscribeToCvWorkspaceFirestore(
  onUpdate: (payload: BazarFirestoreCvPayload) => void
): Unsubscribe {
  const path = `${WORKSPACE_COLLECTION}/${CV_WORKSPACE_DOC_ID}`;
  return onSnapshot(
    doc(db, WORKSPACE_COLLECTION, CV_WORKSPACE_DOC_ID),
    (snap) => {
      if (!snap.exists()) return;
      const data = snap.data();
      if (!data) return;
      try {
        const parsedRecords = JSON.parse(data.cvDiv1RecordsJson || '[]');
        const parsedOps = JSON.parse(data.cvInventoryOpsJson || '[]');
        if (!Array.isArray(parsedRecords) || parsedRecords.length === 0) return;

        const remoteTs = Number(data.updatedAtMs) || 0;
        if (remoteTs > latestKnownCvFirestoreUpdatedAtMs) {
          latestKnownCvFirestoreUpdatedAtMs = remoteTs;
        }

        onUpdate({
          cvDiv1Records: parsedRecords,
          cvInventoryOps: normalizeCvInventoryOpsList(parsedOps),
          updatedAtMs: remoteTs,
          updatedByRole:
            data.updatedByRole === 'report_link' ||
            data.updatedByRole === 'public_sync'
              ? data.updatedByRole
              : 'admin',
        });
      } catch (err) {
        console.error('Error parsing Firestore CV workspace snapshot:', err);
      }
    },
    (error) => {
      if (isPermissionError(error)) {
        handleFirestoreError(error, OperationType.GET, path);
      } else {
        console.warn('Firestore CV real-time sync is waiting for connection/offline mode:', error?.message);
      }
    }
  );
}

export async function signInWithGooglePopup() {
  return signInWithPopup(auth, googleProvider);
}

export async function signOutFirebase() {
  return signOut(auth);
}

