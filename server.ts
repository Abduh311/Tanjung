import express from 'express';
import fs from 'fs';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import {
  INITIAL_WEEKLY_TABLES,
  INITIAL_REVIEW_DOCS,
  INITIAL_SUBMISSION_LOGS,
  PERSISTED_STORE_UPDATED_AT,
  WeeklySalesTable,
  ReviewMonitoringDoc,
  ReportSubmissionLog,
  normalizeAndRepairWeeklyTables,
  normalizeReviewDocs,
  createEmptyInitialWeeklyTables,
  mergeWeeklyTablesNonDestructive,
  mergeReviewDocsNonDestructive,
} from './src/data/initialData.ts';
import {
  generateInitialCvDiv1Records,
  normalizeCvInventoryOpsList,
} from './src/data/cvDiv1Data.ts';

const STORE_FILE = path.resolve(process.cwd(), 'bazar_store.json');
const BACKUP_STORE_FILE = path.resolve(process.cwd(), 'bazar_store_backup.json');

const activePublicTunnelUrl = process.env.PUBLIC_TUNNEL_URL || '';

interface BazarStoreState {
  weeklyTables: WeeklySalesTable[];
  reviewDocs: ReviewMonitoringDoc[];
  submissionLogs: ReportSubmissionLog[];
  cvDiv1Records?: any[];
  cvInventoryOps?: any[];
  adminPassword?: string;
  updatedAt?: number;
  isPersisted?: boolean;
}

function parseStoreCandidate(raw: string): BazarStoreState | null {
  try {
    const parsed = JSON.parse(raw);
    if (parsed && Array.isArray(parsed.weeklyTables) && parsed.weeklyTables.length > 0) {
      const repairedTables = normalizeAndRepairWeeklyTables(parsed.weeklyTables);
      return {
        weeklyTables: repairedTables,
        reviewDocs:
          Array.isArray(parsed.reviewDocs) && parsed.reviewDocs.length > 0
            ? normalizeReviewDocs(parsed.reviewDocs)
            : INITIAL_REVIEW_DOCS,
        submissionLogs: Array.isArray(parsed.submissionLogs)
          ? parsed.submissionLogs
          : INITIAL_SUBMISSION_LOGS,
        cvDiv1Records: Array.isArray(parsed.cvDiv1Records)
          ? parsed.cvDiv1Records
          : undefined,
        cvInventoryOps: Array.isArray(parsed.cvInventoryOps)
          ? parsed.cvInventoryOps
          : undefined,
        adminPassword:
          typeof parsed.adminPassword === 'string' &&
          parsed.adminPassword.trim() &&
          parsed.adminPassword.trim() !== 'admin123'
            ? parsed.adminPassword.trim()
            : 'MDS311',
        updatedAt: Number(parsed.updatedAt) || PERSISTED_STORE_UPDATED_AT || 1,
        isPersisted: true,
      };
    }
  } catch {}
  return null;
}

function loadStore(): BazarStoreState {
  try {
    if (fs.existsSync(STORE_FILE)) {
      const candidate = parseStoreCandidate(fs.readFileSync(STORE_FILE, 'utf-8'));
      if (candidate) return candidate;
    }
    if (fs.existsSync(BACKUP_STORE_FILE)) {
      const candidate = parseStoreCandidate(fs.readFileSync(BACKUP_STORE_FILE, 'utf-8'));
      if (candidate) return candidate;
    }
  } catch (err) {
    console.error('Error loading store:', err);
  }
  return {
    weeklyTables: normalizeAndRepairWeeklyTables(INITIAL_WEEKLY_TABLES),
    reviewDocs: normalizeReviewDocs(INITIAL_REVIEW_DOCS),
    submissionLogs: INITIAL_SUBMISSION_LOGS,
    adminPassword: 'MDS311',
    updatedAt: PERSISTED_STORE_UPDATED_AT || 1,
    isPersisted: Boolean(PERSISTED_STORE_UPDATED_AT > 0),
  };
}

const PERSISTED_TS_FILE = path.resolve(
  process.cwd(),
  'src/data/persistedStore.ts'
);

let cachedStoreState: BazarStoreState | null = null;
let persistedTsWriteTimer: ReturnType<typeof setTimeout> | null = null;

function saveStore(state: BazarStoreState, preserveExactTimestamp: boolean = false) {
  try {
    const withTimestamp: BazarStoreState = {
      ...state,
      updatedAt: preserveExactTimestamp
        ? Number(state.updatedAt || Date.now())
        : Math.max(Date.now(), Number(state.updatedAt || 0)),
      isPersisted: true,
    };
    cachedStoreState = withTimestamp;
    const jsonPayload = JSON.stringify(withTimestamp, null, 2);
    fs.writeFileSync(STORE_FILE, jsonPayload, 'utf-8');
    try {
      fs.writeFileSync(BACKUP_STORE_FILE, jsonPayload, 'utf-8');
    } catch {}
    if (persistedTsWriteTimer) {
      clearTimeout(persistedTsWriteTimer);
    }
    persistedTsWriteTimer = setTimeout(() => {
      try {
        const tsContent =
          '// Auto-synchronized persistent store snapshot so application data stays permanently saved\n' +
          '// across AI prompts, code edits, builds, and server restarts.\n\n' +
          'export const PERSISTED_STORE_SNAPSHOT = ' +
          jsonPayload +
          ';\n';
        fs.writeFileSync(PERSISTED_TS_FILE, tsContent, 'utf-8');
      } catch {}
    }, 2000);
    return withTimestamp;
  } catch (err) {
    console.error('Error saving store:', err);
    return state;
  }
}

async function startServer() {
  const app = express();
  app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader(
      'Access-Control-Allow-Methods',
      'GET, POST, PUT, DELETE, OPTIONS'
    );
    res.setHeader(
      'Access-Control-Allow-Headers',
      'Content-Type, Authorization, X-Requested-With'
    );
    if (req.method === 'OPTIONS') {
      res.status(204).end();
      return;
    }
    next();
  });
  app.use(express.json({ limit: '5mb' }));

  // GET current bazar state (supports ?since=<timestamp> to avoid transmitting large payloads when unchanged)
  app.get('/api/bazar-data', (req, res) => {
    const state = loadStore();
    res.setHeader('Cache-Control', 'no-store');
    const since = Number(req.query.since || 0);
    if (since > 0 && Number(state.updatedAt || 0) <= since) {
      res.json({
        notModified: true,
        updatedAt: state.updatedAt,
        publicTunnelUrl: activePublicTunnelUrl || undefined,
      });
      return;
    }
    res.json({
      ...state,
      publicTunnelUrl: activePublicTunnelUrl || undefined,
    });
  });

  app.get('/api/public-url', (_req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    res.json({
      publicTunnelUrl: activePublicTunnelUrl || '',
    });
  });

  // PUT or POST full or partial bazar state atomically
  const handleSaveBazarData = (req: express.Request, res: express.Response) => {
    const current = loadStore();
    const isInitialSync = Boolean(req.body?.isInitialSync);
    const isMirrorFromFirestore = Boolean(req.body?.isMirrorFromFirestore);

    let mergedTables = current.weeklyTables;
    if (Array.isArray(req.body?.weeklyTables) && req.body.weeklyTables.length > 0) {
      const incomingTables = normalizeAndRepairWeeklyTables(req.body.weeklyTables);
      if (current.isPersisted && isInitialSync && !isMirrorFromFirestore) {
        mergedTables = mergeWeeklyTablesNonDestructive(
          incomingTables,
          current.weeklyTables
        );
      } else {
        mergedTables = incomingTables;
      }
    }

    let mergedLogs = current.submissionLogs;
    if (Array.isArray(req.body?.submissionLogs)) {
      const seenIds = new Set<string>();
      mergedLogs = [];
      const sourceLogs =
        isInitialSync && !isMirrorFromFirestore && current.isPersisted
          ? [...current.submissionLogs, ...req.body.submissionLogs]
          : req.body.submissionLogs;
      for (const lg of sourceLogs) {
        if (lg && lg.id && !seenIds.has(lg.id)) {
          seenIds.add(lg.id);
          mergedLogs.push(lg);
        }
      }
    }

    const updated: BazarStoreState = {
      weeklyTables: mergedTables,
      reviewDocs:
        Array.isArray(req.body?.reviewDocs) && req.body.reviewDocs.length > 0
          ? isInitialSync && !isMirrorFromFirestore && current.isPersisted
            ? mergeReviewDocsNonDestructive(current.reviewDocs, req.body.reviewDocs)
            : normalizeReviewDocs(req.body.reviewDocs)
          : current.reviewDocs,
      submissionLogs: mergedLogs,
      cvDiv1Records: Array.isArray(req.body?.cvDiv1Records)
        ? req.body.cvDiv1Records
        : current.cvDiv1Records,
      cvInventoryOps: Array.isArray(req.body?.cvInventoryOps)
        ? normalizeCvInventoryOpsList(req.body.cvInventoryOps)
        : current.cvInventoryOps,
      adminPassword:
        typeof req.body?.adminPassword === 'string' && req.body.adminPassword.trim()
          ? req.body.adminPassword.trim()
          : current.adminPassword || 'MDS311',
      updatedAt: isMirrorFromFirestore
        ? Number(req.body?.updatedAt || current.updatedAt || Date.now())
        : isInitialSync && current.isPersisted
          ? Number(current.updatedAt || Date.now())
          : Math.max(
              Date.now(),
              Number(current.updatedAt || 0) + 1,
              Number(req.body?.updatedAt || 0)
            ),
      isPersisted: true,
    };
    const saved = saveStore(updated, isMirrorFromFirestore);
    res.json({ ok: true, state: saved });
  };

  app.put('/api/bazar-data', handleSaveBazarData);
  app.post('/api/bazar-data', handleSaveBazarData);

  // POST atomic single-report submission (never resets or overwrites other days/shifts/weeks)
  app.post('/api/bazar-submit-report', (req, res) => {
    const current = loadStore();
    const {
      targetTable,
      dayIdx,
      updatedDay,
      shiftPatch,
      updatedReviewDocs,
      newLog,
    } = req.body || {};

    let nextTables = [...current.weeklyTables];
    if (targetTable && typeof targetTable.weekNumber === 'number') {
      const existingIdx = nextTables.findIndex(
        (t) => t.weekNumber === targetTable.weekNumber || t.id === targetTable.id
      );

      const targetTanggal = (
        shiftPatch?.tanggal ||
        updatedDay?.tanggal ||
        targetTable?.days?.[dayIdx]?.tanggal ||
        ''
      ).trim();

      const mergeDayRecord = (existingDay: any) => {
        if (shiftPatch) {
          const {
            shouldUpdatePagi,
            shouldUpdateSiang,
            pagi,
            siang,
            reasonPagi,
            reasonSiang,
          } = shiftPatch;
          return {
            ...existingDay,
            pagi: shouldUpdatePagi
              ? {
                  nama:
                    pagi?.nama !== undefined
                      ? pagi.nama
                      : existingDay.pagi?.nama || '',
                  sales:
                    pagi?.sales !== undefined && pagi?.sales !== null
                      ? pagi.sales
                      : existingDay.pagi?.sales ?? null,
                  qty:
                    pagi?.qty !== undefined && pagi?.qty !== null
                      ? pagi.qty
                      : existingDay.pagi?.qty ?? null,
                  trx:
                    pagi?.trx !== undefined && pagi?.trx !== null
                      ? pagi.trx
                      : existingDay.pagi?.trx ?? null,
                  traffic:
                    pagi?.traffic !== undefined && pagi?.traffic !== null
                      ? pagi.traffic
                      : existingDay.pagi?.traffic ?? null,
                }
              : existingDay.pagi,
            siang: shouldUpdateSiang
              ? {
                  nama:
                    siang?.nama !== undefined
                      ? siang.nama
                      : existingDay.siang?.nama || '',
                  sales:
                    siang?.sales !== undefined && siang?.sales !== null
                      ? siang.sales
                      : existingDay.siang?.sales ?? null,
                  qty:
                    siang?.qty !== undefined && siang?.qty !== null
                      ? siang.qty
                      : existingDay.siang?.qty ?? null,
                  trx:
                    siang?.trx !== undefined && siang?.trx !== null
                      ? siang.trx
                      : existingDay.siang?.trx ?? null,
                  traffic:
                    siang?.traffic !== undefined && siang?.traffic !== null
                      ? siang.traffic
                      : existingDay.siang?.traffic ?? null,
                }
              : existingDay.siang,
            reasonPagi:
              shouldUpdatePagi && typeof reasonPagi === 'string' && reasonPagi.trim() !== ''
                ? reasonPagi
                : existingDay.reasonPagi || '',
            reasonSiang:
              shouldUpdateSiang && typeof reasonSiang === 'string' && reasonSiang.trim() !== ''
                ? reasonSiang
                : existingDay.reasonSiang || '',
          };
        }
        return updatedDay || existingDay;
      };

      const isMatchingDay = (d: any, idx: number, daysArr: any[]) => {
        if (targetTanggal && d?.tanggal && d.tanggal.trim() === targetTanggal) {
          return true;
        }
        if (
          targetTanggal &&
          daysArr.some((item) => item?.tanggal?.trim() === targetTanggal)
        ) {
          return false;
        }
        return idx === dayIdx;
      };

      if (existingIdx >= 0) {
        const existingTbl = nextTables[existingIdx];
        nextTables[existingIdx] = {
          ...existingTbl,
          days: existingTbl.days.map((d, idx, arr) =>
            isMatchingDay(d, idx, arr) ? mergeDayRecord(d) : d
          ),
        };
      } else {
        const newTbl: WeeklySalesTable = {
          ...targetTable,
          days: Array.isArray(targetTable.days)
            ? targetTable.days.map((d: any, idx: number, arr: any[]) =>
                isMatchingDay(d, idx, arr) ? mergeDayRecord(d) : d
              )
            : [],
        };
        nextTables = [...nextTables, newTbl].sort(
          (a, b) => a.weekNumber - b.weekNumber
        );
      }
    }

    const nextReviews = Array.isArray(updatedReviewDocs)
      ? mergeReviewDocsNonDestructive(updatedReviewDocs, current.reviewDocs)
      : current.reviewDocs;

    const seenIds = new Set<string>();
    const nextLogs: ReportSubmissionLog[] = [];
    const candidateLogs = newLog
      ? [newLog, ...current.submissionLogs]
      : [...current.submissionLogs];
    for (const lg of candidateLogs) {
      if (lg && lg.id && !seenIds.has(lg.id)) {
        seenIds.add(lg.id);
        nextLogs.push(lg);
      }
    }

    const saved = saveStore({
      weeklyTables: normalizeAndRepairWeeklyTables(nextTables),
      reviewDocs: nextReviews,
      submissionLogs: nextLogs,
      cvDiv1Records: current.cvDiv1Records,
      cvInventoryOps: current.cvInventoryOps,
      adminPassword: current.adminPassword || 'MDS311',
      updatedAt: Math.max(Date.now(), Number(current.updatedAt || 0) + 1),
      isPersisted: true,
    });

    res.json({ ok: true, state: saved });
  });

  // POST atomic CV report / Data CV Ops submission from public report links
  app.post('/api/bazar-submit-cv-report', (req, res) => {
    const current = loadStore();
    const { updatedCvRecord, updatedCvOpsEntries, newLog } = req.body || {};

    let nextCvRecords = Array.isArray(current.cvDiv1Records) && current.cvDiv1Records.length > 0
      ? [...current.cvDiv1Records]
      : generateInitialCvDiv1Records();

    if (updatedCvRecord && typeof updatedCvRecord.tanggal === 'string' && updatedCvRecord.tanggal.trim()) {
      const targetTgl = updatedCvRecord.tanggal.trim();
      const idx = nextCvRecords.findIndex((r: any) => String(r?.tanggal || '').trim() === targetTgl);
      if (idx >= 0) {
        nextCvRecords[idx] = updatedCvRecord;
      } else {
        nextCvRecords = [...nextCvRecords, updatedCvRecord];
      }
    }

    let nextCvOps = normalizeCvInventoryOpsList(
      Array.isArray(current.cvInventoryOps) ? current.cvInventoryOps : []
    );
    if (Array.isArray(updatedCvOpsEntries) && updatedCvOpsEntries.length > 0) {
      const incomingMap = new Map<string, any>();
      updatedCvOpsEntries.forEach((entry: any) => {
        if (entry && typeof entry.id === 'string' && entry.id.trim()) {
          incomingMap.set(entry.id, entry);
        }
      });
      nextCvOps = nextCvOps.map((existing) =>
        incomingMap.has(existing.id) ? { ...existing, ...incomingMap.get(existing.id) } : existing
      );
      incomingMap.forEach((entry, id) => {
        if (!nextCvOps.some((e) => e.id === id)) {
          nextCvOps.push(entry);
        }
      });
      nextCvOps = normalizeCvInventoryOpsList(nextCvOps);
    }

    const seenIds = new Set<string>();
    const nextLogs: ReportSubmissionLog[] = [];
    const candidateLogs = newLog
      ? [newLog, ...current.submissionLogs]
      : [...current.submissionLogs];
    for (const lg of candidateLogs) {
      if (lg && lg.id && !seenIds.has(lg.id)) {
        seenIds.add(lg.id);
        nextLogs.push(lg);
      }
    }

    const saved = saveStore({
      weeklyTables: current.weeklyTables,
      reviewDocs: current.reviewDocs,
      submissionLogs: nextLogs,
      cvDiv1Records: nextCvRecords,
      cvInventoryOps: nextCvOps,
      adminPassword: current.adminPassword || 'MDS311',
      updatedAt: Math.max(Date.now(), Number(current.updatedAt || 0) + 1),
      isPersisted: true,
    });

    res.json({ ok: true, state: saved });
  });

  // POST Reset to initial data or empty (0) data
  app.post('/api/bazar-reset', (req, res) => {
    const current = loadStore();
    const mode = req.body?.mode === 'empty' ? 'empty' : 'initial';
    const resetState: BazarStoreState =
      mode === 'empty'
        ? {
            weeklyTables: createEmptyInitialWeeklyTables(),
            reviewDocs: INITIAL_REVIEW_DOCS,
            submissionLogs: [],
            cvDiv1Records: generateInitialCvDiv1Records(),
            cvInventoryOps: normalizeCvInventoryOpsList([]),
            adminPassword: current.adminPassword || 'MDS311',
            updatedAt: Date.now(),
            isPersisted: true,
          }
        : {
            weeklyTables: INITIAL_WEEKLY_TABLES,
            reviewDocs: INITIAL_REVIEW_DOCS,
            submissionLogs: INITIAL_SUBMISSION_LOGS,
            cvDiv1Records: generateInitialCvDiv1Records(),
            cvInventoryOps: normalizeCvInventoryOpsList([]),
            adminPassword: current.adminPassword || 'MDS311',
            updatedAt: Date.now(),
            isPersisted: true,
          };
    const saved = saveStore(resetState);
    res.json({ ok: true, state: saved, ...saved });
  });

  const isProd = process.env.NODE_ENV === 'production';
  const distPath = path.resolve(process.cwd(), 'dist');
  const hasDistIndex = fs.existsSync(path.join(distPath, 'index.html'));

  if (!isProd || !hasDistIndex) {
    const vite = await createViteServer({
      server: { middlewareMode: true, allowedHosts: true, cors: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(distPath));
    app.use((_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  const PORT = Number(process.env.PORT) || 3000;
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
