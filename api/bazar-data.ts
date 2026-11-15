import type { IncomingMessage, ServerResponse } from 'node:http';

function readBody(req: IncomingMessage): Promise<Record<string, any>> {
  return new Promise((resolve) => {
    const preParsed = (req as IncomingMessage & { body?: unknown }).body;
    if (preParsed && typeof preParsed === 'object') {
      resolve(preParsed as Record<string, any>);
      return;
    }
    if (typeof preParsed === 'string') {
      try {
        resolve(JSON.parse(preParsed));
      } catch {
        resolve({});
      }
      return;
    }
    const chunks: Buffer[] = [];
    req.on('data', (chunk: Buffer) => chunks.push(chunk));
    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf-8');
        resolve(raw ? JSON.parse(raw) : {});
      } catch {
        resolve({});
      }
    });
    req.on('error', () => resolve({}));
  });
}

export default async function handler(req: IncomingMessage, res: ServerResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    res.end();
    return;
  }

  const publicTunnelUrl = (
    process.env.PUBLIC_TUNNEL_URL ||
    process.env.APP_URL ||
    process.env.VITE_APP_URL ||
    ''
  )
    .trim()
    .replace(/\/+$/, '');

  if (req.method === 'PUT' || req.method === 'POST') {
    const body = await readBody(req);
    const updatedAt = Math.max(Date.now(), Number(body?.updatedAt) || 0);
    res.statusCode = 200;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.end(
      JSON.stringify({
        ok: true,
        serverless: true,
        state: {
          updatedAt,
          isPersisted: false,
        },
      })
    );
    return;
  }

  const reqUrl = new URL(req.url || '/api/bazar-data', 'http://localhost');
  const since = Number(reqUrl.searchParams.get('since') || 0);

  res.statusCode = 200;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(
    JSON.stringify({
      ok: true,
      serverless: true,
      isPersisted: false,
      notModified: since > 0,
      updatedAt: 0,
      ...(publicTunnelUrl ? { publicTunnelUrl } : {}),
    })
  );
}
