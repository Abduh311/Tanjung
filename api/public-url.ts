import type { IncomingMessage, ServerResponse } from 'node:http';

export default function handler(req: IncomingMessage, res: ServerResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    res.end();
    return;
  }

  const rawUrl = (
    process.env.PUBLIC_TUNNEL_URL ||
    process.env.VITE_APP_URL ||
    (process.env.VERCEL ? '' : process.env.APP_URL) ||
    ''
  )
    .trim()
    .replace(/\/+$/, '');
  const publicTunnelUrl = rawUrl.includes('ais-dev-')
    ? rawUrl.replace('ais-dev-', 'ais-pre-')
    : rawUrl;

  res.statusCode = 200;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(
    JSON.stringify({
      ok: true,
      publicTunnelUrl,
    })
  );
}
