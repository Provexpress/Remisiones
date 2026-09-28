import type { IncomingMessage, ServerResponse } from 'node:http';

const ERP_BASE = 'http://152.200.146.226:50010';

export default async function handler(req: any, res: any) {
  // Manejar CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const path = req.query.path || req.url?.replace(/^\/api\/erp-proxy/, '') || '';
  const targetUrl = `${ERP_BASE}${path.startsWith('/') ? path : `/${path}`}`;

  try {
    const fetchOptions: RequestInit = {
      method: req.method,
      headers: {
        'Content-Type': req.headers['content-type'] || 'application/json',
      },
    };

    if (req.headers.authorization) {
      (fetchOptions.headers as any)['Authorization'] = req.headers.authorization;
    }

    if (req.method === 'POST' && req.body) {
      fetchOptions.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
    }

    const response = await fetch(targetUrl, fetchOptions);
    const data = await response.text();

    res.status(response.status);
    res.setHeader('Content-Type', response.headers.get('content-type') || 'application/json');
    return res.send(data);
  } catch (error: any) {
    return res.status(500).json({ error: error.message || 'Error proxying to ERP API' });
  }
}
