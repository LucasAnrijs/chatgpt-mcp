import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import { createRemoteJWKSet, jwtVerify } from 'jose';

const {
  PORT = '3000',
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  DRIVE_ID,
  FOLDER_ITEM_ID, // optional
  MCP_API_KEY,
  OAUTH_AUDIENCE, // optional override (defaults to CLIENT_ID)
  PUBLIC_BASE_URL, // optional; used for absolute download links
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !DRIVE_ID || !MCP_API_KEY) {
  throw new Error('Missing required env vars (TENANT_ID, CLIENT_ID, CLIENT_SECRET, DRIVE_ID, MCP_API_KEY)');
}

// Resolve public base URL for links returned to clients
const RESOLVED_BASE_URL = PUBLIC_BASE_URL || `http://localhost:${PORT}`;

// --- OAuth token cache ---
let cachedToken: { token: string; exp: number } | null = null;

async function getGraphToken(): Promise<string> {
  const now = Date.now() / 1000;
  if (cachedToken && cachedToken.exp - now > 60) return cachedToken.token;

  const form = new URLSearchParams({
    client_id: CLIENT_ID!,
    client_secret: CLIENT_SECRET!,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const resp = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: form,
  });

  if (!resp.ok) throw new Error(`Token error: ${resp.status} ${await resp.text()}`);

  const data = (await resp.json()) as any;
  cachedToken = { token: data.access_token, exp: Math.floor(Date.now() / 1000) + (data.expires_in || 3600) };
  return cachedToken.token;
}

// --- Graph helper ---
async function g<T>(path: string, init?: RequestInit & { raw?: boolean }): Promise<T> {
  const token = await getGraphToken();
  const fullPath = path.startsWith('/') ? path : `/${path}`;
  const url = `https://graph.microsoft.com/v1.0${fullPath}`;
  console.log(`[Graph] Calling: ${url}`);
  const resp = await fetch(url, {
    ...init,
    headers: { authorization: `Bearer ${token}`, accept: 'application/json', ...(init?.headers || {}) },
  });

  if (!resp.ok) {
    let errBody: any;
    try { errBody = await resp.json(); } catch { errBody = await resp.text(); }
    throw new Error(`Graph ${resp.status} ${resp.statusText}: ${typeof errBody === 'string' ? errBody : JSON.stringify(errBody)}`);
  }

  if ((init as any)?.raw) return resp as any;
  return resp.json() as Promise<T>;
}

// --- Optional folder scope ---
let SEARCH_ROOT_ITEM_ID: string | null = FOLDER_ITEM_ID || null;

// --- Graph wrappers ---
const Graph = {
  async search(q: string, top = 20) {
    const safe = q.replace(/'/g, "''");
    const base = SEARCH_ROOT_ITEM_ID
      ? `/drives/${DRIVE_ID}/items/${SEARCH_ROOT_ITEM_ID}/search(q='${safe}')`
      : `/drives/${DRIVE_ID}/root/search(q='${safe}')`;
    const url = `${base}?$top=${top}`;
    console.log(`[Graph] Search URL: ${url}`);
    return g<any>(url);
  },
  async getItemById(itemId: string) {
    return g<any>(`/drives/${DRIVE_ID}/items/${itemId}`);
  },
  async downloadById(itemId: string): Promise<globalThis.Response> {
    return g<globalThis.Response>(`/drives/${DRIVE_ID}/items/${itemId}/content`, { raw: true, method: 'GET' });
  },
};

// --- Tool handlers (MCP) ---
const toolHandlers: Record<string, (args: any) => Promise<any>> = {
  // Return stable IDs via structuredContent; no Graph URLs here
  search: async ({ query, top }) => {
    console.log('[DEBUG] search tool invoked', { query, top });
    try {
      const res = await Graph.search(query, top ?? 20);
      console.log('[DEBUG] Graph.search returned', res?.value?.length);
      const items = (res?.value ?? []).map((it: any) => ({
        id: it.id,
        name: it.name,
        mimeType: it.file?.mimeType,
      }));
      return {
        content: [{ type: 'text', text: `Found ${items.length} item(s).` }],
        structuredContent: { items },
      };
    } catch (e: any) {
      console.error('[ERROR] search tool failed', e);
      return {
        content: [{ type: 'text', text: `Search failed: ${e?.message || e}` }],
        structuredContent: { items: [] },
      };
    }
  },

  // Accept either { id } or { ids: string[] }
  fetch: async ({ id, ids }: { id?: string; ids?: string[] }) => {
    const targetIds = (Array.isArray(ids) && ids.length > 0) ? ids : (id ? [id] : []);
    if (targetIds.length === 0) {
      return { content: [{ type: 'text', text: 'No id(s) provided' }] };
    }

    const content: any[] = [];

    for (const curId of targetIds) {
      try {
        const meta = await Graph.getItemById(curId);
        if (!meta.file) {
          content.push({ type: 'text', text: `Not a file: ${meta.name || curId}` });
          continue;
        }

        const size = meta.size as number;
        const mime = meta.file.mimeType as string | undefined;
        const isText = !!(mime && (mime.startsWith('text/') || ['application/json', 'application/xml', 'application/javascript'].includes(mime)));
        const under1MB = typeof size === 'number' ? size < 1_000_000 : false;

        if (isText && under1MB) {
          const resp = await Graph.downloadById(curId);
          const text = await resp.text();
          content.push({ type: 'text', text: `# ${meta.name}\n(mime: ${mime}, size: ${size}B)` });
          content.push({ type: 'text', text });
          content.push({ type: 'resource', resource: { uri: `sp://${curId}`, title: meta.name, mimeType: mime, text } });
        } else {
          content.push({ type: 'text', text: `Large/binary file: ${meta.name} (mime: ${mime}, size: ${size}B)` });
          content.push({
            type: 'resource_link',
            uri: `${RESOLVED_BASE_URL}/download/${curId}`,
            name: meta.name,
            mimeType: mime,
            description: `Download ${meta.name}`,
          });
        }
      } catch (e: any) {
        console.error(`[ERROR] fetch failed for ${curId}`, e);
        content.push({ type: 'text', text: `Fetch failed for ${curId}: ${e?.message || e}` });
      }
    }

    return { content };
  },
};

// --- Express wiring ---
const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(cors({
  origin: ['https://chat.openai.com', 'https://chatgpt.com'],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Authorization', 'Content-Type', 'Accept', 'MCP-Protocol-Version', 'Origin'],
  exposedHeaders: ['WWW-Authenticate'],
}));

// PUBLIC endpoints (mount BEFORE auth)
app.get('/', (_req, res) => {
  res.json({ status: 'OK' });
});

app.get('/health', (req, res) => {
  res.json({ ok: true, accept: req.headers.accept, contentType: req.headers['content-type'] });
});

app.get('/version', (_req, res) => {
  res.json({ name: 'sharepoint-drive-connector', version: '1.2.0' });
});

// OAuth Authorization Server metadata (kept for completeness)
app.get('/.well-known/oauth-authorization-server', (_req, res) => {
  const base = `https://login.microsoftonline.com/${TENANT_ID}`;
  res.json({
    issuer: base,
    authorization_endpoint: `${base}/oauth2/v2.0/authorize`,
    token_endpoint: `${base}/oauth2/v2.0/token`,
    jwks_uri: `${base}/discovery/v2.0/keys`,
    token_endpoint_auth_methods_supported: [
      'client_secret_post',
      'client_secret_basic',
    ],
  });
});

// OAuth Protected Resource Metadata (RFC 9728) — consulted by MCP clients
app.get('/.well-known/oauth-protected-resource', (req, res) => {
  const resource = OAUTH_AUDIENCE || CLIENT_ID!; // expected audience your server validates
  res.json({
    resource,
    authorization_servers: [
      `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
    ],
    bearer_methods_supported: ['header'],
    scopes_supported: ['mcp:tools:search', 'mcp:tools:fetch'],
    resource_documentation: `${RESOLVED_BASE_URL}/docs`,
  });
});

// Optional: simple docs placeholder
app.get('/docs', (_req, res) => {
  res.type('text/plain').send('SharePoint MCP Connector: implements /mcp with tools search+fetch.');
});

// Auth middleware for MCP endpoint
const jwks = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`));
const acceptedIssuers = [
  `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
  `https://sts.windows.net/${TENANT_ID}/`,
];
const expectedAudience = OAUTH_AUDIENCE || CLIENT_ID!;

async function auth(req: Request, res: Response, next: NextFunction) {
  try {
    const hdr = req.header('authorization') || '';
    const bearer = hdr.startsWith('Bearer ') ? hdr.slice(7) : '';
    if (!bearer) {
      return res
        .status(401)
        .set('WWW-Authenticate', `Bearer resource_metadata="${RESOLVED_BASE_URL}/.well-known/oauth-protected-resource"`)
        .json({ error: 'Missing Authorization header' });
    }

    // Back-compat: allow static key as bearer
    if (bearer === MCP_API_KEY) return next();

    // Verify JWT issued by Microsoft Entra ID
    const { payload } = await jwtVerify(bearer, jwks, {
      issuer: acceptedIssuers,
      audience: expectedAudience,
    });

    (req as any).oauth = { sub: payload.sub, appid: (payload as any).appid, roles: (payload as any).roles };
    return next();
  } catch (err: any) {
    return res
      .status(401)
      .set('WWW-Authenticate', `Bearer resource_metadata="${RESOLVED_BASE_URL}/.well-known/oauth-protected-resource"`)
      .json({ error: 'Unauthorized', details: err?.message });
  }
}

// MCP endpoint - Handle batch requests with manual JSON-RPC processing
app.post('/mcp', auth, async (req: Request, res: Response) => {
  try {
    console.log('[MCP] Processing request');

    const requests = Array.isArray(req.body) ? req.body : [req.body];
    const responses: any[] = [];

    let initialized = false;

    for (const request of requests) {
      console.log(`[MCP] Processing method: ${request.method}`);

      if (request.method === 'initialize') {
        if (initialized) {
          responses.push({
            jsonrpc: '2.0',
            error: { code: -32600, message: 'Already initialized' },
            id: request.id,
          });
        } else {
          initialized = true;
          const requested: string | undefined = request?.params?.protocolVersion;
          const supported = new Set(['2024-11-05', '2025-03-26', '2025-06-18']);
          const chosen = requested && supported.has(requested) ? requested : '2025-06-18';

          responses.push({
            jsonrpc: '2.0',
            result: {
              protocolVersion: chosen,
              capabilities: {
                tools: { listChanged: true },
                resources: { listChanged: false, subscribe: false },
              },
              serverInfo: { name: 'sharepoint-drive-connector', version: '1.2.0' },
            },
            id: request.id,
          });
        }
      } else if (request.method === 'tools/list') {
        if (!initialized) {
          responses.push({
            jsonrpc: '2.0',
            error: { code: -32600, message: 'Not initialized' },
            id: request.id,
          });
        } else {
          responses.push({
            jsonrpc: '2.0',
            result: {
              tools: [
                {
                  name: 'search',
                  title: 'Search SharePoint drive',
                  description: SEARCH_ROOT_ITEM_ID
                    ? 'Search within the specified folder only'
                    : 'Search the entire document library (drive root)',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      query: { type: 'string', minLength: 1 },
                      top: { type: 'integer', minimum: 1, maximum: 50 },
                    },
                    required: ['query'],
                  },
                  outputSchema: {
                    type: 'object',
                    properties: {
                      items: {
                        type: 'array',
                        items: {
                          type: 'object',
                          properties: {
                            id: { type: 'string' },
                            name: { type: 'string' },
                            mimeType: { type: 'string' },
                          },
                          required: ['id', 'name'],
                        },
                      },
                    },
                    required: ['items'],
                  },
                },
                {
                  name: 'fetch',
                  title: 'Fetch file(s) by item id',
                  description: 'Return inline content for small text; otherwise a proxy download link',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      id: { type: 'string', minLength: 1 },
                      ids: { type: 'array', items: { type: 'string', minLength: 1 } },
                    },
                  },
                },
              ],
            },
            id: request.id,
          });
        }
      } else if (request.method === 'tools/call') {
        if (!initialized) {
          responses.push({
            jsonrpc: '2.0',
            error: { code: -32600, message: 'Not initialized' },
            id: request.id,
          });
        } else {
          try {
            const { name, arguments: args } = request.params;
            console.log(`[MCP] Calling tool: ${name} with args:`, args);
            const handler = toolHandlers[name];
            if (!handler) {
              responses.push({
                jsonrpc: '2.0',
                error: { code: -32601, message: `Tool not found: ${name}` },
                id: request.id,
              });
            } else {
              const result = await handler(args);
              console.log(`[MCP] tools/call handler result for ${name}:`, JSON.stringify(result, null, 2));
              responses.push({ jsonrpc: '2.0', result, id: request.id });
            }
          } catch (err: any) {
            console.error('[MCP] Tool execution error:', err);
            responses.push({
              jsonrpc: '2.0',
              error: { code: -32603, message: err.message },
              id: request.id,
            });
          }
        }
      } else {
        responses.push({
          jsonrpc: '2.0',
          error: { code: -32601, message: `Method not found: ${request.method}` },
          id: request.id,
        });
      }
    }

    if (Array.isArray(req.body)) {
      res.json(responses);
    } else {
      res.json(responses[0]);
    }
  } catch (err: any) {
    console.error('[MCP] Error:', err);
    res.status(500).json({
      jsonrpc: '2.0',
      error: { code: -32603, message: err.message },
      id: null,
    });
  }
});

// PUBLIC proxy for downloads so clients don\'t need Graph auth
app.get('/download/:id', async (req: Request, res: Response) => {
  try {
    const itemId = req.params.id as string | undefined;
    if (!itemId) {
      return res.status(400).json({ error: 'Missing file ID' });
    }
    const meta = await Graph.getItemById(itemId);
    if (!meta.file) {
      return res.status(400).json({ error: `Not a file: ${meta.name || itemId}` });
    }
    const fileResponse = await Graph.downloadById(itemId);
    const buf = Buffer.from(await fileResponse.arrayBuffer());
    res.setHeader('Content-Type', meta.file.mimeType || 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${meta.name}"`);
    res.send(buf);
  } catch (error: any) {
    console.error('[Proxy] download error:', error);
    res.status(500).json({ error: error.message });
  }
});

// REST endpoints for ChatGPT Actions (kept for convenience/testing)
app.post('/tools/search', async (req: Request, res: Response) => {
  try {
    const query = req.body.q || '';
    console.log(`[Tool] search called with query="${query}"`);
    if (!query) return res.json({ results: [] });

    const searchResult = await Graph.search(query, 20);
    const items = searchResult?.value || [];

    const results = items.map((item: any) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      mimeType: item.file?.mimeType,
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      downloadUrl: `${RESOLVED_BASE_URL}/download/${item.id}`,
    }));

    console.log(`[Tool] search("${query}") -> ${results.length} results`);
    return res.json({ results });
  } catch (error: any) {
    console.error('[Tool] search error:', error);
    return res.status(500).json({ error: error.message });
  }
});

app.post('/tools/fetch', async (req: Request, res: Response) => {
  try {
    const itemId = req.body.id;
    console.log(`[Tool] fetch called with id="${itemId}"`);
    if (!itemId) return res.status(400).json({ error: 'Missing file ID' });

    const meta = await Graph.getItemById(itemId);
    if (!meta.file) return res.status(400).json({ error: `Not a file: ${meta.name || itemId}` });

    const fileResponse = await Graph.downloadById(itemId);
    const fileBuffer = await fileResponse.arrayBuffer();

    res.setHeader('Content-Type', meta.file.mimeType || 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${meta.name}"`);
    res.send(Buffer.from(fileBuffer));
  } catch (error: any) {
    console.error('[Tool] fetch error:', error);
    return res.status(500).json({ error: error.message });
  }
});

// Help endpoint
app.get('/mcp/help', (_req, res) => {
  res.json({
    usage: 'Send batch requests with initialization + method call',
    example: {
      endpoint: 'POST /mcp',
      headers: {
        Authorization: 'Bearer YOUR_API_KEY_OR_OAUTH_TOKEN',
        Accept: 'application/json, text/event-stream',
        'Content-Type': 'application/json',
      },
      body: [
        {
          jsonrpc: '2.0',
          id: '0',
          method: 'initialize',
          params: {
            clientInfo: { name: 'your-client', version: '1.0.0' },
            protocolVersion: '2025-06-18',
            capabilities: {},
          },
        },
        {
          jsonrpc: '2.0',
          id: '1',
          method: 'tools/call',
          params: {
            name: 'search',
            arguments: { query: 'your search query', top: 10 },
          },
        },
      ],
    },
    availableTools: [
      {
        name: 'search',
        description: 'Search SharePoint drive',
        parameters: {
          query: 'Search query (required)',
          top: 'Number of results (optional, max 50)',
        },
      },
      {
        name: 'fetch',
        description: 'Fetch file(s) by item id',
        parameters: {
          id: 'SharePoint item ID (optional if ids provided)',
          ids: 'Array of SharePoint item IDs (optional)',
        },
      },
    ],
  });
});

app.listen(Number(PORT), () => {
  console.log(`MCP server on :${PORT} — scoped to ${FOLDER_ITEM_ID ? `folder ${FOLDER_ITEM_ID}` : 'drive root'}`);
  console.log('Server expects batch requests: [initialize, method_call]');
  console.log('Public base URL:', RESOLVED_BASE_URL);
  console.log('See /mcp/help for usage examples');
});
