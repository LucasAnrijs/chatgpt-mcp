import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import { createRemoteJWKSet, jwtVerify } from 'jose';
import { z } from 'zod';

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

console.log('🔧 Checking environment variables...');
const missingVars = [];
if (!TENANT_ID) missingVars.push('TENANT_ID');
if (!CLIENT_ID) missingVars.push('CLIENT_ID');
if (!CLIENT_SECRET) missingVars.push('CLIENT_SECRET');
if (!DRIVE_ID) missingVars.push('DRIVE_ID');
if (!MCP_API_KEY) missingVars.push('MCP_API_KEY');

if (missingVars.length > 0) {
  console.error('❌ Missing required environment variables:', missingVars.join(', '));
  console.error('💡 Please set these in your Render.com environment settings');
  throw new Error(`Missing required env vars: ${missingVars.join(', ')}`);
}
console.log('✅ All required environment variables found');

// Resolve public base URL for links returned to clients
const RESOLVED_BASE_URL = PUBLIC_BASE_URL || `https://chatgpt-mcp.onrender.com`;

// --- OAuth token cache ---
let cachedToken: { token: string; exp: number } | null = null;

// --- Query cache and deduplication ---
const queryCache = new Map<string, { results: any[]; timestamp: number }>();
const inFlightQueries = new Map<string, Promise<any>>();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

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

// --- Graph helper with retry logic ---
async function g<T>(path: string, init?: RequestInit & { raw?: boolean }, retries = 3): Promise<T> {
  const token = await getGraphToken();
  const fullPath = path.startsWith('/') ? path : `/${path}`;
  const url = `https://graph.microsoft.com/v1.0${fullPath}`;
  
  for (let attempt = 0; attempt <= retries; attempt++) {
    console.log(`[Graph] Calling: ${url} (attempt ${attempt + 1})`);
    
  const resp = await fetch(url, {
    ...init,
    headers: { authorization: `Bearer ${token}`, accept: 'application/json', ...(init?.headers || {}) },
  });

    if (resp.ok) {
      if ((init as any)?.raw) return resp as any;
      return resp.json() as Promise<T>;
    }

    // Handle rate limiting with proper Retry-After
    if (resp.status === 429 && attempt < retries) {
      const retryAfter = resp.headers.get('retry-after');
      const delay = retryAfter 
        ? parseInt(retryAfter) * 1000 
        : Math.min(Math.pow(2, attempt) * 1000 + Math.random() * 1000, 30000); // Jitter + cap at 30s
      console.log(`[Graph] Rate limited (429), retrying after ${delay}ms (attempt ${attempt + 1}/${retries})`);
      await new Promise(resolve => setTimeout(resolve, delay));
      continue;
    }
    
    // Handle server errors
    if (resp.status >= 500 && resp.status < 600 && attempt < retries) {
      const delay = Math.min(Math.pow(2, attempt) * 1000 + Math.random() * 1000, 30000);
      console.log(`[Graph] Server error (${resp.status}), retrying after ${delay}ms`);
      await new Promise(resolve => setTimeout(resolve, delay));
      continue;
    }

    let errBody: any;
    try { errBody = await resp.json(); } catch { errBody = await resp.text(); }
    throw new Error(`Graph ${resp.status} ${resp.statusText}: ${typeof errBody === 'string' ? errBody : JSON.stringify(errBody)}`);
  }

  throw new Error(`Graph request failed after ${retries + 1} attempts`);
}

// --- Optional folder scope ---
let SEARCH_ROOT_ITEM_ID: string | null = FOLDER_ITEM_ID || null;

// --- Graph wrappers ---
const Graph = {
  async search(q: string, top = 25) {
    // Normalize query for caching (lowercase, trim)
    const normalizedQuery = q.toLowerCase().trim();
    const cacheKey = `search:${normalizedQuery}:${top}`;
    
    // Check cache first
    const cached = queryCache.get(cacheKey);
    if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
      console.log(`[Graph] Cache hit for "${normalizedQuery}"`);
      return { value: cached.results };
    }
    
    // Check for in-flight duplicate request
    if (inFlightQueries.has(cacheKey)) {
      console.log(`[Graph] Deduplicating query "${normalizedQuery}"`);
      const result = await inFlightQueries.get(cacheKey)!;
      return result;
    }
    
    // Use drive search (tenant-wide search requires region config)
    const searchPromise = this.driveSearch(normalizedQuery, top);
    inFlightQueries.set(cacheKey, searchPromise);
    
    try {
      const result = await searchPromise;
      
      // Cache successful results
      queryCache.set(cacheKey, {
        results: result.value || [],
        timestamp: Date.now()
      });
      
      return result;
    } finally {
      inFlightQueries.delete(cacheKey);
    }
  },



  async driveSearch(q: string, top = 25) {
    const safe = q.replace(/'/g, "''");
    const base = SEARCH_ROOT_ITEM_ID
      ? `/drives/${DRIVE_ID}/items/${SEARCH_ROOT_ITEM_ID}/search(q='${safe}')`
      : `/drives/${DRIVE_ID}/root/search(q='${safe}')`;
    const url = `${base}?$top=${top}`;
    console.log(`[Graph] Drive search: "${q}" -> ${url}`);
    return g<any>(url);
  },

  async getItemById(itemId: string) {
    return g<any>(`/drives/${DRIVE_ID}/items/${itemId}`);
  },
  
  async downloadById(itemId: string): Promise<globalThis.Response> {
    return g<globalThis.Response>(`/drives/${DRIVE_ID}/items/${itemId}/content`, { raw: true, method: 'GET' });
  },
};

// Simple tool implementations matching Python FastMCP pattern
const mcpTools = {
  async search(query: string): Promise<any[]> {
    console.log(`[MCP] search called with query="${query}"`);
    try {
      if (!query) return [];

      const res = await Graph.search(query, 20);
      const items = res?.value || [];
      
      // Return simple array like Python version
      const results = items.map((item: any) => ({
        id: item.id,
        title: item.name,
        text: `${item.name} | ${item.webUrl || ''} | [Call fetch(id) for full content]`.slice(0, 500),
        url: item.webUrl || '',
      }));

      console.log(`[MCP] search("${query}") -> ${results.length} results`);
      return results;
    } catch (e: any) {
      console.error('[MCP] search error:', e);
      
      // Return user-friendly error for rate limiting
      if (e.message?.includes('429') || e.message?.includes('throttled')) {
        return [{
          id: 'error-throttled',
          title: 'Search Temporarily Limited',
          text: 'Microsoft Graph is currently throttling requests. Please try again in a moment.',
          url: ''
        }];
      }
      
      return [{
        id: 'error-general',
        title: 'Search Error',
        text: `Search failed: ${e.message?.slice(0, 100) || 'Unknown error'}`,
        url: ''
      }];
    }
  },

  async fetch(id: string): Promise<any> {
    console.log(`[MCP] fetch called with id="${id}"`);
    try {
      if (!id) {
        return { id: '', title: 'No ID provided', text: '', url: '', metadata: {} };
      }

      const meta = await Graph.getItemById(id);
      if (!meta.file) {
        return {
          id, 
          title: `${meta.name || id} (not a file)`, 
          text: 'Not a file', 
          url: meta.webUrl || '', 
          metadata: {} 
        };
      }

      // Return metadata text like Python version
      const lines = [
        `Name: ${meta.name}`,
        `Modified: ${meta.lastModifiedDateTime}`,
        `Size: ${meta.size}`,
        `Type: ${meta.file.mimeType}`,
      ];
      const text = lines.join('\n');

      return {
        id,
        title: meta.name,
        text,
        url: meta.webUrl || '',
        metadata: {
          created: meta.createdDateTime,
          modified: meta.lastModifiedDateTime,
          size: meta.size,
          mimeType: meta.file.mimeType,
        }
      };
    } catch (e: any) {
      console.error('[MCP] fetch error:', e);
      return { 
        id, 
        title: `Item ${id} (error)`, 
        text: '', 
        url: '', 
        metadata: { error: e.message } 
      };
    }
  },
};



// --- Express wiring ---
const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(cors({
  origin: ['https://chat.openai.com', 'https://chatgpt.com'],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Authorization', 'Content-Type', 'Accept', 'MCP-Protocol-Version', 'Origin', 'Mcp-Session-Id'],
  exposedHeaders: ['WWW-Authenticate', 'Mcp-Session-Id'],
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

// Simple OAuth config for ChatGPT (no auth required)
app.get('/.well-known/oauth-configuration', (_req, res) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.json({
    authorization_url: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`,
    token_url: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    client_id: CLIENT_ID,
    scope: 'https://graph.microsoft.com/.default',
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
            const handler = mcpTools[name as keyof typeof mcpTools];
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

// FastMCP-compatible SSE endpoint (matches your working Python exactly)
app.get('/mcp/sse', (req, res) => {
  console.log('[FastMCP] SSE connection established');
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  });
  
  // Send initial connection message
  res.write('data: {"type": "connection_established"}\n\n');
  
  // Keep connection alive
  const keepAlive = setInterval(() => {
    res.write(': keepalive\n\n');
  }, 30000);
  
  req.on('close', () => {
    console.log('[FastMCP] SSE connection closed');
    clearInterval(keepAlive);
  });
});

app.post('/mcp/sse', async (req, res) => {
  try {
    console.log('[MCP-SSE] Request:', JSON.stringify(req.body, null, 2));
    
    const { jsonrpc, method, params = {}, id } = req.body;
    
    // Handle JSON-RPC 2.0 MCP protocol
    if (method === 'initialize') {
      const requested = params?.protocolVersion;
      const supported = new Set(['2024-11-05', '2025-03-26', '2025-06-18']);
      const chosen = requested && supported.has(requested) ? requested : '2025-06-18';
      
      return res.json({
        jsonrpc: '2.0',
        result: {
          protocolVersion: chosen,
          capabilities: {
            tools: { listChanged: true },
            resources: { listChanged: false, subscribe: false },
          },
          serverInfo: { name: 'sharepoint-drive-connector', version: '1.2.0' },
        },
        id,
      });
    }
    
    if (method === 'tools/list') {
      return res.json({
        jsonrpc: '2.0',
        result: {
          tools: [
            {
              name: 'search',
              description: 'Search SharePoint documents',
              inputSchema: {
                type: 'object',
                properties: {
                  query: { type: 'string', description: 'Search query' }
                },
                required: ['query']
              }
            },
            {
              name: 'fetch', 
              description: 'Fetch document by ID',
              inputSchema: {
                type: 'object',
                properties: {
                  id: { type: 'string', description: 'Document ID' }
                },
                required: ['id']
              }
            }
          ]
        },
        id,
      });
    }
    
    if (method === 'notifications/initialized') {
      // ChatGPT sends this after initialization - just acknowledge
      console.log('[MCP-SSE] Client initialized notification received');
      return res.status(204).send(); // No content response for notifications
    }
    
    if (method === 'tools/call') {
      const { name, arguments: args } = params;
      console.log(`[MCP-SSE] Tool call: ${name}`, args);
      
      if (name === 'search') {
        const results = await mcpTools.search(args?.query || '');
        console.log(`[MCP-SSE] Search results: ${results.length} items`);
        return res.json({
          jsonrpc: '2.0',
          result: { content: [{ type: 'text', text: `Found ${results.length} results` }], structuredContent: { items: results } },
          id,
        });
      }
      
      if (name === 'fetch') {
        const result = await mcpTools.fetch(args?.id || '');
        console.log(`[MCP-SSE] Fetch result:`, result.title);
        return res.json({
          jsonrpc: '2.0',
          result: { content: [{ type: 'text', text: result.text }] },
          id,
        });
      }
      
      return res.json({
        jsonrpc: '2.0',
        error: { code: -32601, message: `Tool not found: ${name}` },
        id,
      });
    }
    
    return res.json({
      jsonrpc: '2.0',
      error: { code: -32601, message: `Method not found: ${method}` },
      id,
    });
  } catch (err: any) {
    console.error('[MCP-SSE] Error:', err);
    return res.json({
      jsonrpc: '2.0',
      error: { code: -32603, message: err.message },
      id: req.body?.id || null,
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
  console.log('🚀 MCP SharePoint Connector starting...');
  console.log(`📡 Server running on port ${PORT}`);
  console.log(`🔗 Public URL: ${RESOLVED_BASE_URL}`);
  console.log(`📁 Drive ID: ${DRIVE_ID}`);
  console.log(`📂 Folder scope: ${FOLDER_ITEM_ID ? `folder ${FOLDER_ITEM_ID}` : 'drive root'}`);
  console.log(`🔑 Tenant: ${TENANT_ID}`);
  console.log('');
  console.log('Available endpoints:');
  console.log('  🔍 GET  / - Health check');
  console.log('  🤖 GET  /mcp/sse - ChatGPT connector SSE stream');
  console.log('  🤖 POST /mcp/sse - ChatGPT connector MCP calls');
  console.log('  📥 GET  /download/:id - File download proxy');
  console.log('');
  console.log('✅ Server ready for connections!');
}).on('error', (err) => {
  console.error('❌ Server failed to start:', err);
  process.exit(1);
});
