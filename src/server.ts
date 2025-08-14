import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';

const {
  PORT = '3000',
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  DRIVE_ID,
  FOLDER_ITEM_ID, // optional
  MCP_API_KEY,
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !DRIVE_ID || !MCP_API_KEY) {
  throw new Error('Missing required env vars (TENANT_ID, CLIENT_ID, CLIENT_SECRET, DRIVE_ID, MCP_API_KEY)');
}

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
  const url = new URL(path, 'https://graph.microsoft.com/v1.0');
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
    const base = SEARCH_ROOT_ITEM_ID
      ? `/drives/${DRIVE_ID}/items/${SEARCH_ROOT_ITEM_ID}/search(q='${q}')`
      : `/drives/${DRIVE_ID}/root/search(q='${q}')`;
    return g<any>(`${base}?$top=${top}`);
  },
  async getItemById(itemId: string) {
    return g<any>(`/drives/${DRIVE_ID}/items/${itemId}`);
  },
  async downloadById(itemId: string): Promise<globalThis.Response> {
    return g<globalThis.Response>(`/drives/${DRIVE_ID}/items/${itemId}/content`, { raw: true, method: 'GET' });
  },
};

// --- MCP server ---
function buildServer() {
  const server = new McpServer({ name: 'sharepoint-drive-connector', version: '1.1.1' });

  server.registerTool(
    'search',
    {
      title: 'Search SharePoint drive',
      description: SEARCH_ROOT_ITEM_ID
        ? 'Search within the specified folder only'
        : 'Search the entire document library (drive root)',
      inputSchema: {
        query: z.string().min(1),
        top: z.number().int().min(1).max(50).optional(),
      },
    },
    async ({ query, top }) => {
      const res = await Graph.search(query, top ?? 20);
      const items = res?.value ?? [];

      const links = items.map((it: any) => ({
        type: 'resource_link' as const,
        uri: `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${it.id}/content`,
        name: it.name,
        mimeType: it.file?.mimeType,
        description: `driveItem ${it.id}`,
      }));

      return {
        content: [
          { type: 'text', text: `Found ${items.length} item(s). Showing up to ${top ?? 20}.` },
          ...links,
        ],
      };
    }
  );

  server.registerTool(
    'fetch',
    {
      title: 'Fetch a file by item id',
      description: 'Return file content inline (if text & small), else a download link',
      inputSchema: { id: z.string().min(1) },
    },
    async ({ id }) => {
      const meta = await Graph.getItemById(id);
      if (!meta.file) return { content: [{ type: 'text', text: `Not a file: ${meta.name || id}` }] };

      const size = meta.size as number;
      const mime = meta.file.mimeType as string | undefined;
      const isText = mime?.startsWith('text/') || ['application/json', 'application/xml', 'application/javascript'].includes(mime || '');
      const under1MB = typeof size === 'number' ? size < 1_000_000 : false;

      if (isText && under1MB) {
        const resp = await Graph.downloadById(id);
        const text = await resp.text();
        return {
          content: [
            { type: 'text', text: `# ${meta.name}\n(mime: ${mime}, size: ${size}B)` },
            { type: 'text', text },
          ],
        };
      }

      return {
        content: [
          { type: 'text', text: `Large/binary file. Returning link.\n(mime: ${mime}, size: ${size}B)` },
          {
            type: 'resource_link',
            uri: `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${id}/content`,
            name: meta.name,
            mimeType: mime,
            description: `Download ${meta.name}`,
          },
        ],
      };
    }
  );

  return server;
}

// --- Express wiring ---
const app = express();
app.use(express.json());

// PUBLIC health/version (mount BEFORE auth)
app.get('/health', (req, res) => {
  res.json({ ok: true, accept: req.headers.accept, contentType: req.headers['content-type'] });
});
app.get('/version', (_req, res) => {
  res.json({ name: 'sharepoint-drive-connector', version: '1.1.1' });
});

// Force sane headers for MCP requests (fix 406)
app.use('/mcp', (req, _res, next) => {
    // Helpful one-line log
    console.log('[MCP] incoming', {
      accept: req.headers.accept,
      type: req.headers['content-type'],
    });
  
    // Ensure the MCP transport-required Accept is present
    req.headers.accept = 'application/json, text/event-stream';
  
    // Ensure Content-Type when a body is posted
    if (!req.headers['content-type']) {
      req.headers['content-type'] = 'application/json';
    }
    next();
});
  
  
  
// Auth AFTER public routes (so /health is open)
function auth(req: Request, res: Response, next: NextFunction) {
  const hdr = req.header('authorization') || '';
  const token = hdr.startsWith('Bearer ') ? hdr.slice(7) : '';
  if (token !== MCP_API_KEY) return res.status(401).json({ error: 'Unauthorized' });
  next();
}
app.use(auth);

// MCP endpoint
app.post('/mcp', async (req: Request, res: Response) => {
    try {
      const server = buildServer();
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => Math.random().toString(36).slice(2),
      });
  
      await server.connect(transport);
  
      const incoming = req.body;
  
      const hasInitialize =
        (Array.isArray(incoming) && incoming.some((m: any) => m?.method === 'initialize')) ||
        (!Array.isArray(incoming) && incoming?.method === 'initialize');
  
      // Build a single payload: [initialize, <original>] to keep the same response stream
      const initFrame = {
        jsonrpc: '2.0',
        id: '__auto__',
        method: 'initialize',
        params: {
          clientInfo: { name: 'auto-init', version: '1.0.0' },
          protocolVersion: '2024-11-05',
          capabilities: {},
        },
      };
  
      const payload = hasInitialize
        ? incoming
        : (Array.isArray(incoming) ? [initFrame, ...incoming] : [initFrame, incoming]);
  
      await transport.handleRequest(req, res, payload);
    } catch (err: any) {
      console.error(err);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: '2.0',
          error: { code: -32603, message: err.message },
          id: null,
        });
      }
    }
});
  

app.listen(Number(PORT), () => {
  console.log(`MCP server on :${PORT} — scoped to ${FOLDER_ITEM_ID ? `folder ${FOLDER_ITEM_ID}` : 'drive root'}`);
});
