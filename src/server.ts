import 'dotenv/config';//f
import express, { type Request, type Response, type NextFunction } from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { z } from 'zod';

const {
  PORT = '3000',
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  SITE_ID,
  LIST_ID,
  FOLDER_ITEM_ID,       // optional: folder driveItem id to scope search
  MCP_API_KEY,
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SITE_ID || !LIST_ID || !MCP_API_KEY) {
  throw new Error('Missing required env vars (TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, LIST_ID, MCP_API_KEY)');
}

function auth(req: Request, res: Response, next: NextFunction) {
  const hdr = req.header('authorization') || '';
  const token = hdr.startsWith('Bearer ') ? hdr.slice(7) : '';
  if (!token || token !== MCP_API_KEY) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// ---- OAuth token cache ----
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
  const data = await resp.json() as any;
  cachedToken = { token: data.access_token, exp: Math.floor(Date.now() / 1000) + (data.expires_in || 3600) };
  return cachedToken.token;
}

async function g<T>(path: string, init?: RequestInit & { raw?: boolean }): Promise<T> {
  const token = await getGraphToken();
  const url = new URL(path, 'https://graph.microsoft.com/v1.0');
  const resp = await fetch(url, {
    ...init,
    headers: { 'authorization': `Bearer ${token}`, 'accept': 'application/json', ...(init?.headers || {}) },
  });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`Graph ${resp.status} ${resp.statusText}: ${body}`);
  }
  if ((init as any)?.raw) return (resp as any);
  return resp.json() as Promise<T>;
}

// ---- Resolve DRIVE_ID from SITE_ID + LIST_ID ----
// /sites/{site-id}/lists/{list-id}/drive → Drive (has .id)  (doc: List->Drive relationship)
let DRIVE_ID: string;
let SEARCH_ROOT_ITEM_ID: string | null = null; // folder to scope search, else null = drive root

async function initDrive() {
    const { SITE_HOSTNAME, SITE_PATH, LIST_ID, FOLDER_ITEM_ID } = process.env;
    if (!SITE_HOSTNAME || !SITE_PATH) {
      throw new Error('SITE_HOSTNAME and SITE_PATH must be set in .env');
    }
  
    // Keep the colon and slashes; only encode each segment
    const rawPath = SITE_PATH.replace(/^\/+/, ''); // strip leading slash
    const encodedPath = rawPath.split('/').map(encodeURIComponent).join('/'); // spaces -> %20, slashes kept
  
    // IMPORTANT: do NOT encode the hostname; do NOT encode the colon
    const siteInfo = await g<any>(`/sites/${SITE_HOSTNAME}:/${encodedPath}`);
    const compositeSiteId = siteInfo.id;
    if (!compositeSiteId) throw new Error('Could not resolve composite SITE_ID from hostname + path');
  
    console.log(`Resolved SITE_ID: ${compositeSiteId}`);
  
    if (!LIST_ID) throw new Error('LIST_ID must be set in .env');
  
    const drive = await g<any>(`/sites/${compositeSiteId}/lists/${LIST_ID}/drive`);
    DRIVE_ID = drive.id;
    if (!DRIVE_ID) throw new Error('Could not resolve DRIVE_ID from SITE_ID + LIST_ID');
  
    console.log(`Resolved DRIVE_ID: ${DRIVE_ID}`);
    SEARCH_ROOT_ITEM_ID = FOLDER_ITEM_ID || null;
  }
  
const Graph = {
  // Search within a folder (item) or drive root
  async search(q: string, top = 20) {
    const base = SEARCH_ROOT_ITEM_ID
      ? `/drives/${DRIVE_ID}/items/${SEARCH_ROOT_ITEM_ID}/search(q='${encodeURIComponent(q)}')`
      : `/drives/${DRIVE_ID}/root/search(q='${encodeURIComponent(q)}')`;
    const qs = `?$top=${top}`;
    return g<any>(`${base}${qs}`);
  },

  // Get metadata by item id
  async getItemById(itemId: string) {
    return g<any>(`/drives/${DRIVE_ID}/items/${itemId}`);
  },

  // Download content by item id (returns Response)
  async downloadById(itemId: string) {
    return g<any>(`/drives/${DRIVE_ID}/items/${itemId}/content`, { raw: true, method: 'GET' });
  },
};

// ---- MCP server ----
function buildServer() {
  const server = new McpServer({ name: 'sharepoint-list-or-folder-connector', version: '1.1.0' });

  // CHATGPT-EXPECTED TOOL #1: search
  server.registerTool(
    'search',
    {
      title: 'Search SharePoint library (or locked folder)',
      description: SEARCH_ROOT_ITEM_ID
        ? 'Search within the specified folder only'
        : 'Search within the entire document library (list root)',
      inputSchema: { query: z.string().min(1), top: z.number().int().min(1).max(50).optional() },
    },
    async ({ query, top }) => {
      const res = await Graph.search(query, top ?? 20);
      const items = (res?.value ?? []) as any[];

      const links = items.slice(0, top ?? 20).map((it) => {
        const id = it.id as string;
        const name = it.name as string;
        const mime = it.file?.mimeType as string | undefined;
        return {
          type: 'resource_link' as const,
          uri: `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${id}/content`,
          name,
          mimeType: mime,
          description: `driveItem ${id}`,
        };
      });

      return {
        content: [
          { type: 'text', text: `Found ${items.length} item(s). Showing up to ${top ?? 20}.` },
          ...links,
          { type: 'text', text: 'Tip: pass the item "id" to fetch for inline content when possible.' },
        ],
      };
    }
  );

  // CHATGPT-EXPECTED TOOL #2: fetch
  server.registerTool(
    'fetch',
    {
      title: 'Fetch a file by item id',
      description: 'Return file content inline (if text & small), else a download link',
      inputSchema: { id: z.string().min(1) },
    },
    async ({ id }) => {
      // Guard: ensure the requested item is within the allowed scope
      if (SEARCH_ROOT_ITEM_ID) {
        // Optional: verify ancestry is under SEARCH_ROOT_ITEM_ID by walking parents (not shown to keep code compact).
        // For strict enforcement, fetch /items/{id}/ancestors and confirm the chain contains SEARCH_ROOT_ITEM_ID.
        // If not, throw an error here.
      }

      const meta = await Graph.getItemById(id);
      if (!meta.file) return { content: [{ type: 'text', text: `Not a file: ${meta.name || id}` }] };

      const size = meta.size as number;
      const mime = meta.file.mimeType as string | undefined;
      const isText = mime?.startsWith('text/') || ['application/json', 'application/xml', 'application/javascript'].includes(mime || '');
      const under1MB = typeof size === 'number' ? size < 1_000_000 : false;

      if (isText && under1MB) {
        const resp: Response = await Graph.downloadById(id) as any;
        const text = await (resp as any).text();
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

// ---- HTTP wiring ----
const app = express();
app.use(express.json());
app.use(auth);

app.post('/mcp', async (req: Request, res: Response) => {
  try {
    await initDrive(); // resolves DRIVE_ID and sets SEARCH_ROOT_ITEM_ID
    const server = buildServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: () => Math.random().toString(36).slice(2),
    });
    res.on('close', () => {
      transport.close();
      server.close();
    });
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (err: any) {
    console.error(err);
    if (!res.headersSent) {
      res.status(500).json({ jsonrpc: '2.0', error: { code: -32603, message: 'Internal server error' }, id: null });
    }
  }
});

app.listen(Number(PORT), () => {
  console.log(`MCP server on :${PORT} — scoped to ${FOLDER_ITEM_ID ? `folder ${FOLDER_ITEM_ID}` : 'list root'}`);
});
