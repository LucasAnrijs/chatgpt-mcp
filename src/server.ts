import 'dotenv/config';
import express, { type Request, type Response, type NextFunction } from 'express';
import cors from 'cors';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
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
  // Fix: Ensure path starts with / and construct full URL properly
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
    // Escape single quotes per OData specification
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

// --- Store tool handlers separately for manual access ---
const toolHandlers: Record<string, (args: any) => Promise<any>> = {
  search: async ({ query, top }) => {
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
  },
  
  fetch: async ({ id }) => {
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
  },
};

// --- Express wiring ---
const app = express();
app.use(express.json());
app.use(cors({ 
  origin: ['https://chat.openai.com', 'https://chatgpt.com'],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Authorization', 'Content-Type', 'Accept']
}));

// PUBLIC endpoints (mount BEFORE auth)
app.get('/', (_req, res) => {
  res.json({ status: 'OK' });
});

app.get('/health', (req, res) => {
  res.json({ ok: true, accept: req.headers.accept, contentType: req.headers['content-type'] });
});

app.get('/version', (_req, res) => {
  res.json({ name: 'sharepoint-drive-connector', version: '1.1.1' });
});

app.get('/.well-known/oauth-authorization-server', (_req, res) => {
  const base = `https://login.microsoftonline.com/${TENANT_ID}`;
  res.json({
    issuer: base,
    authorization_endpoint: `${base}/oauth2/v2.0/authorize`,
    token_endpoint: `${base}/oauth2/v2.0/token`,
    jwks_uri: `${base}/discovery/v2.0/keys`,
    token_endpoint_auth_methods_supported: [
      "client_secret_post", 
      "client_secret_basic"
    ]
  });
});

app.get('/.well-known/ai-plugin.json', (req, res) => {
  res.json({
    schema_version: "v1",
    name_for_human: "SharePoint Knowledge Base",
    name_for_model: "sharepoint_kb",
    description_for_human: "Search and retrieve files from SharePoint document library",
    description_for_model: "Search for documents in a SharePoint library using keywords and fetch file contents by ID",
    auth: {
      type: "oauth",
      client_url: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`,
      scope: "https://graph.microsoft.com/.default",
      authorization_url: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      authorization_content_type: "application/x-www-form-urlencoded",
      verification_tokens: {
        openai: process.env.OPENAI_VERIFICATION_TOKEN || "verification-token"
      }
    },
    api: {
      type: "openapi",
      url: `${req.protocol}://${req.get('host')}/openapi.json`
    },
    logo_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg/512px-Microsoft_Office_SharePoint_%282019%E2%80%93present%29.svg.png",
    contact_email: "support@example.com",
    legal_info_url: "https://example.com/legal"
  });
});

app.get('/openapi.json', (req, res) => {
  const baseUrl = `${req.protocol}://${req.get('host')}`;
  res.json({
    openapi: "3.0.1",
    info: {
      title: "SharePoint Knowledge Base API",
      description: "Search and retrieve files from SharePoint document library",
      version: "1.1.1"
    },
    servers: [{ url: baseUrl }],
    components: {
      securitySchemes: {
        OAuth2: {
          type: "oauth2",
          flows: {
            clientCredentials: {
              tokenUrl: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
              scopes: {
                "https://graph.microsoft.com/.default": "Access Microsoft Graph"
              }
            }
          }
        }
      }
    },
    security: [{ OAuth2: ["https://graph.microsoft.com/.default"] }],
    paths: {
      "/tools/search": {
        post: {
          operationId: "searchDocuments",
          summary: "Search for documents",
          description: "Search for documents in the SharePoint library using keywords",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    q: {
                      type: "string",
                      description: "Search query"
                    }
                  },
                  required: ["q"]
                }
              }
            }
          },
          responses: {
            "200": {
              description: "Search results",
              content: {
                "application/json": {
                  schema: {
                    type: "object",
                    properties: {
                      results: {
                        type: "array",
                        items: {
                          type: "object",
                          properties: {
                            id: { type: "string" },
                            name: { type: "string" },
                            webUrl: { type: "string" },
                            mimeType: { type: "string" },
                            size: { type: "number" },
                            lastModified: { type: "string" },
                            downloadUrl: { type: "string" }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      },
      "/tools/fetch": {
        post: {
          operationId: "fetchDocument",
          summary: "Download a document",
          description: "Download a document by its SharePoint item ID",
          requestBody: {
            required: true,
            content: {
              "application/json": {
                schema: {
                  type: "object",
                  properties: {
                    id: {
                      type: "string",
                      description: "SharePoint item ID"
                    }
                  },
                  required: ["id"]
                }
              }
            }
          },
          responses: {
            "200": {
              description: "File content",
              content: {
                "application/octet-stream": {
                  schema: {
                    type: "string",
                    format: "binary"
                  }
                }
              }
            }
          }
        }
      }
    }
  });
});

// Auth middleware
function auth(req: Request, res: Response, next: NextFunction) {
  const hdr = req.header('authorization') || '';
  const token = hdr.startsWith('Bearer ') ? hdr.slice(7) : '';
  if (token !== MCP_API_KEY) return res.status(401).json({ error: 'Unauthorized' });
  next();
}

// MCP endpoint - Handle batch requests with manual JSON-RPC processing
app.post('/mcp', auth, async (req: Request, res: Response) => {
  try {
    console.log('[MCP] Processing request');
    
    // Handle both single and batch requests
    const requests = Array.isArray(req.body) ? req.body : [req.body];
    const responses: any[] = [];
    
    let initialized = false;
    
    for (const request of requests) {
      console.log(`[MCP] Processing method: ${request.method}`);
      
      if (request.method === 'initialize') {
        // Handle initialization
        if (initialized) {
          responses.push({
            jsonrpc: '2.0',
            error: { code: -32600, message: 'Already initialized' },
            id: request.id,
          });
        } else {
          initialized = true;
          responses.push({
            jsonrpc: '2.0',
            result: {
              protocolVersion: '2024-11-05',
              capabilities: { tools: { listChanged: true } },
              serverInfo: { name: 'sharepoint-drive-connector', version: '1.1.1' },
            },
            id: request.id,
          });
        }
      } else if (request.method === 'tools/list') {
        // Handle tools list
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
                      top: { type: 'number', minimum: 1, maximum: 50 },
                    },
                    required: ['query'],
                  },
                },
                {
                  name: 'fetch',
                  title: 'Fetch a file by item id',
                  description: 'Return file content inline (if text & small), else a download link',
                  inputSchema: {
                    type: 'object',
                    properties: {
                      id: { type: 'string', minLength: 1 },
                    },
                    required: ['id'],
                  },
                },
              ],
            },
            id: request.id,
          });
        }
      } else if (request.method === 'tools/call') {
        // Handle tool calls
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
            
            // Get the tool handler
            const handler = toolHandlers[name];
            
            if (!handler) {
              responses.push({
                jsonrpc: '2.0',
                error: { code: -32601, message: `Tool not found: ${name}` },
                id: request.id,
              });
            } else {
              // Execute the tool
              const result = await handler(args);
              responses.push({
                jsonrpc: '2.0',
                result,
                id: request.id,
              });
            }
          } catch (err: any) {
            console.error(`[MCP] Tool execution error:`, err);
            responses.push({
              jsonrpc: '2.0',
              error: { code: -32603, message: err.message },
              id: request.id,
            });
          }
        }
      } else {
        // Unknown method
        responses.push({
          jsonrpc: '2.0',
          error: { code: -32601, message: `Method not found: ${request.method}` },
          id: request.id,
        });
      }
    }
    
    // Send response - match the input format
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

// REST endpoints for ChatGPT Actions
app.post('/tools/search', async (req: Request, res: Response) => {
  try {
    const query = req.body.q || '';
    console.log(`[Tool] search called with query="${query}"`);
    
    if (!query) {
      return res.json({ results: [] });
    }
    
    // Search SharePoint
    const searchResult = await Graph.search(query, 20);
    const items = searchResult?.value || [];
    
    // Convert to simple format for ChatGPT
    const results = items.map((item: any) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      mimeType: item.file?.mimeType,
      size: item.size,
      lastModified: item.lastModifiedDateTime,
      downloadUrl: `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${item.id}/content`
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
    
    if (!itemId) {
      return res.status(400).json({ error: 'Missing file ID' });
    }
    
    // Get file metadata
    const meta = await Graph.getItemById(itemId);
    if (!meta.file) {
      return res.status(400).json({ error: `Not a file: ${meta.name || itemId}` });
    }
    
    // Download the file content
    const fileResponse = await Graph.downloadById(itemId);
    const fileBuffer = await fileResponse.arrayBuffer();
    
    // Send file to client
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
        'Authorization': 'Bearer YOUR_API_KEY',
        'Accept': 'application/json, text/event-stream',
        'Content-Type': 'application/json',
      },
      body: [
        {
          jsonrpc: '2.0',
          id: '0',
          method: 'initialize',
          params: {
            clientInfo: { name: 'your-client', version: '1.0.0' },
            protocolVersion: '2024-11-05',
            capabilities: {},
          },
        },
        {
          jsonrpc: '2.0',
          id: '1',
          method: 'tools/call',
          params: {
            name: 'search',
            arguments: {
              query: 'your search query',
              top: 10,
            },
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
        description: 'Fetch a file by item id',
        parameters: {
          id: 'SharePoint item ID (required)',
        },
      },
    ],
  });
});

app.listen(Number(PORT), () => {
  console.log(`MCP server on :${PORT} — scoped to ${FOLDER_ITEM_ID ? `folder ${FOLDER_ITEM_ID}` : 'drive root'}`);
  console.log('Server expects batch requests: [initialize, method_call]');
  console.log('See /mcp/help for usage examples');
});