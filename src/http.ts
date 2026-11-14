/**
 * Streamable-HTTP server mode.
 *
 * Runs the same MCP server over HTTP (JSON-RPC 2.0 at POST/GET/DELETE /mcp) so
 * it can be reached by Microsoft Copilot (Cowork / Copilot Studio). Those hosts
 * require a public HTTPS URL; for local use, expose this local HTTP port with a
 * dev tunnel (which terminates TLS) — see the README.
 *
 * Sessions: a fresh McpServer is created per MCP session (keyed by the
 * Mcp-Session-Id header). Tenant registry and cached tokens are disk-backed, so
 * sessions share them.
 *
 * NOTE: this endpoint is unauthenticated. Only expose it through a tunnel you
 * trust (ideally one requiring sign-in), and prefer app-only or a pre-completed
 * `login` so no interactive prompt is needed mid-request.
 */

import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { randomUUID } from "node:crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { buildServer } from "./server.js";
import { SERVER_NAME, SERVER_VERSION } from "./config.js";

const MCP_PATH = "/mcp";
const MAX_BODY_BYTES = 4 * 1024 * 1024; // 4 MiB

function readBody(req: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    req.on("data", (c: Buffer) => {
      size += c.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Request body too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      if (!raw) return resolve(undefined);
      try {
        resolve(JSON.parse(raw));
      } catch {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  const payload = JSON.stringify(body);
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(payload);
}

export interface HttpServerHandle {
  port: number;
  host: string;
  url: string;
  close: () => Promise<void>;
}

export function startHttpServer(port: number, host: string): Promise<HttpServerHandle> {
  const transports = new Map<string, StreamableHTTPServerTransport>();

  const httpServer = createServer((req, res) => {
    void handle(req, res).catch((err) => {
      if (!res.headersSent) {
        sendJson(res, 500, {
          jsonrpc: "2.0",
          error: { code: -32603, message: err instanceof Error ? err.message : "Internal error" },
          id: null,
        });
      }
    });
  });

  async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = (req.url ?? "/").split("?")[0];

    if (url === "/healthz" || url === "/") {
      sendJson(res, 200, { name: SERVER_NAME, version: SERVER_VERSION, mcp: MCP_PATH });
      return;
    }
    if (url !== MCP_PATH) {
      sendJson(res, 404, { error: "Not found", mcp: MCP_PATH });
      return;
    }

    const sessionId = req.headers["mcp-session-id"] as string | undefined;

    if (req.method === "POST") {
      const body = await readBody(req);
      let transport = sessionId ? transports.get(sessionId) : undefined;

      if (!transport) {
        if (sessionId || !isInitializeRequest(body)) {
          sendJson(res, 400, {
            jsonrpc: "2.0",
            error: { code: -32000, message: "Bad Request: no valid session, and not an initialize request." },
            id: null,
          });
          return;
        }
        // New session: fresh server + transport.
        const built = buildServer();
        const created: StreamableHTTPServerTransport = new StreamableHTTPServerTransport({
          sessionIdGenerator: () => randomUUID(),
          onsessioninitialized: (id: string) => {
            transports.set(id, created);
          },
        });
        created.onclose = () => {
          if (created.sessionId) transports.delete(created.sessionId);
        };
        await built.server.connect(created);
        transport = created;
      }

      await transport.handleRequest(req, res, body);
      return;
    }

    if (req.method === "GET" || req.method === "DELETE") {
      const transport = sessionId ? transports.get(sessionId) : undefined;
      if (!transport) {
        sendJson(res, 400, {
          jsonrpc: "2.0",
          error: { code: -32000, message: "Missing or invalid Mcp-Session-Id header." },
          id: null,
        });
        return;
      }
      await transport.handleRequest(req, res);
      return;
    }

    sendJson(res, 405, {
      jsonrpc: "2.0",
      error: { code: -32000, message: `Method ${req.method} not allowed.` },
      id: null,
    });
  }

  return new Promise((resolve, reject) => {
    httpServer.once("error", reject);
    const onListening = () => {
      httpServer.off("error", reject);
      const shown = host === "::" || host === "0.0.0.0" ? "localhost" : host;
      resolve({
        port,
        host,
        url: `http://${shown}:${port}${MCP_PATH}`,
        close: () =>
          new Promise<void>((res) => {
            for (const t of transports.values()) void t.close();
            httpServer.close(() => res());
          }),
      });
    };
    // With the default "::", omit the host so Node binds dual-stack IPv6+IPv4
    // and falls back to IPv4 automatically where IPv6 is unavailable. An
    // explicit host is honoured as given.
    if (host === "::") httpServer.listen(port, onListening);
    else httpServer.listen(port, host, onListening);
  });
}
