import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { GraphClient } from "../graph.js";
import type { TenantManager } from "../tenants.js";

/** Everything a tool module needs to register its tools. */
export interface ToolContext {
  server: McpServer;
  /** GraphClient facade that always targets the active tenant. */
  client: GraphClient;
  /** Multi-tenant registry + per-tenant credentials. */
  manager: TenantManager;
  /** OAuth scope requested at sign-in. */
  scope: string;
  /** When true, mutating tools must not be registered. */
  readOnly: boolean;
}

/** A tool module registers one logical group of tools. */
export type ToolModule = (ctx: ToolContext) => void;
