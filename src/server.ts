/**
 * Assembles the MCP server: builds the multi-tenant manager and registers every
 * tool module. Tools talk to a GraphClient facade that always targets the
 * currently-active tenant.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { readBaseAuth } from "./auth.js";
import { loadConfig, SERVER_NAME, SERVER_VERSION, type Config } from "./config.js";
import { TenantManager } from "./tenants.js";
import type { ToolContext } from "./tools/context.js";
import { registerTenantTools } from "./tools/tenants.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerBalanceTools } from "./tools/balances.js";
import { registerPolicyTools } from "./tools/policies.js";
import { registerPolicyConfigTools } from "./tools/policyConfig.js";
import { registerCreditRequestTools } from "./tools/creditRequests.js";
import { registerUsageReportTools } from "./tools/usageReports.js";
import { registerRawTool } from "./tools/raw.js";

export interface ServerContext {
  manager: TenantManager;
  config: Config;
}

export interface BuiltServer {
  server: McpServer;
  manager: TenantManager;
  config: Config;
}

/** Resolve config + base auth and build the tenant manager (no server yet). */
export function createContext(): ServerContext {
  const config = loadConfig();
  const base = readBaseAuth();
  const manager = new TenantManager(config, base);
  return { manager, config };
}

export function buildServer(): BuiltServer {
  const { manager, config } = createContext();

  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    {
      instructions:
        "Tools for the Microsoft 365 Copilot Cost Management API (Microsoft Graph beta, " +
        "/copilot/costManagement). Govern and observe tenant consumption of Copilot/AI CREDITS. " +
        "MULTI-TENANT: every tool runs against the ACTIVE tenant — use list_tenants / add_tenant / " +
        "use_tenant to manage and switch, and start_login to sign in to the active tenant " +
        "(device-code). Everything is credit-denominated — no currency or invoice amounts. IDs are " +
        "Entra object GUIDs, not UPNs/emails. Usage reports need a time selector (a single completed " +
        "UTC day, a calendar month, or a ≤24h timestamp window). Mutations use optimistic " +
        "concurrency (If-Match); write tools fetch the current ETag when you omit it." +
        (config.readOnly ? " This server is running in READ-ONLY mode; mutating tools are disabled." : ""),
    },
  );

  const ctx: ToolContext = {
    server,
    client: manager.dynamicClient(),
    manager,
    scope: config.scope,
    readOnly: config.readOnly,
  };

  registerTenantTools(ctx);
  registerAuthTools(ctx);
  registerBalanceTools(ctx);
  registerPolicyTools(ctx);
  registerPolicyConfigTools(ctx);
  registerCreditRequestTools(ctx);
  registerUsageReportTools(ctx);
  registerRawTool(ctx);

  return { server, manager, config };
}
