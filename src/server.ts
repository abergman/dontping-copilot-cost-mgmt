/**
 * Assembles the MCP server: resolves credentials, builds the Graph client, and
 * registers every tool module.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { resolveCredential } from "./auth.js";
import { loadConfig, SERVER_NAME, SERVER_VERSION } from "./config.js";
import { GraphClient } from "./graph.js";
import type { ToolContext } from "./tools/context.js";
import { registerBalanceTools } from "./tools/balances.js";
import { registerPolicyTools } from "./tools/policies.js";
import { registerPolicyConfigTools } from "./tools/policyConfig.js";
import { registerCreditRequestTools } from "./tools/creditRequests.js";
import { registerUsageReportTools } from "./tools/usageReports.js";
import { registerRawTool } from "./tools/raw.js";

export interface BuiltServer {
  server: McpServer;
  describeAuth: string;
  readOnly: boolean;
}

export function buildServer(): BuiltServer {
  const config = loadConfig();
  const auth = resolveCredential();
  const client = new GraphClient(auth.credential, config);

  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    {
      instructions:
        "Tools for the Microsoft 365 Copilot Cost Management API (Microsoft Graph beta, " +
        "/copilot/costManagement). Govern and observe tenant consumption of Copilot/AI CREDITS: " +
        "tenant and per-user balances, spending policies and their services/alerts, credit-request " +
        "approvals, and usage reports by user, group, service, and funding class. Everything is " +
        "credit-denominated — no currency, price, or invoice amounts. IDs are Entra object GUIDs, " +
        "not UPNs/emails. Usage reports need a time selector (a single completed UTC day, a " +
        "calendar month, or a ≤24h timestamp window). Mutations use optimistic concurrency " +
        "(If-Match); the write tools fetch the current ETag when you omit it." +
        (config.readOnly ? " This server is running in READ-ONLY mode; mutating tools are disabled." : ""),
    },
  );

  const ctx: ToolContext = { server, client, readOnly: config.readOnly };

  registerBalanceTools(ctx);
  registerPolicyTools(ctx);
  registerPolicyConfigTools(ctx);
  registerCreditRequestTools(ctx);
  registerUsageReportTools(ctx);
  registerRawTool(ctx);

  return { server, describeAuth: auth.describe, readOnly: config.readOnly };
}
