#!/usr/bin/env node
/**
 * Entry point / small CLI.
 *
 *   copilot-cost-management-mcp            run the MCP server over stdio (default)
 *   copilot-cost-management-mcp --check    verify credentials + reach the API
 *   copilot-cost-management-mcp --help
 *   copilot-cost-management-mcp --version
 *
 * All diagnostics go to stderr so stdout stays a clean JSON-RPC channel when the
 * server runs. A local ./.env (or COPILOT_COST_MGMT_ENV_FILE) is loaded first;
 * env vars already set by the MCP client take precedence.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer, createContext } from "./server.js";
import { SERVER_NAME, SERVER_VERSION } from "./config.js";
import { GraphError } from "./graph.js";
import { maybeLoadEnvFile } from "./env.js";

function printHelp(): void {
  process.stdout.write(
    `${SERVER_NAME} v${SERVER_VERSION} — MCP server for the Microsoft 365 Copilot Cost Management API\n\n` +
      `Usage:\n` +
      `  copilot-cost-management-mcp            Run the MCP server over stdio (for your MCP client)\n` +
      `  copilot-cost-management-mcp --check    Verify credentials and reach the API, then exit\n` +
      `  copilot-cost-management-mcp --help      Show this help\n` +
      `  copilot-cost-management-mcp --version   Print the version\n\n` +
      `Configuration (environment variables, or a .env in the working directory):\n` +
      `  Auth — pick one:\n` +
      `    GRAPH_ACCESS_TOKEN                               raw bearer token\n` +
      `    AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET   app-only (balance + reports)\n` +
      `    AZURE_TENANT_ID + AZURE_CLIENT_ID                         delegated device-code (full access)\n` +
      `  Optional:\n` +
      `    COPILOT_COST_MGMT_READONLY=true    expose only read tools\n` +
      `    GRAPH_API_VERSION, GRAPH_BASE_URL, GRAPH_SCOPE, COPILOT_COST_MGMT_ENV_FILE\n`,
  );
}

async function runServer(): Promise<void> {
  const { server, describeAuth, readOnly } = buildServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write(
    `[${SERVER_NAME} v${SERVER_VERSION}] ready on stdio — auth: ${describeAuth}${readOnly ? " — READ-ONLY" : ""}\n`,
  );
}

async function runCheck(): Promise<void> {
  const { client, describeAuth, readOnly, config } = createContext();
  process.stderr.write(`${SERVER_NAME} v${SERVER_VERSION} — configuration check\n`);
  process.stderr.write(`  Graph:   ${config.graphBaseUrl}/${config.apiVersion}/copilot/costManagement\n`);
  process.stderr.write(`  Auth:    ${describeAuth}${readOnly ? " (read-only mode)" : ""}\n`);

  process.stderr.write("\nStep 1/2 — acquiring a Microsoft Graph token…\n");
  await client.ensureToken();
  process.stderr.write("  ✓ Token acquired.\n");

  process.stderr.write("\nStep 2/2 — calling getTenantCreditBalance…\n");
  try {
    const res = await client.cmGet("/getTenantCreditBalance");
    process.stderr.write("  ✓ API reachable.\n\nTenant credit balance:\n");
    process.stdout.write(`${JSON.stringify(res.data, null, 2)}\n`);
    process.stderr.write("\nAll good — you can wire this server into your MCP client.\n");
  } catch (err) {
    const detail = err instanceof GraphError ? err.toJSON() : { error: String(err) };
    process.stderr.write(
      "  ⚠ The token works, but the API call failed. This is usually a missing API " +
        "permission or directory role on the app/account (not a config error).\n\n",
    );
    process.stdout.write(`${JSON.stringify(detail, null, 2)}\n`);
    process.exitCode = 2;
  }
}

async function main(): Promise<void> {
  const loaded = maybeLoadEnvFile();
  const arg = process.argv[2];

  if (arg === "--version" || arg === "-v") {
    process.stdout.write(`${SERVER_VERSION}\n`);
    return;
  }
  if (arg === "--help" || arg === "-h") {
    printHelp();
    return;
  }
  if (loaded) process.stderr.write(`[config] loaded env from ${loaded}\n`);

  if (arg === "--check" || arg === "check" || arg === "doctor") {
    await runCheck();
    return;
  }
  await runServer();
}

main().catch((err) => {
  process.stderr.write(`[fatal] ${err instanceof Error ? err.stack ?? err.message : String(err)}\n`);
  process.exit(1);
});
