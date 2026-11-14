#!/usr/bin/env node
/**
 * Entry point / small CLI. Commands operate on the ACTIVE tenant (or a tenant
 * id passed as an argument, which is registered + activated first).
 *
 *   copilot-cost-management-mcp                 run the MCP server over stdio (default)
 *   copilot-cost-management-mcp tenants         list configured tenants
 *   copilot-cost-management-mcp login [tenant]  device-code sign-in; caches the token
 *   copilot-cost-management-mcp logout [tenant] remove the cached token
 *   copilot-cost-management-mcp --check [tenant] verify credentials + reach the API
 *   copilot-cost-management-mcp --help / --version
 *
 * All server diagnostics go to stderr so stdout stays a clean JSON-RPC channel.
 * A local ./.env (or COPILOT_COST_MGMT_ENV_FILE) is loaded first; env vars
 * already set by the MCP client take precedence.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer, createContext, type ServerContext } from "./server.js";
import { SERVER_NAME, SERVER_VERSION } from "./config.js";
import { GraphError } from "./graph.js";
import { NotSignedInError, PersistentDeviceCodeCredential } from "./auth.js";
import { NoActiveTenantError } from "./tenants.js";
import { startHttpServer } from "./http.js";
import { maybeLoadEnvFile } from "./env.js";

/** Read `--flag value` from argv, if present. */
function getFlag(name: string): string | undefined {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

function printHelp(): void {
  process.stdout.write(
    `${SERVER_NAME} v${SERVER_VERSION} — MCP server for the Microsoft 365 Copilot Cost Management API\n\n` +
      `Usage:\n` +
      `  copilot-cost-management-mcp                  Run the MCP server over stdio (for your MCP client)\n` +
      `  copilot-cost-management-mcp http [--port N]  Run as a Streamable-HTTP server (for Copilot Cowork/Studio via a tunnel)\n` +
      `  copilot-cost-management-mcp tenants          List configured tenants and the active one\n` +
      `  copilot-cost-management-mcp login [tenant]   Device-code sign-in (caches the token) for a tenant\n` +
      `  copilot-cost-management-mcp logout [tenant]  Remove a tenant's cached token\n` +
      `  copilot-cost-management-mcp --check [tenant] Verify credentials and reach the API, then exit\n` +
      `  copilot-cost-management-mcp --help / --version\n\n` +
      `Multi-tenant: all tools run against the ACTIVE tenant. Manage tenants from the chat with\n` +
      `list_tenants / add_tenant / use_tenant, or from the CLI (a [tenant] arg is registered + activated).\n\n` +
      `Configuration (environment variables, or a .env in the working directory):\n` +
      `  Shared app:\n` +
      `    AZURE_CLIENT_ID                    app (client) id used for all tenants\n` +
      `    AZURE_CLIENT_SECRET               optional — app-only auth (balance + reports) across tenants\n` +
      `    AZURE_TENANT_ID                   optional — seeds the first/active tenant\n` +
      `    GRAPH_ACCESS_TOKEN               optional — raw bearer token (single tenant)\n` +
      `  Optional:\n` +
      `    COPILOT_COST_MGMT_READONLY=true   expose only read tools\n` +
      `    COPILOT_COST_MGMT_CACHE_DIR=dir   registry + token caches location\n` +
      `    GRAPH_API_VERSION, GRAPH_BASE_URL, GRAPH_SCOPE, COPILOT_COST_MGMT_ENV_FILE\n`,
  );
}

/** Register + activate a tenant id passed on the CLI, if any. */
function selectTenant(ctx: ServerContext, tenantArg?: string): void {
  if (!tenantArg) return;
  ctx.manager.add({ id: tenantArg });
  ctx.manager.use(tenantArg);
}

async function runServer(): Promise<void> {
  const { server, manager } = buildServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  const active = manager.activeEntry();
  let status = "no active tenant (add_tenant to begin)";
  if (active) {
    const resolved = manager.resolvedFor(active.id);
    status = `active tenant ${active.label ?? active.id} — ${resolved.describe}`;
    if (resolved.credential instanceof PersistentDeviceCodeCredential) {
      const who = await resolved.credential.signedInAs().catch(() => undefined);
      status += who ? ` — signed in as ${who}` : " — NOT signed in (start_login)";
    }
  }
  const tenants = manager.list().length;
  process.stderr.write(
    `[${SERVER_NAME} v${SERVER_VERSION}] ready on stdio — ${tenants} tenant(s); ${status}\n`,
  );
}

async function activeTenantStatus(ctx: ServerContext): Promise<string> {
  const active = ctx.manager.activeEntry();
  if (!active) return "no active tenant (use add_tenant / use_tenant)";
  const resolved = ctx.manager.resolvedFor(active.id);
  let s = `active tenant ${active.label ?? active.id} — ${resolved.mode}`;
  if (resolved.credential instanceof PersistentDeviceCodeCredential) {
    const who = await resolved.credential.signedInAs().catch(() => undefined);
    s += who ? `, signed in as ${who}` : " (NOT signed in — run `login`)";
  }
  return s;
}

async function runHttp(): Promise<void> {
  const ctx = createContext();
  const port = Number(getFlag("--port") ?? ctx.config.httpPort);
  const host = getFlag("--host") ?? ctx.config.httpHost;
  const handle = await startHttpServer(port, host);
  const status = await activeTenantStatus(ctx);

  process.stderr.write(
    `[${SERVER_NAME} v${SERVER_VERSION}] Streamable-HTTP server listening\n` +
      `  Local MCP endpoint:  ${handle.url}\n` +
      `  ${status}\n\n` +
      `To use from Microsoft Copilot (Cowork / Studio), expose this over HTTPS with a dev tunnel:\n` +
      `  devtunnel port create <name> -p ${port} --protocol http\n` +
      `  devtunnel host <name>\n` +
      `then use the tunnel's HTTPS URL (ending in ${"/mcp"}) as the connector's mcpServerUrl.\n` +
      `This endpoint is unauthenticated — only expose it through a tunnel you trust. Ctrl+C to stop.\n`,
  );

  const stop = async () => {
    await handle.close();
    process.exit(0);
  };
  process.on("SIGINT", stop);
  process.on("SIGTERM", stop);
}

function promptBox(info: { verificationUri: string; userCode: string; message: string }): string {
  const line = "─".repeat(54);
  return (
    `\n${line}\n` +
    `  Sign in to Microsoft to authorize this server\n\n` +
    `  Open:   ${info.verificationUri}\n` +
    `  Code:   ${info.userCode}\n` +
    `${line}\n` +
    `  ${info.message}\n\n` +
    `  Waiting for you to finish signing in…\n`
  );
}

function runTenants(): void {
  const ctx = createContext();
  const active = ctx.manager.activeId();
  const rows = ctx.manager.list();
  if (rows.length === 0) {
    process.stderr.write("No tenants registered. Add one with AZURE_TENANT_ID, the add_tenant tool, or `login <tenantId>`.\n");
    return;
  }
  process.stderr.write(`Tenants (${rows.length}):\n`);
  for (const t of rows) {
    const mark = t.id === active ? "* " : "  ";
    process.stderr.write(`${mark}${t.label ?? t.id}  [${t.id}]\n`);
  }
  process.stderr.write(`\n* = active\n`);
}

async function runLogin(tenantArg?: string): Promise<void> {
  const ctx = createContext();
  selectTenant(ctx, tenantArg);
  const entry = ctx.manager.activeEntry();
  if (!entry) {
    process.stderr.write("No tenant to sign in to. Pass a tenant id: `login <tenantId>`.\n");
    process.exitCode = 1;
    return;
  }
  const resolved = ctx.manager.activeResolved();
  if (!(resolved.credential instanceof PersistentDeviceCodeCredential)) {
    process.stderr.write(`Active tenant ${entry.id} uses ${resolved.mode} auth — interactive login isn't applicable.\n`);
    return;
  }
  process.stderr.write(`${SERVER_NAME} v${SERVER_VERSION} — sign in to tenant ${entry.label ?? entry.id}\n  ${resolved.describe}\n`);
  try {
    const result = await resolved.credential.login(ctx.config.scope, (info) => {
      process.stderr.write(promptBox(info));
    });
    const who = result?.account?.username ?? (await resolved.credential.signedInAs());
    process.stderr.write(
      `\n  ✓ Signed in${who ? ` as ${who}` : ""} to tenant ${entry.id}.\n` +
        `  Token cached at ${resolved.credential.cacheFile}\n`,
    );
  } catch (err) {
    process.stderr.write(`\n  ✗ Sign-in failed: ${err instanceof Error ? err.message : String(err)}\n`);
    process.exitCode = 1;
  }
}

async function runLogout(tenantArg?: string): Promise<void> {
  const ctx = createContext();
  selectTenant(ctx, tenantArg);
  const entry = ctx.manager.activeEntry();
  if (!entry) {
    process.stderr.write("No tenant selected.\n");
    return;
  }
  const resolved = ctx.manager.activeResolved();
  if (!(resolved.credential instanceof PersistentDeviceCodeCredential)) {
    process.stderr.write(`Tenant ${entry.id} (${resolved.mode}) has no cached token.\n`);
    return;
  }
  const removed = await resolved.credential.logout();
  process.stderr.write(removed ? `✓ Signed out of tenant ${entry.id}.\n` : `No cached token for tenant ${entry.id}.\n`);
}

async function runCheck(tenantArg?: string): Promise<void> {
  const ctx = createContext();
  selectTenant(ctx, tenantArg);
  const entry = ctx.manager.activeEntry();
  process.stderr.write(`${SERVER_NAME} v${SERVER_VERSION} — configuration check\n`);
  process.stderr.write(`  Graph:   ${ctx.config.graphBaseUrl}/${ctx.config.apiVersion}/copilot/costManagement\n`);
  if (!entry) {
    process.stderr.write("  ✗ No tenant configured. Set AZURE_TENANT_ID (+ AZURE_CLIENT_ID), or run `login <tenantId>`.\n");
    process.exitCode = 1;
    return;
  }
  const resolved = ctx.manager.activeResolved();
  process.stderr.write(`  Tenant:  ${entry.label ?? entry.id} (${entry.id})\n`);
  process.stderr.write(`  Auth:    ${resolved.describe}${ctx.config.readOnly ? " (read-only mode)" : ""}\n`);

  process.stderr.write("\nStep 1/2 — acquiring a Microsoft Graph token…\n");
  try {
    await ctx.manager.activeClient().ensureToken();
  } catch (err) {
    if (err instanceof NotSignedInError) {
      process.stderr.write(`  ✗ ${err.message}\n`);
      process.exitCode = 1;
      return;
    }
    throw err;
  }
  process.stderr.write("  ✓ Token acquired.\n");

  process.stderr.write("\nStep 2/2 — calling getTenantCreditBalance…\n");
  try {
    const res = await ctx.manager.activeClient().cmGet("/getTenantCreditBalance");
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
  const arg2 = process.argv[3];

  if (arg === "--version" || arg === "-v") {
    process.stdout.write(`${SERVER_VERSION}\n`);
    return;
  }
  if (arg === "--help" || arg === "-h") {
    printHelp();
    return;
  }
  if (loaded) process.stderr.write(`[config] loaded env from ${loaded}\n`);

  switch (arg) {
    case "http":
    case "serve":
      await runHttp();
      return;
    case "tenants":
      runTenants();
      return;
    case "login":
      await runLogin(arg2);
      return;
    case "logout":
      await runLogout(arg2);
      return;
    case "--check":
    case "check":
    case "doctor":
      await runCheck(arg2);
      return;
    default:
      await runServer();
  }
}

main().catch((err) => {
  if (err instanceof NoActiveTenantError) {
    process.stderr.write(`[error] ${err.message}\n`);
    process.exit(1);
  }
  process.stderr.write(`[fatal] ${err instanceof Error ? err.stack ?? err.message : String(err)}\n`);
  process.exit(1);
});
