#!/usr/bin/env node
/**
 * Entry point / small CLI.
 *
 *   copilot-cost-management-mcp            run the MCP server over stdio (default)
 *   copilot-cost-management-mcp login      interactive device-code sign-in (persists the token)
 *   copilot-cost-management-mcp logout     remove the cached token
 *   copilot-cost-management-mcp --check    verify credentials + reach the API
 *   copilot-cost-management-mcp --help / --version
 *
 * All server diagnostics go to stderr so stdout stays a clean JSON-RPC channel.
 * A local ./.env (or COPILOT_COST_MGMT_ENV_FILE) is loaded first; env vars
 * already set by the MCP client take precedence.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer, createContext } from "./server.js";
import { SERVER_NAME, SERVER_VERSION } from "./config.js";
import { GraphError } from "./graph.js";
import { NotSignedInError, PersistentDeviceCodeCredential } from "./auth.js";
import { maybeLoadEnvFile } from "./env.js";

function printHelp(): void {
  process.stdout.write(
    `${SERVER_NAME} v${SERVER_VERSION} — MCP server for the Microsoft 365 Copilot Cost Management API\n\n` +
      `Usage:\n` +
      `  copilot-cost-management-mcp            Run the MCP server over stdio (for your MCP client)\n` +
      `  copilot-cost-management-mcp login      Sign in with Microsoft (device code) and cache the token\n` +
      `  copilot-cost-management-mcp logout     Remove the cached token\n` +
      `  copilot-cost-management-mcp --check    Verify credentials and reach the API, then exit\n` +
      `  copilot-cost-management-mcp --help      Show this help\n` +
      `  copilot-cost-management-mcp --version   Print the version\n\n` +
      `Configuration (environment variables, or a .env in the working directory):\n` +
      `  Auth — pick one:\n` +
      `    GRAPH_ACCESS_TOKEN                               raw bearer token\n` +
      `    AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET   app-only (balance + reports)\n` +
      `    AZURE_TENANT_ID + AZURE_CLIENT_ID                         delegated device-code (full access)\n` +
      `  Optional:\n` +
      `    COPILOT_COST_MGMT_READONLY=true      expose only read tools\n` +
      `    COPILOT_COST_MGMT_TOKEN_CACHE=path   token-cache file (device-code)\n` +
      `    COPILOT_COST_MGMT_CACHE_DIR=dir      token-cache directory (device-code)\n` +
      `    GRAPH_API_VERSION, GRAPH_BASE_URL, GRAPH_SCOPE, COPILOT_COST_MGMT_ENV_FILE\n`,
  );
}

async function runServer(): Promise<void> {
  const { server, auth, describeAuth, readOnly } = buildServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  let status = "";
  if (auth.credential instanceof PersistentDeviceCodeCredential) {
    const who = await auth.credential.signedInAs().catch(() => undefined);
    status = who
      ? ` — signed in as ${who}`
      : " — NOT signed in (run `copilot-cost-management-mcp login`)";
  }
  process.stderr.write(
    `[${SERVER_NAME} v${SERVER_VERSION}] ready on stdio — auth: ${describeAuth}${readOnly ? " — READ-ONLY" : ""}${status}\n`,
  );
}

function promptBox(info: { verificationUri: string; userCode: string; message: string }): string {
  const line = "─".repeat(54);
  return (
    `\n${line}\n` +
    `  Sign in to Microsoft to authorize this server\n\n` +
    `  Open:   ${info.verificationUri}\n` +
    `  Code:   ${info.userCode}\n` +
    `${line}\n` +
    // MSAL's own message always carries the URL + code — shown as a fallback.
    `  ${info.message}\n\n` +
    `  Waiting for you to finish signing in…\n`
  );
}

async function runLogin(): Promise<void> {
  const { auth, config } = createContext();
  if (!(auth.credential instanceof PersistentDeviceCodeCredential)) {
    process.stderr.write(
      `Auth mode '${auth.mode}' doesn't use interactive login — nothing to do.\n` +
        `(Interactive login applies to the delegated device-code flow: set AZURE_TENANT_ID + AZURE_CLIENT_ID only.)\n`,
    );
    return;
  }

  process.stderr.write(`${SERVER_NAME} v${SERVER_VERSION} — sign in\n  ${auth.describe}\n`);
  try {
    const result = await auth.credential.login(config.scope, (info) => {
      process.stderr.write(promptBox(info));
    });
    const who = result?.account?.username ?? (await auth.credential.signedInAs());
    process.stderr.write(
      `\n  ✓ Signed in${who ? ` as ${who}` : ""}.\n` +
        `  Token cached at ${auth.credential.cacheFile}\n` +
        `  Your MCP client can now use the tools without prompting (until the refresh token expires).\n`,
    );
  } catch (err) {
    process.stderr.write(`\n  ✗ Sign-in failed: ${err instanceof Error ? err.message : String(err)}\n`);
    process.exitCode = 1;
  }
}

async function runLogout(): Promise<void> {
  const { auth } = createContext();
  if (!(auth.credential instanceof PersistentDeviceCodeCredential)) {
    process.stderr.write(`Auth mode '${auth.mode}' has no cached token to remove.\n`);
    return;
  }
  const removed = await auth.credential.logout();
  process.stderr.write(
    removed
      ? `✓ Signed out — removed ${auth.credential.cacheFile}\n`
      : `Nothing to remove (no cached token at ${auth.credential.cacheFile}).\n`,
  );
}

async function runCheck(): Promise<void> {
  const { client, describeAuth, readOnly, config } = createContext();
  process.stderr.write(`${SERVER_NAME} v${SERVER_VERSION} — configuration check\n`);
  process.stderr.write(`  Graph:   ${config.graphBaseUrl}/${config.apiVersion}/copilot/costManagement\n`);
  process.stderr.write(`  Auth:    ${describeAuth}${readOnly ? " (read-only mode)" : ""}\n`);

  process.stderr.write("\nStep 1/2 — acquiring a Microsoft Graph token…\n");
  try {
    await client.ensureToken();
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

  switch (arg) {
    case "login":
      await runLogin();
      return;
    case "logout":
      await runLogout();
      return;
    case "--check":
    case "check":
    case "doctor":
      await runCheck();
      return;
    default:
      await runServer();
  }
}

main().catch((err) => {
  process.stderr.write(`[fatal] ${err instanceof Error ? err.stack ?? err.message : String(err)}\n`);
  process.exit(1);
});
