#!/usr/bin/env node
/**
 * Entry point: connect the Copilot Cost Management MCP server over stdio.
 *
 * All diagnostics go to stderr so stdout stays a clean JSON-RPC channel.
 */

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { buildServer } from "./server.js";
import { SERVER_NAME, SERVER_VERSION } from "./config.js";

async function main(): Promise<void> {
  const { server, describeAuth, readOnly } = buildServer();

  const transport = new StdioServerTransport();
  await server.connect(transport);

  process.stderr.write(
    `[${SERVER_NAME} v${SERVER_VERSION}] ready on stdio — auth: ${describeAuth}` +
      `${readOnly ? " — READ-ONLY" : ""}\n`,
  );
}

main().catch((err) => {
  process.stderr.write(`[fatal] ${err instanceof Error ? err.stack ?? err.message : String(err)}\n`);
  process.exit(1);
});
