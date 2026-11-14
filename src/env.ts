/**
 * Minimal, dependency-free .env loader for local runs.
 *
 * Looks for an env file (explicit path → COPILOT_COST_MGMT_ENV_FILE → ./.env),
 * parses simple KEY=VALUE lines, and sets any key NOT already present in the
 * process environment. Variables passed by the MCP client (its `env` block)
 * therefore always win over the file.
 */

import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";

export function maybeLoadEnvFile(explicit?: string): string | undefined {
  const candidate =
    explicit ?? process.env.COPILOT_COST_MGMT_ENV_FILE ?? resolve(process.cwd(), ".env");
  if (!existsSync(candidate)) return undefined;
  try {
    parseInto(readFileSync(candidate, "utf8"));
    return candidate;
  } catch {
    return undefined;
  }
}

function parseInto(content: string): void {
  for (const raw of content.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq < 0) continue;
    const key = line.slice(0, eq).trim();
    if (!key || key in process.env) continue;
    let val = line.slice(eq + 1).trim();
    if (
      (val.startsWith('"') && val.endsWith('"')) ||
      (val.startsWith("'") && val.endsWith("'"))
    ) {
      val = val.slice(1, -1);
    }
    process.env[key] = val;
  }
}
