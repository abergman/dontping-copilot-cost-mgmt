/**
 * Escape hatch: issue an arbitrary request against the Copilot Cost Management
 * API root. Useful for following opaque nextLinks, trying newly-added beta
 * endpoints, or sending request shapes the typed tools don't cover yet.
 *
 * Scoped to /copilot/costManagement to keep the server focused; a full
 * graph.microsoft.com URL (e.g. an @odata.nextLink) is also accepted as long as
 * it targets the costManagement path.
 */

import { z } from "zod";
import { GraphClient, type HttpMethod } from "../graph.js";
import type { ToolContext } from "./context.js";
import { handle, respond } from "./helpers.js";

const READ_METHODS: HttpMethod[] = ["GET"];
const WRITE_METHODS: HttpMethod[] = ["POST", "PATCH", "PUT", "DELETE"];

export function registerRawTool({ server, client, readOnly }: ToolContext): void {
  const allowed = readOnly ? READ_METHODS : [...READ_METHODS, ...WRITE_METHODS];

  server.registerTool(
    "cost_management_request",
    {
      title: "Raw Copilot Cost Management request",
      description:
        "Advanced escape hatch: issue a raw request to the Copilot Cost Management API. Provide " +
        "either a sub-path (relative to /copilot/costManagement, e.g. '/spendingPolicies') or a full " +
        "graph.microsoft.com URL that targets the costManagement path (e.g. an @odata.nextLink). " +
        "Use the typed tools when one exists; reach for this to follow nextLinks or call endpoints " +
        "not yet wrapped." +
        (readOnly ? " (Read-only mode: only GET is permitted.)" : ""),
      inputSchema: {
        method: z
          .enum(["GET", "POST", "PATCH", "PUT", "DELETE"])
          .describe("HTTP method."),
        path: z
          .string()
          .min(1)
          .describe(
            "Sub-path under /copilot/costManagement (leading '/' optional), or a full graph URL " +
              "targeting that path.",
          ),
        query: z
          .record(z.union([z.string(), z.number(), z.boolean()]))
          .optional()
          .describe("Query parameters (e.g. { \"$top\": 5 })."),
        body: z.any().optional().describe("JSON request body for POST/PATCH/PUT."),
        ifMatch: z.string().optional().describe("If-Match ETag header for concurrency-guarded writes."),
      },
      annotations: { readOnlyHint: readOnly, openWorldHint: true },
    },
    ({ method, path, query, body, ifMatch }) =>
      handle(async () => {
        if (!allowed.includes(method)) {
          throw new Error(
            `Method ${method} is not permitted (read-only mode). Set COPILOT_COST_MGMT_READONLY=false to enable writes.`,
          );
        }

        const target = resolveTarget(path);
        const res = await client.request(method, target, { query, body, ifMatch });
        return respond(res);
      }),
  );
}

/** Resolve a user-supplied path/URL and guard it to the costManagement root. */
function resolveTarget(input: string): string {
  if (/^https?:\/\//i.test(input)) {
    let url: URL;
    try {
      url = new URL(input);
    } catch {
      throw new Error(`Invalid URL: ${input}`);
    }
    if (!/\/copilot\/costManagement(\/|$)/i.test(url.pathname)) {
      throw new Error("Absolute URLs must target the /copilot/costManagement path.");
    }
    return url.toString();
  }
  const sub = input.startsWith("/") ? input : `/${input}`;
  // Fold an accidental duplicate root (e.g. "/copilot/costManagement/...") so
  // both "/spendingPolicies" and "/copilot/costManagement/spendingPolicies" work.
  const normalized = sub.replace(/^\/copilot\/costManagement/i, "");
  return GraphClient.path(normalized.startsWith("/") ? normalized : `/${normalized}`);
}
