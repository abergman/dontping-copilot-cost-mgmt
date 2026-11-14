/**
 * Shared helpers for tool modules: MCP result formatting, error mapping,
 * ETag resolution, and reference-data constants.
 */

import { z } from "zod";
import { GraphClient, GraphError, type GraphResponse } from "../graph.js";

/**
 * Minimal shape of an MCP tool result (text content + optional error flag).
 * The index signature matches the SDK's CallToolResult contract.
 */
export interface ToolResult {
  content: { type: "text"; text: string }[];
  isError?: boolean;
  structuredContent?: Record<string, unknown>;
  [key: string]: unknown;
}

/** Wrap arbitrary JSON data as a successful MCP tool result. */
export function ok(data: unknown, note?: string): ToolResult {
  const body = typeof data === "string" ? data : JSON.stringify(data, null, 2);
  const text = note ? `${note}\n\n${body}` : body;
  const result: ToolResult = { content: [{ type: "text", text }] };
  if (data && typeof data === "object" && !Array.isArray(data)) {
    result.structuredContent = data as Record<string, unknown>;
  }
  return result;
}

/** Format an error (Graph or otherwise) as an MCP error result. */
export function fail(err: unknown): ToolResult {
  let payload: unknown;
  if (err instanceof GraphError) {
    payload = err.toJSON();
  } else if (err instanceof Error) {
    payload = { error: err.message };
  } else {
    payload = { error: String(err) };
  }
  return {
    content: [{ type: "text", text: JSON.stringify(payload, null, 2) }],
    isError: true,
  };
}

/** Run a handler and convert any thrown error into a well-formed error result. */
export async function handle(fn: () => Promise<ToolResult>): Promise<ToolResult> {
  try {
    return await fn();
  } catch (err) {
    return fail(err);
  }
}

/**
 * Return the ETag to use for a mutating request: the caller-supplied value, or
 * — when omitted — the resource's current ETag fetched via GET. The subPath is
 * the cost-management sub-path of the resource whose ETag governs concurrency.
 */
export async function etagFor(
  client: GraphClient,
  subPath: string,
  provided?: string,
): Promise<string | undefined> {
  if (provided && provided.trim()) return provided.trim();
  return client.resolveEtag(subPath);
}

/** Pretty-print a GraphResponse, folding in the ETag when present. */
export function respond(res: GraphResponse, note?: string): ToolResult {
  if (res.status === 204 || res.data === undefined) {
    return ok({ status: res.status, ok: true, etag: res.etag }, note ?? "Operation succeeded (no content).");
  }
  if (res.etag && res.data && typeof res.data === "object" && !("@odata.etag" in res.data)) {
    return ok({ ...(res.data as object), "@odata.etag": res.etag }, note);
  }
  return ok(res.data, note);
}

// ── Reference data (beta; known values — the API may add more) ────────────────

/** Known governed service identifiers. Free-form string is still accepted. */
export const KNOWN_SERVICE_IDS = ["cowork", "workIq"] as const;

/** Known billing methods. */
export const KNOWN_BILLING_METHODS = ["payAsYouGo", "prepaid"] as const;

/** The only consumption unit the API emits today. */
export const CREDIT_UNIT = "copilotCredit";

// ── Common zod fragments ──────────────────────────────────────────────────────

export const guidParam = z
  .string()
  .min(1)
  .describe("Entra object ID (GUID). Not a UPN/email.");

export const etagParam = z
  .string()
  .optional()
  .describe(
    "ETag for optimistic concurrency (If-Match). If omitted, the server reads the " +
      "resource's current ETag first — convenient, but loses the stale-write guard.",
  );

export const completedDaySelectors = {
  startDate: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/)
    .describe("Inclusive start, YYYY-MM-DD (UTC)."),
  endDate: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/)
    .describe("Exclusive end, YYYY-MM-DD. Must equal startDate + 1 day; not in the future."),
};

export const pagingSelectors = {
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .optional()
    .describe("Page size ($top), 1–100 (default 10 for usage records)."),
  skipToken: z
    .string()
    .optional()
    .describe("Server-driven paging token ($skiptoken) from a prior page's nextLink."),
  nextLink: z
    .string()
    .optional()
    .describe("A full @odata.nextLink URL to follow unchanged (overrides other selectors)."),
  orderByConsumedQuantity: z
    .enum(["asc", "desc"])
    .optional()
    .describe("Order by consumedQuantity (the only orderable field)."),
};
