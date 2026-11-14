/**
 * Usage reports — consumption breakdowns (read-only). All live under
 * /copilot/costManagement/usageReports and require a time selector.
 *
 * Completed-UTC-day records (startDate + endDate = startDate+1day):
 *   GET /usageReports/userUsageRecords            (+ /{userId})
 *   GET /usageReports/groupUsageRecords           (+ /{groupId})
 *   GET /usageReports/serviceUsageRecords         (+ /{serviceId})
 * Calendar-month, not paged:
 *   GET /usageReports/fundingUsageRecords         (month=YYYY-MM)
 * Timestamp-window raw detail (≤24h, within one UTC month):
 *   GET /usageReports/userUsageDetails/{userId}
 *   GET /usageReports/userUsageDetails/{userId}/services   (not paged)
 *   GET /usageReports/userUsageDetails/{userId}/sessions   (paged, default 20)
 *
 * Paging: follow @odata.nextLink unchanged ($skiptoken; $skip unsupported).
 * Ordering: only consumedQuantity. Units: copilotCredit.
 */

import { z } from "zod";
import { GraphClient } from "../graph.js";
import type { ToolContext } from "./context.js";
import { completedDaySelectors, guidParam, handle, pagingSelectors, respond } from "./helpers.js";

const ur = (p: string) => GraphClient.path(`/usageReports${p}`);

const isoDateTime = z
  .string()
  .regex(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{1,7})?Z$/)
  .describe("UTC ISO 8601 timestamp, e.g. 2026-11-13T00:00:00Z (up to 7 fractional digits).");

function dayQuery(startDate: string, endDate: string, extra?: Record<string, unknown>) {
  return { startDate, endDate, ...(extra ?? {}) };
}

function orderByClause(dir?: "asc" | "desc"): string | undefined {
  return dir ? `consumedQuantity ${dir}` : undefined;
}

export function registerUsageReportTools({ server, client }: ToolContext): void {
  // ── Per-user, completed UTC day ──────────────────────────────────────────────
  server.registerTool(
    "list_user_usage_records",
    {
      title: "List per-user usage for a day",
      description:
        "List processed consumption per user for exactly one completed UTC day (endDate must equal " +
        "startDate + 1 day). Rankable by consumedQuantity. $top 1–100 (default 10). Requires " +
        "CopilotCostManagement-Reports.Read.All or CopilotCostManagement.Read.All.",
      inputSchema: {
        ...completedDaySelectors,
        ...pagingSelectors,
        filterUserId: guidParam.optional().describe("Filter to one user ($filter userId eq)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ startDate, endDate, top, skipToken, nextLink, orderByConsumedQuantity, filterUserId }) =>
      handle(async () => {
        if (nextLink) return respond(await client.request("GET", nextLink));
        const res = await client.request("GET", ur("/userUsageRecords"), {
          query: dayQuery(startDate, endDate, {
            $top: top,
            $skiptoken: skipToken,
            $orderby: orderByClause(orderByConsumedQuantity),
            $filter: filterUserId ? `userId eq '${filterUserId}'` : undefined,
          }),
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "get_user_usage_record",
    {
      title: "Get one user's usage for a day",
      description:
        "Retrieve one user's processed consumption for one completed UTC day. 404 " +
        "usageDataUnavailable means no represented usage (not proof of zero or of existence).",
      inputSchema: { userId: guidParam, ...completedDaySelectors },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, startDate, endDate }) =>
      handle(async () =>
        respond(
          await client.request("GET", ur(`/userUsageRecords/${encodeURIComponent(userId)}`), {
            query: dayQuery(startDate, endDate),
          }),
        ),
      ),
  );

  // ── Per-group, completed UTC day ─────────────────────────────────────────────
  server.registerTool(
    "list_group_usage_records",
    {
      title: "List per-group usage for a day",
      description:
        "List processed consumption per security group for one completed UTC day. Rankable by " +
        "consumedQuantity. $top 1–100 (default 10). Requires CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        ...completedDaySelectors,
        ...pagingSelectors,
        filterGroupId: guidParam.optional().describe("Filter to one group ($filter groupId eq)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ startDate, endDate, top, skipToken, nextLink, orderByConsumedQuantity, filterGroupId }) =>
      handle(async () => {
        if (nextLink) return respond(await client.request("GET", nextLink));
        const res = await client.request("GET", ur("/groupUsageRecords"), {
          query: dayQuery(startDate, endDate, {
            $top: top,
            $skiptoken: skipToken,
            $orderby: orderByClause(orderByConsumedQuantity),
            $filter: filterGroupId ? `groupId eq '${filterGroupId}'` : undefined,
          }),
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "get_group_usage_record",
    {
      title: "Get one group's usage for a day",
      description: "Retrieve one security group's processed consumption for one completed UTC day.",
      inputSchema: { groupId: guidParam, ...completedDaySelectors },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ groupId, startDate, endDate }) =>
      handle(async () =>
        respond(
          await client.request("GET", ur(`/groupUsageRecords/${encodeURIComponent(groupId)}`), {
            query: dayQuery(startDate, endDate),
          }),
        ),
      ),
  );

  // ── Per-service, completed UTC day ───────────────────────────────────────────
  server.registerTool(
    "list_service_usage_records",
    {
      title: "List per-service usage for a day",
      description:
        "List processed consumption per governed service for one completed UTC day — see where " +
        "credits went across AI experiences. $top 1–100 (default 10). Requires " +
        "CopilotCostManagement-Reports.Read.All or CopilotCostManagement.Read.All.",
      inputSchema: {
        ...completedDaySelectors,
        ...pagingSelectors,
        filterServiceId: z.string().optional().describe("Filter to one service ($filter serviceId eq)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ startDate, endDate, top, skipToken, nextLink, orderByConsumedQuantity, filterServiceId }) =>
      handle(async () => {
        if (nextLink) return respond(await client.request("GET", nextLink));
        const res = await client.request("GET", ur("/serviceUsageRecords"), {
          query: dayQuery(startDate, endDate, {
            $top: top,
            $skiptoken: skipToken,
            $orderby: orderByClause(orderByConsumedQuantity),
            $filter: filterServiceId ? `serviceId eq '${filterServiceId}'` : undefined,
          }),
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "get_service_usage_record",
    {
      title: "Get one service's usage for a day",
      description: "Retrieve one governed service's processed consumption for one completed UTC day.",
      inputSchema: {
        serviceId: z.string().min(1).describe('Governed service id, e.g. "cowork".'),
        ...completedDaySelectors,
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ serviceId, startDate, endDate }) =>
      handle(async () =>
        respond(
          await client.request("GET", ur(`/serviceUsageRecords/${encodeURIComponent(serviceId)}`), {
            query: dayQuery(startDate, endDate),
          }),
        ),
      ),
  );

  // ── Funding class, calendar month ────────────────────────────────────────────
  server.registerTool(
    "list_funding_usage_records",
    {
      title: "List usage by funding class for a month",
      description:
        "List a calendar month's consumption by funding class (prepaid, pay-as-you-go, unbilled " +
        "overage, non-billable). Not paged. Requires CopilotCostManagement-Reports.Read.All or " +
        "CopilotCostManagement.Read.All.",
      inputSchema: {
        month: z
          .string()
          .regex(/^\d{4}-\d{2}$/)
          .describe("Calendar month, YYYY-MM (date ranges are rejected)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ month }) =>
      handle(async () =>
        respond(await client.request("GET", ur("/fundingUsageRecords"), { query: { month } })),
      ),
  );

  // ── Raw per-user detail, timestamp window ────────────────────────────────────
  server.registerTool(
    "get_user_usage_detail",
    {
      title: "Get a user's raw usage detail",
      description:
        "Retrieve one user's raw, billable-flagged usage over a timestamp window (inclusive start, " +
        "exclusive end). The window must be ≤24h, start<end, end not in the future, and must not " +
        "cross a UTC month boundary. Requires CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        userId: guidParam,
        startDateTime: isoDateTime.describe("Inclusive window start (UTC ISO 8601)."),
        endDateTime: isoDateTime.describe("Exclusive window end (UTC ISO 8601)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, startDateTime, endDateTime }) =>
      handle(async () =>
        respond(
          await client.request("GET", ur(`/userUsageDetails/${encodeURIComponent(userId)}`), {
            query: { startDateTime, endDateTime },
          }),
        ),
      ),
  );

  server.registerTool(
    "list_user_usage_service_details",
    {
      title: "List a user's raw usage by service",
      description:
        "List one user's raw consumption broken down by governed service over a timestamp window " +
        "(same window rules as get_user_usage_detail). Not paged. Requires " +
        "CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        userId: guidParam,
        startDateTime: isoDateTime,
        endDateTime: isoDateTime,
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, startDateTime, endDateTime }) =>
      handle(async () =>
        respond(
          await client.request(
            "GET",
            ur(`/userUsageDetails/${encodeURIComponent(userId)}/services`),
            { query: { startDateTime, endDateTime } },
          ),
        ),
      ),
  );

  server.registerTool(
    "list_user_usage_session_details",
    {
      title: "List a user's raw usage by session",
      description:
        "List one user's raw usage grouped by session over a timestamp window (same window rules). " +
        "Paged: default page size 20, max 100; follow nextLink. Requires " +
        "CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        userId: guidParam,
        startDateTime: isoDateTime,
        endDateTime: isoDateTime,
        top: z.number().int().min(1).max(100).optional().describe("Page size ($top), default 20."),
        skipToken: z.string().optional().describe("Paging token ($skiptoken)."),
        nextLink: z.string().optional().describe("A full @odata.nextLink URL to follow unchanged."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, startDateTime, endDateTime, top, skipToken, nextLink }) =>
      handle(async () => {
        if (nextLink) return respond(await client.request("GET", nextLink));
        const res = await client.request(
          "GET",
          ur(`/userUsageDetails/${encodeURIComponent(userId)}/sessions`),
          { query: { startDateTime, endDateTime, $top: top, $skiptoken: skipToken } },
        );
        return respond(res);
      }),
  );
}
