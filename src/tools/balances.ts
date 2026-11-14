/**
 * Tenant credit balance + per-user consumption diagnostics (read-only).
 *
 *   GET /copilot/costManagement/getTenantCreditBalance
 *   GET /copilot/costManagement/userBalances/{userId}
 *   GET /copilot/costManagement/userBalances/{userId}/serviceBalances
 *   GET /copilot/costManagement/userBalances/{userId}/serviceBalances/{serviceId}
 */

import { z } from "zod";
import type { ToolContext } from "./context.js";
import { guidParam, handle, respond } from "./helpers.js";

export function registerBalanceTools({ server, client }: ToolContext): void {
  server.registerTool(
    "get_tenant_credit_balance",
    {
      title: "Get tenant credit balance",
      description:
        "Get the tenant-wide Copilot/AI credit balance snapshot, one row per billing method " +
        "(e.g. prepaid, payAsYouGo). Credit-denominated; no currency or invoice amounts. " +
        "This is the only balance read that also supports app-only permissions " +
        "(CopilotCostManagement.Read.All).",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const res = await client.cmGet("/getTenantCreditBalance");
        return respond(res);
      }),
  );

  server.registerTool(
    "get_user_balance",
    {
      title: "Get a user's consumption balance",
      description:
        "Retrieve one user's credit consumption balance and the spending policy governing them " +
        "(spendingPolicyId). Pass expandServiceBalances=true to include the per-service breakdown " +
        "inline. Requires CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        userId: guidParam,
        expandServiceBalances: z
          .boolean()
          .optional()
          .describe("Include per-service balance rows inline via $expand=serviceBalances."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, expandServiceBalances }) =>
      handle(async () => {
        const res = await client.cmGet(`/userBalances/${encodeURIComponent(userId)}`, {
          query: expandServiceBalances ? { $expand: "serviceBalances" } : undefined,
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "list_user_service_balances",
    {
      title: "List a user's per-service balances",
      description:
        "List a user's per-service consumption balance rows — consumed, configured limit, and " +
        "remaining per governed service — useful to diagnose why a user can or can't consume a " +
        "service. Requires CopilotCostManagement-UserData.Read.All.",
      inputSchema: { userId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId }) =>
      handle(async () => {
        const res = await client.cmGet(
          `/userBalances/${encodeURIComponent(userId)}/serviceBalances`,
        );
        return respond(res);
      }),
  );

  server.registerTool(
    "get_user_service_balance",
    {
      title: "Get a user's balance for one service",
      description:
        "Retrieve one governed service's balance row for a user (consumed / limit / remaining). " +
        "Requires CopilotCostManagement-UserData.Read.All.",
      inputSchema: {
        userId: guidParam,
        serviceId: z.string().min(1).describe('Governed service id, e.g. "cowork" or "workIq".'),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ userId, serviceId }) =>
      handle(async () => {
        const res = await client.cmGet(
          `/userBalances/${encodeURIComponent(userId)}/serviceBalances/${encodeURIComponent(serviceId)}`,
        );
        return respond(res);
      }),
  );
}
