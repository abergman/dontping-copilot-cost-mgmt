/**
 * Spending policies — caps on Copilot Credit consumption for an audience.
 *
 *   GET    /copilot/costManagement/spendingPolicies
 *   POST   /copilot/costManagement/spendingPolicies
 *   GET    /copilot/costManagement/spendingPolicies/{id}
 *   PATCH  /copilot/costManagement/spendingPolicies/{id}          (If-Match)
 *   DELETE /copilot/costManagement/spendingPolicies/{id}          (If-Match)
 *
 * A policy is either tenant-wide (allUsersSpendingPolicy) or group-scoped
 * (selectedGroupsSpendingPolicy). The audience type is fixed at creation.
 * Policies are created DISABLED; group-scoped policies need an assigned group
 * (add_policy_assigned_group) before they can be enabled.
 */

import { z } from "zod";
import { GraphClient } from "../graph.js";
import type { ToolContext } from "./context.js";
import {
  CREDIT_UNIT,
  KNOWN_BILLING_METHODS,
  KNOWN_SERVICE_IDS,
  etagFor,
  guidParam,
  handle,
  respond,
} from "./helpers.js";

const policies = () => GraphClient.path("/spendingPolicies");
const policy = (id: string) => GraphClient.path(`/spendingPolicies/${encodeURIComponent(id)}`);
const policySub = (id: string) => `/spendingPolicies/${encodeURIComponent(id)}`;

function quantity(value: number, unit: string) {
  return {
    "@odata.type": "#microsoft.graph.limitedCostManagementQuantity",
    unit,
    value,
  };
}

const serviceSettingInput = z.object({
  serviceId: z
    .string()
    .min(1)
    .describe(`Governed service id, e.g. ${KNOWN_SERVICE_IDS.map((s) => `"${s}"`).join(", ")}.`),
  state: z
    .enum(["enabled", "disabled"])
    .default("enabled")
    .describe("Whether the service is governed by this policy."),
});

export function registerPolicyTools({ server, client, readOnly }: ToolContext): void {
  // ── Reads ───────────────────────────────────────────────────────────────────

  server.registerTool(
    "list_spending_policies",
    {
      title: "List spending policies",
      description:
        "List the tenant's Copilot cost-management spending policies (tenant-wide and " +
        "group-scoped), with their limits and enablement state. Requires " +
        "CopilotCostManagement-Policy.Read.All.",
      inputSchema: {
        top: z.number().int().min(1).max(999).optional().describe("Max items to return ($top)."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ top }) =>
      handle(async () => {
        const res = await client.cmGet("/spendingPolicies", {
          query: top ? { $top: top } : undefined,
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "get_spending_policy",
    {
      title: "Get a spending policy",
      description:
        "Retrieve a single spending policy by id, including its @odata.etag (needed as the " +
        "If-Match value for updates/deletes). Service settings and alert settings are read via " +
        "their own tools. Requires CopilotCostManagement-Policy.Read.All.",
      inputSchema: { policyId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId }) =>
      handle(async () => {
        const res = await client.request("GET", policy(policyId));
        return respond(res);
      }),
  );

  if (readOnly) return;

  // ── Writes ──────────────────────────────────────────────────────────────────

  server.registerTool(
    "create_spending_policy",
    {
      title: "Create a spending policy",
      description:
        "Create a Copilot Credit spending policy. Policies are created DISABLED. For a " +
        "group-scoped policy, next add at least one group (add_policy_assigned_group), then enable " +
        "it (update_spending_policy state=enabled). displayName must be ≥10 characters with no " +
        "special characters. Requires CopilotCostManagement-Policy.ReadWrite.All and a " +
        "Billing/Global Administrator role.",
      inputSchema: {
        scope: z
          .enum(["allUsers", "selectedGroups"])
          .describe("Audience: all tenant users, or selected groups. Fixed after creation."),
        displayName: z.string().min(10).describe("Policy name (≥10 chars, no special characters)."),
        periodLimit: z.number().positive().describe("Total credit cap for the billing period."),
        perUserLimit: z.number().positive().describe("Per-user credit cap within the period."),
        serviceSettings: z
          .array(serviceSettingInput)
          .min(1)
          .describe("Governed services and their state (at least one required)."),
        billingMethods: z
          .string()
          .optional()
          .describe(
            `Optional billing method, e.g. ${KNOWN_BILLING_METHODS.map((m) => `"${m}"`).join(", ")}.`,
          ),
        unit: z.string().optional().describe(`Credit unit (default "${CREDIT_UNIT}").`),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ scope, displayName, periodLimit, perUserLimit, serviceSettings, billingMethods, unit }) =>
      handle(async () => {
        const u = unit ?? CREDIT_UNIT;
        const odataType =
          scope === "allUsers"
            ? "#microsoft.graph.allUsersSpendingPolicy"
            : "#microsoft.graph.selectedGroupsSpendingPolicy";

        const body: Record<string, unknown> = {
          "@odata.type": odataType,
          displayName,
          periodLimit: quantity(periodLimit, u),
          perUserLimit: quantity(perUserLimit, u),
          serviceSettings: serviceSettings.map((s) => ({ serviceId: s.serviceId, state: s.state })),
        };
        if (billingMethods) body.billingMethods = billingMethods;

        const res = await client.request("POST", policies(), { body });
        const note =
          scope === "selectedGroups"
            ? "Spending policy created (disabled, no groups). Add groups with add_policy_assigned_group, then enable."
            : "Spending policy created (disabled). Enable with update_spending_policy state=enabled.";
        return respond(res, note);
      }),
  );

  server.registerTool(
    "update_spending_policy",
    {
      title: "Update a spending policy",
      description:
        "Update a spending policy's editable fields (e.g. displayName, state, limits). Enabling a " +
        "group-scoped policy requires it to already have an assigned group. Uses optimistic " +
        "concurrency: pass the ETag from get_spending_policy, or omit it to let the server read the " +
        "current one. Requires CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        displayName: z.string().min(10).optional().describe("New policy name (≥10 chars)."),
        state: z.enum(["enabled", "disabled"]).optional().describe("Enable or disable the policy."),
        periodLimit: z.number().positive().optional().describe("New total period credit cap."),
        perUserLimit: z.number().positive().optional().describe("New per-user credit cap."),
        billingMethods: z.string().optional().describe("New billing method."),
        unit: z.string().optional().describe(`Credit unit for limit changes (default "${CREDIT_UNIT}").`),
        rawPatch: z
          .record(z.any())
          .optional()
          .describe("Advanced: extra fields merged into the PATCH body (takes precedence)."),
        etag: z
          .string()
          .optional()
          .describe("If-Match ETag. If omitted, the current ETag is fetched first."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, displayName, state, periodLimit, perUserLimit, billingMethods, unit, rawPatch, etag }) =>
      handle(async () => {
        const u = unit ?? CREDIT_UNIT;
        const patch: Record<string, unknown> = {};
        if (displayName !== undefined) patch.displayName = displayName;
        if (state !== undefined) patch.state = state;
        if (billingMethods !== undefined) patch.billingMethods = billingMethods;
        if (periodLimit !== undefined) patch.periodLimit = quantity(periodLimit, u);
        if (perUserLimit !== undefined) patch.perUserLimit = quantity(perUserLimit, u);
        if (rawPatch) Object.assign(patch, rawPatch);

        if (Object.keys(patch).length === 0) {
          throw new Error("Provide at least one field to update.");
        }

        const ifMatch = await etagFor(client, policySub(policyId), etag);
        const res = await client.request("PATCH", policy(policyId), { body: patch, ifMatch });
        return respond(res, "Spending policy updated.");
      }),
  );

  server.registerTool(
    "delete_spending_policy",
    {
      title: "Delete a spending policy",
      description:
        "Delete a spending policy. Uses If-Match (pass etag, or omit to fetch current). Note: the " +
        "beta service may reject deletion of policies created via the API — prefer disabling " +
        "(update_spending_policy state=disabled) if deletion is not supported. Requires " +
        "CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        etag: z.string().optional().describe("If-Match ETag. If omitted, fetched first."),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    ({ policyId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, policySub(policyId), etag);
        const res = await client.request("DELETE", policy(policyId), { ifMatch });
        return respond(res, "Spending policy deleted.");
      }),
  );
}
