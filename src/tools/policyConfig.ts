/**
 * Child configuration of a spending policy: group audience, per-service
 * settings, and notify-only email threshold alerts + recipients.
 *
 * Group audience (selectedGroupsSpendingPolicy only):
 *   GET    /spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups
 *   POST   .../assignedGroups/$ref
 *   DELETE .../assignedGroups/{groupId}/$ref
 *
 * Service settings:
 *   GET   /spendingPolicies/{id}/serviceSettings
 *   POST  /spendingPolicies/{id}/serviceSettings
 *   GET   /spendingPolicies/{id}/serviceSettings/{serviceId}
 *   PATCH /spendingPolicies/{id}/serviceSettings/{serviceId}   (no DELETE — disable instead)
 *
 * Alert settings:
 *   GET    /spendingPolicies/{id}/emailThresholdSettings
 *   PATCH  /spendingPolicies/{id}/emailThresholdSettings
 *   DELETE /spendingPolicies/{id}/emailThresholdSettings
 *   GET    /spendingPolicies/{id}/emailThresholdSettings/recipients
 *   POST   .../recipients/$ref
 *   DELETE .../recipients/{userId}/$ref
 *
 * All mutations use the PARENT POLICY's ETag as If-Match (child resources have
 * no ETag of their own).
 */

import { z } from "zod";
import { GraphClient } from "../graph.js";
import type { ToolContext } from "./context.js";
import { KNOWN_SERVICE_IDS, etagFor, guidParam, handle, respond } from "./helpers.js";

const SELECTED_GROUPS_CAST = "microsoft.graph.selectedGroupsSpendingPolicy";

const parentSub = (id: string) => `/spendingPolicies/${encodeURIComponent(id)}`;
const assignedGroups = (id: string) =>
  GraphClient.path(`/spendingPolicies/${encodeURIComponent(id)}/${SELECTED_GROUPS_CAST}/assignedGroups`);
const serviceSettings = (id: string) =>
  GraphClient.path(`/spendingPolicies/${encodeURIComponent(id)}/serviceSettings`);
const serviceSetting = (id: string, serviceId: string) =>
  GraphClient.path(
    `/spendingPolicies/${encodeURIComponent(id)}/serviceSettings/${encodeURIComponent(serviceId)}`,
  );
const alertSettings = (id: string) =>
  GraphClient.path(`/spendingPolicies/${encodeURIComponent(id)}/emailThresholdSettings`);
const recipients = (id: string) =>
  GraphClient.path(`/spendingPolicies/${encodeURIComponent(id)}/emailThresholdSettings/recipients`);

export function registerPolicyConfigTools({ server, client, readOnly }: ToolContext): void {
  // ── Group audience (reads) ────────────────────────────────────────────────────
  server.registerTool(
    "list_policy_assigned_groups",
    {
      title: "List a policy's assigned groups",
      description:
        "List the Entra groups assigned to a group-scoped spending policy. Requires " +
        "CopilotCostManagement-Assignment.Read.All.",
      inputSchema: { policyId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId }) =>
      handle(async () => respond(await client.request("GET", assignedGroups(policyId)))),
  );

  // ── Service settings (reads) ──────────────────────────────────────────────────
  server.registerTool(
    "list_policy_service_settings",
    {
      title: "List a policy's service settings",
      description:
        "List the per-service settings of a spending policy (which AI services it governs and " +
        "their state). Requires CopilotCostManagement-Policy.Read.All.",
      inputSchema: { policyId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId }) =>
      handle(async () => respond(await client.request("GET", serviceSettings(policyId)))),
  );

  server.registerTool(
    "get_policy_service_setting",
    {
      title: "Get a policy's service setting",
      description:
        "Retrieve one service's setting on a spending policy. The parentEtag field is the policy " +
        "ETag used for If-Match on child mutations. Requires CopilotCostManagement-Policy.Read.All.",
      inputSchema: {
        policyId: guidParam,
        serviceId: z.string().min(1).describe('Governed service id, e.g. "cowork".'),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId, serviceId }) =>
      handle(async () => respond(await client.request("GET", serviceSetting(policyId, serviceId)))),
  );

  // ── Alert settings (reads) ────────────────────────────────────────────────────
  server.registerTool(
    "get_policy_alert_settings",
    {
      title: "Get a policy's alert settings",
      description:
        "Retrieve the notify-only email threshold alert configuration of a spending policy " +
        "(isEnabled + thresholds). Returns 404 until alerts have been configured at least once. " +
        "Requires CopilotCostManagement-Policy.Read.All.",
      inputSchema: { policyId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId }) =>
      handle(async () => respond(await client.request("GET", alertSettings(policyId)))),
  );

  server.registerTool(
    "list_policy_alert_recipients",
    {
      title: "List a policy's alert recipients",
      description:
        "List the users who receive this policy's threshold alert emails. Requires " +
        "CopilotCostManagement-Policy.Read.All.",
      inputSchema: { policyId: guidParam },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ policyId }) =>
      handle(async () => respond(await client.request("GET", recipients(policyId)))),
  );

  if (readOnly) return;

  // ── Group audience (writes) ───────────────────────────────────────────────────
  server.registerTool(
    "add_policy_assigned_group",
    {
      title: "Assign a group to a policy",
      description:
        "Add an Entra group to a group-scoped spending policy's audience (reference link). Uses the " +
        "parent policy's ETag for If-Match (pass etag or omit to fetch). Requires " +
        "CopilotCostManagement-Assignment.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        groupId: guidParam.describe("Entra group object ID to assign."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, groupId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request("POST", `${assignedGroups(policyId)}/$ref`, {
          body: { "@odata.id": client.resourceUrl(`/groups/${groupId}`) },
          ifMatch,
        });
        return respond(res, "Group assigned to policy.");
      }),
  );

  server.registerTool(
    "remove_policy_assigned_group",
    {
      title: "Remove a group from a policy",
      description:
        "Remove a group from a group-scoped spending policy's audience. Removing the last group is " +
        "rejected (409 lastAssignedGroup). Requires CopilotCostManagement-Assignment.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        groupId: guidParam.describe("Entra group object ID to remove."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    ({ policyId, groupId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request(
          "DELETE",
          `${assignedGroups(policyId)}/${encodeURIComponent(groupId)}/$ref`,
          { ifMatch },
        );
        return respond(res, "Group removed from policy.");
      }),
  );

  // ── Service settings (writes) ─────────────────────────────────────────────────
  server.registerTool(
    "add_policy_service_setting",
    {
      title: "Add a service to a policy",
      description:
        "Add a governed AI service to a spending policy. Uses the parent policy's ETag for " +
        "If-Match. Requires CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        serviceId: z
          .string()
          .min(1)
          .describe(`Service id, e.g. ${KNOWN_SERVICE_IDS.map((s) => `"${s}"`).join(", ")}.`),
        state: z.enum(["enabled", "disabled"]).default("enabled").describe("Service state."),
        rawBody: z
          .record(z.any())
          .optional()
          .describe("Advanced: extra fields merged into the service-setting body."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, serviceId, state, rawBody, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const body = { serviceId, state, ...(rawBody ?? {}) };
        const res = await client.request("POST", serviceSettings(policyId), { body, ifMatch });
        return respond(res, "Service added to policy.");
      }),
  );

  server.registerTool(
    "update_policy_service_setting",
    {
      title: "Update a policy's service setting",
      description:
        "Update a service's setting on a policy. There is no DELETE for service settings — to " +
        "remove a service, set state=disabled (it stays in that state). Disabling the last enabled " +
        "service is rejected (409 atLeastOneServiceRequired). Requires " +
        "CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        serviceId: z.string().min(1).describe("Service id to update."),
        state: z.enum(["enabled", "disabled"]).optional().describe("New service state."),
        rawPatch: z.record(z.any()).optional().describe("Advanced: extra fields merged into the PATCH."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, serviceId, state, rawPatch, etag }) =>
      handle(async () => {
        const patch: Record<string, unknown> = {};
        if (state !== undefined) patch.state = state;
        if (rawPatch) Object.assign(patch, rawPatch);
        if (Object.keys(patch).length === 0) throw new Error("Provide at least one field to update.");

        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request("PATCH", serviceSetting(policyId, serviceId), {
          body: patch,
          ifMatch,
        });
        return respond(res, "Service setting updated.");
      }),
  );

  // ── Alert settings (writes) ───────────────────────────────────────────────────
  server.registerTool(
    "update_policy_alert_settings",
    {
      title: "Create or update a policy's alert settings",
      description:
        "Create or update the notify-only email threshold alerts for a policy. Alerts are email-only " +
        "to individual users and require at least one recipient before isEnabled=true " +
        "(add_policy_alert_recipient). Thresholds is a raw array; each item is either\n" +
        '  { "@odata.type": "#microsoft.graph.absoluteCostManagementAlertThreshold", "value": <n>, "unit": "copilotCredit" }\n' +
        "or a percentage threshold whose integer value is 1–100. Requires " +
        "CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        isEnabled: z.boolean().optional().describe("Enable/disable alerting (needs a recipient to enable)."),
        thresholds: z
          .array(z.record(z.any()))
          .optional()
          .describe("Raw threshold objects (see description for the two shapes)."),
        rawBody: z.record(z.any()).optional().describe("Advanced: extra fields merged into the body."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, isEnabled, thresholds, rawBody, etag }) =>
      handle(async () => {
        const body: Record<string, unknown> = {};
        if (isEnabled !== undefined) body.isEnabled = isEnabled;
        if (thresholds !== undefined) body.thresholds = thresholds;
        if (rawBody) Object.assign(body, rawBody);

        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request("PATCH", alertSettings(policyId), { body, ifMatch });
        return respond(res, "Alert settings saved.");
      }),
  );

  server.registerTool(
    "delete_policy_alert_settings",
    {
      title: "Delete a policy's alert settings",
      description:
        "Delete the alert configuration of a policy. Requires CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    ({ policyId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request("DELETE", alertSettings(policyId), { ifMatch });
        return respond(res, "Alert settings deleted.");
      }),
  );

  server.registerTool(
    "add_policy_alert_recipient",
    {
      title: "Add an alert recipient",
      description:
        "Add a user as a recipient of a policy's threshold alert emails (reference link). The user " +
        "must resolve to a real tenant user. Requires CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        userId: guidParam.describe("Entra user object ID to add as a recipient."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ policyId, userId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request("POST", `${recipients(policyId)}/$ref`, {
          body: { "@odata.id": client.resourceUrl(`/users/${userId}`) },
          ifMatch,
        });
        return respond(res, "Alert recipient added.");
      }),
  );

  server.registerTool(
    "remove_policy_alert_recipient",
    {
      title: "Remove an alert recipient",
      description:
        "Remove a user from a policy's alert recipients. Removing the last recipient while alerting " +
        "is enabled is rejected (400 alertRecipientsRequired). Requires " +
        "CopilotCostManagement-Policy.ReadWrite.All.",
      inputSchema: {
        policyId: guidParam,
        userId: guidParam.describe("Entra user object ID to remove."),
        etag: z.string().optional().describe("Parent policy If-Match ETag (optional)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    ({ policyId, userId, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, parentSub(policyId), etag);
        const res = await client.request(
          "DELETE",
          `${recipients(policyId)}/${encodeURIComponent(userId)}/$ref`,
          { ifMatch },
        );
        return respond(res, "Alert recipient removed.");
      }),
  );
}
