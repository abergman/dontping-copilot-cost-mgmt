/**
 * Credit requests — end users ask for more Copilot credits; admins approve/reject.
 * Requests can't be created or deleted via the API (users create them elsewhere).
 *
 *   GET   /copilot/costManagement/creditRequests
 *   GET   /copilot/costManagement/creditRequests/{id}
 *   PATCH /copilot/costManagement/creditRequests/{id}             (If-Match, strong ETag)
 *   POST  /copilot/costManagement/creditRequests/batchRecordDecisions
 *
 * Supported transitions: pending → approved, pending → rejected.
 */

import { z } from "zod";
import { GraphClient } from "../graph.js";
import type { ToolContext } from "./context.js";
import { etagFor, guidParam, handle, respond } from "./helpers.js";

const creditRequests = () => GraphClient.path("/creditRequests");
const creditRequest = (id: string) =>
  GraphClient.path(`/creditRequests/${encodeURIComponent(id)}`);
const creditRequestSub = (id: string) => `/creditRequests/${encodeURIComponent(id)}`;
const batchDecisions = () => GraphClient.path("/creditRequests/batchRecordDecisions");

export function registerCreditRequestTools({ server, client, readOnly }: ToolContext): void {
  server.registerTool(
    "list_credit_requests",
    {
      title: "List credit requests",
      description:
        "List the tenant's Copilot credit requests (pending, approved, rejected). Use pageSize / " +
        "skipToken (or follow nextLink) for paging. Readable by Billing/Global/AI Administrators and " +
        "Global/AI Readers. Requires CopilotCostManagement-CreditReq.Read.All.",
      inputSchema: {
        pageSize: z.number().int().min(1).max(100).optional().describe("Page size ($top)."),
        skipToken: z.string().optional().describe("Paging token ($skiptoken) from a prior page."),
        filter: z.string().optional().describe("Raw OData $filter expression (advanced)."),
        nextLink: z.string().optional().describe("A full @odata.nextLink URL to follow unchanged."),
      },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ pageSize, skipToken, filter, nextLink }) =>
      handle(async () => {
        if (nextLink) return respond(await client.request("GET", nextLink));
        const res = await client.request("GET", creditRequests(), {
          query: { $top: pageSize, $skiptoken: skipToken, $filter: filter },
        });
        return respond(res);
      }),
  );

  server.registerTool(
    "get_credit_request",
    {
      title: "Get a credit request",
      description:
        "Retrieve a single credit request by id, including its strong @odata.etag (needed as the " +
        "If-Match for approve/reject). Requires CopilotCostManagement-CreditReq.Read.All.",
      inputSchema: { requestId: guidParam.describe("Credit request id.") },
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    ({ requestId }) =>
      handle(async () => respond(await client.request("GET", creditRequest(requestId)))),
  );

  if (readOnly) return;

  server.registerTool(
    "decide_credit_request",
    {
      title: "Approve or reject a credit request",
      description:
        "Approve or reject a single pending credit request (sets state to approved or rejected). " +
        "Uses a strong-ETag If-Match: pass etag from get_credit_request, or omit to fetch the " +
        "current one. A same-state decision is a no-op (204) but the ETag is still checked. " +
        "Requires CopilotCostManagement-CreditReq.ReadWrite.All and a Global/Billing/AI " +
        "Administrator role.",
      inputSchema: {
        requestId: guidParam.describe("Credit request id."),
        decision: z
          .enum(["approved", "rejected"])
          .describe("Target state for the request."),
        etag: z
          .string()
          .optional()
          .describe("Strong If-Match ETag. If omitted, fetched first (must not be weak/*)."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ requestId, decision, etag }) =>
      handle(async () => {
        const ifMatch = await etagFor(client, creditRequestSub(requestId), etag);
        const res = await client.request("PATCH", creditRequest(requestId), {
          body: { state: decision },
          ifMatch,
        });
        return respond(res, `Credit request ${decision}.`);
      }),
  );

  server.registerTool(
    "batch_decide_credit_requests",
    {
      title: "Approve/reject credit requests in bulk",
      description:
        "Approve or reject up to 100 credit requests in one call. Each item carries its own strong " +
        "ETag; the action does NOT use If-Match. Items are processed independently (not atomic) and " +
        "results come back in input order with a per-item status (succeeded, notFound, failed, " +
        "preconditionFailed, invalidStateTransition). requestId values must be unique. Requires " +
        "CopilotCostManagement-CreditReq.ReadWrite.All.",
      inputSchema: {
        decisions: z
          .array(
            z.object({
              requestId: guidParam.describe("Credit request id (unique within the batch)."),
              state: z.enum(["approved", "rejected"]).describe("Target state."),
              etag: z.string().describe("The request's current strong ETag."),
            }),
          )
          .min(1)
          .max(100)
          .describe("1–100 decisions with unique requestId values."),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    ({ decisions }) =>
      handle(async () => {
        const ids = decisions.map((d) => d.requestId);
        if (new Set(ids).size !== ids.length) {
          throw new Error("Duplicate requestId values in the batch; each must be unique.");
        }
        const res = await client.request("POST", batchDecisions(), {
          body: { requests: decisions },
        });
        return respond(res);
      }),
  );
}
