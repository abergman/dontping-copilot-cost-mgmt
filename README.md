# Copilot Cost Management MCP

An [MCP](https://modelcontextprotocol.io) server for the **Microsoft 365 Copilot Cost Management API** — the credit-denominated governance and observability layer for Copilot/AI consumption in Microsoft Graph (`/copilot/costManagement`, currently `beta`).

> The Copilot Cost Management API is the building block for your own agent. This server turns it into tools any MCP client (Claude, or your own agent) can call: read balances, shape spending policies, triage credit requests, and pull usage reports — all in Copilot **credits**, with no currency, price, or invoice amounts anywhere in the contract.

## What it does

| Area | Tools |
| --- | --- |
| **Balances & diagnostics** | `get_tenant_credit_balance`, `get_user_balance`, `list_user_service_balances`, `get_user_service_balance` |
| **Spending policies** | `list_spending_policies`, `get_spending_policy`, `create_spending_policy`, `update_spending_policy`, `delete_spending_policy` |
| **Policy audience** | `list_policy_assigned_groups`, `add_policy_assigned_group`, `remove_policy_assigned_group` |
| **Policy services** | `list_policy_service_settings`, `get_policy_service_setting`, `add_policy_service_setting`, `update_policy_service_setting` |
| **Policy alerts** | `get_policy_alert_settings`, `update_policy_alert_settings`, `delete_policy_alert_settings`, `list_policy_alert_recipients`, `add_policy_alert_recipient`, `remove_policy_alert_recipient` |
| **Credit requests** | `list_credit_requests`, `get_credit_request`, `decide_credit_request`, `batch_decide_credit_requests` |
| **Usage reports** | `list_user_usage_records`, `get_user_usage_record`, `list_group_usage_records`, `get_group_usage_record`, `list_service_usage_records`, `get_service_usage_record`, `list_funding_usage_records`, `get_user_usage_detail`, `list_user_usage_service_details`, `list_user_usage_session_details` |
| **Tenants** | `list_tenants`, `current_tenant`, `add_tenant`, `use_tenant`, `remove_tenant` |
| **Authentication** | `start_login`, `check_login_status`, `sign_out` |
| **Escape hatch** | `cost_management_request` (raw request, e.g. to follow a `nextLink` or reach a brand-new beta endpoint) |

45 tools in all (32 in read-only mode). All of the balance/policy/credit/usage tools run against the **active tenant** (see [Multiple tenants](#multiple-tenants)).

## Quick start

**Claude Desktop (easiest):** install the prebuilt **Desktop Extension** — drag `…-<version>.mcpb` into **Settings → Extensions**, enter your Entra tenant/client ID in the form, and sign in from the chat with `start_login`. No terminal, no JSON. Build it yourself with `bash scripts/pack-mcpb.sh`.

**From source:**

```bash
npm install
npm run build
```

Then point your MCP client at `node dist/index.js` (stdio transport) with credentials in the environment. Nothing listens on a port — it speaks JSON-RPC over stdin/stdout, and all logs go to stderr.

### CLI commands

```
copilot-cost-management-mcp            run the MCP server over stdio (default)
copilot-cost-management-mcp login      device-code sign-in; caches the token (device-code auth)
copilot-cost-management-mcp logout     remove the cached token
copilot-cost-management-mcp --check    verify credentials + reach the API, then exit
copilot-cost-management-mcp --help     full usage and env vars
```

See **[INSTALL.md](./INSTALL.md)** for the full local walkthrough (Entra app registration → install → sign in → connect to Claude).

### Claude Desktop / Claude Code

Add to your MCP config (e.g. `claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "copilot-cost-management": {
      "command": "node",
      "args": ["/absolute/path/to/dontping-copilot-cost-mgmt/dist/index.js"],
      "env": {
        "AZURE_TENANT_ID": "<tenant-guid>",
        "AZURE_CLIENT_ID": "<app-client-id>",
        "AZURE_CLIENT_SECRET": "<app-secret>"
      }
    }
  }
}
```

## Authentication

The server talks to Microsoft Graph and needs a token carrying the relevant `CopilotCostManagement-*` scopes. Pick **one** strategy (checked in this order):

1. **Raw bearer token** — `GRAPH_ACCESS_TOKEN`. Easiest for testing; not refreshed, so it expires in ~an hour.
2. **App-only (client credentials)** — `AZURE_TENANT_ID` + `AZURE_CLIENT_ID` + `AZURE_CLIENT_SECRET`. Only works for the operations that support **application** permissions: `get_tenant_credit_balance` and the service/funding usage reports.
3. **Delegated (device code, persistent)** — `AZURE_TENANT_ID` + `AZURE_CLIENT_ID` (no secret). Works for **all** operations (most are delegated-only). Sign in once — either from a terminal (`copilot-cost-management-mcp login`) **or from inside Claude** by invoking the `start_login` tool (it returns a URL + code in the chat). Either way the token is cached on disk and reused across restarts (no re-prompting) until the refresh token expires. The signed-in admin must hold the required directory role (Billing/AI/Global Administrator, etc.).

   In device-code mode the server also exposes `start_login`, `check_login_status`, and `sign_out` as tools, so an agent can authenticate without touching a terminal.

See [`.env.example`](./.env.example) for every variable, plus optional overrides (`GRAPH_BASE_URL`, `GRAPH_API_VERSION`, `GRAPH_SCOPE`).

> **Most of this API is delegated-only.** If you want full coverage (policies, credit-request decisions, per-user diagnostics), use the device-code flow. App-only credentials are limited to tenant balance and aggregate reports.

### Permission scopes

| Scope | Covers |
| --- | --- |
| `CopilotCostManagement.Read.All` | Tenant credit balance (also app-only); all usage reports (delegated) |
| `CopilotCostManagement-Reports.Read.All` | Service & funding usage reports (also app-only) |
| `CopilotCostManagement-Policy.Read.All` / `.ReadWrite.All` | Policies, service settings, alerts |
| `CopilotCostManagement-Assignment.Read.All` / `.ReadWrite.All` | Policy group audience |
| `CopilotCostManagement-UserData.Read.All` | Per-user balances, user/group usage records, raw per-user detail |
| `CopilotCostManagement-CreditReq.Read.All` / `.ReadWrite.All` | Credit requests (read / approve-reject) |

Each tool's description names the least-privileged scope it needs.

## Multiple tenants

The server manages any number of Microsoft 365 tenants and runs every tool against the **active** tenant. The tenant list and a separate token cache per tenant are persisted under the cache dir (`~/.copilot-cost-management-mcp/`), so it all survives restarts.

### Which model to choose

|  | One multi-tenant app *(recommended)* | A separate app per tenant |
| --- | --- | --- |
| Entra registrations | one, shared | one per customer tenant |
| `AZURE_CLIENT_ID` | the shared app | a default (optional); override per tenant |
| Onboarding a tenant | that tenant's admin consents once | register + consent in that tenant |
| Best for | MSPs / many tenants | strict per-customer isolation |

Auth per tenant: with **`AZURE_CLIENT_SECRET` set**, tenants use **app-only** — unattended, but limited to balances + usage reports (the rest of the API is delegated-only). **Without a secret**, each tenant uses **device-code** — you sign in per tenant, with full access. This matches most management workflows.

### Setup — one multi-tenant app (recommended)

1. **Register the app as multi-tenant.** In the [Entra admin center](https://entra.microsoft.com) of *your own* tenant → **App registrations → New registration**. Under **Supported account types**, choose **Accounts in any organizational directory (multitenant)**. Leave the redirect URI empty. Register, and copy the **Application (client) ID** — this one ID is used for every customer tenant.
2. **Enable device-code** (for full-access management): **Authentication → Allow public client flows → Yes**.
3. **Add API permissions** (see [Permission scopes](#permission-scopes)) as **delegated** (for device-code) and/or **application** (for app-only). You consent to these in *your* tenant; each customer tenant consents separately in the next step.
4. **Onboard each customer tenant** — a Global/Billing Administrator of that tenant grants consent once, via this URL (substitute the customer's tenant ID or domain and your client ID):

   ```
   https://login.microsoftonline.com/<customer-tenant-id>/adminconsent?client_id=<your-client-id>
   ```

   (For device-code, an admin can alternatively just consent interactively the first time they sign in with `start_login`.)
5. **Point the server at the shared app.** In the Desktop Extension form set **App (client) ID** to your client id and leave **Default tenant ID** blank (or set one to seed the first tenant). For the npm/CLI path, set `AZURE_CLIENT_ID` (and optionally `AZURE_CLIENT_SECRET` for app-only).
6. **Register tenants and sign in — from the chat:**

   ```text
   add_tenant   { "tenantId": "<contoso-guid>", "label": "Contoso" }
   start_login                      # opens the sign-in URL + code for Contoso
   add_tenant   { "tenantId": "<fabrikam-guid>", "label": "Fabrikam" }
   start_login                      # sign in to Fabrikam
   ```

   Each sign-in is cached separately. (App-only needs no `start_login`.)

### Setup — a separate app per tenant

Register an app in each customer tenant as usual, then register each with its own client id:

```text
add_tenant { "tenantId": "<contoso-guid>", "label": "Contoso", "clientId": "<contoso-app-id>" }
```

The per-tenant `clientId` overrides the shared `AZURE_CLIENT_ID` for that tenant.

### Day-to-day

- `list_tenants` / `current_tenant` — see what's registered and which is active
- `use_tenant { "tenantId": "<guid>" }` — switch; subsequent tools run against it
- `remove_tenant` — unregister (sign out first with `sign_out` to clear its token)

CLI equivalents: `copilot-cost-management-mcp tenants`, `… login [tenantId]`, `… logout [tenantId]`, `… --check [tenantId]`.

## Read-only mode

Set `COPILOT_COST_MGMT_READONLY=true` to register **only** read tools — no create/update/delete, no credit-request decisions, and the raw tool is limited to `GET`. Ideal for handing an agent observability without the ability to change anything.

## Concepts worth knowing

- **Credits, not money.** Every quantity carries a `unit` (today always `copilotCredit`). No currency/invoice/price data is returned.
- **IDs are GUIDs.** Users and groups are Entra **object IDs**, not UPNs or emails.
- **Optimistic concurrency.** Mutations use `If-Match` ETags. The write tools accept an `etag`; omit it and the server fetches the current one first (convenient, but it drops the stale-write guard — pass the ETag from the matching `get_*` tool when that matters). Policy child resources (service settings, group/recipient links) are guarded by the **parent policy's** ETag.
- **Spending policies** are tenant-wide (`allUsers`) or group-scoped (`selectedGroups`); the audience is fixed at creation. Policies are **created disabled** — a group-scoped policy needs an assigned group before it can be enabled.
- **Service settings have no DELETE** — set `state: "disabled"` to remove a service from a policy.
- **Credit requests** can't be created or deleted via the API (users raise them elsewhere); you can only list/read and approve/reject (`pending → approved | rejected`).
- **Usage reports require a time selector:**
  - User/group/service records → `startDate` + `endDate`, exactly one completed UTC day (`endDate = startDate + 1 day`).
  - Funding records → `month` (`YYYY-MM`).
  - Raw user detail → `startDateTime` + `endDateTime`, a ≤24h window that doesn't cross a UTC month boundary.
  - Paging is server-driven (`$skiptoken`); follow the returned `@odata.nextLink` unchanged (each list tool accepts a `nextLink` argument for this). Ordering is only by `consumedQuantity`.

## Caveats

- The underlying API is **`beta`** and subject to change; it isn't supported for production use by Microsoft.
- Deleting a spending policy created via the API may be rejected by the service — prefer disabling it.
- Some admin-center settings (e.g. "allow new agents as they become available") aren't exposed by the API.

## Development

```bash
npm run typecheck   # tsc --noEmit
npm run build       # compile to dist/
npm run dev         # watch mode (tsx)
npm start           # run the built server
```

## References

- [Copilot Cost Management API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview)
- [Usage-based billing / Copilot credits](https://learn.microsoft.com/en-us/microsoft-365/copilot/usage-based-billing-overview-copilot-credits)
- [Model Context Protocol](https://modelcontextprotocol.io)

## License

MIT — see [LICENSE](./LICENSE).
