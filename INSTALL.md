# Local install & setup

Run the Copilot Cost Management MCP server on your own machine and connect it to Claude Desktop (or Claude Code). ~10 minutes.

There are three parts:

1. **A Microsoft Entra app** — so the server can sign in to Microsoft Graph (one-time).
2. **Install the package** — gives you the `copilot-cost-management-mcp` command.
3. **Connect it to your Claude client.**

---

## 0. Prerequisites

- **Node.js 20 or newer** — check with `node -v`. Get it from <https://nodejs.org> if needed.
- A **Microsoft 365 tenant** with Copilot usage-based billing, and an account with an admin role (Billing / AI / Global Administrator) for the write and per-user operations.

---

## 1. Register a Microsoft Entra app (one-time)

This gives the server a **tenant ID** and **client ID** to sign in with. Pick the auth style that fits how you'll run it:

- **Device code (recommended, full access)** — you sign in interactively in a browser; no secret is stored. Covers every operation (policies, credit decisions, per-user data). → do steps **A–D**.
- **App-only (unattended, limited)** — no prompts ever, but only `get_tenant_credit_balance` and the service/funding usage reports work. → do steps **A, C (Application), E**.

### A. Create the registration
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it e.g. `Copilot Cost Management MCP`.
3. Supported account types: **Accounts in this organizational directory only**.
4. Leave **Redirect URI** empty. Click **Register**.
5. On the **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID** — you'll need both.

### B. Enable device-code sign-in *(device-code auth only)*
1. **Authentication** → **Advanced settings** → **Allow public client flows** → **Yes** → **Save**.

### C. Add API permissions
1. **API permissions** → **Add a permission** → **Microsoft Graph**.
2. Choose **Delegated permissions** (for device code) **or** **Application permissions** (for app-only).
3. Search **`CopilotCostManagement`** and add the scopes you want. A good full set for delegated:
   - `CopilotCostManagement.Read.All`
   - `CopilotCostManagement-Reports.Read.All`
   - `CopilotCostManagement-Policy.Read.All` / `CopilotCostManagement-Policy.ReadWrite.All`
   - `CopilotCostManagement-Assignment.Read.All` / `CopilotCostManagement-Assignment.ReadWrite.All`
   - `CopilotCostManagement-UserData.Read.All`
   - `CopilotCostManagement-CreditReq.Read.All` / `CopilotCostManagement-CreditReq.ReadWrite.All`

   For app-only, add just `CopilotCostManagement.Read.All` and `CopilotCostManagement-Reports.Read.All`.
4. Click **Grant admin consent for \<your tenant\>**.

> **If some `CopilotCostManagement-*` scopes aren't in the picker yet** (the API is in beta): a tenant admin can still consent to them by scope name via an [admin-consent request](https://learn.microsoft.com/en-us/entra/identity-platform/v2-admin-consent#request-the-permissions-from-a-directory-admin) using the app's client ID. You can also set `GRAPH_SCOPE` to a specific scope (e.g. `https://graph.microsoft.com/CopilotCostManagement-Policy.Read.All`) so the server requests exactly that at sign-in. See the [API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview#permissions).

### D. *(device code)* — nothing else to do. Skip E.

### E. Create a client secret *(app-only only)*
1. **Certificates & secrets** → **Client secrets** → **New client secret**.
2. Copy the **Value** immediately (shown once). This is your `AZURE_CLIENT_SECRET`.

---

## 2. Install the package

You were given a file like `dontping-copilot-cost-management-mcp-0.1.0.tgz`. Install it globally:

```bash
npm install -g ./dontping-copilot-cost-management-mcp-0.1.0.tgz
```

This adds the `copilot-cost-management-mcp` command. Confirm:

```bash
copilot-cost-management-mcp --version
```

> Prefer not to install globally? You can run it with `npx ./dontping-copilot-cost-management-mcp-0.1.0.tgz` or point your client's `command` at `node /path/to/unpacked/dist/index.js`.

---

## 3. Configure & verify

Create a `.env` file (anywhere you'll run the check from) with your values:

**Device code (full access):**
```dotenv
AZURE_TENANT_ID=<your-tenant-guid>
AZURE_CLIENT_ID=<your-client-id>
```

**App-only (limited):**
```dotenv
AZURE_TENANT_ID=<your-tenant-guid>
AZURE_CLIENT_ID=<your-client-id>
AZURE_CLIENT_SECRET=<your-client-secret>
```

Optional: add `COPILOT_COST_MGMT_READONLY=true` to expose only read tools.

**Device code — sign in once:**
```bash
copilot-cost-management-mcp login
```
A Microsoft sign-in URL + code prints in the terminal; open it, enter the code, sign in. The token is then cached on disk (`~/.copilot-cost-management-mcp/`) and reused across restarts — your MCP client won't prompt again until the refresh token expires.

Then verify (and, for app-only, this is the only step needed):
```bash
copilot-cost-management-mcp --check
```
Success prints your tenant credit balance. A permission error (with a clear code) means sign-in worked but a scope/role is missing — revisit step 1C. If it says "Not signed in", run `login` first.

To sign out and delete the cached token: `copilot-cost-management-mcp logout`.

---

## 4. Connect it to Claude

### Claude Desktop
1. Open **Settings → Developer → Edit Config** (or edit `claude_desktop_config.json` directly:
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`).
2. Add an entry under `mcpServers` (put your real values in `env`):

```json
{
  "mcpServers": {
    "copilot-cost-management": {
      "command": "copilot-cost-management-mcp",
      "env": {
        "AZURE_TENANT_ID": "<your-tenant-guid>",
        "AZURE_CLIENT_ID": "<your-client-id>"
      }
    }
  }
}
```
   For app-only, also add `"AZURE_CLIENT_SECRET": "<secret>"`. To keep secrets out of this file, instead set `"COPILOT_COST_MGMT_ENV_FILE": "/absolute/path/to/.env"` and the server will read your `.env`.

3. **Fully quit and reopen Claude Desktop.** The Copilot Cost Management tools now appear in the tool picker. With device-code auth, as long as you ran `login` first, the server picks up the cached token silently — no prompt.

> **Prefer not to use a terminal?** Skip the `login` step and sign in from the chat instead: ask Claude to run **`start_login`**. It replies with a Microsoft URL and a short code — open the URL, enter the code, sign in, then ask Claude to continue. `check_login_status` confirms you're signed in, and `sign_out` clears the token to switch accounts.

> If `command: "copilot-cost-management-mcp"` isn't found (Claude Desktop may not use your shell's PATH), use the absolute path — run `which copilot-cost-management-mcp` (macOS/Linux) or `where copilot-cost-management-mcp` (Windows) and put that in `command`, or use `"command": "node", "args": ["/full/path/to/dist/index.js"]`.

### Claude Code
From your project (or with `-s user` for global):

```bash
claude mcp add copilot-cost-management \
  --env AZURE_TENANT_ID=<tenant> \
  --env AZURE_CLIENT_ID=<client> \
  -- copilot-cost-management-mcp
```

Then reopen Claude Code.

---

## Notes

- **Sign in once.** The device-code token is cached on disk and refreshed automatically, so you only run `login` again when the refresh token finally expires (or after `logout`). If a tool ever reports "Not signed in", just run `copilot-cost-management-mcp login` again.
- **Token cache security.** The cache file (`~/.copilot-cost-management-mcp/token-cache-<clientid>.json`) holds refresh tokens in plain text, written owner-only (0600). It is not encrypted at rest — treat it like an SSH private key, and use `logout` to remove it. Override its location with `COPILOT_COST_MGMT_CACHE_DIR` / `COPILOT_COST_MGMT_TOKEN_CACHE`.
- **Fully unattended?** App-only auth (client secret) never prompts at all — at the cost of narrower coverage (balance + reports only).
- **Updating:** install the new `.tgz` the same way (`npm install -g ./…tgz`) and restart your client.
- **Uninstall:** `npm uninstall -g @dontping/copilot-cost-management-mcp`.
- The underlying API is **beta** and may change.
