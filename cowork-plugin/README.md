# Copilot Cowork plugin (local server)

This folder is a portable plugin source for **Microsoft 365 Copilot Cowork**. It points a Cowork connector at the Copilot Cost Management MCP server **running locally on your machine**, exposed to Cowork over HTTPS with a dev tunnel.

> Cowork is a cloud service: it can't reach `localhost` directly, and it requires an HTTPS connector URL. A dev tunnel terminates TLS and relays to your local HTTP server — Microsoft's documented way to test a local MCP server with Cowork.

## Prerequisites

- The server installed (`npm install -g ./dontping-copilot-cost-management-mcp-<version>.tgz`).
- You've signed in once: `copilot-cost-management-mcp login` (device-code), or set `AZURE_CLIENT_SECRET` for app-only.
- [`devtunnel`](https://learn.microsoft.com/azure/developer/dev-tunnels/get-started) and Microsoft 365 Agents Toolkit CLI (`atk`) installed.

## Steps

1. **Run the server in HTTP mode:**
   ```bash
   copilot-cost-management-mcp http --port 3000
   ```

2. **Expose it with a dev tunnel** (in another terminal). Use `--protocol http` — it describes the *local* service; the public tunnel URL is HTTPS regardless:
   ```bash
   devtunnel port create cost-mgmt -p 3000 --protocol http
   devtunnel host cost-mgmt
   ```
   Copy the tunnel's HTTPS URL and append `/mcp` (e.g. `https://abc123-3000.euw.devtunnels.ms/mcp`).

3. **Set the URL** in [`.mcp.json`](./.mcp.json), replacing `REPLACE-WITH-YOUR-DEVTUNNEL-URL/mcp`.

4. **Build the Cowork app package** from this folder (Agents Toolkit converts the `.claude-plugin/plugin.json` + `.mcp.json` into an M365 app package, generating placeholder icons). Because the tunnel URL is HTTPS, pass `--default-auth-type None` so the connector doesn't require OAuth (the server is unauthenticated for local testing):
   ```bash
   atk import openplugin --path . \
     --privacy-url https://example.com/privacy \
     --terms-url https://example.com/terms \
     --default-auth-type None
   atk package --manifest-file ./appPackage/manifest.json \
     --output-package-file ./appPackage/build/appPackage.zip \
     --output-folder ./appPackage/build
   ```

5. **Add it to Cowork:** in Cowork → **Customize / Sources & Skills → Plugins**, upload/sideload the `appPackage.zip` (or `atk` can provision/share it). Cowork calls `initialize` + `tools/list` on your tunnel URL and the tools appear.

## Notes & cautions

- **Security:** the local HTTP endpoint is unauthenticated. Anyone with the tunnel URL can call it, which means acting as *your* signed-in Microsoft account. Keep the tunnel private (dev tunnels can require sign-in), stop it when you're done, and prefer read-only mode (`COPILOT_COST_MGMT_READONLY=true`) while testing.
- **This is a local/testing pattern.** For a shared or published Cowork plugin, host the server on a stable HTTPS endpoint (Azure App Service/Functions/Container Apps) and switch the connector auth to `OAuthPluginVault`.
- `macOS`: if tunneled requests return `502`, the server must accept IPv6 — it binds dual-stack by default; keep the default host.
- Keep the version in `.claude-plugin/plugin.json` roughly in step with the server version.
