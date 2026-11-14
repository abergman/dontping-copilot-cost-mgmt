---
name: connect-cowork
description: |
  Connect this Copilot Cost Management MCP server to Microsoft 365 Copilot
  Cowork (or Copilot Studio). Use when the user wants to "set up Cowork",
  "connect to Copilot Cowork", "expose the server to Copilot", "build the
  Cowork plugin package", "create a dev tunnel for Copilot", or "add this to
  Copilot Cowork/Studio".
license: MIT
---

# Connect to Microsoft Copilot (Cowork / Studio)

Copilot hosts reach MCP servers only over **remote Streamable HTTP (HTTPS)**, so
this server must run in HTTP mode and be exposed over HTTPS. Walk the user
through it; run commands for them where you can, but never paste or handle their
secrets.

## 1. Pre-flight

- Confirm the build is current: `npm run build`.
- Confirm Graph auth works: `copilot-cost-management-mcp --check` (or
  `node dist/index.js --check`). If it says "not signed in", have them run
  `copilot-cost-management-mcp login` and complete the browser sign-in. For an
  unattended host, app-only (`AZURE_CLIENT_SECRET`) is an alternative with
  narrower coverage.
- Ensure `devtunnel` is installed (https://aka.ms/devtunnel). For the Cowork
  package step, `atk` (`npm i -g @microsoft/m365agentstoolkit-cli`) is optional.

## 2. Serve over HTTP + open a tunnel

Use the helper, which starts the HTTP server and a dev tunnel and prints the
public `/mcp` URL:

```bash
bash scripts/cowork-serve.sh --port 3000
```

Or do it manually:

```bash
copilot-cost-management-mcp http --port 3000
devtunnel port create cost-mgmt -p 3000 --protocol http   # http = the LOCAL service
devtunnel host cost-mgmt
```

Capture the tunnel's HTTPS URL and append `/mcp`. Remind the user this endpoint
is **unauthenticated** — keep the tunnel private, stop it when done, and consider
`COPILOT_COST_MGMT_READONLY=true` while testing.

## 3a. Copilot Studio

In the agent: **Tools → Add a tool → Model Context Protocol**, paste the tunnel
`/mcp` URL, auth **None** (dev). Generative orchestration must be on. Done.

## 3b. Copilot Cowork — build and add the plugin

The plugin source is in `cowork-plugin/` (manifest v1.29, a `configure-cost-management`
skill, and a connector pointing at your server).

1. Set the connector URL — either edit `cowork-plugin/manifest.json`
   (`mcpServerUrl`) or pass it to the build script below.
2. Build the uploadable M365 app package `.zip`:

   ```bash
   COWORK_MCP_URL="https://<tunnel>/mcp" bash scripts/pack-cowork-plugin.sh
   ```

   (This hand-assembles the `.zip` — no extra tooling needed. Alternatively, use
   Microsoft's Agents Toolkit on the `.claude-plugin/plugin.json` + `.mcp.json`
   source: `atk import openplugin --path cowork-plugin --default-auth-type None …`
   then `atk package`.)
3. Add it to Cowork: upload/sideload the `.zip` (or `atk install --file-path … --scope Personal`).
   Cowork calls `initialize` + `tools/list` on the URL and the tools appear; the
   `configure-cost-management` skill shows up via `/`.

## Notes

- Keep `cowork-plugin/manifest.json` `version` in step with `package.json`.
- Before anything shared/published: host the server on a stable HTTPS endpoint
  (Azure App Service/Functions/Container Apps) and switch the connector auth from
  `None` to `OAuthPluginVault` with a registered client. Placeholder
  `privacyUrl`/`termsOfUseUrl` in the manifest must be replaced too.
- Full reference: `cowork-plugin/README.md`.
