#!/usr/bin/env bash
#
# Build the Microsoft 365 Copilot Cowork plugin package (.zip) by hand — no extra
# tooling required. Assembles manifest.json + icons + skills/ from cowork-plugin/.
#
# The connector URL comes from COWORK_MCP_URL (your hosted/tunnel .../mcp URL);
# if unset, the manifest's placeholder is kept and the package must be edited
# before it will work. The manifest version is synced from package.json.
#
# Usage:  COWORK_MCP_URL="https://<host>/mcp" bash scripts/pack-cowork-plugin.sh
set -euo pipefail
cd "$(dirname "$0")/.."

VERSION="$(node -p "require('./package.json').version")"
SRC="cowork-plugin"
OUT_ABS="$PWD/dontping-copilot-cost-management-cowork-plugin-${VERSION}.zip"

# Make sure icons exist (regenerate placeholders if missing).
if [ ! -f "$SRC/color.png" ] || [ ! -f "$SRC/outline.png" ]; then
  node scripts/gen-icons.mjs
fi

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
cp "$SRC/manifest.json" "$SRC/color.png" "$SRC/outline.png" "$STAGE/"
cp -r "$SRC/skills" "$STAGE/skills"

# Sync version and (optionally) the connector URL into the staged manifest.
node -e '
  const fs = require("fs");
  const p = process.argv[1];
  const m = JSON.parse(fs.readFileSync(p, "utf8"));
  m.version = require("./package.json").version;
  const url = process.env.COWORK_MCP_URL;
  if (url) m.agentConnectors[0].toolSource.remoteMcpServer.mcpServerUrl = url;
  fs.writeFileSync(p, JSON.stringify(m, null, 2) + "\n");
  const current = m.agentConnectors[0].toolSource.remoteMcpServer.mcpServerUrl;
  if (/REPLACE-WITH/.test(current)) {
    console.warn("  ! mcpServerUrl is still a placeholder — set COWORK_MCP_URL or edit the manifest before uploading.");
  } else {
    console.log("  mcpServerUrl = " + current);
  }
' "$STAGE/manifest.json"

rm -f "$OUT_ABS"
( cd "$STAGE" && zip -r -q -X "$OUT_ABS" manifest.json color.png outline.png skills )
echo "Built $(basename "$OUT_ABS") ($(du -h "$OUT_ABS" | cut -f1))"
