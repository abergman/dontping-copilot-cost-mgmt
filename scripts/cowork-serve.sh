#!/usr/bin/env bash
#
# Run the server in HTTP mode and expose it over HTTPS with a dev tunnel, so
# Microsoft Copilot (Cowork / Studio) can reach it. Prints the public /mcp URL.
#
# Usage:  bash scripts/cowork-serve.sh [--port 3000]
#
# SECURITY: the tunnel is anonymous for easy testing — anyone with the URL can
# call the server as your signed-in account. Stop it (Ctrl+C) when done, and
# consider COPILOT_COST_MGMT_READONLY=true while testing.
set -euo pipefail
cd "$(dirname "$0")/.."

PORT=3000
while [ $# -gt 0 ]; do
  case "$1" in
    --port) PORT="$2"; shift 2 ;;
    *) echo "Unknown arg: $1" >&2; exit 2 ;;
  esac
done

if ! command -v devtunnel >/dev/null 2>&1; then
  echo "devtunnel not found. Install it: https://aka.ms/devtunnel" >&2
  exit 1
fi

# Prefer the installed CLI; fall back to the local build.
if command -v copilot-cost-management-mcp >/dev/null 2>&1; then
  copilot-cost-management-mcp http --port "$PORT" &
else
  node dist/index.js http --port "$PORT" &
fi
SERVER_PID=$!
trap 'kill "$SERVER_PID" 2>/dev/null || true' EXIT INT TERM
sleep 1

echo ""
echo "Local MCP endpoint: http://localhost:${PORT}/mcp"
echo "Opening an anonymous dev tunnel for port ${PORT} (use --protocol http: it describes the LOCAL service)…"
echo "Copy the printed HTTPS URL and append /mcp for the Copilot connector."
echo ""
# Hosts a temporary tunnel and prints its public HTTPS URL; blocks until Ctrl+C.
devtunnel host -p "$PORT" --protocol http --allow-anonymous
