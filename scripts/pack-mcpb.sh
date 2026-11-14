#!/usr/bin/env bash
#
# Build the Claude Desktop Extension bundle (.mcpb).
#
# Produces a single installable file that users drag into
# Claude Desktop → Settings → Extensions. It contains the compiled server, its
# production dependencies, and mcpb/manifest.json (which renders the config form).
#
# Usage:  bash scripts/pack-mcpb.sh
set -euo pipefail
cd "$(dirname "$0")/.."

VERSION="$(node -p "require('./package.json').version")"
OUT="dontping-copilot-cost-management-mcp-${VERSION}.mcpb"
BUILD=".mcpb-build"

echo "Building server…"
npm run build

echo "Staging bundle in ${BUILD}…"
rm -rf "$BUILD" "$OUT"
mkdir -p "$BUILD/server"
cp -r dist/. "$BUILD/server/"
# Runtime doesn't need source maps or type declarations.
find "$BUILD/server" -name '*.map' -delete
find "$BUILD/server" -name '*.d.ts' -delete
cp package.json package-lock.json "$BUILD/"
cp mcpb/manifest.json "$BUILD/manifest.json"

# Keep the manifest version in lockstep with package.json.
node -e "const fs=require('fs');const p='${BUILD}/manifest.json';const m=JSON.parse(fs.readFileSync(p));m.version=require('./package.json').version;fs.writeFileSync(p,JSON.stringify(m,null,2)+'\n')"

echo "Installing production dependencies…"
( cd "$BUILD" && npm ci --omit=dev --ignore-scripts >/dev/null 2>&1 )

echo "Zipping ${OUT}…"
if command -v zip >/dev/null 2>&1; then
  ( cd "$BUILD" && zip -r -q -X "../$OUT" . )
else
  # Fallback: zip with Node if the `zip` binary is unavailable.
  ( cd "$BUILD" && node -e "
    const {execSync}=require('child_process');
    // Prefer python's zipfile if present; otherwise fail loudly.
    try { execSync('python3 -c \"import zipfile,os;z=zipfile.ZipFile(\\'../${OUT}\\',\\'w\\',zipfile.ZIP_DEFLATED);[z.write(os.path.join(r,f),os.path.relpath(os.path.join(r,f),\\'.\\')) for r,_,fs in os.walk(\\'.\\') for f in fs];z.close()\"',{stdio:'inherit'}); }
    catch(e){ console.error('Need zip or python3 to build the bundle'); process.exit(1); }
  " )
fi

rm -rf "$BUILD"
SIZE="$(du -h "$OUT" | cut -f1)"
echo "Built ${OUT} (${SIZE})"
