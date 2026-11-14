/**
 * Pure-JS persistent token cache for the delegated (device-code) flow.
 *
 * MSAL serialises its token cache to a single JSON file we own, so a device-code
 * sign-in survives process restarts (the server, and Claude relaunching it, read
 * the cached refresh token and get access tokens silently). No native keychain
 * dependency — which keeps installation trivial.
 *
 * SECURITY: the file contains refresh tokens in plain text. It's written with
 * 0600 permissions (owner-only) in a 0700 directory, but it is NOT encrypted at
 * rest. Treat it like an SSH private key. Use `logout` to remove it.
 */

import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import type { Config } from "./config.js";

/** Resolve the cache file path for a given client id. */
export function resolveCacheFile(config: Config, clientId: string): string {
  if (config.tokenCacheFile) return config.tokenCacheFile;
  const safeClient = clientId.replace(/[^a-zA-Z0-9._-]/g, "_");
  return join(config.cacheDir, `token-cache-${safeClient}.json`);
}

/** Build an MSAL cache plugin backed by the given file. */
export function makeCachePlugin(file: string): ICachePlugin {
  return {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      try {
        if (existsSync(file)) {
          ctx.tokenCache.deserialize(readFileSync(file, "utf8"));
        }
      } catch {
        // Corrupt/unreadable cache: start fresh rather than crash.
      }
    },
    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (!ctx.cacheHasChanged) return;
      try {
        mkdirSync(dirname(file), { recursive: true, mode: 0o700 });
        writeFileSync(file, ctx.tokenCache.serialize(), { mode: 0o600 });
      } catch {
        // Non-fatal: tokens simply won't persist this run.
      }
    },
  };
}

/** Delete the cache file. Returns true if a file was removed. */
export function deleteCacheFile(file: string): boolean {
  if (!existsSync(file)) return false;
  rmSync(file, { force: true });
  return true;
}
