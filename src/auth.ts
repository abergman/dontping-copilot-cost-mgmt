/**
 * Credential resolution for Microsoft Graph.
 *
 * Strategies, chosen by which environment variables are present (in order):
 *
 *   1. GRAPH_ACCESS_TOKEN                      → static bearer token (no refresh)
 *   2. AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET
 *                                              → client credentials (app-only)
 *   3. AZURE_TENANT_ID + AZURE_CLIENT_ID       → persistent device-code (delegated)
 *
 * The device-code credential (3) is MSAL-backed with a persistent on-disk token
 * cache, so the user signs in ONCE (via the `login` command) and the token is
 * reused across restarts. For the running server, token acquisition is SILENT:
 * if nothing is cached (or it expired), getToken throws NotSignedInError with an
 * actionable message rather than blocking on an interactive prompt.
 */

import type { AccessToken, TokenCredential, GetTokenOptions } from "@azure/identity";
import { ClientSecretCredential } from "@azure/identity";
import {
  PublicClientApplication,
  LogLevel,
  type AuthenticationResult,
  type AccountInfo,
} from "@azure/msal-node";
import type { Config } from "./config.js";
import { deleteCacheFile, makeCachePlugin, resolveCacheFile } from "./tokenCache.js";

/** The fields of MSAL's device-code response that we surface to the user. */
export interface DeviceCodePrompt {
  userCode: string;
  verificationUri: string;
  message: string;
}

export const NOT_SIGNED_IN =
  "Not signed in to Microsoft (no cached credentials, or the session expired). " +
  "Call the `start_login` tool to sign in from here (it returns a sign-in URL + code), " +
  "or run `copilot-cost-management-mcp login` in a terminal; then retry.";

/** Status of an in-progress or last in-chat login attempt. */
export interface LoginStatus {
  signedIn: boolean;
  username?: string;
  pending: boolean;
  last?: { status: "succeeded" | "failed"; username?: string; error?: string };
}

/** Thrown by the device-code credential when no usable cached token exists. */
export class NotSignedInError extends Error {
  constructor(message: string = NOT_SIGNED_IN) {
    super(message);
    this.name = "NotSignedInError";
  }
}

/** A fixed, pre-minted bearer token. Does not refresh. */
class StaticTokenCredential implements TokenCredential {
  constructor(private readonly token: string) {}
  async getToken(_scopes: string | string[], _options?: GetTokenOptions): Promise<AccessToken> {
    return { token: this.token, expiresOnTimestamp: Date.now() + 5 * 60 * 1000 };
  }
}

/**
 * Read an env var for credential use: trim it, and treat an empty value — or an
 * unexpanded "${...}" template (e.g. a blank optional field substituted by a
 * Desktop Extension config form) — as absent.
 */
function cleanEnv(value: string | undefined): string | undefined {
  const v = value?.trim();
  if (!v || v.includes("${")) return undefined;
  return v;
}

function normalizeScopes(scopes: string | string[]): string[] {
  const arr = Array.isArray(scopes) ? scopes : [scopes];
  return arr.flatMap((s) => s.split(/\s+/)).filter(Boolean);
}

/**
 * Delegated device-code credential with a persistent token cache.
 * `getToken` is silent-only; `login` performs the interactive device-code flow.
 */
export class PersistentDeviceCodeCredential implements TokenCredential {
  private readonly pca: PublicClientApplication;

  /** In-flight in-chat login (device code issued, awaiting the user). */
  private pending?: { prompt: DeviceCodePrompt; completion: Promise<AuthenticationResult | null> };
  /** Outcome of the most recent in-chat login attempt. */
  private lastResult?: { status: "succeeded" | "failed"; username?: string; error?: string };

  constructor(tenantId: string, clientId: string, readonly cacheFile: string) {
    this.pca = new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: { cachePlugin: makeCachePlugin(cacheFile) },
      system: {
        loggerOptions: {
          // Route MSAL logs to stderr only — stdout is the JSON-RPC channel.
          loggerCallback: (_level, message) => process.stderr.write(`[msal] ${message}\n`),
          piiLoggingEnabled: false,
          logLevel: LogLevel.Error,
        },
      },
    });
  }

  private async firstAccount(): Promise<AccountInfo | null> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts[0] ?? null;
  }

  /** Silent token acquisition from the cache (refreshing as needed). */
  async getToken(scopes: string | string[], _options?: GetTokenOptions): Promise<AccessToken> {
    const account = await this.firstAccount();
    if (!account) throw new NotSignedInError();
    try {
      const result = await this.pca.acquireTokenSilent({
        account,
        scopes: normalizeScopes(scopes),
      });
      if (!result?.accessToken) throw new NotSignedInError();
      return {
        token: result.accessToken,
        expiresOnTimestamp: result.expiresOn?.getTime() ?? Date.now() + 5 * 60 * 1000,
      };
    } catch {
      // Refresh token missing/expired → interaction required.
      throw new NotSignedInError();
    }
  }

  /** Interactive device-code sign-in; persists the result to the cache. */
  async login(
    scopes: string | string[],
    onPrompt: (info: DeviceCodePrompt) => void,
  ): Promise<AuthenticationResult | null> {
    return this.pca.acquireTokenByDeviceCode({
      scopes: normalizeScopes(scopes),
      deviceCodeCallback: onPrompt,
    });
  }

  /**
   * Start an interactive device-code login WITHOUT blocking: returns the
   * verification URL + user code as soon as the device code is issued, while the
   * flow keeps polling in the background and persists the token on success.
   * Safe to call again while one is pending — it returns the existing prompt.
   */
  async beginLogin(
    scopes: string | string[],
  ): Promise<{ prompt: DeviceCodePrompt; alreadyPending: boolean }> {
    if (this.pending) return { prompt: this.pending.prompt, alreadyPending: true };

    let onPromptResolve!: (p: DeviceCodePrompt) => void;
    const promptReady = new Promise<DeviceCodePrompt>((resolve) => {
      onPromptResolve = resolve;
    });

    const completion = this.login(scopes, (info) =>
      onPromptResolve({
        userCode: info.userCode,
        verificationUri: info.verificationUri,
        message: info.message,
      }),
    );

    // Record the outcome and clear the pending slot when it finishes.
    completion
      .then((r) => {
        this.lastResult = { status: "succeeded", username: r?.account?.username };
      })
      .catch((e) => {
        this.lastResult = { status: "failed", error: e instanceof Error ? e.message : String(e) };
      })
      .finally(() => {
        this.pending = undefined;
      });

    // Resolve once the code is issued, or reject if the flow fails first.
    const prompt = await new Promise<DeviceCodePrompt>((resolve, reject) => {
      let settled = false;
      void promptReady.then((p) => {
        settled = true;
        resolve(p);
      });
      completion.then(
        () => {
          if (!settled) reject(new Error("Login finished before a device code was issued."));
        },
        (e) => {
          if (!settled) reject(e);
        },
      );
    });

    this.pending = { prompt, completion };
    return { prompt, alreadyPending: false };
  }

  /** Current sign-in / in-chat-login status. */
  async loginStatus(): Promise<LoginStatus> {
    const username = await this.signedInAs();
    return {
      signedIn: Boolean(username),
      username,
      pending: Boolean(this.pending),
      last: this.lastResult,
    };
  }

  /** The signed-in account's username, if any. */
  async signedInAs(): Promise<string | undefined> {
    return (await this.firstAccount())?.username;
  }

  /** Remove cached accounts and delete the cache file. */
  async logout(): Promise<boolean> {
    const tc = this.pca.getTokenCache();
    for (const account of await tc.getAllAccounts()) {
      await tc.removeAccount(account);
    }
    return deleteCacheFile(this.cacheFile);
  }
}

export type AuthMode = "static-token" | "client-credentials" | "device-code";

export interface ResolvedAuth {
  credential: TokenCredential;
  mode: AuthMode;
  /** Human description for logs (no secrets). */
  describe: string;
}

/** Shared (tenant-independent) auth material read from the environment. */
export interface BaseAuth {
  staticToken?: string;
  clientId?: string;
  clientSecret?: string;
  /** Seed tenant from AZURE_TENANT_ID, used to pre-populate the registry. */
  defaultTenantId?: string;
}

export function readBaseAuth(): BaseAuth {
  return {
    // A blank optional secret — or an unexpanded "${...}" template left by a
    // Desktop Extension config form — must NOT force app-only mode.
    staticToken: cleanEnv(process.env.GRAPH_ACCESS_TOKEN),
    clientId: cleanEnv(process.env.AZURE_CLIENT_ID),
    clientSecret: cleanEnv(process.env.AZURE_CLIENT_SECRET),
    defaultTenantId: cleanEnv(process.env.AZURE_TENANT_ID),
  };
}

/**
 * Build the credential for a specific tenant from the shared base auth.
 * - static token  → single-tenant bearer (tenant id is informational)
 * - client secret → app-only (per tenant via that tenant's authority)
 * - else          → persistent device-code (per tenant), full access
 *
 * `clientIdOverride` supports the "separate app per tenant" model.
 */
export function buildTenantCredential(opts: {
  tenantId: string;
  base: BaseAuth;
  config: Config;
  clientIdOverride?: string;
}): ResolvedAuth {
  const { tenantId, base, config, clientIdOverride } = opts;

  if (base.staticToken) {
    return {
      credential: new StaticTokenCredential(base.staticToken),
      mode: "static-token",
      describe: "static bearer token (GRAPH_ACCESS_TOKEN)",
    };
  }

  const clientId = clientIdOverride ?? base.clientId;
  if (!clientId) {
    throw new Error(
      "No client ID configured. Set AZURE_CLIENT_ID (shared app), or add the tenant " +
        "with a per-tenant client id.",
    );
  }

  if (base.clientSecret) {
    return {
      credential: new ClientSecretCredential(tenantId, clientId, base.clientSecret),
      mode: "client-credentials",
      describe: `client credentials (app ${clientId} in tenant ${tenantId})`,
    };
  }

  const cacheFile = resolveCacheFile(config, tenantId, clientId);
  return {
    credential: new PersistentDeviceCodeCredential(tenantId, clientId, cacheFile),
    mode: "device-code",
    describe: `device code (app ${clientId} in tenant ${tenantId}); token cache: ${cacheFile}`,
  };
}
