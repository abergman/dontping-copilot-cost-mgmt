/**
 * Credential resolution for Microsoft Graph.
 *
 * Three mutually-exclusive strategies are supported, chosen by which
 * environment variables are present (checked in this order):
 *
 *   1. GRAPH_ACCESS_TOKEN                      → static bearer token (no refresh)
 *   2. AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET
 *                                              → client credentials (app-only)
 *   3. AZURE_TENANT_ID + AZURE_CLIENT_ID       → device code (delegated, interactive)
 *
 * All strategies expose the standard Azure `TokenCredential` surface so the
 * Graph client can treat them uniformly.
 */

import type { AccessToken, TokenCredential, GetTokenOptions } from "@azure/identity";
import { ClientSecretCredential, DeviceCodeCredential } from "@azure/identity";

/** A fixed, pre-minted bearer token. Does not refresh. */
class StaticTokenCredential implements TokenCredential {
  constructor(private readonly token: string) {}

  // Expires are unknown for an injected token; report a short horizon so the
  // Graph client's cache re-reads it rather than trusting it forever.
  async getToken(_scopes: string | string[], _options?: GetTokenOptions): Promise<AccessToken> {
    return {
      token: this.token,
      expiresOnTimestamp: Date.now() + 5 * 60 * 1000,
    };
  }
}

export type AuthMode = "static-token" | "client-credentials" | "device-code";

export interface ResolvedAuth {
  credential: TokenCredential;
  mode: AuthMode;
  /** Human description for logs (no secrets). */
  describe: string;
}

export function resolveCredential(): ResolvedAuth {
  const token = process.env.GRAPH_ACCESS_TOKEN?.trim();
  const tenantId = process.env.AZURE_TENANT_ID?.trim();
  const clientId = process.env.AZURE_CLIENT_ID?.trim();
  const clientSecret = process.env.AZURE_CLIENT_SECRET?.trim();

  if (token) {
    return {
      credential: new StaticTokenCredential(token),
      mode: "static-token",
      describe: "static bearer token (GRAPH_ACCESS_TOKEN)",
    };
  }

  if (tenantId && clientId && clientSecret) {
    return {
      credential: new ClientSecretCredential(tenantId, clientId, clientSecret),
      mode: "client-credentials",
      describe: `client credentials (app ${clientId} in tenant ${tenantId})`,
    };
  }

  if (tenantId && clientId) {
    return {
      credential: new DeviceCodeCredential({
        tenantId,
        clientId,
        // Route the prompt to stderr so it never corrupts the stdio JSON-RPC
        // channel on stdout.
        userPromptCallback: (info) => {
          process.stderr.write(`\n[auth] ${info.message}\n`);
        },
      }),
      mode: "device-code",
      describe: `device code (app ${clientId} in tenant ${tenantId})`,
    };
  }

  throw new Error(
    "No Microsoft Graph credentials found. Set one of:\n" +
      "  • GRAPH_ACCESS_TOKEN (raw bearer token), or\n" +
      "  • AZURE_TENANT_ID + AZURE_CLIENT_ID + AZURE_CLIENT_SECRET (app-only), or\n" +
      "  • AZURE_TENANT_ID + AZURE_CLIENT_ID (delegated device-code sign-in).\n" +
      "See .env.example for details.",
  );
}
