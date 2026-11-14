/**
 * Thin Microsoft Graph client specialised for the Copilot Cost Management API.
 *
 * Responsibilities:
 *   - acquire + cache an access token from the resolved credential
 *   - build URLs under /{version}/copilot/costManagement
 *   - serialise OData query options ($filter, $top, $skiptoken, $orderby, $expand)
 *   - set Content-Type / If-Match headers
 *   - transparently retry on 429 (honouring Retry-After)
 *   - surface Graph's OData error envelope as a typed {@link GraphError}
 */

import type { AccessToken, TokenCredential } from "@azure/identity";
import type { Config } from "./config.js";

export type HttpMethod = "GET" | "POST" | "PATCH" | "DELETE" | "PUT";

export interface RequestOptions {
  /** OData / arbitrary query options. Values are stringified; undefined/null dropped. */
  query?: Record<string, string | number | boolean | undefined | null>;
  /** JSON request body. */
  body?: unknown;
  /** ETag for optimistic concurrency → sent as the If-Match header. */
  ifMatch?: string;
  /** Extra headers (lowest precedence). */
  headers?: Record<string, string>;
}

export interface GraphResponse<T = unknown> {
  status: number;
  headers: Headers;
  /** Parsed JSON body, or undefined for 204/empty responses. */
  data: T | undefined;
  /** Strong ETag from the body (@odata.etag) or the ETag response header, if any. */
  etag?: string;
}

/** Structured error carrying Graph's OData `error`/`innerError` detail. */
export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code?: string,
    readonly innerCode?: string,
    readonly requestId?: string,
    readonly body?: unknown,
  ) {
    super(message);
    this.name = "GraphError";
  }

  toJSON() {
    return {
      error: this.message,
      status: this.status,
      code: this.code,
      innerErrorCode: this.innerCode,
      requestId: this.requestId,
      details: this.body,
    };
  }
}

const MAX_RETRIES = 3;
const COST_MGMT_ROOT = "copilot/costManagement";

export class GraphClient {
  private cachedToken?: AccessToken;
  private readonly versionRoot: string;

  constructor(
    private readonly credential: TokenCredential,
    private readonly config: Config,
  ) {
    this.versionRoot = `${config.graphBaseUrl}/${config.apiVersion}`;
  }

  /** Force a token acquisition (used by `--check` to trigger interactive sign-in). */
  async ensureToken(): Promise<void> {
    await this.getToken();
  }

  /** Acquire a token, reusing the cache until ~2 minutes before expiry. */
  private async getToken(): Promise<string> {
    const now = Date.now();
    if (this.cachedToken && this.cachedToken.expiresOnTimestamp - now > 120_000) {
      return this.cachedToken.token;
    }
    const token = await this.credential.getToken(this.config.scope);
    if (!token) {
      throw new Error("Failed to acquire a Microsoft Graph access token.");
    }
    this.cachedToken = token;
    return token.token;
  }

  /**
   * Resolve a path or absolute URL into a request URL.
   * - Absolute http(s) URLs (e.g. an @odata.nextLink) are used as-is.
   * - A path starting with "/" is resolved against the version root.
   * - Anything else is treated as relative to the cost-management root.
   */
  private buildUrl(pathOrUrl: string, query?: RequestOptions["query"]): string {
    let url: string;
    if (/^https?:\/\//i.test(pathOrUrl)) {
      url = pathOrUrl;
    } else if (pathOrUrl.startsWith("/")) {
      url = `${this.versionRoot}${pathOrUrl}`;
    } else {
      url = `${this.versionRoot}/${pathOrUrl}`;
    }

    if (query) {
      const u = new URL(url);
      for (const [key, value] of Object.entries(query)) {
        if (value === undefined || value === null || value === "") continue;
        u.searchParams.set(key, String(value));
      }
      url = u.toString();
    }
    return url;
  }

  async request<T = unknown>(
    method: HttpMethod,
    pathOrUrl: string,
    options: RequestOptions = {},
  ): Promise<GraphResponse<T>> {
    const url = this.buildUrl(pathOrUrl, options.query);

    for (let attempt = 0; ; attempt++) {
      const token = await this.getToken();
      const headers: Record<string, string> = {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...(options.headers ?? {}),
      };
      if (options.ifMatch) headers["If-Match"] = options.ifMatch;

      let payload: string | undefined;
      if (options.body !== undefined) {
        payload = JSON.stringify(options.body);
        headers["Content-Type"] = "application/json";
      }

      const res = await fetch(url, { method, headers, body: payload });

      // Throttling: honour Retry-After and retry a bounded number of times.
      if (res.status === 429 && attempt < MAX_RETRIES) {
        const retryAfter = Number(res.headers.get("retry-after")) || 2 ** attempt;
        await delay(retryAfter * 1000);
        continue;
      }

      const text = await res.text();
      const data = text ? safeJsonParse(text) : undefined;

      if (!res.ok) {
        throw toGraphError(res, data);
      }

      const etag =
        extractEtag(data) ?? res.headers.get("etag") ?? undefined;

      return { status: res.status, headers: res.headers, data: data as T, etag };
    }
  }

  // ── Convenience wrappers ───────────────────────────────────────────────────

  /**
   * Build a full path under /{version}/copilot/costManagement.
   * Exposed so tool modules can compose child paths without repeating the root.
   * e.g. GraphClient.path("/spendingPolicies") → "/copilot/costManagement/spendingPolicies"
   */
  static path(subPath: string): string {
    const clean = subPath.startsWith("/") ? subPath : `/${subPath}`;
    return `/${COST_MGMT_ROOT}${clean}`;
  }

  /** GET a cost-management sub-path (e.g. "/spendingPolicies"). */
  cmGet<T = unknown>(subPath: string, options?: RequestOptions) {
    return this.request<T>("GET", GraphClient.path(subPath), options);
  }

  /**
   * Build an absolute Graph resource URL under the configured version, e.g.
   * resourceUrl("/groups/{id}") → "https://graph.microsoft.com/beta/groups/{id}".
   * Used for @odata.id reference-link bodies.
   */
  resourceUrl(path: string): string {
    const clean = path.startsWith("/") ? path : `/${path}`;
    return `${this.versionRoot}${clean}`;
  }

  /**
   * Resolve the current ETag of a cost-management resource by GETting it.
   * Used when a caller omits the ETag on a mutating operation.
   */
  async resolveEtag(subPath: string): Promise<string | undefined> {
    const res = await this.request("GET", GraphClient.path(subPath));
    return res.etag;
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function safeJsonParse(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text };
  }
}

function extractEtag(data: unknown): string | undefined {
  if (data && typeof data === "object" && "@odata.etag" in data) {
    const v = (data as Record<string, unknown>)["@odata.etag"];
    return typeof v === "string" ? v : undefined;
  }
  return undefined;
}

function toGraphError(res: Response, data: unknown): GraphError {
  const requestId = res.headers.get("request-id") ?? res.headers.get("client-request-id") ?? undefined;
  let message = `Microsoft Graph request failed with HTTP ${res.status} ${res.statusText}`;
  let code: string | undefined;
  let innerCode: string | undefined;

  if (data && typeof data === "object" && "error" in data) {
    const err = (data as Record<string, any>).error;
    if (err && typeof err === "object") {
      code = typeof err.code === "string" ? err.code : undefined;
      if (typeof err.message === "string") message = err.message;
      const inner = err.innerError;
      if (inner && typeof inner === "object" && typeof inner.code === "string") {
        innerCode = inner.code;
      }
    }
  }

  return new GraphError(message, res.status, code, innerCode, requestId, data);
}
