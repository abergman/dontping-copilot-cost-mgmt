/**
 * Runtime configuration, read from environment variables.
 *
 * Nothing here is secret-logged; auth material is only ever handed to the
 * Azure Identity credential chain in {@link ./auth.ts}.
 */

export interface Config {
  /** Graph host, e.g. https://graph.microsoft.com */
  graphBaseUrl: string;
  /** Graph API version segment, e.g. "beta". */
  apiVersion: string;
  /** OAuth scope requested from the token credential. */
  scope: string;
  /** When true, only read-only tools are exposed. */
  readOnly: boolean;
}

function envFlag(value: string | undefined): boolean {
  if (!value) return false;
  return ["1", "true", "yes", "on"].includes(value.trim().toLowerCase());
}

export function loadConfig(): Config {
  const graphBaseUrl = (process.env.GRAPH_BASE_URL ?? "https://graph.microsoft.com").replace(/\/+$/, "");
  const apiVersion = (process.env.GRAPH_API_VERSION ?? "beta").replace(/^\/+|\/+$/g, "");
  const scope = process.env.GRAPH_SCOPE ?? "https://graph.microsoft.com/.default";
  const readOnly = envFlag(process.env.COPILOT_COST_MGMT_READONLY);

  return { graphBaseUrl, apiVersion, scope, readOnly };
}

export const SERVER_NAME = "copilot-cost-management";
export const SERVER_VERSION = "0.1.0";
