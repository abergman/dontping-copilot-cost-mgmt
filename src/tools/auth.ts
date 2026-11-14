/**
 * In-chat authentication tools, operating on the ACTIVE tenant.
 *
 *   start_login         → begin device-code sign-in for the active tenant
 *   check_login_status  → sign-in status of the active tenant
 *   sign_out            → clear the active tenant's cached token
 *
 * Login applies to device-code tenants; for app-only / static-token tenants the
 * tools report that interactive sign-in isn't applicable.
 */

import { PersistentDeviceCodeCredential } from "../auth.js";
import type { ToolContext } from "./context.js";
import { handle, ok } from "./helpers.js";

export function registerAuthTools({ server, manager, scope }: ToolContext): void {
  server.registerTool(
    "start_login",
    {
      title: "Sign in to Microsoft (active tenant)",
      description:
        "Begin interactive Microsoft sign-in for the ACTIVE tenant. Returns a verification URL and " +
        "a short code: open the URL, enter the code, and sign in with an admin of that tenant. The " +
        "token is cached per tenant and reused. Switch tenants with use_tenant first if needed. " +
        "No-op if already signed in; not applicable for app-only/static-token tenants.",
      inputSchema: {},
      annotations: { readOnlyHint: false, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const resolved = manager.activeResolved();
        const tenant = manager.activeId();
        if (!(resolved.credential instanceof PersistentDeviceCodeCredential)) {
          return ok(
            { tenant, mode: resolved.mode },
            `Active tenant (${tenant}) uses ${resolved.mode} auth — interactive login isn't applicable.`,
          );
        }
        const cred = resolved.credential;
        const who = await cred.signedInAs();
        if (who) {
          return ok(
            { tenant, signedIn: true, username: who },
            `Already signed in as ${who} on tenant ${tenant}. Use sign_out to switch accounts.`,
          );
        }
        const { prompt, alreadyPending } = await cred.beginLogin(scope);
        return ok(
          {
            tenant,
            verificationUri: prompt.verificationUri,
            userCode: prompt.userCode,
            instructions: prompt.message,
            alreadyPending,
          },
          `To sign in to tenant ${tenant}:\n  1. Open ${prompt.verificationUri}\n  2. Enter the code: ${prompt.userCode}\n  3. Sign in with an admin account.\n\nThen retry your request, or call check_login_status.`,
        );
      }),
  );

  server.registerTool(
    "check_login_status",
    {
      title: "Check sign-in status (active tenant)",
      description:
        "Report whether the active tenant is signed in, the username, whether a device-code login is " +
        "pending, and the last login outcome.",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const resolved = manager.activeResolved();
        const tenant = manager.activeId();
        if (!(resolved.credential instanceof PersistentDeviceCodeCredential)) {
          return ok(
            { tenant, mode: resolved.mode, signedIn: resolved.mode !== "static-token" ? "n/a (app-only)" : "n/a" },
            `Active tenant (${tenant}) uses ${resolved.mode} auth; no interactive sign-in state.`,
          );
        }
        const status = await resolved.credential.loginStatus();
        const note = status.signedIn
          ? `Tenant ${tenant}: signed in as ${status.username}.`
          : status.pending
            ? `Tenant ${tenant}: sign-in pending — finish it in your browser, then retry.`
            : `Tenant ${tenant}: not signed in. Call start_login.`;
        return ok({ tenant, ...status }, note);
      }),
  );

  server.registerTool(
    "sign_out",
    {
      title: "Sign out of the active tenant",
      description:
        "Remove the active tenant's cached Microsoft token. The next request needs a fresh " +
        "start_login. Use this to switch accounts within a tenant.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const resolved = manager.activeResolved();
        const tenant = manager.activeId();
        if (!(resolved.credential instanceof PersistentDeviceCodeCredential)) {
          return ok({ tenant, mode: resolved.mode }, `Active tenant (${tenant}) has no cached token to remove.`);
        }
        const removed = await resolved.credential.logout();
        return ok(
          { tenant, signedOut: true, cacheRemoved: removed },
          removed ? `Signed out of tenant ${tenant}.` : `No cached token for tenant ${tenant}.`,
        );
      }),
  );
}
