/**
 * In-chat authentication tools for the delegated device-code flow.
 *
 *   start_login         → begin device-code sign-in; returns a URL + code
 *   check_login_status  → report whether we're signed in (and any pending login)
 *   sign_out            → clear the cached token
 *
 * These are registered only when the server is using device-code auth; for
 * static-token or client-credentials auth there's nothing to sign in to.
 */

import { PersistentDeviceCodeCredential } from "../auth.js";
import type { ToolContext } from "./context.js";
import { handle, ok } from "./helpers.js";

export function registerAuthTools({ server, auth, scope }: ToolContext): void {
  const cred = auth.credential;
  if (!(cred instanceof PersistentDeviceCodeCredential)) return;

  server.registerTool(
    "start_login",
    {
      title: "Sign in to Microsoft (device code)",
      description:
        "Begin interactive Microsoft sign-in for this server. Returns a verification URL and a " +
        "short code: open the URL in a browser, enter the code, and complete sign-in with an admin " +
        "account. Sign-in finishes in the background and the token is cached for reuse. After you " +
        "sign in, retry your request (or call check_login_status). If already signed in, this is a " +
        "no-op.",
      inputSchema: {},
      annotations: { readOnlyHint: false, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const who = await cred.signedInAs();
        if (who) {
          return ok(
            { signedIn: true, username: who },
            `Already signed in as ${who}. Use sign_out to switch accounts.`,
          );
        }
        const { prompt, alreadyPending } = await cred.beginLogin(scope);
        return ok(
          {
            verificationUri: prompt.verificationUri,
            userCode: prompt.userCode,
            instructions: prompt.message,
            alreadyPending,
          },
          `To sign in:\n  1. Open ${prompt.verificationUri}\n  2. Enter the code: ${prompt.userCode}\n  3. Sign in with an admin account.\n\nThen retry your request, or call check_login_status to confirm.`,
        );
      }),
  );

  server.registerTool(
    "check_login_status",
    {
      title: "Check sign-in status",
      description:
        "Report whether the server is signed in to Microsoft, the signed-in username, whether a " +
        "device-code login is still pending, and the outcome of the last login attempt.",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const status = await cred.loginStatus();
        const note = status.signedIn
          ? `Signed in as ${status.username}.`
          : status.pending
            ? "A sign-in is pending — finish it in your browser, then retry."
            : "Not signed in. Call start_login to begin.";
        return ok(status, note);
      }),
  );

  server.registerTool(
    "sign_out",
    {
      title: "Sign out (clear cached token)",
      description:
        "Remove this server's cached Microsoft token. The next request will require a fresh " +
        "start_login. Use this to switch accounts.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const removed = await cred.logout();
        return ok(
          { signedOut: true, cacheRemoved: removed },
          removed ? "Signed out — cached token removed." : "No cached token to remove.",
        );
      }),
  );
}
