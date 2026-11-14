/**
 * Multi-tenant management tools. These operate on the local tenant registry
 * (not on cloud data), so they're available even in read-only mode.
 *
 *   list_tenants     → show configured tenants + which is active
 *   current_tenant   → the active tenant + its auth mode / sign-in state
 *   add_tenant       → register a tenant (id, label, optional per-tenant client id)
 *   use_tenant       → switch the active tenant
 *   remove_tenant    → unregister a tenant
 */

import { z } from "zod";
import { PersistentDeviceCodeCredential } from "../auth.js";
import type { TenantManager } from "../tenants.js";
import type { ToolContext } from "./context.js";
import { handle, ok } from "./helpers.js";

function safeMode(manager: TenantManager, id: string): string {
  try {
    return manager.resolvedFor(id).mode;
  } catch {
    return "unconfigured (no client id)";
  }
}

export function registerTenantTools({ server, manager }: ToolContext): void {
  server.registerTool(
    "list_tenants",
    {
      title: "List configured tenants",
      description:
        "List the Microsoft 365 tenants registered with this server, which one is active, and each " +
        "tenant's auth mode. All tool calls run against the active tenant (switch with use_tenant).",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    () =>
      handle(async () => {
        const active = manager.activeId();
        const tenants = manager.list().map((t) => ({
          id: t.id,
          label: t.label,
          clientId: t.clientId,
          active: t.id === active,
          authMode: safeMode(manager, t.id),
        }));
        return ok(
          { active, count: tenants.length, tenants },
          tenants.length ? `${tenants.length} tenant(s); active: ${active ?? "none"}.` : "No tenants registered yet. Use add_tenant.",
        );
      }),
  );

  server.registerTool(
    "current_tenant",
    {
      title: "Show the active tenant",
      description: "Show the active tenant and its auth mode / sign-in status.",
      inputSchema: {},
      annotations: { readOnlyHint: true, openWorldHint: true },
    },
    () =>
      handle(async () => {
        const entry = manager.activeEntry();
        if (!entry) return ok({ active: null }, "No active tenant. Use add_tenant then use_tenant.");
        const resolved = manager.resolvedFor(entry.id);
        let signedIn: unknown = undefined;
        if (resolved.credential instanceof PersistentDeviceCodeCredential) {
          signedIn = (await resolved.credential.signedInAs()) ?? false;
        }
        return ok(
          { id: entry.id, label: entry.label, clientId: entry.clientId, authMode: resolved.mode, signedIn },
          `Active tenant: ${entry.label ?? entry.id} (${entry.id}), auth: ${resolved.mode}.`,
        );
      }),
  );

  server.registerTool(
    "add_tenant",
    {
      title: "Register a tenant",
      description:
        "Register a Microsoft 365 tenant by its directory (tenant) ID. Optionally give it a label " +
        "and a per-tenant app client id (for the 'separate app per tenant' model; otherwise the " +
        "shared AZURE_CLIENT_ID is used). By default the new tenant becomes active — then call " +
        "start_login to sign in to it (device-code auth).",
      inputSchema: {
        tenantId: z.string().min(1).describe("Directory (tenant) ID (GUID)."),
        label: z.string().optional().describe("Friendly name, e.g. the customer's name."),
        clientId: z.string().optional().describe("Per-tenant app client id (optional override)."),
        activate: z.boolean().optional().describe("Make this the active tenant (default true)."),
      },
      annotations: { readOnlyHint: false, openWorldHint: false },
    },
    ({ tenantId, label, clientId, activate }) =>
      handle(async () => {
        const entry = manager.add({ id: tenantId, label, clientId });
        if (activate !== false) manager.use(entry.id);
        const mode = safeMode(manager, entry.id);
        return ok(
          { tenant: entry, active: manager.activeId(), authMode: mode },
          `Registered ${entry.label ?? entry.id}.` +
            (activate !== false
              ? mode === "device-code"
                ? " It's now active — call start_login to sign in."
                : " It's now active."
              : ""),
        );
      }),
  );

  server.registerTool(
    "use_tenant",
    {
      title: "Switch the active tenant",
      description:
        "Make a registered tenant the active one; subsequent tool calls run against it. If it uses " +
        "device-code auth and isn't signed in yet, call start_login next.",
      inputSchema: {
        tenantId: z.string().min(1).describe("Tenant ID (or the label is not accepted — use the ID)."),
      },
      annotations: { readOnlyHint: false, openWorldHint: false },
    },
    ({ tenantId }) =>
      handle(async () => {
        const entry = manager.use(tenantId);
        const resolved = manager.resolvedFor(entry.id);
        let note = `Active tenant is now ${entry.label ?? entry.id}.`;
        if (resolved.credential instanceof PersistentDeviceCodeCredential) {
          const who = await resolved.credential.signedInAs();
          note += who ? ` Signed in as ${who}.` : " Not signed in — call start_login.";
        }
        return ok({ active: entry.id, authMode: resolved.mode }, note);
      }),
  );

  server.registerTool(
    "remove_tenant",
    {
      title: "Unregister a tenant",
      description:
        "Remove a tenant from the registry. This does not delete its cached token file; sign out " +
        "first with sign_out if you want to clear credentials.",
      inputSchema: {
        tenantId: z.string().min(1).describe("Tenant ID to remove."),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
    },
    ({ tenantId }) =>
      handle(async () => {
        const removed = manager.remove(tenantId);
        return ok(
          { removed, active: manager.activeId() },
          removed ? `Removed tenant ${tenantId}. Active is now ${manager.activeId() ?? "none"}.` : `Tenant ${tenantId} was not registered.`,
        );
      }),
  );
}
