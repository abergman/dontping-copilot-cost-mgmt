/**
 * Multi-tenant support.
 *
 * A persisted registry of tenants plus a per-tenant credential/GraphClient
 * cache. One "active" tenant drives every tool call; switching is done from the
 * chat (use_tenant) or the CLI. Each tenant gets its own token cache file, so
 * device-code sign-ins don't collide across tenants.
 *
 * Credential selection per tenant comes from the shared base auth:
 *   - static token  → single tenant
 *   - client secret → app-only (per tenant via that tenant's authority)
 *   - else          → persistent device-code (per tenant), full access
 * A tenant may override the client id (the "separate app per tenant" model).
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { join } from "node:path";
import { buildTenantCredential, type BaseAuth, type ResolvedAuth } from "./auth.js";
import type { Config } from "./config.js";
import { GraphClient } from "./graph.js";

export interface TenantEntry {
  id: string;
  label?: string;
  /** Optional per-tenant app client id (overrides the shared AZURE_CLIENT_ID). */
  clientId?: string;
  addedAt: string;
}

interface RegistryData {
  active?: string;
  tenants: TenantEntry[];
}

/** Thrown when a tool needs a tenant but none is selected/configured. */
export class NoActiveTenantError extends Error {
  constructor() {
    super(
      "No active tenant selected. Use the `add_tenant` tool to register a tenant " +
        "(tenant id + label), then `use_tenant` to select it. `list_tenants` shows what's configured.",
    );
    this.name = "NoActiveTenantError";
  }
}

export class TenantManager {
  private data: RegistryData = { tenants: [] };
  private readonly file: string;
  private readonly resolvedCache = new Map<string, ResolvedAuth>();
  private readonly clientCache = new Map<string, GraphClient>();
  private readonly proxy: GraphClient;

  constructor(
    private readonly config: Config,
    private readonly base: BaseAuth,
  ) {
    this.file = join(config.cacheDir, "tenants.json");
    this.load();
    this.seed();
    this.proxy = this.makeDynamicClient();
  }

  // ── Persistence ───────────────────────────────────────────────────────────

  private load(): void {
    try {
      if (existsSync(this.file)) {
        const parsed = JSON.parse(readFileSync(this.file, "utf8")) as RegistryData;
        if (parsed && Array.isArray(parsed.tenants)) this.data = parsed;
      }
    } catch {
      this.data = { tenants: [] };
    }
  }

  private save(): void {
    try {
      mkdirSync(dirname(this.file), { recursive: true, mode: 0o700 });
      writeFileSync(this.file, JSON.stringify(this.data, null, 2) + "\n", { mode: 0o600 });
    } catch {
      // Non-fatal: registry just won't persist this run.
    }
  }

  /** Seed from env on first run, and keep `active` valid. */
  private seed(): void {
    if (this.data.tenants.length === 0) {
      if (this.base.staticToken) {
        this.data.tenants.push({ id: "static-token", label: "Static token", addedAt: now() });
      } else if (this.base.defaultTenantId) {
        this.data.tenants.push({ id: this.base.defaultTenantId, label: "Default", addedAt: now() });
      }
    }
    if (!this.data.active || !this.getEntry(this.data.active)) {
      this.data.active = this.data.tenants[0]?.id;
    }
    this.save();
  }

  // ── Registry operations ─────────────────────────────────────────────────────

  list(): TenantEntry[] {
    return [...this.data.tenants];
  }

  getEntry(id: string): TenantEntry | undefined {
    return this.data.tenants.find((t) => t.id === id);
  }

  activeId(): string | undefined {
    return this.data.active;
  }

  activeEntry(): TenantEntry | undefined {
    return this.data.active ? this.getEntry(this.data.active) : undefined;
  }

  add(entry: { id: string; label?: string; clientId?: string }): TenantEntry {
    const id = entry.id.trim();
    if (!id) throw new Error("Tenant id is required.");
    let existing = this.getEntry(id);
    if (existing) {
      if (entry.label !== undefined) existing.label = entry.label;
      if (entry.clientId !== undefined) existing.clientId = entry.clientId;
    } else {
      existing = { id, label: entry.label, clientId: entry.clientId, addedAt: now() };
      this.data.tenants.push(existing);
    }
    if (!this.data.active) this.data.active = id;
    this.invalidate(id);
    this.save();
    return existing;
  }

  remove(id: string): boolean {
    const before = this.data.tenants.length;
    this.data.tenants = this.data.tenants.filter((t) => t.id !== id);
    const removed = this.data.tenants.length < before;
    if (removed) {
      if (this.data.active === id) this.data.active = this.data.tenants[0]?.id;
      this.invalidate(id);
      this.save();
    }
    return removed;
  }

  use(id: string): TenantEntry {
    const entry = this.getEntry(id);
    if (!entry) throw new Error(`Tenant '${id}' is not registered. Use add_tenant first.`);
    this.data.active = id;
    this.save();
    return entry;
  }

  private invalidate(id: string): void {
    this.resolvedCache.delete(id);
    this.clientCache.delete(id);
  }

  // ── Credentials / clients ────────────────────────────────────────────────────

  resolvedFor(id: string): ResolvedAuth {
    let resolved = this.resolvedCache.get(id);
    if (!resolved) {
      const entry = this.getEntry(id);
      resolved = buildTenantCredential({
        tenantId: id,
        base: this.base,
        config: this.config,
        clientIdOverride: entry?.clientId,
      });
      this.resolvedCache.set(id, resolved);
    }
    return resolved;
  }

  clientFor(id: string): GraphClient {
    let client = this.clientCache.get(id);
    if (!client) {
      client = new GraphClient(this.resolvedFor(id).credential, this.config);
      this.clientCache.set(id, client);
    }
    return client;
  }

  /** The active tenant's resolved auth, or throw if none active. */
  activeResolved(): ResolvedAuth {
    const id = this.data.active;
    if (!id) throw new NoActiveTenantError();
    return this.resolvedFor(id);
  }

  /** The active tenant's GraphClient, or throw if none active. */
  activeClient(): GraphClient {
    const id = this.data.active;
    if (!id) throw new NoActiveTenantError();
    return this.clientFor(id);
  }

  /** A stable GraphClient facade that always routes to the active tenant. */
  dynamicClient(): GraphClient {
    return this.proxy;
  }

  private makeDynamicClient(): GraphClient {
    const self = this;
    return new Proxy({} as GraphClient, {
      get(_target, prop, receiver) {
        const client = self.activeClient(); // throws NoActiveTenantError if none
        const value = Reflect.get(client as object, prop, receiver);
        return typeof value === "function" ? value.bind(client) : value;
      },
    });
  }
}

function now(): string {
  return new Date().toISOString();
}
