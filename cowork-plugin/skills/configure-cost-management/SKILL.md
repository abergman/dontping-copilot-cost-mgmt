---
name: configure-cost-management
description: |
  Set up and manage Microsoft 365 Copilot cost governance using the Copilot
  Cost Management tools. Use when the user wants to "set up Copilot spending
  limits", "create a spending policy", "cap Copilot credits for a group",
  "add a usage/threshold alert", "review or approve credit requests", "check
  the tenant credit balance", or "see where Copilot credits are going".
license: MIT
metadata:
  author: Andreas Bergman
  version: "0.5"
---

# Configure Copilot Cost Management

Guides an administrator through governing and observing Microsoft 365 Copilot /
AI **credit** consumption with the Copilot Cost Management connector tools. All
amounts are Copilot credits — never currency.

## Before you start

- These actions need the right Microsoft Entra role (Billing / AI / Global
  Administrator for writes; readers for reports).
- Tenancy: tools run against the **active tenant**. If the user manages several,
  confirm it first with `list_tenants` / `current_tenant`, and switch with
  `use_tenant`. If a tool reports "not signed in", run `start_login` and have the
  admin complete the browser sign-in, then retry.
- IDs are Entra **object GUIDs** (users, groups), not emails/UPNs.

## Decide what the user wants

1. **See the current picture** → go to *Observe*.
2. **Put a cap in place** → go to *Create a spending policy*.
3. **Get alerted near a limit** → go to *Add threshold alerts*.
4. **Handle users asking for more credits** → go to *Triage credit requests*.

## Observe (read-only)

- Tenant-wide snapshot: `get_tenant_credit_balance` (one row per billing method).
- Where credits went for a day: `list_service_usage_records` (by service) or
  `list_user_usage_records` / `list_group_usage_records` (rank by consumedQuantity).
  These need a time selector: `startDate` and `endDate` = exactly one completed
  UTC day (endDate = startDate + 1 day).
- A month by funding class (prepaid / pay-as-you-go / unbilled overage /
  non-billable): `list_funding_usage_records` with `month` = `YYYY-MM`.
- Diagnose one user: `get_user_balance` (add expandServiceBalances) or
  `list_user_service_balances`.

Summarize findings as a short table and call out the highest consumers.

## Create a spending policy

A policy caps credits for an audience (tenant-wide or selected groups) and is
**created disabled**.

1. Gather: scope (`allUsers` or `selectedGroups`), a display name (≥10 chars, no
   special characters), the total `periodLimit`, the `perUserLimit`, and which
   services it governs (e.g. `cowork`, `workIq`).
2. `create_spending_policy` with those values.
3. For a group-scoped policy, add at least one group:
   `add_policy_assigned_group` (needs the group's object ID).
4. Enable it: `update_spending_policy` with `state: "enabled"`.
5. Confirm with `get_spending_policy`.

Notes: audience type is fixed at creation; to remove a service later, set it to
`disabled` with `update_policy_service_setting` (there is no delete). Watch for
conflicts (e.g. another non-disabled policy already governs the same service).

## Add threshold alerts

Notify-only email alerts when usage approaches a limit.

1. Add at least one recipient first: `add_policy_alert_recipient` (user object ID)
   — alerting can't be enabled without one.
2. `update_policy_alert_settings` with `isEnabled: true` and `thresholds`, e.g.
   an absolute threshold `{ "@odata.type": "#microsoft.graph.absoluteCostManagementAlertThreshold", "value": <n>, "unit": "copilotCredit" }`,
   or a percentage threshold whose integer value is 1–100.
3. Verify with `get_policy_alert_settings` and `list_policy_alert_recipients`.

## Triage credit requests

Users request more credits outside Graph; admins approve or reject.

1. `list_credit_requests` (filter/paging as needed) to see pending ones.
2. Review each with `get_credit_request` (note its ETag).
3. Decide one: `decide_credit_request` (`approved` or `rejected`). For many at
   once: `batch_decide_credit_requests` with up to 100 items, each carrying its
   current ETag.

Only `pending → approved` and `pending → rejected` are allowed. Summarize what
you decided and anything you left pending.

## Good practice

- Prefer the least-privileged read tools when the user only wants to *see*
  something.
- Echo back the exact limits, services, groups, and recipients before making a
  change, and confirm destructive actions (enabling/disabling a policy,
  approving/rejecting requests).
- After any change, read the resource back and report the new state plainly.
