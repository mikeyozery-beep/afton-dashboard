# Admin-consent request for IT — Executive Dashboard (read-only SharePoint)

**Requested by:** Michael Ozery (mozery@aftonprop.com)
**What I need:** Admin approval of a pending consent request so an automated internal
dashboard can read one SharePoint list.

## Approve this
- **Application:** Microsoft Graph Command Line Tools
- **Application (client) ID:** `14d82eec-204b-4c2f-b7e8-296a70dab67e`  (Microsoft's first-party public client)
- **Permission to grant:** Microsoft Graph → **`Sites.Read.All`** → **Delegated** type
  ("Read items in all site collections")
- **Tenant:** aftonprop.com  (Tenant ID `3f73f992-5f90-448f-a139-a3ceb1ad94e5`)

## How to approve (either method)
**A) Approve the pending request**
1. Microsoft Entra admin center (entra.microsoft.com) → Identity → Applications →
   **Enterprise applications** → **Admin consent requests**.
2. Find the request for *Microsoft Graph Command Line Tools* / `Sites.Read.All` (requested by mozery).
3. **Review → Approve**.

**B) Or grant consent directly**
1. Enterprise applications → search **Microsoft Graph Command Line Tools**.
2. **Security → Permissions** → **Grant admin consent for aftonprop.com**.
3. Confirm Microsoft Graph **Delegated** `Sites.Read.All` is included.

## Why this is low-risk
- **Delegated** (not Application) permission: access is limited to what the *signed-in
  user* (mozery) can already see — it grants no broader data access than that user has.
- **Read-only** (`*.Read.*`). No write/delete.
- **Public client, no secret**: there is no client secret or certificate to manage.
- Used only to count open positions from the "Open Positions" SharePoint list for an
  internal executive dashboard.

## After approval
Please notify Michael — no further action needed from IT. He re-runs a one-line sign-in
and the dashboard begins reading the list live.
