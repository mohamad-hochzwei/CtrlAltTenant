# Schuler M365 Tenant-Standardisierung

Step 1: inventory of all 7 tenants. Read-only.

```
tenants.json ─► Export-Inventory.ps1 ─► exports/<tenant>/*.json ─► Compare-Inventory.ps1 ─► matrix.xlsx
                 Graph, Exchange,        _errors.json    = failed areas (403 = permission)
                 Purview, Teams, DNS     _workloads.json = exported / not licensed
```

## Run

```powershell
./Export-Inventory.ps1 -Tenant SyTech                    # one tenant, all workloads
./Export-Inventory.ps1 -Tenant SyTech -Workload Teams    # one workload
./Export-Inventory.ps1 -Tenant SyTech -Broker Browser    # if WAM fails for EXO/Purview/Teams (embedded terminals)
./Export-Inventory.ps1                                   # all tenants
./Compare-Inventory.ps1                                  # builds matrix.xlsx
```

`tenants.json`: per tenant `clientId` (our app, same value everywhere; tenants without it are skipped) and optional `account` (skips account selection for Exchange and Purview).

Modules: `Microsoft.Graph.Authentication`, `ExchangeOnlineManagement`, `MicrosoftTeams`, `ImportExcel`.

## Coverage

| Workload | Areas | Source |
|---|---|---|
| Entra | Users, guests, groups, apps, consent, PIM, MFA registration, auth methods, device join, cross-tenant, lockout | Graph |
| Conditional Access | Policies, named locations, auth strengths | Graph |
| Intune | Compliance, profiles, settings catalog, ADMX, enrollment, app protection, apps, Autopilot, updates, scripts, devices | Graph |
| SharePoint / OneDrive | Tenant sharing and session settings | Graph |
| Security | Secure Score | Graph |
| Copilot | Licenses, seats | Graph |
| Exchange Online | Org, transport, connectors, rules, domains, DKIM, audit, OWA/mobile/auth policies, mailbox counts | EXO PowerShell |
| Defender for Office 365 | Anti-spam/phish/malware, Safe Links/Attachments, presets, quarantine | EXO PowerShell |
| Purview | DLP, sensitivity labels, auto-labeling, retention | Security & Compliance PowerShell |
| Teams | Teams, federation, guest, meeting/messaging/calling/channel/app policies | Teams PowerShell |
| DNS | MX, SPF, DKIM, DMARC | Public DNS |

Manual (no read-only API): SharePoint unmanaged-device access and site-level sharing, Copilot settings in the M365 admin center. Global Reader can view both in the portals.

## Permissions

Interactive sign-in with MFA, delegated permissions only. Effective rights = consented permissions ∩ role (Global Reader), so nothing can be changed. The script requests no Graph scopes: the list lives on the app registration in Entra.

| What | Role | App approval |
|---|---|---|
| Graph | Global Reader | Our app "hochzwei Inventory", admin consent once per tenant |
| Exchange + Defender | Global Reader | None (Microsoft first-party app) |
| Purview | Global Reader | None |
| Teams | Global Reader | None |
| DNS | None | None |

### App registration (once, hochzwei tenant)

- Supported account types: **Multitenant**.
- Platform "Mobile and desktop applications", redirect URIs `ms-appx-web://microsoft.aad.brokerplugin/<clientId>` (WAM) and `http://localhost`. "Allow public client flows" stays **No** (no device code).
- No certificates or secrets.
- API permissions, Microsoft Graph, type **Delegated**: `Organization.Read.All`, `Directory.Read.All`, `Policy.Read.All`, `Policy.Read.DeviceConfiguration`, `RoleManagement.Read.Directory`, `AuditLog.Read.All`, `SecurityEvents.Read.All`, `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`, `DeviceManagementManagedDevices.Read.All`, `DeviceManagementServiceConfig.Read.All`, `DeviceManagementScripts.Read.All`, `SharePointTenantSettings.Read.All`.

A new permission needs admin consent again in every tenant.

### Per Schuler tenant (Global Admin)

1. Invite our hochzwei account as guest, assign **Global Reader** (PIM-eligible is fine: activate before running).
2. Admin consent: `https://login.microsoftonline.com/<tenant>.onmicrosoft.com/adminconsent?client_id=<clientId>`.
3. Enterprise applications ► our app ► Properties: **Assignment required = Yes**, then assign only our guest account.
4. Cross-tenant access settings ► hochzwei tenant ► Inbound ► Trust settings: trust MFA from Entra tenants.
5. Conditional Access must not block our guest sign-in (compliant device, location). Possibly an exclusion.

Open: Exchange and Purview PowerShell with a guest account is untested. Fallback is a member account in that tenant, set as `account` in `tenants.json`.

Prerequisites besides roles:

- PIM data needs Entra ID P2, MFA registration report needs P1/P2. Missing license shows up in `_errors.json`.
- WAM windows may open behind the terminal.
- SharePoint PowerShell is skipped on purpose: it needs SharePoint Administrator (write role).
