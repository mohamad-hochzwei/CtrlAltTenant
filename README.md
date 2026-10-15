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

`tenants.json`: set `account` to skip account selection, `clientId` once Schuler created the app registration.

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

## Permissions (per tenant, least privilege)

| What | Role | App approval |
|---|---|---|
| Graph | Global Reader | One-time admin consent, see below |
| Exchange + Defender | Global Reader | None (Microsoft first-party app) |
| Purview | Global Reader | None |
| Teams | Global Reader | None |
| DNS | None | None |

Admin consent for the Graph scopes in `Export-Inventory.ps1`, granted by **Cloud Application Administrator** (least privileged for delegated consent) or Global Admin:

- Recommended: dedicated app registration "hochzwei Inventory", public client, redirect URIs `ms-appx-web://microsoft.aad.brokerplugin/<clientId>` (WAM) and `http://localhost`, delegated scopes only, "assignment required" with only our accounts assigned. Client ID goes into `tenants.json`.
- Alternative: consent on "Microsoft Graph Command Line Tools". Simpler, but the grant applies to every Graph PowerShell user in the tenant.

Delegated rights = scopes ∩ role. Global Reader cannot write, regardless of scopes.

Prerequisites besides roles:

- PIM data needs Entra ID P2, MFA registration report needs P1/P2. Missing license shows up in `_errors.json`.
- Conditional Access must not block our sign-in (compliant device, location). Possibly an exclusion for our accounts.
- If Global Reader is PIM-eligible: activate before running.
- Sign-in is interactive (WAM). Device code flow is not used: Microsoft recommends blocking it via Conditional Access. WAM windows may open behind the terminal.
- SharePoint PowerShell is skipped on purpose: it needs SharePoint Administrator (write role).
