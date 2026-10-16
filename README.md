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
./Export-Inventory.ps1                                   # all tenants
./Compare-Inventory.ps1                                  # builds matrix.xlsx
```

`tenants.json`: per tenant `clientId` and `certificate` (thumbprint) of that tenant's own app. Tenants with empty values, missing or expired certificate are skipped.

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

App-only sign-in, no prompts. Each tenant has its **own** app registration and certificate, so a leaked key affects one tenant only. All permissions are read-only.

| What | Permission on the tenant's app | Role of the app |
|---|---|---|
| Graph | Graph application permissions (list in `CreateEntraApp.ps1`) | None needed |
| Exchange + Defender | `Exchange.ManageAsApp` | Global Reader |
| Purview | `Exchange.ManageAsApp` | Global Reader |
| Teams | Graph application permissions | Global Reader (untested, may need Teams Administrator) |
| DNS | None | None |

### Setup per tenant

```
us:           ./CreateCertificate.ps1 -Tenant SyTech            ──► certs/SyTech.cer
tenant admin: ./CreateEntraApp.ps1 -TenantId sytechschuler.onmicrosoft.com `
                  -AppName 'hochzwei Inventory' -CertificatePath certs/SyTech.cer
us:           clientId + certificate into tenants.json
```

- `CreateCertificate.ps1`: private key non-exportable in `Cert:\CurrentUser\My` of the export machine, 6 months. Only the `.cer` (public key) leaves the machine.
- `CreateEntraApp.ps1`: needs Global Admin, or Application Administrator + Privileged Role Administrator. Creates the single-tenant app, grants the permissions (admin consent) and assigns Global Reader. Checks all permission names first, creates nothing on error.

Whoever runs code as our user on the export machine has read access to every set-up tenant, without MFA or Conditional Access. Keep that machine clean, and remove apps and certificates when the project ends (or after a tenant is done):

- Tenant: Entra ► App registrations ► "hochzwei Inventory" ► Delete.
- Us: `Remove-Item Cert:\CurrentUser\My\<thumbprint>`.

Monitoring: Entra ► Sign-in logs ► Service principal sign-ins, filter on the app.

Prerequisites besides roles:

- PIM data needs Entra ID P2, MFA registration report needs P1/P2. Missing license shows up in `_errors.json`.
- SharePoint PowerShell is skipped on purpose: it needs SharePoint Administrator (write role).
